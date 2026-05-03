/**
 * PHASE 16a — Push đơn từ tracking sheet sang OMS Onflow
 *
 * Schema CONFIRMED từ DevTools 03/05:
 *   POST https://oms.onflow.vn/api/v1/orders/create
 *   Header: Authorization: NH <JWT>  (prefix "NH " + JWT)
 *   Body: { store_id, pickup_id, volume, config{}, items[], receiver{}, ... }
 *
 * ⚠️ JWT có expiry ~24h. Phải refresh manual qua action `set_oms_jwt&token=NH ey...`
 *    OMS_API_KEY 76 chars (từ trang config-hook) là webhook secret KHÁC, không dùng cho POST này.
 *
 * Defaults từ DevTools (đơn 1913486):
 *   store_id = 111 (Wellhome VN)
 *   pickup_id = 44 (Kho N&H An Phú Đông)
 *   courier_integration_id = 104
 *   delivery_service = "SOF_STD"  (Self-Owned Fleet Standard)
 *   box_dimension = 41
 *
 * Mapping cần lưu trong ScriptProperties (chị paste 1 lần qua action):
 *   OMS_PROVINCE_MAP_JSON  — { "ha noi": 1, "ho chi minh": 79, ... }
 *   OMS_DISTRICT_MAP_JSON  — { "nam tu liem": 19, "quan 12": 761, ... }
 *   OMS_WARD_MAP_JSON      — { "phu my": 625, "an phu dong": 26779, ... }
 *   OMS_PRODUCT_MAP_JSON   — { "7211005447": 353791, ... } (SKU → product_id)
 *
 * Workflow:
 *   1. Cron 30 phút (đã cài). Skip nếu JWT chưa set / expired / endpoint chưa cấu hình.
 *   2. Quét đơn TT xử lý = "Đã xác nhận" + cột AH (Mã OMS) trống.
 *   3. Build payload + lookup IDs từ map.
 *   4. POST /api/v1/orders/create
 *   5. Response: data.tracking_code (NHOV...) → lưu vào AH; data.order_id → note vào Y.
 *   6. Lỗi → mark "Lỗi OMS" + email alert.
 */

const OMS_PUSH_CFG_FIXED = {
  ENDPOINT: 'https://oms.onflow.vn/api/v1/orders/create',
  AUTH_HEADER_NAME: 'Authorization',
  AUTH_PREFIX: 'NH ',

  // Properties keys
  PROP_JWT: 'OMS_JWT',
  PROP_PROVINCE_MAP: 'OMS_PROVINCE_MAP_JSON',
  PROP_DISTRICT_MAP: 'OMS_DISTRICT_MAP_JSON',
  PROP_WARD_MAP: 'OMS_WARD_MAP_JSON',
  PROP_PRODUCT_MAP: 'OMS_PRODUCT_MAP_JSON',

  // Defaults từ DevTools 03/05
  STORE_ID: 111,
  PICKUP_ID: 44,
  COURIER_INTEGRATION_ID: 104,
  DELIVERY_SERVICE: 'SOF_STD',
  BOX_DIMENSION: 41,

  MAX_PER_RUN: 10,
  ALERT_EMAIL: 'admin@khomes.com.vn',
};

// ============================================================
// MAIN
// ============================================================

function pushOrdersToOms() {
  const props = PropertiesService.getScriptProperties();
  const jwt = props.getProperty(OMS_PUSH_CFG_FIXED.PROP_JWT);
  if (!jwt) {
    Logger.log('🛑 OMS_JWT chưa set. Action set_oms_jwt&token=NH...');
    return { ok: false, error: 'JWT not set' };
  }

  const expInfo = decodeOmsJwtExpiry_(jwt);
  if (expInfo.expired) {
    Logger.log('🛑 JWT expired tại ' + expInfo.expDate);
    sendOmsJwtAlertEmail_(expInfo);
    return { ok: false, error: 'JWT expired at ' + expInfo.expDate };
  }
  if (expInfo.expiresInHours < 2) {
    Logger.log('⚠️ JWT sắp expire: ' + expInfo.expiresInHours.toFixed(1) + 'h');
    sendOmsJwtAlertEmail_(expInfo);
  }

  const ss = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(GHN_CFG.TARGET_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const data = sheet.getRange(2, 1, lastRow - 1, OMS_CFG.COL_FFM_AT).getValues();
  const groups = {};
  data.forEach(function (row, idx) {
    const status = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    const omsMa  = String(row[OMS_CFG.COL_OMS_MA - 1] || '').trim();
    if (status !== 'Đã xác nhận' || omsMa) return;
    const haravan = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim();
    if (!groups[haravan]) groups[haravan] = { rows: [], rowIdxs: [] };
    groups[haravan].rows.push(row);
    groups[haravan].rowIdxs.push(idx + 2);
  });

  const orderNames = Object.keys(groups).slice(0, OMS_PUSH_CFG_FIXED.MAX_PER_RUN);
  if (!orderNames.length) return { ok: true, count: 0 };

  let okCount = 0, failCount = 0;
  const errors = [];

  orderNames.forEach(function (orderName) {
    const grp = groups[orderName];
    try {
      const payload = buildOmsPayload_(grp.rows, false);  // preview=false → tạo thật
      if (!payload) {
        grp.rowIdxs.forEach(function (r) {
          sheet.getRange(r, GHN_COL.TT_XU_LY).setValue('Lỗi OMS');
          appendNote_(sheet, r, 'OMS payload build fail (location/product lookup)');
        });
        failCount++;
        return;
      }
      const result = postOmsCreate_(payload, jwt);
      if (result.ok && result.tracking_code) {
        grp.rowIdxs.forEach(function (r) {
          sheet.getRange(r, OMS_CFG.COL_OMS_MA).setValue(result.tracking_code);
          sheet.getRange(r, OMS_CFG.COL_OMS_AT).setValue(new Date());
          if (result.order_id) appendNote_(sheet, r, 'OMS order_id ' + result.order_id);
        });
        okCount++;
      } else {
        errors.push({ orderName, error: result.error, body: result.body });
        grp.rowIdxs.forEach(function (r) {
          sheet.getRange(r, GHN_COL.TT_XU_LY).setValue('Lỗi OMS');
          appendNote_(sheet, r, 'OMS: ' + (result.error || '').slice(0, 100));
        });
        failCount++;
      }
      Utilities.sleep(400);
    } catch (err) {
      errors.push({ orderName, error: err.message });
      failCount++;
    }
  });

  Logger.log('OMS push: ok=' + okCount + ' fail=' + failCount);
  return { ok: true, count: orderNames.length, okCount, failCount, errors };
}

function postOmsCreate_(payload, jwt) {
  const headers = {
    'Authorization': OMS_PUSH_CFG_FIXED.AUTH_PREFIX + jwt,
    'system': 'oms',
    'country': 'VN',
    'lang': 'vi',
    'origin': 'https://oms.onflow.vn',
  };
  const res = UrlFetchApp.fetch(OMS_PUSH_CFG_FIXED.ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) return { ok: false, error: 'HTTP ' + code, body: body.slice(0, 500) };
  try {
    const json = JSON.parse(body);
    if (json.error || (json.status_code && json.status_code !== 200)) {
      return { ok: false, error: json.messages || 'OMS error', body: body.slice(0, 500) };
    }
    return {
      ok: true,
      tracking_code: (json.data && json.data.tracking_code) || '',
      order_id: (json.data && json.data.order_id) || '',
      raw: json,
    };
  } catch (err) {
    return { ok: false, error: 'JSON parse: ' + err.message, body: body.slice(0, 500) };
  }
}

// ============================================================
// PAYLOAD BUILDER
// ============================================================

function buildOmsPayload_(rows, preview) {
  if (!rows || !rows.length) return null;
  const r0 = rows[0];

  const items = rows.map(function (row) {
    const sku = String(row[GHN_COL.MA_SP - 1] || '').trim();
    if (!sku) return null;
    const sl = Number(row[GHN_COL.SL - 1] || 1);
    const price = Number(row[GHN_COL.DON_GIA - 1] || 0);
    const productId = lookupOmsProductId_(sku);
    const item = {
      sku,
      sale_price: price,
      discounted_price: price,
      customs_value: price,
      quantity: sl,
    };
    if (productId) item.product_id = productId;
    return item;
  }).filter(Boolean);
  if (!items.length) return null;

  const province = String(r0[GHN_COL.TINH - 1] || '').trim();
  const district = String(r0[GHN_COL.QUAN - 1] || '').trim();
  const ward     = String(r0[GHN_COL.PHUONG - 1] || '').trim();
  const loc = lookupOmsLocation_(province, district, ward);
  if (!loc.province_id || !loc.district_id || !loc.ward_id) {
    Logger.log('⚠️ Location lookup miss: ' + province + ' / ' + district + ' / ' + ward + ' → ' + JSON.stringify(loc));
    return null;
  }

  const tongDon = Number(r0[GHN_COL.TONG_DON - 1] || 0);
  const pttt = String(r0[GHN_COL.PTTT - 1] || '').toLowerCase();
  const isCod = /cod|khi giao|cash on/i.test(pttt);

  // Volume từ Scheme nếu có (max DxRxC), fallback 30x30x30
  let volume = '30x30x30';
  try {
    if (typeof loadSchemeFromLocal_ === 'function' || typeof loadSchemeCache_ === 'function') {
      // Skip lookup phức tạp — dùng default. Refine sau khi confirm OMS schema.
    }
  } catch (e) {}

  return {
    store_id: OMS_PUSH_CFG_FIXED.STORE_ID,
    pickup_id: OMS_PUSH_CFG_FIXED.PICKUP_ID,
    volume,
    config: {
      order_type: 'b2c',
      approve: 1,
      courier_integration_id: OMS_PUSH_CFG_FIXED.COURIER_INTEGRATION_ID,
      delivery_service: OMS_PUSH_CFG_FIXED.DELIVERY_SERVICE,
      additional_services: '',
      box_dimension: OMS_PUSH_CFG_FIXED.BOX_DIMENSION,
      fee_paid_by: 'sender',
      fulfill_now: 0,
      preview: preview ? 1 : 0,
      protective_packaging: 0,
      ship_now: 0,
      tax_paid_by: 'sender',
      use_insurance: 0,
    },
    discounts: [],
    documents: [],
    extra_info: {
      note: String(r0[GHN_COL.NOTE_DON - 1] || '').slice(0, 500),
      packaging_note: '',
      order_number: String(r0[GHN_COL.MA_HARAVAN - 1] || ''),
    },
    fees: [{ code: 'buyer_shipping_fee', amount: 0 }],
    items,
    payments: [],
    receiver: {
      country: 'VN',
      fullname: String(r0[GHN_COL.TEN_KH - 1] || ''),
      email: String(r0[GHN_COL.EMAIL - 1] || ''),
      phone: String(r0[GHN_COL.SDT - 1] || ''),
      address: String(r0[GHN_COL.DIA_CHI - 1] || ''),
      province_id: loc.province_id,
      district_id: loc.district_id,
      ward_id: loc.ward_id,
      zipcode: null,
    },
  };
}

// ============================================================
// LOOKUPS
// ============================================================

function lookupOmsProductId_(sku) {
  const map = readOmsMap_(OMS_PUSH_CFG_FIXED.PROP_PRODUCT_MAP);
  return map[String(sku).trim()] || null;
}

function lookupOmsLocation_(provinceName, districtName, wardName) {
  const provinceMap = readOmsMap_(OMS_PUSH_CFG_FIXED.PROP_PROVINCE_MAP);
  const districtMap = readOmsMap_(OMS_PUSH_CFG_FIXED.PROP_DISTRICT_MAP);
  const wardMap = readOmsMap_(OMS_PUSH_CFG_FIXED.PROP_WARD_MAP);
  return {
    province_id: provinceMap[normalizeOmsName_(provinceName)] || null,
    district_id: districtMap[normalizeOmsName_(districtName)] || null,
    ward_id: wardMap[normalizeOmsName_(wardName)] || null,
  };
}

function readOmsMap_(propKey) {
  const json = PropertiesService.getScriptProperties().getProperty(propKey) || '{}';
  try { return JSON.parse(json); } catch (e) { return {}; }
}

function normalizeOmsName_(name) {
  return String(name || '').toLowerCase().trim()
    .replace(/^(tỉnh|thành phố|tp\.?|tp|huyện|quận|q\.?|q|phường|p\.?|p|xã|thị trấn)\s+/i, '')
    .replace(/\s+/g, ' ')
    .normalize('NFD').replace(/[̀-ͯ]/g, '');  // strip Vietnamese accents
}

// ============================================================
// JWT UTILS
// ============================================================

function decodeOmsJwtExpiry_(jwt) {
  try {
    const cleanJwt = String(jwt).replace(/^NH\s+/i, '').trim();
    const parts = cleanJwt.split('.');
    if (parts.length !== 3) return { expired: false, error: 'invalid JWT format' };
    let payload = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    while (payload.length % 4) payload += '=';
    const decoded = Utilities.base64Decode(payload);
    const json = JSON.parse(Utilities.newBlob(decoded).getDataAsString());
    const exp = json.exp;
    if (!exp) return { expired: false };
    const now = Math.floor(Date.now() / 1000);
    return {
      expired: now >= exp,
      expDate: new Date(exp * 1000),
      expiresInHours: (exp - now) / 3600,
      payload: json,
    };
  } catch (err) {
    return { expired: false, error: err.message };
  }
}

function sendOmsJwtAlertEmail_(expInfo) {
  const subject = expInfo.expired
    ? '[K-Homes] 🚨 OMS JWT đã hết hạn — cần refresh ngay'
    : '[K-Homes] ⚠️ OMS JWT sắp hết hạn (' + expInfo.expiresInHours.toFixed(1) + 'h)';
  const html = '<h2>🔑 OMS JWT alert</h2>' +
    '<p>Trạng thái: <b>' + (expInfo.expired ? 'EXPIRED' : 'EXPIRING SOON') + '</b></p>' +
    '<p>Hết hạn: ' + expInfo.expDate + '</p>' +
    '<p><b>Cách refresh:</b></p>' +
    '<ol><li>Mở https://oms.onflow.vn → login</li>' +
    '<li>F12 → Network → Filter "create" → Click 1 request bất kỳ → tab Headers</li>' +
    '<li>Copy giá trị header <code>authorization</code> (phải bắt đầu bằng "NH ey...")</li>' +
    '<li>Gọi action: <code>set_oms_jwt&token=NH eyJ...</code></li></ol>';
  MailApp.sendEmail({ to: OMS_PUSH_CFG_FIXED.ALERT_EMAIL, subject, htmlBody: html });
}

// ============================================================
// DRY-RUN + SETUP HELPERS
// ============================================================

function dryRunPushOms() {
  const props = PropertiesService.getScriptProperties();
  const jwt = props.getProperty(OMS_PUSH_CFG_FIXED.PROP_JWT);
  const expInfo = jwt ? decodeOmsJwtExpiry_(jwt) : { error: 'no JWT' };
  Logger.log('JWT: ' + (jwt ? 'set (' + jwt.length + ' chars)' : 'NOT set'));
  Logger.log('Expiry: ' + JSON.stringify(expInfo));

  const sheet = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID).getSheetByName(GHN_CFG.TARGET_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0, jwt_expiry: expInfo };

  const data = sheet.getRange(2, 1, lastRow - 1, OMS_CFG.COL_FFM_AT).getValues();
  const groups = {};
  data.forEach(function (row, idx) {
    const status = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    const omsMa = String(row[OMS_CFG.COL_OMS_MA - 1] || '').trim();
    if (status !== 'Đã xác nhận' || omsMa) return;
    const haravan = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim();
    if (!groups[haravan]) groups[haravan] = { rows: [], rowIdxs: [] };
    groups[haravan].rows.push(row);
    groups[haravan].rowIdxs.push(idx + 2);
  });
  const orderNames = Object.keys(groups);
  const samples = orderNames.slice(0, 3).map(function (n) {
    return { orderName: n, rowCount: groups[n].rows.length, payload: buildOmsPayload_(groups[n].rows, true) };
  });
  return { ok: true, count: orderNames.length, sample: samples, jwt_expiry: expInfo };
}

function setupOmsPushTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pushOrdersToOms') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pushOrdersToOms').timeBased().everyMinutes(30).create();
  Logger.log('✅ Trigger pushOrdersToOms 30p (skip nếu JWT chưa set)');
}
