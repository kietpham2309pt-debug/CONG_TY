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
  const provN = normalizeOmsName_(provinceName);
  const distN = normalizeOmsName_(districtName);
  const wardN = normalizeOmsName_(wardName);

  const province_id = provinceMap[provN] || null;
  // Try province-prefixed key first, fallback to plain name
  const district_id = districtMap[province_id + ':' + distN] || districtMap[distN] || null;

  // Wards: lazy fetch nếu chưa có
  let wardMap = readOmsMap_(OMS_PUSH_CFG_FIXED.PROP_WARD_MAP);
  let ward_id = wardMap[district_id + ':' + wardN] || wardMap[wardN] || null;
  if (!ward_id && district_id) {
    wardMap = ensureWardsForDistrict_(district_id);
    ward_id = wardMap[district_id + ':' + wardN] || wardMap[wardN] || null;
  }

  return { province_id, district_id, ward_id };
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

// ============================================================
// AUTO-FETCH MAPPING FROM OMS API
// ============================================================

/**
 * Probe 1 endpoint OMS bất kỳ với JWT — return status + body preview.
 * Dùng để discover endpoint locations/products.
 */
function omsProbeEndpoint(path, queryString) {
  const jwt = PropertiesService.getScriptProperties().getProperty(OMS_PUSH_CFG_FIXED.PROP_JWT);
  if (!jwt) return { ok: false, error: 'OMS_JWT chưa set' };
  const qs = queryString ? ('?' + queryString) : '';
  const url = 'https://oms.onflow.vn' + path + qs;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'NH ' + jwt,
        'system': 'oms',
        'country': 'VN',
        'lang': 'vi',
        'Accept': 'application/json',
      },
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    const body = res.getContentText();
    return {
      ok: code === 200,
      code,
      body_preview: body.slice(0, 2000),
      body_size: body.length,
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Endpoints CONFIRMED từ DevTools 03/05:
 *   GET /api/v1/addresses/province/list                 (deduce — pattern)
 *   GET /api/v1/addresses/district/list?province_id=X
 *   GET /api/v1/addresses/ward/list?district_id=Y
 *   GET /api/v1/products/list-variant-or-single?store_id=111&page=1&page_size=50&status=200,100
 */

function omsGet_(path, qs) {
  const jwt = PropertiesService.getScriptProperties().getProperty(OMS_PUSH_CFG_FIXED.PROP_JWT);
  if (!jwt) return { ok: false, error: 'OMS_JWT not set' };
  const url = 'https://oms.onflow.vn' + path + (qs ? ('?' + qs) : '');
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'NH ' + jwt,
        'system': 'oms', 'country': 'VN', 'lang': 'vi',
        'Accept': 'application/json',
      },
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code !== 200) return { ok: false, code, body: body.slice(0, 500) };
    try { return { ok: true, json: JSON.parse(body), size: body.length }; }
    catch (err) { return { ok: false, error: 'parse: ' + err.message, body: body.slice(0, 500) }; }
  } catch (err) { return { ok: false, error: err.message }; }
}

function fetchAllOmsProvinces() {
  const r = omsGet_('/api/v1/addresses/province/list', 'country=VN&page_size=100');
  if (!r.ok) return r;
  const list = r.json.data || r.json.results || r.json;
  if (!Array.isArray(list)) return { ok: false, error: 'not array', sample: JSON.stringify(r.json).slice(0, 300) };
  const map = {};
  list.forEach(function (p) {
    const name = p.province_name || p.name;
    if (name && p.id) {
      map[normalizeOmsName_(name)] = p.id;
      // Add code as alias (HCM, HN...)
      if (p.code) map[normalizeOmsName_(p.code)] = p.id;
    }
  });
  return { ok: true, count: Object.keys(map).length, map, raw_sample: list.slice(0, 3) };
}

function fetchOmsDistricts(provinceId) {
  const r = omsGet_('/api/v1/addresses/district/list', 'province_id=' + provinceId);
  if (!r.ok) return r;
  const list = r.json.data || r.json.results || r.json;
  if (!Array.isArray(list)) return { ok: false, error: 'not array' };
  return { ok: true, list };
}

function fetchOmsWards(districtId) {
  const r = omsGet_('/api/v1/addresses/ward/list', 'district_id=' + districtId);
  if (!r.ok) return r;
  const list = r.json.data || r.json.results || r.json;
  if (!Array.isArray(list)) return { ok: false, error: 'not array' };
  return { ok: true, list };
}

/**
 * Bulk fetch toàn bộ provinces + districts (~63 calls). Wards = lazy on-demand.
 * Save vào ScriptProperties:
 *   OMS_PROVINCE_MAP_JSON: { "ha noi": 24, "hcm": 79, ... }
 *   OMS_DISTRICT_MAP_JSON: { "<provinceId>:<districtName>": <districtId>, ... }
 *     ← key có prefix province_id để tránh collision (vd "Quận 1" có ở nhiều tỉnh)
 *   OMS_WARD_MAP_JSON: { "<districtId>:<wardName>": <wardId>, ... }
 */
function autoFetchOmsLocations() {
  const props = PropertiesService.getScriptProperties();
  const stats = { provinces: 0, districts: 0, wards: 0, errors: [] };

  const provRes = fetchAllOmsProvinces();
  if (!provRes.ok) return { ok: false, error: 'Fetch provinces fail', detail: provRes };
  props.setProperty(OMS_PUSH_CFG_FIXED.PROP_PROVINCE_MAP, JSON.stringify(provRes.map));
  stats.provinces = provRes.count;

  // Inverse map để loop districts theo từng province
  const districtMap = {};
  const rawProv = omsGet_('/api/v1/addresses/province/list', 'country=VN&page_size=100');
  const provList = rawProv.ok ? (rawProv.json.data || rawProv.json) : [];

  provList.forEach(function (p, idx) {
    if (!p.id) return;
    if (idx > 0 && idx % 10 === 0) Utilities.sleep(500);
    const dr = fetchOmsDistricts(p.id);
    if (!dr.ok) { stats.errors.push({ province: p.id, error: dr.error || dr.code }); return; }
    dr.list.forEach(function (d) {
      const dname = d.district_name || d.name;
      if (dname && d.id) {
        const key = p.id + ':' + normalizeOmsName_(dname);
        districtMap[key] = d.id;
        // Plain name (no province prefix) — fallback nếu duplicate sẽ overwrite
        districtMap[normalizeOmsName_(dname)] = d.id;
        stats.districts++;
      }
    });
    Utilities.sleep(150);
  });
  props.setProperty(OMS_PUSH_CFG_FIXED.PROP_DISTRICT_MAP, JSON.stringify(districtMap));

  Logger.log('✅ Provinces: ' + stats.provinces + ', Districts: ' + stats.districts);
  Logger.log('⏸ Wards lazy fetch on-demand (12k entries quá nhiều cho bulk)');
  return { ok: true, stats, sample_provinces: Object.keys(provRes.map).slice(0, 5) };
}

/**
 * Lazy fetch wards của 1 district + cache vào OMS_WARD_MAP_JSON.
 * Gọi từ lookupOmsLocation_ khi cache miss.
 */
function ensureWardsForDistrict_(districtId) {
  const props = PropertiesService.getScriptProperties();
  const cached = readOmsMap_(OMS_PUSH_CFG_FIXED.PROP_WARD_MAP);
  const sentinel = districtId + ':_loaded';
  if (cached[sentinel]) return cached;
  const wr = fetchOmsWards(districtId);
  if (!wr.ok) return cached;
  wr.list.forEach(function (w) {
    const wname = w.ward_name || w.name;
    if (wname && w.id) {
      cached[districtId + ':' + normalizeOmsName_(wname)] = w.id;
      cached[normalizeOmsName_(wname)] = w.id;  // fallback no district prefix
    }
  });
  cached[sentinel] = 1;
  props.setProperty(OMS_PUSH_CFG_FIXED.PROP_WARD_MAP, JSON.stringify(cached));
  return cached;
}

/**
 * Fetch products list (pagination), build SKU → product_id map.
 */
function autoFetchOmsProducts() {
  const props = PropertiesService.getScriptProperties();
  const map = {};
  let page = 1;
  const pageSize = 50;
  let total = 0;
  while (page <= 30) {
    const qs = 'page=' + page + '&page_size=' + pageSize + '&status=200,100&store_id=' +
               OMS_PUSH_CFG_FIXED.STORE_ID + '&conversion_currency=VND';
    const r = omsGet_('/api/v1/products/list-variant-or-single', qs);
    if (!r.ok) { Logger.log('Page ' + page + ' fail: ' + JSON.stringify(r)); break; }
    const list = r.json.data || r.json.results || r.json;
    if (!Array.isArray(list) || !list.length) break;
    list.forEach(function (item) {
      // Variants có thể nest — try common shapes
      const sku = item.sku || (item.variant && item.variant.sku) || '';
      const id = item.id || (item.variant && item.variant.id) || (item.product && item.product.id);
      if (sku && id) {
        map[String(sku).trim()] = id;
        total++;
      }
    });
    if (list.length < pageSize) break;
    page++;
    Utilities.sleep(300);
  }
  props.setProperty(OMS_PUSH_CFG_FIXED.PROP_PRODUCT_MAP, JSON.stringify(map));
  Logger.log('✅ Products: ' + total + ' SKUs mapped');
  return { ok: true, count: total, pages: page, sample: Object.keys(map).slice(0, 5) };
}

/**
 * Test header với endpoint cần auth — POST /orders/create với body invalid để chỉ check auth.
 * Nếu auth pass → trả 400 (validation fail) hoặc 200.
 * Nếu auth fail → trả 401/403.
 */
function testOmsApiKeyAuthOnPost() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('OMS_API_KEY');
  if (!apiKey) return { ok: false, error: 'OMS_API_KEY not set' };

  const url = 'https://oms.onflow.vn/api/v1/orders/create';
  const baseHeaders = { 'system': 'oms', 'country': 'VN', 'lang': 'vi', 'Content-Type': 'application/json' };
  const dummyBody = JSON.stringify({ store_id: 111, _test_auth_only: true });

  const tries = [
    { name: 'NO_AUTH (baseline)', extra: {} },
    { name: 'X-API-Key',           extra: { 'X-API-Key': apiKey } },
    { name: 'apikey',              extra: { 'apikey': apiKey } },
    { name: 'Api-Key dash',        extra: { 'Api-Key': apiKey } },
    { name: 'Auth ApiKey',         extra: { 'Authorization': 'ApiKey ' + apiKey } },
    { name: 'Auth Bearer',         extra: { 'Authorization': 'Bearer ' + apiKey } },
    { name: 'Auth NH',             extra: { 'Authorization': 'NH ' + apiKey } },
    { name: 'X-Auth-Token',        extra: { 'X-Auth-Token': apiKey } },
  ];

  const results = [];
  tries.forEach(function (t) {
    const headers = Object.assign({}, baseHeaders, t.extra);
    try {
      const res = UrlFetchApp.fetch(url, {
        method: 'post', headers, payload: dummyBody, muteHttpExceptions: true,
      });
      const code = res.getResponseCode();
      const body = res.getContentText();
      // 401/403 = auth fail. 400/422 = auth ok but validation fail. 200 = success.
      const authPassed = (code !== 401 && code !== 403);
      results.push({
        pattern: t.name,
        code,
        auth_passed: authPassed,
        body_preview: body.slice(0, 200),
      });
    } catch (err) {
      results.push({ pattern: t.name, error: err.message });
    }
    Utilities.sleep(300);
  });
  return { ok: true, note: '401/403 = auth fail; 400 = auth ok validation fail; 200 = ok', results };
}

/**
 * Test 7 header patterns với OMS_API_KEY 76 chars để xem có thay được JWT không.
 * Mục đích: nếu API key permanent work → KHÔNG cần refresh JWT mỗi 24h.
 */
function testOmsApiKeyAuth() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('OMS_API_KEY');
  if (!apiKey) return { ok: false, error: 'OMS_API_KEY not set' };

  const baseUrl = 'https://oms.onflow.vn/api/v1/addresses/province/list?country=VN';
  const baseHeaders = { 'system': 'oms', 'country': 'VN', 'lang': 'vi', 'Accept': 'application/json' };

  const tries = [
    { name: 'X-API-Key',           extra: { 'X-API-Key': apiKey } },
    { name: 'apikey lowercase',     extra: { 'apikey': apiKey } },
    { name: 'Api-Key dash',         extra: { 'Api-Key': apiKey } },
    { name: 'Authorization ApiKey', extra: { 'Authorization': 'ApiKey ' + apiKey } },
    { name: 'Authorization Bearer', extra: { 'Authorization': 'Bearer ' + apiKey } },
    { name: 'Authorization NH',     extra: { 'Authorization': 'NH ' + apiKey } },
    { name: 'X-Auth-Token',         extra: { 'X-Auth-Token': apiKey } },
    { name: 'X-Onflow-API-Key',     extra: { 'X-Onflow-API-Key': apiKey } },
  ];

  const results = [];
  tries.forEach(function (t) {
    const headers = Object.assign({}, baseHeaders, t.extra);
    try {
      const res = UrlFetchApp.fetch(baseUrl, {
        method: 'get', headers, muteHttpExceptions: true,
      });
      const code = res.getResponseCode();
      const body = res.getContentText();
      const hasData = body.indexOf('"province_name"') >= 0;
      results.push({
        pattern: t.name,
        code,
        works: code === 200 && hasData,
        body_preview: body.slice(0, 150),
      });
    } catch (err) {
      results.push({ pattern: t.name, error: err.message });
    }
    Utilities.sleep(200);
  });
  return { ok: true, results };
}

/**
 * Discover login endpoint — try 7 paths × 2 body shapes.
 * Email = op.dept@wellhome.asia (hardcode default), password = OMS_PASSWORD property.
 */
function discoverOmsLoginEndpoint() {
  const props = PropertiesService.getScriptProperties();
  const password = props.getProperty('OMS_PASSWORD');
  const email = props.getProperty('OMS_EMAIL') || 'op.dept@wellhome.asia';
  if (!password) return { ok: false, error: 'OMS_PASSWORD chưa set. Action set_oms_password&password=...' };

  const endpoints = [
    '/api/v1/auth/login',
    '/api/v1/auth/token',
    '/api/v1/auth/sign-in',
    '/api/v1/users/login',
    '/api/v1/login',
    '/api/auth/login',
    '/api/v1/auth/jwt/create',
  ];
  const bodyShapes = [
    { email, password },
    { username: email, password },
    { email, password, country: 'VN' },
  ];

  const results = [];
  endpoints.forEach(function (ep) {
    bodyShapes.forEach(function (body) {
      try {
        const res = UrlFetchApp.fetch('https://oms.onflow.vn' + ep, {
          method: 'post',
          contentType: 'application/json',
          headers: { 'system': 'oms', 'country': 'VN', 'lang': 'vi', 'Accept': 'application/json' },
          payload: JSON.stringify(body),
          muteHttpExceptions: true,
        });
        const code = res.getResponseCode();
        const respBody = res.getContentText();
        if (code !== 404 && code !== 405) {  // skip endpoints không tồn tại / sai method
          const hasJwt = respBody.indexOf('access') >= 0 || respBody.indexOf('token') >= 0 || respBody.indexOf('eyJ') >= 0;
          results.push({
            endpoint: ep,
            body_keys: Object.keys(body).join(','),
            code,
            has_jwt_marker: hasJwt,
            preview: respBody.slice(0, 250),
          });
        }
      } catch (err) {
        // skip
      }
      Utilities.sleep(200);
    });
  });
  return { ok: true, results };
}

/**
 * Auto-login dùng email/password lấy JWT mới — cron daily 6h sáng.
 *
 * Endpoint CONFIRMED từ DevTools 03/05:
 *   POST /api/v1/users/token
 *   Body: { email, password, country: "VN", storeData: { country: "VN" } }
 *   Response: { status_code: 200, data: { access, refresh } }
 *
 * Refresh token có exp = +365 ngày → có thể dùng để renew access không cần password.
 * Nhưng login lại bằng password mỗi 24h cũng OK (1 request/ngày).
 */
function autoRefreshOmsJwt() {
  const props = PropertiesService.getScriptProperties();
  const password = props.getProperty('OMS_PASSWORD');
  const email = props.getProperty('OMS_EMAIL') || 'op.dept@wellhome.asia';
  if (!password) {
    Logger.log('🛑 OMS_PASSWORD chưa set');
    return { ok: false, error: 'Password not set' };
  }

  const url = 'https://oms.onflow.vn/api/v1/users/token';
  const body = { email, password, country: 'VN', storeData: { country: 'VN' } };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'system': 'oms', 'country': 'VN', 'lang': 'vi', 'Accept': 'application/json' },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    const respText = res.getContentText();
    if (code !== 200) {
      Logger.log('Login HTTP ' + code + ': ' + respText.slice(0, 300));
      MailApp.sendEmail({
        to: OMS_PUSH_CFG_FIXED.ALERT_EMAIL,
        subject: '[K-Homes] 🚨 OMS auto-refresh JWT FAIL',
        htmlBody: '<p>Login fail HTTP ' + code + '. Có thể password OMS đã đổi hoặc account bị khóa.</p>' +
                  '<p>Action cần làm: refresh password OMS rồi gọi <code>set_oms_password&password=NEW</code></p>' +
                  '<pre>' + respText.slice(0, 500) + '</pre>',
      });
      return { ok: false, code, body: respText.slice(0, 300) };
    }
    const json = JSON.parse(respText);
    const access = json.data && json.data.access;
    const refresh = json.data && json.data.refresh;
    if (!access) {
      Logger.log('Login OK but no access in response: ' + respText.slice(0, 500));
      return { ok: false, error: 'No access in response', preview: respText.slice(0, 500) };
    }
    props.setProperty(OMS_PUSH_CFG_FIXED.PROP_JWT, access);
    if (refresh) props.setProperty('OMS_REFRESH_TOKEN', refresh);
    const expInfo = decodeOmsJwtExpiry_(access);
    Logger.log('✅ JWT refreshed, expires ' + expInfo.expDate);
    return { ok: true, expiry: expInfo, has_refresh: !!refresh };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function setupOmsAutoRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoRefreshOmsJwt') ScriptApp.deleteTrigger(t);
  });
  // Refresh mỗi ngày 6h sáng — JWT 24h nên refresh trước khi expire
  ScriptApp.newTrigger('autoRefreshOmsJwt').timeBased().atHour(6).everyDays(1).create();
  Logger.log('✅ Trigger autoRefreshOmsJwt daily 6h');
}

function setupOmsPushTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pushOrdersToOms') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pushOrdersToOms').timeBased().everyMinutes(30).create();
  Logger.log('✅ Trigger pushOrdersToOms 30p (skip nếu JWT chưa set)');
}
