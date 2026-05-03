/**
 * PHASE 16a — Push đơn từ tracking sheet sang OMS Onflow
 *
 * v2 refactor: config (endpoint + auth header) lưu trong ScriptProperties
 * → user paste DevTools info qua action set_oms_endpoint / set_oms_auth_*
 * → KHÔNG cần redeploy code mỗi lần đổi config.
 *
 * Workflow khi đầy đủ config:
 *   1. Cron 30 phút (cùng nhịp Phase 2 GHN)
 *   2. Quét tab "tracking haravan" filter cột X = "Đã xác nhận" + cột AH (Mã OMS) trống
 *   3. Build payload từ row data (parse địa chỉ, lookup KL/DxRxC từ Scheme — tận dụng Phase 2)
 *   4. POST → OMS
 *   5. Lưu Mã đơn OMS vào AH + timestamp AI
 *   6. Mark "Lỗi OMS" cột X + email alert nếu fail
 *
 * Hàm `dryRunPushOms()` chạy được KHÔNG CẦN endpoint — log danh sách đơn sẽ push.
 */

const OMS_PUSH_KEYS = {
  ENDPOINT: 'OMS_PUSH_ENDPOINT',
  AUTH_HEADER_NAME: 'OMS_PUSH_AUTH_HEADER_NAME',
  AUTH_PREFIX: 'OMS_PUSH_AUTH_PREFIX',
  EXTRA_HEADERS_JSON: 'OMS_PUSH_EXTRA_HEADERS_JSON',
};

// Defaults metadata captured từ DevTools 03/05 (response GET detail order 1913486):
//   pickup_address.id = 44 (Kho N&H An Phú Đông, code PK100044)
//   store.id = 111 (WELLHOME VIET NAM)
//   courier = 16 (Tự Vận Chuyển — sof prefix)
//   country=VN, currency=VND, platform=retail, order_type=b2c
const OMS_PUSH_DEFAULTS = {
  country: 'VN',
  currency: 'VND',
  platform: 'retail',
  source: 'app',
  order_type: 'b2c',
  pickup_address: 44,
  store: 111,
  courier: 16,
  fulfill_type: 2,
  shipping_provider: 'normal',
  tax_paid_by: 'sender',
  fee_paid_by: 'sender',
};

function omsPushReadConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    endpoint: props.getProperty(OMS_PUSH_KEYS.ENDPOINT) || '',
    authHeaderName: props.getProperty(OMS_PUSH_KEYS.AUTH_HEADER_NAME) || 'X-API-Key',
    authPrefix: props.getProperty(OMS_PUSH_KEYS.AUTH_PREFIX) || '',
    extraHeadersJson: props.getProperty(OMS_PUSH_KEYS.EXTRA_HEADERS_JSON) || '{}',
  };
}

function pushOrdersToOms() {
  const cfg = omsPushReadConfig_();
  if (!cfg.endpoint) {
    Logger.log('🛑 OMS_PUSH_ENDPOINT chưa được set. Gọi action=set_oms_endpoint&endpoint=<URL>');
    return { ok: false, error: 'Endpoint not configured' };
  }
  const apiKey = PropertiesService.getScriptProperties().getProperty('OMS_API_KEY');
  if (!apiKey) {
    Logger.log('🛑 OMS_API_KEY chưa được set');
    return { ok: false, error: 'API key not set' };
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

  const orderNames = Object.keys(groups);
  if (!orderNames.length) return { ok: true, count: 0 };

  const limited = orderNames.slice(0, 20);
  let okCount = 0, failCount = 0;
  const errors = [];

  limited.forEach(function (orderName) {
    const grp = groups[orderName];
    const payload = buildOmsPayload_(grp.rows);
    if (!payload) {
      grp.rowIdxs.forEach(function (r) {
        sheet.getRange(r, GHN_COL.TT_XU_LY).setValue('Lỗi OMS');
      });
      failCount++;
      return;
    }
    const headers = { 'Content-Type': 'application/json' };
    headers[cfg.authHeaderName] = (cfg.authPrefix || '') + apiKey;
    try {
      const extra = JSON.parse(cfg.extraHeadersJson);
      Object.keys(extra).forEach(function (k) { headers[k] = extra[k]; });
    } catch (e) {}

    try {
      const res = UrlFetchApp.fetch(cfg.endpoint, {
        method: 'post',
        headers,
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });
      const code = res.getResponseCode();
      const body = res.getContentText();
      if (code >= 200 && code < 300) {
        const json = JSON.parse(body);
        const omsMaDon = pickFirst_(json, ['tracking_code', 'order_code', 'id', 'reference_code', 'data'])
                     || pickFirst_(json.data || {}, ['tracking_code', 'order_code', 'id', 'reference_code']);
        grp.rowIdxs.forEach(function (r) {
          if (omsMaDon) sheet.getRange(r, OMS_CFG.COL_OMS_MA).setValue(omsMaDon);
          sheet.getRange(r, OMS_CFG.COL_OMS_AT).setValue(new Date());
        });
        okCount++;
      } else {
        errors.push({ orderName, code, body: body.slice(0, 200) });
        grp.rowIdxs.forEach(function (r) {
          sheet.getRange(r, GHN_COL.TT_XU_LY).setValue('Lỗi OMS');
          appendNote_(sheet, r, 'OMS push fail HTTP ' + code + ': ' + body.slice(0, 100));
        });
        failCount++;
      }
      Utilities.sleep(300);
    } catch (err) {
      errors.push({ orderName, error: err.message });
      failCount++;
    }
  });

  Logger.log('OMS push: ok=' + okCount + ' fail=' + failCount);
  return { ok: true, count: limited.length, okCount, failCount, errors };
}

/**
 * Build payload OMS — placeholder shape generic. Khi DevTools confirm, refine theo schema thật.
 * Hiện dùng tập hợp common fields (customer/items/address/cod) — phần lớn OMS đều support.
 */
function buildOmsPayload_(rows) {
  if (!rows || !rows.length) return null;
  const r0 = rows[0];
  const items = rows.map(function (r) {
    return {
      sku: String(r[GHN_COL.MA_SP - 1] || '').trim(),
      name: String(r[GHN_COL.TEN_SP - 1] || '').trim(),
      quantity: Number(r[GHN_COL.SL - 1] || 1),
      price: Number(r[GHN_COL.DON_GIA - 1] || 0),
    };
  }).filter(function (it) { return it.sku && it.quantity > 0; });

  if (!items.length) return null;

  const tongDon = Number(r0[GHN_COL.TONG_DON - 1] || 0);
  const pttt = String(r0[GHN_COL.PTTT - 1] || '').toLowerCase();
  const isCod = /cod|khi giao|cash on/i.test(pttt);

  return {
    reference_code: String(r0[GHN_COL.MA_HARAVAN - 1] || ''),
    customer: {
      name: String(r0[GHN_COL.TEN_KH - 1] || ''),
      phone: String(r0[GHN_COL.SDT - 1] || ''),
      email: String(r0[GHN_COL.EMAIL - 1] || ''),
    },
    address: {
      detail: String(r0[GHN_COL.DIA_CHI - 1] || ''),
      ward: String(r0[GHN_COL.PHUONG - 1] || ''),
      district: String(r0[GHN_COL.QUAN - 1] || ''),
      province: String(r0[GHN_COL.TINH - 1] || ''),
    },
    items,
    cod_amount: isCod ? tongDon : 0,
    total_amount: tongDon,
    note: String(r0[GHN_COL.NOTE_DON - 1] || ''),
  };
}

function dryRunPushOms() {
  Logger.log('=== Phase 16a Dry-Run ===');
  const cfg = omsPushReadConfig_();
  Logger.log('Config: ' + JSON.stringify(cfg));

  const sheet = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID).getSheetByName(GHN_CFG.TARGET_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng'); return { ok: true, count: 0, config: cfg }; }
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
  Logger.log('Candidates: ' + orderNames.length);
  const samples = orderNames.slice(0, 3).map(function (n) {
    return { orderName: n, rowCount: groups[n].rows.length, payload: buildOmsPayload_(groups[n].rows) };
  });
  samples.forEach(function (s) { Logger.log(JSON.stringify(s, null, 2)); });
  return { ok: true, count: orderNames.length, sample: samples, config: cfg };
}

function setupOmsPushTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pushOrdersToOms') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pushOrdersToOms').timeBased().everyMinutes(30).create();
  Logger.log('✅ Trigger pushOrdersToOms 30p (CHỈ kích hoạt khi config đã đủ — pushOrdersToOms tự skip)');
}
