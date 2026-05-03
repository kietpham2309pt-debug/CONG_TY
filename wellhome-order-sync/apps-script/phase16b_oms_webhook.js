/**
 * PHASE 16b — OMS Onflow Webhook Receiver  (refactored 03/05/2026 evening)
 *
 * Schema THẬT đã capture từ OMS test webhook 03/05 19:52:
 *   - Field event = `event_code` (KHÔNG phải `event`)
 *   - status có thể là NUMBER (vd 215) hoặc string
 *   - status_update_time, updated_time, created_time là Unix timestamp (số float, sec từ epoch)
 *   - shipment events có tracking_code, courier_tracking_code, partner_tracking_code
 *   - stock.update có sku, nhsin, pickup_address_id, stock_info{stock, available_stock, waiting, ...}
 *
 * Webhook URL paste vào OMS UI: https://wellhome-oms-proxy.kietpham2309-pt.workers.dev
 *   (Cloudflare Worker proxy embed action+token, forward → Apps Script doPost)
 *
 * Tabs file 1Bn4C0Ud:
 *   tracking haravan → cột AH-AK (OMS Mã đơn / OMS tạo lúc / FFM trạng thái / FFM cập nhật lúc)
 *   oms_webhook_log  → A-F (Timestamp / Event / Order ref / Tracking / Status / Raw JSON)
 *   oms_stock_live   → A-I (SKU / NHSIN / Stock / Available / Waiting / Reserved / Damaged / Pickup / Updated)
 *
 * Public functions:
 *   setupOmsApiKey, setupOmsWebhookToken, setupOmsExtHeader_, setupOms16All
 *   doGet                — cho OMS xác thực URL
 *   handleOmsWebhook_    — main entry, gọi từ phase2_ghn.js doPost
 *   handleStockUpdate_   — sub-handler cho event stock.update
 *   findTrackingRow_     — match row qua orderRef HOẶC tracking_code
 *   omsStatusLabel_      — map status NUMBER → label Việt
 *   omsConvertTime_      — Unix → Date VN
 *   testOmsWebhookFake, testOmsRealisticEvents, showOmsProps
 */

const OMS_CFG = {
  PROP_API_KEY: 'OMS_API_KEY',
  PROP_WEBHOOK_TOKEN: 'OMS_WEBHOOK_TOKEN',

  TARGET_SHEET_ID: '1Bn4C0UdvX2hT82p1pput3VtBDkdJe_jc1-w8ZjKfdGw',
  TARGET_TAB: 'tracking haravan',
  LOG_TAB: 'oms_webhook_log',
  STOCK_TAB: 'oms_stock_live',

  COL_OMS_MA: 34,        // AH
  COL_OMS_AT: 35,        // AI
  COL_FFM_STATUS: 36,    // AJ
  COL_FFM_AT: 37,        // AK

  EXT_HEADERS: ['OMS Mã đơn', 'OMS tạo lúc', 'FFM trạng thái', 'FFM cập nhật lúc'],
  LOG_HEADERS: ['Timestamp', 'Event', 'Order ref', 'Tracking', 'Status', 'Raw JSON'],
  STOCK_HEADERS: ['SKU', 'NHSIN', 'Stock', 'Available', 'Waiting', 'Reserved', 'Damaged', 'Pickup Address', 'Updated'],

  LOG_MAX_ROWS: 5000,
  LOG_KEEP_ROWS: 1000,
};

// Map event code → label tiếng Việt cho cột AJ.
//   null      = lấy status từ payload thay vì label cứng
//   undefined = log only, không update sheet đơn (vd stock/inbound)
const OMS_EVENT_LABELS = {
  'order.created':              'FFM đã nhận đơn',
  'order.update_status':         null,
  'shipment.created':           'FFM đã tạo vận đơn',
  'shipment.update_status':      null,
  'order_return.created':       'FFM hoàn trả',
  'order_return.update_status':  null,
};

// Map status NUMBER → label tiếng Việt — best-guess theo logic shipping pipeline.
// REFINE khi confirm với Mai Onflow. Status 215 đã observed trong test webhook.
// Số ngoài map → hiển thị "FFM mã <N>" để user nhận diện cần update map.
const OMS_STATUS_CODE_MAP = {
  100: 'Đơn mới (chờ xác nhận)',
  101: 'Đã xác nhận',
  105: 'Đang xử lý',
  110: 'Đã in tem',
  150: 'Đã đóng gói (chờ shipper)',
  200: 'Đã tạo vận đơn',
  205: 'Đang chuẩn bị',
  210: 'Đang đóng gói',
  215: 'Đã đóng gói (chờ shipper)',
  220: 'Shipper đã lấy',
  225: 'Đang giao',
  230: 'Đã giao',
  300: 'Trả hàng',
  305: 'Hoàn tất trả hàng',
  400: 'Lỗi',
  500: 'Hủy',
};

// ============================================================
// SETUP — chạy 1 lần qua Web App POST action=oms_setup_all
// ============================================================

function setupOmsApiKey() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(OMS_CFG.PROP_API_KEY);
  if (existing) {
    Logger.log('✅ OMS_API_KEY đã có (' + existing.length + ' chars). Bỏ qua.');
    return existing;
  }
  throw new Error('OMS_API_KEY chưa được set. Gọi action=set_oms_api_key&api_key=<KEY> trước.');
}

function setupOmsWebhookToken() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(OMS_CFG.PROP_WEBHOOK_TOKEN);
  if (existing) {
    Logger.log('⚠️ OMS_WEBHOOK_TOKEN đã có: ' + existing);
    return existing;
  }
  const raw = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const token = raw.slice(0, 48);
  props.setProperty(OMS_CFG.PROP_WEBHOOK_TOKEN, token);
  Logger.log('✅ Sinh OMS_WEBHOOK_TOKEN: ' + token);
  return token;
}

function setupOmsExtHeader_() {
  const ss = SpreadsheetApp.openById(OMS_CFG.TARGET_SHEET_ID);

  const sheet = ss.getSheetByName(OMS_CFG.TARGET_TAB);
  if (!sheet) throw new Error('Tab "' + OMS_CFG.TARGET_TAB + '" chưa tồn tại');

  const curCols = sheet.getMaxColumns();
  if (curCols < OMS_CFG.COL_FFM_AT) {
    sheet.insertColumnsAfter(curCols, OMS_CFG.COL_FFM_AT - curCols);
  }
  sheet.getRange(1, OMS_CFG.COL_OMS_MA, 1, 4).setValues([OMS_CFG.EXT_HEADERS]);
  sheet.getRange(1, OMS_CFG.COL_OMS_MA, 1, 4)
    .setFontWeight('bold')
    .setBackground('#fce5cd');

  let log = ss.getSheetByName(OMS_CFG.LOG_TAB);
  if (!log) {
    log = ss.insertSheet(OMS_CFG.LOG_TAB);
    log.getRange(1, 1, 1, OMS_CFG.LOG_HEADERS.length).setValues([OMS_CFG.LOG_HEADERS]);
    log.getRange(1, 1, 1, OMS_CFG.LOG_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#cfe2f3');
    log.setFrozenRows(1);
    log.setColumnWidth(1, 150);
    log.setColumnWidth(2, 180);
    log.setColumnWidth(3, 130);
    log.setColumnWidth(4, 130);
    log.setColumnWidth(5, 200);
    log.setColumnWidth(6, 600);
  }

  let stock = ss.getSheetByName(OMS_CFG.STOCK_TAB);
  if (!stock) {
    stock = ss.insertSheet(OMS_CFG.STOCK_TAB);
    stock.getRange(1, 1, 1, OMS_CFG.STOCK_HEADERS.length).setValues([OMS_CFG.STOCK_HEADERS]);
    stock.getRange(1, 1, 1, OMS_CFG.STOCK_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#d9ead3');
    stock.setFrozenRows(1);
    stock.setColumnWidth(1, 160);
    stock.setColumnWidth(2, 140);
  }

  Logger.log('✅ Setup AH-AK + 2 tabs (oms_webhook_log + oms_stock_live) xong');
}

function setupOms16All() {
  Logger.log('===== Bắt đầu setup Phase 16 OMS =====');
  setupOmsApiKey();
  const token = setupOmsWebhookToken();
  setupOmsExtHeader_();
  const url = ScriptApp.getService().getUrl();
  Logger.log('===== ✅ Setup xong =====');
  Logger.log('📋 Webhook URL paste vào OMS UI:');
  Logger.log('   ' + url + '?action=oms_webhook&oms_token=' + token);
  return { ok: true, webhook_url: url + '?action=oms_webhook&oms_token=' + token };
}

// ============================================================
// doGet — cho OMS Onflow "Xác thực" callback URL pass
// ============================================================

function doGet(e) {
  const params = (e && e.parameter) || {};

  // Public lookup cho KH tra cứu đơn theo SĐT hoặc Mã đơn
  if (params.lookup === '1' || params.action === 'lookup_order') {
    return HtmlService.createHtmlOutput(buildLookupHtml_(params))
      .setTitle('Tra cứu đơn hàng — Wellhome')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (params.action === 'oms_webhook') {
    const expected = PropertiesService.getScriptProperties().getProperty('OMS_WEBHOOK_TOKEN');
    if (params.oms_token && expected && params.oms_token === expected) {
      return ContentService.createTextOutput(JSON.stringify({
        ok: true, verified: true,
        message: 'OMS webhook endpoint READY. Use POST for actual events.',
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({
      ok: false, error: 'Forbidden (oms_webhook GET without valid token)',
    })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({
    ok: true, name: 'K-Homes Wellhome Order Sync Web App', version: 'v42',
    note: 'POST only — see phase2_ghn.js doPost for actions.',
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Build trang HTML public cho KH tra cứu đơn.
 * Match theo SĐT (cột G) HOẶC Mã đơn Haravan (cột B). KHÔNG show: tổng đơn, COD, email.
 */
function buildLookupHtml_(params) {
  const phone = String(params.phone || '').replace(/[^\d]/g, '');
  const order = String(params.order || '').trim().toUpperCase();
  const baseStyle = '<style>body{font-family:Arial,sans-serif;max-width:600px;margin:30px auto;padding:20px;background:#f5f5f5}h2{color:#1a73e8}.box{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.1);margin:15px 0}input{width:100%;padding:10px;font-size:16px;border:1px solid #ccc;border-radius:4px;margin:5px 0}button{background:#1a73e8;color:#fff;padding:12px 24px;border:none;border-radius:4px;font-size:16px;cursor:pointer}button:hover{background:#0d5cb6}.status{padding:8px 12px;border-radius:4px;display:inline-block;font-weight:bold}.s-ok{background:#d9ead3;color:#274e13}.s-pending{background:#fff2cc;color:#7f6000}.s-err{background:#f4cccc;color:#990000}table{width:100%;border-collapse:collapse}td{padding:8px;border-bottom:1px solid #eee}td:first-child{color:#666;width:40%}</style>';

  let body = '<h2>📦 Tra cứu đơn Wellhome</h2>';
  body += '<div class="box"><form method="get">';
  body += '<input type="hidden" name="lookup" value="1">';
  body += '<label><b>Số điện thoại đặt đơn</b></label><input name="phone" value="' + phone + '" placeholder="0901234567">';
  body += '<label><b>HOẶC mã đơn</b></label><input name="order" value="' + order + '" placeholder="EC325103860">';
  body += '<button type="submit">Tra cứu</button></form></div>';

  if (!phone && !order) {
    return baseStyle + body + '<div class="box"><i>Nhập SĐT hoặc Mã đơn để tra cứu.</i></div>';
  }

  try {
    const ss = SpreadsheetApp.openById(OMS_CFG.TARGET_SHEET_ID);
    const sheet = ss.getSheetByName(OMS_CFG.TARGET_TAB);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return baseStyle + body + '<div class="box">Không có dữ liệu</div>';
    }
    const data = sheet.getRange(2, 1, lastRow - 1, GHN_COL.GHN_AT).getValues();
    const matches = [];
    data.forEach(function (row) {
      const sdt = String(row[GHN_COL.SDT - 1] || '').replace(/[^\d]/g, '');
      const ma = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim().toUpperCase();
      if ((phone && sdt === phone) || (order && ma === order)) {
        matches.push(row);
      }
    });

    if (!matches.length) {
      body += '<div class="box"><b>Không tìm thấy đơn.</b><br>Vui lòng kiểm tra lại SĐT/Mã đơn hoặc liên hệ Wellhome.</div>';
      return baseStyle + body;
    }

    // Group by Mã Haravan
    const grouped = {};
    matches.forEach(function (r) {
      const k = String(r[GHN_COL.MA_HARAVAN - 1] || '').trim();
      if (!grouped[k]) grouped[k] = [];
      grouped[k].push(r);
    });

    Object.keys(grouped).forEach(function (orderCode) {
      const rows = grouped[orderCode];
      const r0 = rows[0];
      const ttGiao = String(r0[GHN_COL.TT_GIAO - 1] || '').trim();
      const ttXuLy = String(r0[GHN_COL.TT_XU_LY - 1] || '').trim();
      const maVc = String(r0[GHN_COL.GHN_MA - 1] || '').trim();
      const ngayDat = r0[GHN_COL.NGAY_DAT - 1];

      let displayStatus = ttGiao || ttXuLy || 'Chưa rõ';
      let statusClass = 's-pending';
      if (/đã giao|delivered/i.test(displayStatus)) statusClass = 's-ok';
      else if (/lỗi|hủy|cancel|fail/i.test(displayStatus)) statusClass = 's-err';

      body += '<div class="box">';
      body += '<table>';
      body += '<tr><td>Mã đơn</td><td><b>' + orderCode + '</b></td></tr>';
      body += '<tr><td>Ngày đặt</td><td>' + (ngayDat || '') + '</td></tr>';
      body += '<tr><td>Trạng thái</td><td><span class="status ' + statusClass + '">' + displayStatus + '</span></td></tr>';
      if (maVc) {
        body += '<tr><td>Mã vận đơn GHN</td><td><a href="https://tracking.ghn.dev/?order_code=' + maVc + '" target="_blank">' + maVc + '</a></td></tr>';
      }
      body += '<tr><td>Số sản phẩm</td><td>' + rows.length + ' SP</td></tr>';
      body += '</table></div>';
    });

    return baseStyle + body;
  } catch (err) {
    return baseStyle + body + '<div class="box">Lỗi: ' + err.message + '</div>';
  }
}

// ============================================================
// HANDLER — gọi từ doPost (phase2_ghn.js)
// ============================================================

function handleOmsWebhook_(e) {
  let payload = {};
  let rawText = '';
  try {
    rawText = (e.postData && e.postData.contents) || '';
    payload = rawText ? JSON.parse(rawText) : {};
  } catch (err) {
    appendOmsLog_('parse_error', '', '', err.message, rawText.slice(0, 4000));
    return { ok: false, error: 'Invalid JSON: ' + err.message };
  }

  // event_code đặt đầu tiên — schema thật của OMS Onflow
  const event = pickFirst_(payload, ['event_code', 'event', 'event_type', 'code', 'type']) || 'unknown';
  const data = payload.data || payload.payload || payload;

  const orderRef = pickFirst_(data, [
    'reference_code', 'reference_no', 'order_code', 'order_ref',
    'external_code', 'haravan_code', 'partner_code',
  ]) || pickFirst_(payload, ['reference_code', 'order_code']) || '';

  const tracking = pickFirst_(data, [
    'tracking_code', 'tracking_number', 'shipment_code',
  ]) || pickFirst_(payload, ['tracking_code']) || '';

  // status có thể là NUMBER hoặc string — pickFirst_ cast sang String
  const statusVal = pickFirst_(data, ['status_name', 'status_text', 'state_name'])
    || (data.status !== undefined && data.status !== null ? String(data.status) : '')
    || (data.state !== undefined && data.state !== null ? String(data.state) : '');

  appendOmsLog_(event, orderRef, tracking, statusVal, rawText.slice(0, 4000));

  // ============ STOCK EVENT — sync vào tab oms_stock_live ============
  if (event === 'stock.update') {
    try {
      const synced = handleStockUpdate_(data);
      return { ok: true, event, sku: pickFirst_(data, ['sku']), synced, message: 'Stock processed' };
    } catch (err) {
      return { ok: true, event, error: err.message, message: 'Stock fail, logged' };
    }
  }

  // ============ INBOUND — log only ============
  if (String(event).indexOf('inbound.') === 0) {
    return { ok: true, event, message: 'Inbound logged only' };
  }

  // ============ ORDER / SHIPMENT / RETURN ============
  if (!(event in OMS_EVENT_LABELS)) {
    return { ok: true, event, message: 'Event không subscribe, logged' };
  }

  const rowIdx = findTrackingRow_(orderRef, tracking);
  if (rowIdx < 0) {
    return { ok: true, event, message: 'No matching row', order_ref: orderRef, tracking };
  }

  const ss = SpreadsheetApp.openById(OMS_CFG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(OMS_CFG.TARGET_TAB);

  // shipment.created / order.created CÓ tracking_code → lưu AH (Mã OMS) + AI (timestamp)
  const isCreated = (event === 'shipment.created' || event === 'order.created');
  if (isCreated && tracking) {
    sheet.getRange(rowIdx, OMS_CFG.COL_OMS_MA).setValue(tracking);
    sheet.getRange(rowIdx, OMS_CFG.COL_OMS_AT).setValue(new Date());
  }

  // AJ — status label (NUMBER → map / string → as-is / fallback event label)
  const label = omsStatusLabel_(statusVal) || OMS_EVENT_LABELS[event] || ('OMS event: ' + event);
  sheet.getRange(rowIdx, OMS_CFG.COL_FFM_STATUS).setValue(label);

  // AK — timestamp from event payload (Unix sec) hoặc now
  const eventTs = pickFirst_(data, ['status_update_time', 'updated_time', 'timestamp', 'created_time']);
  sheet.getRange(rowIdx, OMS_CFG.COL_FFM_AT).setValue(omsConvertTime_(eventTs));

  return {
    ok: true, event, order_ref: orderRef, tracking, label,
    row: rowIdx, is_created: isCreated,
  };
}

function findTrackingRow_(orderRef, tracking) {
  const ss = SpreadsheetApp.openById(OMS_CFG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(OMS_CFG.TARGET_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const haravanCodes = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  let omsCodes = [];
  try { omsCodes = sheet.getRange(2, OMS_CFG.COL_OMS_MA, lastRow - 1, 1).getValues(); } catch (err) {}
  const tgtRef = String(orderRef || '').trim();
  const tgtTr  = String(tracking || '').trim();
  for (let i = 0; i < haravanCodes.length; i++) {
    const hv = String(haravanCodes[i][0] || '').trim();
    const oms = omsCodes[i] ? String(omsCodes[i][0] || '').trim() : '';
    // Match cột B (Mã Haravan) hoặc cột AH (Mã OMS) qua orderRef HOẶC tracking_code
    if ((tgtRef && (hv === tgtRef || oms === tgtRef)) ||
        (tgtTr  && oms === tgtTr)) {
      return i + 2;
    }
  }
  return -1;
}

function handleStockUpdate_(data) {
  const sku = String(pickFirst_(data, ['sku']) || '').trim();
  if (!sku) return false;
  const stockInfo = data.stock_info || {};
  const ss = SpreadsheetApp.openById(OMS_CFG.TARGET_SHEET_ID);
  let tab = ss.getSheetByName(OMS_CFG.STOCK_TAB);
  if (!tab) {
    tab = ss.insertSheet(OMS_CFG.STOCK_TAB);
    tab.getRange(1, 1, 1, OMS_CFG.STOCK_HEADERS.length).setValues([OMS_CFG.STOCK_HEADERS]);
    tab.getRange(1, 1, 1, OMS_CFG.STOCK_HEADERS.length).setFontWeight('bold').setBackground('#d9ead3');
    tab.setFrozenRows(1);
    tab.setColumnWidth(1, 160);
  }
  const lastRow = tab.getLastRow();
  let rowIdx = -1;
  if (lastRow > 1) {
    const skus = tab.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < skus.length; i++) {
      if (String(skus[i][0]).trim() === sku) { rowIdx = i + 2; break; }
    }
  }
  const values = [[
    sku,
    String(pickFirst_(data, ['nhsin']) || ''),
    Number(stockInfo.stock || 0),
    Number(stockInfo.available_stock || 0),
    Number(stockInfo.waiting || 0),
    Number(stockInfo.reserved_stock || 0),
    Number(stockInfo.damaged_stock || 0),
    String(pickFirst_(data, ['pickup_address_id']) || ''),
    new Date(),
  ]];
  if (rowIdx > 0) {
    tab.getRange(rowIdx, 1, 1, OMS_CFG.STOCK_HEADERS.length).setValues(values);
  } else {
    tab.getRange(lastRow + 1, 1, 1, OMS_CFG.STOCK_HEADERS.length).setValues(values);
  }
  return true;
}

// ============================================================
// HELPERS
// ============================================================

function pickFirst_(obj, keys) {
  if (!obj) return '';
  for (let i = 0; i < keys.length; i++) {
    const v = obj[keys[i]];
    if (v !== undefined && v !== null && v !== '') return String(v);
  }
  return '';
}

function omsStatusLabel_(statusVal) {
  if (statusVal === undefined || statusVal === null || statusVal === '') return '';
  const s = String(statusVal).trim();
  if (/^\d+$/.test(s)) {
    const n = parseInt(s, 10);
    return OMS_STATUS_CODE_MAP[n] || ('FFM mã ' + n);
  }
  return s;
}

function omsConvertTime_(val) {
  if (val === undefined || val === null || val === '') return new Date();
  const num = parseFloat(val);
  if (!isNaN(num) && num > 1000000000 && num < 9999999999) {
    return new Date(num * 1000);
  }
  try {
    const d = new Date(val);
    if (!isNaN(d.getTime())) return d;
  } catch (err) {}
  return new Date();
}

function appendOmsLog_(event, orderRef, tracking, status, rawJson) {
  try {
    const ss = SpreadsheetApp.openById(OMS_CFG.TARGET_SHEET_ID);
    const log = ss.getSheetByName(OMS_CFG.LOG_TAB);
    if (!log) return;
    log.appendRow([
      new Date(),
      String(event).slice(0, 100),
      String(orderRef).slice(0, 60),
      String(tracking).slice(0, 60),
      String(status).slice(0, 100),
      String(rawJson).slice(0, 5000),
    ]);
    const rows = log.getLastRow();
    if (rows > OMS_CFG.LOG_MAX_ROWS) {
      const deleteCount = rows - OMS_CFG.LOG_KEEP_ROWS;
      log.deleteRows(2, deleteCount);
    }
  } catch (err) {
    Logger.log('⚠️ appendOmsLog_ fail: ' + err.message);
  }
}

// ============================================================
// TEST
// ============================================================

function testOmsWebhookFake() {
  const fake = {
    postData: { contents: JSON.stringify({
      event_code: 'order.update_status',
      data: { reference_code: 'TEST-EC-FAKE-001', status_name: 'Đã đóng gói' },
      timestamp: new Date().toISOString(),
    })},
  };
  const r = handleOmsWebhook_(fake);
  Logger.log('Fake result: ' + JSON.stringify(r));
  return r;
}

// Test 4 case schema THẬT đã observed trong webhook log của OMS Onflow
function testOmsRealisticEvents() {
  const cases = [
    { desc: 'stock.update real schema',
      payload: {
        event_code: 'stock.update', sku: 'V00315WSXK009', nhsin: 'NHSINV43-fake',
        pickup_address_id: 5700,
        stock_info: { stock: 10, available_stock: -60, waiting: 70, damaged_stock: 0, reserved_stock: 0 },
      }
    },
    { desc: 'shipment.update_status NUMBER 215',
      payload: {
        event_code: 'shipment.update_status', tracking_code: 'NHSV-TEST-001',
        status: 215, status_update_time: 1733216743.0,
      }
    },
    { desc: 'shipment.created tracking_code',
      payload: {
        event_code: 'shipment.created', reference_code: 'EC-FAKE-CREATE-001',
        tracking_code: 'NHSV-CREATED-001',
      }
    },
    { desc: 'order.update_status string status',
      payload: {
        event_code: 'order.update_status', reference_code: 'EC-FAKE-002',
        status_name: 'Đã xác nhận',
      }
    },
    { desc: 'shipment.update_status status=230 (delivered)',
      payload: {
        event_code: 'shipment.update_status', tracking_code: 'NHSV-CREATED-001',
        status: 230, status_update_time: 1733300000.0,
      }
    },
  ];
  const out = [];
  cases.forEach(function (c) {
    const fake = { postData: { contents: JSON.stringify(c.payload) } };
    try {
      const r = handleOmsWebhook_(fake);
      out.push({ desc: c.desc, result: r });
    } catch (err) {
      out.push({ desc: c.desc, error: err.message });
    }
  });
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

function showOmsProps() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const out = {};
  Object.keys(props).filter(function (k) { return k.indexOf('OMS_') === 0; })
    .forEach(function (k) {
      const v = props[k];
      if (k === OMS_CFG.PROP_API_KEY) {
        out[k] = v.slice(0, 8) + '...' + v.slice(-4) + ' (' + v.length + ' chars)';
      } else { out[k] = v; }
    });
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}
