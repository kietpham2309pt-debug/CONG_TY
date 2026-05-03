/**
 * PHASE 2 — AUTO TẠO ĐƠN GHN  (bản COMBINED — paste vào CÙNG project Wellhome Order Sync)
 * Tất cả const/hàm public của Phase 2 đã đổi tên (prefix GHN_ hoặc Ghn) để tránh trùng Phase 1.
 * Quét tab "tracking haravan", với mỗi đơn cột X = "Đã xác nhận" → tạo đơn GHN.
 * Cron 30 phút/lần. Idempotent (đơn đã "Đã tạo GHN" sẽ skip).
 *
 * Public functions (gọi tay hoặc trigger):
 *   setupGhnToken           — lưu token GHN vào ScriptProperties
 *   setupGhnAll             — gộp: header + test + cache district + sender (chạy 1 lần)
 *   testGhnConnection       — test token GHN còn sống
 *   setupDistrictCache      — cache ~700 district VN vào Properties
 *   setupSenderAddress      — lookup district_id + ward_code kho Q.12 HCM
 *   createGhnOrders         — hàm chính (cron 30 phút gọi)
 *   retryFailedOrders       — reset đơn "Lỗi GHN" → "Đã xác nhận"
 *   dryRunOrder("EC123")    — log payload sẽ gửi (KHÔNG gọi API)
 *   showGhnProps            — show ScriptProperties của Phase 2 (mask token)
 *   setupGhnTrigger         — cài cron 30 phút cho createGhnOrders
 *   removeGhnTrigger        — xóa cron của createGhnOrders
 */

const GHN_CFG = {
  // Sheet đích — cùng file Phase 1 wellhome
  TARGET_SHEET_ID: '1Bn4C0UdvX2hT82p1pput3VtBDkdJe_jc1-w8ZjKfdGw',
  TARGET_TAB: 'tracking haravan',

  // Scheme — SP master (KL/Dài/Rộng/Cao)
  SCHEME_SHEET_ID: '1BWUAMv6P3SFM94SyZRLNeo6vaS2rE2jj3o9A8KBh9TU',
  SCHEME_TAB_GID: 1684440583,

  // GHN config
  API_BASE: 'https://online-gateway.ghn.vn/shiip/public-api',
  SHOP_ID: 5146557,                            // TEFALxHANNAH OLALA
  FROM_PROVINCE: 'Hồ Chí Minh',
  FROM_DISTRICT: 'Quận 12',
  FROM_WARD: 'Phường An Phú Đông',
  FROM_ADDRESS: '3A Đường An Phú Đông 13, An Phú Đông, Quận 12, TP HCM',
  FROM_PHONE: '0522935286',
  SENDER_TEFAL: 'TEFALxHANNAH OLALA',
  SENDER_WELLHOME: 'WELLHOME - FFM',

  DEFAULT_BY_BRAND: {
    'bosch':  { weight: 5000, length: 30, width: 30, height: 30 },
    'tefal':  { weight: 3000, length: 30, width: 25, height: 25 },
    'finish': { weight: 2000, length: 25, width: 20, height: 15 },
    'default':{ weight: 4000, length: 30, width: 30, height: 30 },
  },

  PROP_TOKEN: 'GHN_TOKEN',
  PROP_FROM_DISTRICT_ID: 'GHN_FROM_DISTRICT_ID',
  PROP_FROM_WARD_CODE: 'GHN_FROM_WARD_CODE',
  PROP_DISTRICT_CACHE: 'GHN_DISTRICT_CACHE',
  PROP_PROVINCE_CACHE: 'GHN_PROVINCE_CACHE',
  PROP_DISTRICT_CACHE_AT: 'GHN_DISTRICT_CACHE_AT',

  MAX_ORDERS_PER_RUN: 30,
};

const GHN_EXT_HEADERS = ['Mã VC GHN', 'Phí GHN', 'Tạo GHN lúc'];   // Z, AA, AB

// Cột index trong tab tracking haravan (1-based, theo Phase 1)
const GHN_COL = {
  NGAY_DAT: 1, MA_HARAVAN: 2, MA_HO: 3, LOAI_DON: 4, CAMP: 5,
  TEN_KH: 6, SDT: 7, EMAIL: 8, DIA_CHI: 9, PHUONG: 10, QUAN: 11, TINH: 12,
  VUNG: 13, PTTT: 14, MA_SP: 15, TEN_SP: 16, SL: 17, DON_GIA: 18,
  TONG_DON: 19, TT_TT: 20, TT_GIAO: 21, NOTE_DON: 22, CAP_NHAT: 23,
  TT_XU_LY: 24, GHI_CHU_NV: 25,
  GHN_MA: 26, GHN_PHI: 27, GHN_AT: 28,
};

// Cột index trong tab Scheme (1-based) — VERIFIED 03/05/2026 với scheme_master:
// A=1 CAMP | B=2 CMMF HANNAH | C=3 NAME HANNAH | D=4 SKU CP | E=5 NAME CP
// F=6 SKU GIFT | G=7 NAME GIFT | H=8 SKU CB | I=9 NAME CB
// J=10 STOCK CP | K=11 STOCK GIFT | L=12 Stock Combo | M=13 GIÁ BÁN | N=14 GIÁ VỐN
// O=15 KL (gram) | P=16 Dài | Q=17 Rộng | R=18 Cao | S=19 KL quy đổi
const GHN_SCHEME_COL = {
  CAMP_KEY: 2,    // B - CMMF HANNAH
  SKU_CP: 4,      // D
  SKU_GIFT: 6,    // F (Phase 13 dùng)
  NAME_GIFT: 7,   // G
  SKU_CB: 8,      // H
  NAME_CB: 9,     // I
  KL: 15,         // O (gram) — KL thực
  DAI: 16,        // P (cm)
  RONG: 17,       // Q (cm)
  CAO: 18,        // R (cm)
};

// ============================================================
// SETUP — chạy LẦN ĐẦU sau khi paste code
// ============================================================

function setupGhnToken() {
  const TOKEN = 'e5542293-32a5-11ef-a745-9a89d3653f31';
  PropertiesService.getScriptProperties().setProperty(GHN_CFG.PROP_TOKEN, TOKEN);
  Logger.log('✅ Đã lưu token GHN');
}

function setupGhnExtHeader_() {
  const sheet = openGhnTargetTab_();
  sheet.getRange(1, GHN_COL.GHN_MA, 1, GHN_EXT_HEADERS.length).setValues([GHN_EXT_HEADERS])
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setColumnWidth(GHN_COL.GHN_MA,  130);
  sheet.setColumnWidth(GHN_COL.GHN_PHI, 100);
  sheet.setColumnWidth(GHN_COL.GHN_AT,  140);
  const N = Math.max(sheet.getLastRow(), 1000);
  sheet.getRange(2, GHN_COL.GHN_PHI, N - 1, 1).setNumberFormat('#,##0" đ"');
  Logger.log(`✅ Đã thêm 3 cột mở rộng (Z, AA, AB) cho tab "${GHN_CFG.TARGET_TAB}"`);
}

/**
 * Defensive: kiểm tra dropdown cột X có chứa "Lỗi GHN" + "Đã tạo GHN" không.
 * Nếu thiếu → tự động chạy updateGhnStatusDropdown() để tránh setValue bị reject.
 * Gọi 1 lần ở đầu createGhnOrders.
 */
function ensureGhnStatusDropdown_(sheet) {
  try {
    const v = sheet.getRange(2, GHN_COL.TT_XU_LY).getDataValidation();
    if (!v) { updateGhnStatusDropdown(); return; }
    const criteria = v.getCriteriaValues();
    if (!criteria || !criteria[0]) { updateGhnStatusDropdown(); return; }
    const list = criteria[0];
    const hasErrLabel  = list.indexOf('Lỗi GHN') >= 0;
    const hasOkLabel   = list.indexOf('Đã tạo GHN') >= 0;
    if (!hasErrLabel || !hasOkLabel) {
      Logger.log('🔧 Dropdown cột X thiếu "Lỗi GHN" hoặc "Đã tạo GHN" — tự động chạy updateGhnStatusDropdown');
      updateGhnStatusDropdown();
    }
  } catch (e) {
    Logger.log(`⚠️ ensureGhnStatusDropdown_ exception: ${e.message} — chạy updateGhnStatusDropdown để chắc.`);
    updateGhnStatusDropdown();
  }
}

/** Cập nhật dropdown cột X (Trạng thái xử lý) — thêm "Lỗi GHN" + format màu đỏ */
function updateGhnStatusDropdown() {
  const sheet = openGhnTargetTab_();
  const N = Math.max(sheet.getLastRow(), 1000);
  const FULL_STATUS = ['Chưa xử lý', 'Đang gọi KH', 'Đã xác nhận', 'Đã tạo GHN', 'Lỗi GHN', 'Lỗi data', 'Thiếu hàng', 'Pickup tại SR', 'Hủy', 'Đã giao'];

  // 1. Update dropdown — allow invalid để tool ghi tạm các status mới phòng tương lai
  sheet.getRange(2, GHN_COL.TT_XU_LY, N - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(FULL_STATUS, true)
      .setAllowInvalid(true)
      .build()
  );

  // 2. Append conditional format "Lỗi GHN" → đỏ nhạt (giữ nguyên rules cũ)
  const rules = sheet.getConditionalFormatRules();
  const colX = sheet.getRange(`X2:X${N}`);
  // Bỏ rule cũ cho "Lỗi GHN" nếu có (tránh duplicate)
  const filtered = rules.filter(r => {
    const cond = r.getBooleanCondition();
    return !(cond && cond.getCriteriaValues && cond.getCriteriaValues()[0] === 'Lỗi GHN');
  });
  filtered.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Lỗi GHN')
    .setBackground('#F8CECC')
    .setRanges([colX]).build());
  sheet.setConditionalFormatRules(filtered);

  Logger.log(`✅ Cập nhật dropdown cột X (7 giá trị: ${FULL_STATUS.join(', ')}) + format đỏ cho "Lỗi GHN"`);
}

function testGhnConnection() {
  const res = ghnFetch_('/master-data/province', {});
  if (!res.ok) { Logger.log(`❌ Test fail: ${res.error}`); return; }
  Logger.log(`✅ Token GHN OK — ${res.data.length} tỉnh/thành load được`);
  Logger.log(`   Sample: ${res.data.slice(0,3).map(p => p.ProvinceName).join(', ')}`);
}

function setupDistrictCache() {
  // Cache province + district trong 1 lần setup. lookupDistrictId_ đọc cache, không call API mỗi đơn.
  const provRes = ghnFetch_('/master-data/province', {});
  if (!provRes.ok) { Logger.log(`❌ Province: ${provRes.error}`); return; }
  const provCompact = provRes.data.map(p => ({
    id: p.ProvinceID, name: p.ProvinceName, ext: p.NameExtension || [],
  }));

  const distRes = ghnFetch_('/master-data/district', {});
  if (!distRes.ok) { Logger.log(`❌ District: ${distRes.error}`); return; }
  const distCompact = distRes.data.map(d => ({
    id: d.DistrictID, name: d.DistrictName, pid: d.ProvinceID, ext: d.NameExtension || [],
  }));

  const props = PropertiesService.getScriptProperties();
  props.setProperty(GHN_CFG.PROP_PROVINCE_CACHE, JSON.stringify(provCompact));
  props.setProperty(GHN_CFG.PROP_DISTRICT_CACHE, JSON.stringify(distCompact));
  props.setProperty(GHN_CFG.PROP_DISTRICT_CACHE_AT, new Date().toISOString());
  Logger.log(`✅ Cache ${provCompact.length} province + ${distCompact.length} district vào Properties`);
}

function setupSenderAddress() {
  const districtId = lookupDistrictId_(GHN_CFG.FROM_PROVINCE, GHN_CFG.FROM_DISTRICT);
  if (!districtId) { Logger.log(`❌ Không tìm được district_id của "${GHN_CFG.FROM_DISTRICT}" - "${GHN_CFG.FROM_PROVINCE}"`); return; }
  const wardCode = lookupWardCode_(districtId, GHN_CFG.FROM_WARD);
  if (!wardCode) { Logger.log(`❌ Không tìm được ward_code của "${GHN_CFG.FROM_WARD}" trong district ${districtId}`); return; }
  PropertiesService.getScriptProperties().setProperty(GHN_CFG.PROP_FROM_DISTRICT_ID, String(districtId));
  PropertiesService.getScriptProperties().setProperty(GHN_CFG.PROP_FROM_WARD_CODE, String(wardCode));
  Logger.log(`✅ Sender: district_id=${districtId}, ward_code=${wardCode}`);
}

function setupGhnTrigger() {
  removeGhnTrigger();
  ScriptApp.newTrigger('createGhnOrders').timeBased().everyHours(1).create();
  Logger.log('✅ Đã cài trigger: createGhnOrders chạy mỗi 1 giờ');
}

function removeGhnTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'createGhnOrders') ScriptApp.deleteTrigger(t);
  });
  Logger.log('✅ Đã xóa trigger cũ của createGhnOrders');
}

// ============================================================
// WEB APP ENDPOINT (cho bound script trong Sheet gọi vào)
// ============================================================

const GHN_WEBAPP_SECRET = 'REPLACE_WITH_WEBAPP_SECRET';   // shared với bound script

// Web App handler — phải sync code mới
function doPost(e) {
  // ===== OMS Onflow webhook — bypass GHN_WEBAPP_SECRET =====
  // OMS không biết secret riêng của Web App, nên dùng oms_token query param.
  if (e && e.parameter && e.parameter.action === 'oms_webhook') {
    const omsToken = e.parameter.oms_token || '';
    const expected = PropertiesService.getScriptProperties().getProperty('OMS_WEBHOOK_TOKEN');
    if (!expected || omsToken !== expected) {
      return ghnJsonRes_({ ok: false, error: 'Forbidden (oms_webhook)' });
    }
    try {
      const r = handleOmsWebhook_(e);
      return ghnJsonRes_(r);
    } catch (err) {
      // Vẫn trả 200 OK (qua ContentService) để OMS không retry vô tận, nhưng log lỗi.
      Logger.log('handleOmsWebhook_ exception: ' + err.message + ' / stack: ' + err.stack);
      return ghnJsonRes_({ ok: false, error: err.message });
    }
  }

  if (!e || !e.parameter || e.parameter.secret !== GHN_WEBAPP_SECRET) {
    return ghnJsonRes_({ ok: false, error: 'Forbidden' });
  }
  const action = e.parameter.action;
  try {
    if (action === 'create_ghn_orders') {
      createGhnOrders();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy createGhnOrders' });
    }
    if (action === 'retry_failed') {
      retryFailedOrders();
      return ghnJsonRes_({ ok: true, message: 'Đã reset đơn lỗi' });
    }
    if (action === 'update_ghn_status') {
      // Phase 3 — updateGhnStatuses ở file phase3_status (cùng project)
      updateGhnStatuses();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy updateGhnStatuses' });
    }
    if (action === 'fulfill_haravan') {
      // Phase 4 — fulfillHaravanOrders ở file phase4_fulfill (cùng project)
      fulfillHaravanOrders();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy fulfillHaravanOrders' });
    }
    if (action === 'print_labels') {
      // Task E — gen token GHN cho đơn chưa in, trả URL print
      const r = generatePrintUrlForUnprinted_();
      return ghnJsonRes_(r);
    }
    if (action === 'run_incremental_setup') {
      runIncrementalSetup();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy runIncrementalSetup — xem log Apps Script' });
    }
    if (action === 'run_all_setups') {
      runAllSetups();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy runAllSetups' });
    }
    if (action === 'setup_all_triggers') {
      setupAllTriggers();
      return ghnJsonRes_({ ok: true, message: 'Đã setup tất cả triggers' });
    }
    if (action === 'sync_inventory') {
      syncHaravanInventory();
      return ghnJsonRes_({ ok: true, message: 'Đã pull tồn kho Haravan' });
    }
    if (action === 'audit_columns') {
      const r = auditTrackingColumns();
      return ghnJsonRes_({ ok: true, audit: r });
    }
    if (action === 'repair_columns') {
      const r = repairColumnsAfterPhase3Fix();
      return ghnJsonRes_({ ok: true, stats: r });
    }
    if (action === 'daily_summary') {
      sendDailyGhnFailedSummary();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy sendDailyGhnFailedSummary' });
    }
    if (action === 'daily_recon') {
      runDailyRecon();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy runDailyRecon' });
    }
    if (action === 'setup_recon_header') {
      setupReconHeader();
      return ghnJsonRes_({ ok: true, message: 'Đã setup tab "đối soát"' });
    }
    if (action === 'health_check') {
      const r = healthCheck();
      return ghnJsonRes_({ ok: true, health: r });
    }
    if (action === 'sync_orders') {
      syncOrders();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy syncOrders Phase 1' });
    }
    if (action === 'dry_run_create_ghn') {
      // Verify createGhnOrders không crash + show đơn nào sắp tạo
      dryRunCreateGhnOrders();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy dryRunCreateGhnOrders — xem log Apps Script' });
    }
    if (action === 'auto_retry_ghn') {
      autoRetryFailedGhnOrders();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy autoRetryFailedGhnOrders' });
    }
    if (action === 'weekly_report') {
      sendWeeklyReport();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy sendWeeklyReport' });
    }
    if (action === 'health_monitor') {
      const r = runHealthMonitor();
      return ghnJsonRes_({ ok: true, monitor: r });
    }
    if (action === 'sync_scheme') {
      syncSchemeToLocal();
      return ghnJsonRes_({ ok: true, message: 'Đã sync Scheme master về tab local' });
    }
    if (action === 'monthly_report') {
      forceSendMonthlyReport();
      return ghnJsonRes_({ ok: true, message: 'Đã chạy forceSendMonthlyReport' });
    }
    if (action === 'set_oms_api_key') {
      const key = e.parameter.api_key || '';
      if (!key || key.length < 20) {
        return ghnJsonRes_({ ok: false, error: 'api_key thiếu hoặc quá ngắn (cần >= 20 chars)' });
      }
      PropertiesService.getScriptProperties().setProperty('OMS_API_KEY', key);
      return ghnJsonRes_({ ok: true, message: 'Đã lưu OMS_API_KEY (' + key.length + ' chars)' });
    }
    if (action === 'oms_setup_all') {
      const r = setupOms16All();
      return ghnJsonRes_(r);
    }
    if (action === 'dry_run_oms_webhook') {
      const r = testOmsWebhookFake();
      return ghnJsonRes_({ ok: true, fake_result: r });
    }
    if (action === 'show_oms_props') {
      const r = showOmsProps();
      return ghnJsonRes_({ ok: true, props: r });
    }
    if (action === 'test_oms_realistic') {
      const r = testOmsRealisticEvents();
      return ghnJsonRes_({ ok: true, results: r });
    }
    if (action === 'clear_oms_log') {
      const ss = SpreadsheetApp.openById(OMS_CFG.TARGET_SHEET_ID);
      const log = ss.getSheetByName(OMS_CFG.LOG_TAB);
      if (!log) return ghnJsonRes_({ ok: false, error: 'No log tab' });
      const lastRow = log.getLastRow();
      if (lastRow < 2) return ghnJsonRes_({ ok: true, deleted: 0 });
      log.deleteRows(2, lastRow - 1);
      return ghnJsonRes_({ ok: true, deleted: lastRow - 1 });
    }
    if (action === 'dry_run_push_oms') {
      const r = dryRunPushOms();
      return ghnJsonRes_({ ok: true, dry_run: r });
    }
    if (action === 'push_oms_orders') {
      const r = pushOrdersToOms();
      return ghnJsonRes_(r);
    }
    if (action === 'set_oms_endpoint') {
      const ep = e.parameter.endpoint || '';
      if (!ep || ep.indexOf('http') !== 0) return ghnJsonRes_({ ok: false, error: 'endpoint required (http/https URL)' });
      PropertiesService.getScriptProperties().setProperty('OMS_PUSH_ENDPOINT', ep);
      return ghnJsonRes_({ ok: true, message: 'Set OMS_PUSH_ENDPOINT', endpoint: ep });
    }
    if (action === 'set_oms_auth_header') {
      const name = e.parameter.header_name || '';
      const prefix = e.parameter.header_prefix || '';
      if (!name) return ghnJsonRes_({ ok: false, error: 'header_name required (vd X-API-Key, Authorization)' });
      PropertiesService.getScriptProperties().setProperty('OMS_PUSH_AUTH_HEADER_NAME', name);
      PropertiesService.getScriptProperties().setProperty('OMS_PUSH_AUTH_PREFIX', prefix);
      return ghnJsonRes_({ ok: true, message: 'Set auth header', name, prefix });
    }
    if (action === 'set_oms_jwt') {
      const token = e.parameter.token || '';
      if (!token || token.length < 50) return ghnJsonRes_({ ok: false, error: 'token required (>= 50 chars, dạng "NH eyJ..." hoặc "eyJ...")' });
      // Cleanup prefix nếu user paste cả "NH eyJ..." hoặc "Bearer eyJ..."
      const clean = String(token).replace(/^(NH|Bearer)\s+/i, '').trim();
      PropertiesService.getScriptProperties().setProperty('OMS_JWT', clean);
      // Decode expiry để confirm
      const exp = decodeOmsJwtExpiry_(clean);
      return ghnJsonRes_({ ok: true, message: 'Set OMS_JWT', length: clean.length, expiry: exp });
    }
    if (action === 'decode_oms_jwt') {
      const jwt = PropertiesService.getScriptProperties().getProperty('OMS_JWT');
      if (!jwt) return ghnJsonRes_({ ok: false, error: 'OMS_JWT chưa set' });
      const exp = decodeOmsJwtExpiry_(jwt);
      return ghnJsonRes_({ ok: true, expiry: exp });
    }
    if (action === 'set_oms_locations_map') {
      const ptype = e.parameter.type || '';   // province | district | ward
      const json = e.parameter.json || '';
      if (['province','district','ward'].indexOf(ptype) < 0) return ghnJsonRes_({ ok: false, error: 'type required (province|district|ward)' });
      try { JSON.parse(json); } catch (err) { return ghnJsonRes_({ ok: false, error: 'invalid JSON: ' + err.message }); }
      const key = 'OMS_' + ptype.toUpperCase() + '_MAP_JSON';
      PropertiesService.getScriptProperties().setProperty(key, json);
      const map = JSON.parse(json);
      return ghnJsonRes_({ ok: true, message: 'Set ' + key, entries: Object.keys(map).length });
    }
    if (action === 'set_oms_products_map') {
      const json = e.parameter.json || '';
      try { JSON.parse(json); } catch (err) { return ghnJsonRes_({ ok: false, error: 'invalid JSON: ' + err.message }); }
      PropertiesService.getScriptProperties().setProperty('OMS_PRODUCT_MAP_JSON', json);
      const map = JSON.parse(json);
      return ghnJsonRes_({ ok: true, message: 'Set OMS_PRODUCT_MAP_JSON', entries: Object.keys(map).length });
    }
    if (action === 'oms_probe') {
      const path = e.parameter.path || '';
      const qs = e.parameter.qs || '';
      if (!path) return ghnJsonRes_({ ok: false, error: 'path required (vd /api/v1/locations/provinces)' });
      const r = omsProbeEndpoint(path, qs);
      return ghnJsonRes_(r);
    }
    if (action === 'oms_discover_locations') {
      const r = discoverOmsLocationEndpoints();
      return ghnJsonRes_({ ok: true, results: r });
    }
    if (action === 'oms_auto_fetch_locations') {
      const r = autoFetchOmsLocations();
      return ghnJsonRes_(r);
    }
    if (action === 'oms_auto_fetch_products') {
      const r = autoFetchOmsProducts();
      return ghnJsonRes_(r);
    }
    if (action === 'oms_fetch_provinces') {
      const r = fetchAllOmsProvinces();
      return ghnJsonRes_(r);
    }
    if (action === 'oms_test_api_key') {
      const r = testOmsApiKeyAuth();
      return ghnJsonRes_(r);
    }
    if (action === 'oms_test_api_key_post') {
      const r = testOmsApiKeyAuthOnPost();
      return ghnJsonRes_(r);
    }
    if (action === 'set_oms_password') {
      const pw = e.parameter.password || '';
      if (!pw) return ghnJsonRes_({ ok: false, error: 'password required' });
      PropertiesService.getScriptProperties().setProperty('OMS_PASSWORD', pw);
      const email = e.parameter.email || 'op.dept@wellhome.asia';
      PropertiesService.getScriptProperties().setProperty('OMS_EMAIL', email);
      return ghnJsonRes_({ ok: true, message: 'Saved OMS_PASSWORD + OMS_EMAIL', email });
    }
    if (action === 'oms_discover_login') {
      const r = discoverOmsLoginEndpoint();
      return ghnJsonRes_(r);
    }
    if (action === 'oms_auto_refresh_jwt') {
      const r = autoRefreshOmsJwt();
      return ghnJsonRes_(r);
    }
    if (action === 'oms_setup_auto_refresh_trigger') {
      setupOmsAutoRefreshTrigger();
      return ghnJsonRes_({ ok: true, message: 'Trigger autoRefreshOmsJwt daily 6h cài đặt' });
    }
    if (action === 'cod_recon') {
      const r = runCodRecon();
      return ghnJsonRes_(r);
    }
    if (action === 'pickup_notify') {
      const r = notifyNewPickupOrders();
      return ghnJsonRes_(r);
    }
    if (action === 'pickup_fulfill') {
      const r = fulfillPickupOrders();
      return ghnJsonRes_(r);
    }
    if (action === 'fee_alert') {
      const r = runFeeAlertReport();
      return ghnJsonRes_(r);
    }
    if (action === 'kol_aggregate') {
      const r = aggregateKolPerformance();
      return ghnJsonRes_(r);
    }
    if (action === 'recover_order') {
      const orderName = e.parameter.order_name || '';
      if (!orderName) return ghnJsonRes_({ ok: false, error: 'order_name required' });
      const r = recoverOrderByName(orderName);
      return ghnJsonRes_(r);
    }
    if (action === 'read_oms_log') {
      const limit = parseInt(e.parameter.limit || '10', 10);
      const ss = SpreadsheetApp.openById(OMS_CFG.TARGET_SHEET_ID);
      const log = ss.getSheetByName(OMS_CFG.LOG_TAB);
      if (!log) return ghnJsonRes_({ ok: false, error: 'Tab oms_webhook_log chưa tồn tại' });
      const lastRow = log.getLastRow();
      if (lastRow < 2) return ghnJsonRes_({ ok: true, total: 0, rows: [] });
      const start = Math.max(2, lastRow - limit + 1);
      const count = lastRow - start + 1;
      const data = log.getRange(start, 1, count, 6).getValues();
      const rows = data.map(function (r) {
        return {
          ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
          event: r[1], order_ref: r[2], tracking: r[3], status: r[4],
          raw: String(r[5] || '').slice(0, 500),
        };
      }).reverse();
      return ghnJsonRes_({ ok: true, total: lastRow - 1, rows: rows });
    }
    if (action === 'inspect_scheme_master') {
      const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
      const sh = ss.getSheetByName('scheme_master');
      if (!sh) return ghnJsonRes_({ ok: false, error: 'Tab scheme_master chưa tồn tại' });
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      const header1 = lastRow >= 1 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      const header2 = lastRow >= 2 ? sh.getRange(2, 1, 1, lastCol).getValues()[0] : [];
      const sample = lastRow >= 3 ? sh.getRange(3, 1, Math.min(3, lastRow - 2), lastCol).getValues() : [];
      return ghnJsonRes_({ ok: true, lastRow, lastCol, header1, header2, sample });
    }
    return ghnJsonRes_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return ghnJsonRes_({ ok: false, error: err.message });
  }
}

function ghnJsonRes_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupGhnAll() {
  // Tổng: 3 UrlFetch (province + district + ward).
  Logger.log('===== Bắt đầu setup Phase 2 GHN =====');
  setupGhnExtHeader_();
  updateGhnStatusDropdown();
  setupDistrictCache();
  setupSenderAddress();
  Logger.log('===== ✅ Setup xong. Test bằng createGhnOrders, sau đó setupGhnTrigger =====');
}

// ============================================================
// MAIN — hàm cron chính
// ============================================================

function createGhnOrders() {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng, skip'); return; }

  // Defensive: đảm bảo dropdown cột X có "Lỗi GHN" + "Đã tạo GHN" trước khi loop.
  // Nếu thiếu → setValue sẽ throw → đơn fail không được mark "Lỗi GHN" → cron sau RETRY → tạo đơn GHN trùng.
  ensureGhnStatusDropdown_(sheet);

  const data = sheet.getRange(2, 1, lastRow - 1, GHN_COL.GHN_AT).getValues();

  const groups = {};
  const pickupRows = [];   // đơn nhận tại showroom — không tạo GHN
  data.forEach((row, idx) => {
    const status = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    if (status !== 'Đã xác nhận') return;
    if (row[GHN_COL.GHN_MA - 1]) return;
    const orderId = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim();
    if (!orderId) return;

    // Filter pickup: KH chọn nhận tại showroom → không tạo GHN
    const pttt = String(row[GHN_COL.PTTT - 1] || '').toLowerCase();
    const diaChi = String(row[GHN_COL.DIA_CHI - 1] || '').trim();
    if (isPickupShipping_(pttt, diaChi)) {
      pickupRows.push({ row: idx + 2, orderId });
      return;
    }

    if (!groups[orderId]) groups[orderId] = { orderId, rows: [], rowIndices: [] };
    groups[orderId].rows.push(row);
    groups[orderId].rowIndices.push(idx + 2);
  });

  // Mark đơn pickup: cột X = "Pickup tại SR" + note (idempotent — cron sau skip vì status != "Đã xác nhận")
  if (pickupRows.length > 0) {
    const pickupNow = formatGhnNow_();
    pickupRows.forEach(p => {
      try {
        sheet.getRange(p.row, GHN_COL.TT_XU_LY).setValue('Pickup tại SR');
        sheet.getRange(p.row, GHN_COL.GHI_CHU_NV).setValue(
          `Đơn pickup tại showroom — không tạo GHN (${pickupNow})`);
      } catch (e) {
        Logger.log(`⚠️ Không mark được pickup row ${p.row}: ${e.message}`);
      }
    });
    Logger.log(`📦 ${pickupRows.length} đơn pickup tại SR — đã mark, skip GHN`);
  }

  // Pre-flight validate: SĐT/tên/địa chỉ/tổng đơn — fail trước khi gọi GHN API
  const dataInvalidOrders = [];
  Object.keys(groups).forEach(orderId => {
    const g = groups[orderId];
    const errors = validateOrderData_(g);
    if (errors.length > 0) {
      dataInvalidOrders.push({ orderId, group: g, errors });
      delete groups[orderId];
    }
  });
  if (dataInvalidOrders.length > 0) {
    const invNow = formatGhnNow_();
    dataInvalidOrders.forEach(d => {
      const noteText = 'Lỗi data: ' + d.errors.join('; ');
      d.group.rowIndices.forEach(r => {
        try {
          sheet.getRange(r, GHN_COL.TT_XU_LY).setValue('Lỗi data');
          sheet.getRange(r, GHN_COL.GHI_CHU_NV).setValue(`${noteText} (${invNow})`);
        } catch (e) {
          Logger.log(`⚠️ Mark lỗi data row ${r}: ${e.message}`);
        }
      });
    });
    Logger.log(`⚠️ ${dataInvalidOrders.length} đơn LỖI DATA — đã mark, skip GHN`);
    try { sendDataInvalidAlertEmail_(dataInvalidOrders); }
    catch (e) { Logger.log(`⚠️ sendDataInvalidAlertEmail: ${e.message}`); }
  }

  // Stock check: load tồn kho Haravan, lọc đơn thiếu hàng → mark "Thiếu hàng"
  // Defensive: nếu Phase 6 chưa được paste (loadStockMap_ undefined) → bypass stock check
  const stockMap = (typeof loadStockMap_ === 'function') ? loadStockMap_() : {};
  const stockShortOrders = [];
  Object.keys(groups).forEach(orderId => {
    const g = groups[orderId];
    const shortItems = checkStockForGroup_(g, stockMap);
    if (shortItems.length > 0) {
      stockShortOrders.push({ orderId, group: g, shortItems });
      delete groups[orderId];   // bỏ khỏi danh sách tạo GHN
    }
  });
  if (stockShortOrders.length > 0) {
    const shortNow = formatGhnNow_();
    stockShortOrders.forEach(s => {
      const noteText = 'Thiếu hàng: ' + s.shortItems.map(it => `${it.sku} (cần ${it.need}, tồn ${it.have})`).join('; ');
      s.group.rowIndices.forEach(r => {
        try {
          sheet.getRange(r, GHN_COL.TT_XU_LY).setValue('Thiếu hàng');
          sheet.getRange(r, GHN_COL.GHI_CHU_NV).setValue(`${noteText} (${shortNow})`);
        } catch (e) {
          Logger.log(`⚠️ Mark thiếu hàng row ${r}: ${e.message}`);
        }
      });
    });
    Logger.log(`📦 ${stockShortOrders.length} đơn THIẾU HÀNG — đã mark, skip GHN`);
    try { sendStockShortAlertEmail_(stockShortOrders); }
    catch (e) { Logger.log(`⚠️ sendStockShortAlertEmail: ${e.message}`); }
  }

  const orderIds = Object.keys(groups);
  Logger.log(`Số đơn cần tạo GHN: ${orderIds.length}`);
  if (orderIds.length === 0) return;

  const scheme = loadSchemeCache_();
  Logger.log(`Scheme cache: ${scheme.bySku.size} SKU + ${scheme.byCamp.size} camp`);

  const fromDistrictId = parseInt(getGhnProp_(GHN_CFG.PROP_FROM_DISTRICT_ID), 10);
  const fromWardCode   = String(getGhnProp_(GHN_CFG.PROP_FROM_WARD_CODE) || '');
  if (!fromDistrictId || !fromWardCode) {
    Logger.log(`❌ Chưa setup sender. Chạy setupDistrictCache() + setupSenderAddress() trước.`);
    return;
  }

  let okCount = 0, failCount = 0;
  const writes = [];

  for (let i = 0; i < Math.min(orderIds.length, GHN_CFG.MAX_ORDERS_PER_RUN); i++) {
    const g = groups[orderIds[i]];
    try {
      const result = buildAndCreateGhnOrder_(g, scheme, fromDistrictId, fromWardCode);
      if (result.ok) {
        okCount++;
        writes.push({ rows: g.rowIndices, status: 'Đã tạo GHN', ghnCode: result.code, ghnFee: result.fee, note: '' });
      } else {
        failCount++;
        writes.push({ rows: g.rowIndices, status: 'Lỗi GHN', ghnCode: '', ghnFee: '', note: result.error });
      }
    } catch (e) {
      failCount++;
      writes.push({ rows: g.rowIndices, status: 'Lỗi GHN', ghnCode: '', ghnFee: '', note: 'Exception: ' + e.message });
    }
    Utilities.sleep(300);
  }

  const now = formatGhnNow_();
  writes.forEach(w => {
    w.rows.forEach(r => {
      // Defensive: nếu cột X có data validation cứng không chứa status đích → ghi vào cột Y
      try {
        sheet.getRange(r, GHN_COL.TT_XU_LY).setValue(w.status);
      } catch (e) {
        Logger.log(`⚠️ Không set được X="${w.status}" cho row ${r}: ${e.message}. Chạy updateGhnStatusDropdown() để fix.`);
        w.note = (w.note ? w.note + ' | ' : '') + `Status="${w.status}" bị reject: ${e.message}`;
      }
      if (w.ghnCode) sheet.getRange(r, GHN_COL.GHN_MA).setValue(w.ghnCode);
      if (w.ghnFee !== '') sheet.getRange(r, GHN_COL.GHN_PHI).setValue(w.ghnFee);
      sheet.getRange(r, GHN_COL.GHN_AT).setValue(now);
      if (w.note) {
        const cur = String(sheet.getRange(r, GHN_COL.GHI_CHU_NV).getValue() || '');
        sheet.getRange(r, GHN_COL.GHI_CHU_NV).setValue(cur ? cur + ' | ' + w.note : w.note);
      }
    });
  });

  // Bonus G3 — email alert nếu có đơn lỗi GHN (không gồm đơn pickup/thiếu hàng)
  const failedWrites = writes.filter(w => w.status === 'Lỗi GHN');
  if (failedWrites.length > 0) {
    try {
      sendGhnFailedAlertEmail_(failedWrites, sheet);
      Logger.log(`📧 Đã gửi alert ${failedWrites.length} đơn lỗi GHN`);
    } catch (e) {
      Logger.log(`⚠️ sendGhnFailedAlertEmail: ${e.message}`);
    }
  }

  Logger.log(`✅ Done. OK: ${okCount}, Lỗi: ${failCount}, Pickup: ${pickupRows.length}, Thiếu hàng: ${stockShortOrders.length}`);
}

/** Gửi email tổng hợp đơn lỗi GHN — PG xử lý bằng FFM hoặc fix tay */
function sendGhnFailedAlertEmail_(failedWrites, sheet) {
  const ALERT_EMAIL = 'admin@khomes.com.vn';
  const subject = `⚠️ [K-Homes] ${failedWrites.length} đơn LỖI GHN — cần xử lý thủ công (FFM/fix)`;
  const rowsHtml = failedWrites.map((w, i) => {
    const firstRow = w.rows[0];
    const maHaravan = sheet.getRange(firstRow, GHN_COL.MA_HARAVAN).getValue();
    const tenKH = sheet.getRange(firstRow, GHN_COL.TEN_KH).getValue();
    const sdt = sheet.getRange(firstRow, GHN_COL.SDT).getValue();
    const tinh = sheet.getRange(firstRow, GHN_COL.TINH).getValue();
    const diaChi = sheet.getRange(firstRow, GHN_COL.DIA_CHI).getValue();
    return `
      <tr>
        <td>${i + 1}</td>
        <td><b>${maHaravan}</b></td>
        <td>${tenKH || ''}</td>
        <td><a href="tel:${sdt}">${sdt || ''}</a></td>
        <td>${tinh || ''}</td>
        <td style="font-size:12px;">${diaChi || ''}</td>
        <td style="color:#c0504d;font-size:12px;">${w.note || '(chưa có lý do)'}</td>
      </tr>`;
  }).join('');
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h2 style="color:#c0504d;">⚠️ ${failedWrites.length} đơn KHÔNG TẠO ĐƯỢC GHN</h2>
      <p>Theo nhân viên video: <i>"đơn lỗi GHN → bắt buộc làm thủ công bằng FFM (Fbox)"</i>. Các đơn dưới đây cần PG xử lý:</p>
      <ol>
        <li>Đăng nhập Fbox <a href="https://asia.qianyierp.com/">https://asia.qianyierp.com/</a> → tạo đơn FFM tay</li>
        <li>HOẶC fix lỗi (vd địa chỉ thiếu Phường/Quận, sửa cột J/K/L) → set lại X = "Đã xác nhận" → cron retry</li>
      </ol>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
        <tr style="background:#1f4e78;color:#fff;">
          <th>#</th><th>Mã đơn</th><th>Tên KH</th><th>SĐT</th>
          <th>Tỉnh</th><th>Địa chỉ</th><th>Lý do lỗi</th>
        </tr>
        ${rowsHtml}
      </table>
      <p style="margin-top:15px;color:#666;">➤ Cron Phase 2 chạy 1h/lần. Sau khi fix tay, đợi cron hoặc dùng menu "🚀 K-Homes > Retry đơn lỗi GHN".</p>
    </div>`;
  MailApp.sendEmail({ to: ALERT_EMAIL, subject, htmlBody: html });
}

// ============================================================
// BUILD ORDER + GỌI GHN
// ============================================================

function buildAndCreateGhnOrder_(group, scheme, fromDistrictId, fromWardCode) {
  const r0 = group.rows[0];
  const camp    = String(r0[GHN_COL.CAMP - 1] || '');
  const tenKH   = String(r0[GHN_COL.TEN_KH - 1] || '').trim();
  const sdt     = normalizeGhnPhone_(String(r0[GHN_COL.SDT - 1] || ''));
  const diaChi  = String(r0[GHN_COL.DIA_CHI - 1] || '').trim();
  const pttt    = String(r0[GHN_COL.PTTT - 1] || '').toLowerCase();
  const tongDon = Number(r0[GHN_COL.TONG_DON - 1] || 0);

  // Ưu tiên cột I (Địa chỉ) — parse Phường/Quận/Tỉnh từ chuỗi đầy đủ.
  // Nếu cột I thiếu trường nào → fallback đọc bù từ cột J (Phường), K (Quận), L (Tỉnh) trên sheet.
  const resolved = resolveAddress_(r0, diaChi);
  const phuong = resolved.ward;
  const quan   = resolved.district;
  const tinh   = resolved.province;
  Logger.log(`Đơn ${group.orderId}: parse "${diaChi}" → P="${phuong}" | Q="${quan}" | T="${tinh}" | nguồn=${resolved.source}`);

  if (!tenKH) return { ok: false, error: 'Thiếu Tên KH' };
  if (!sdt || sdt.length < 9) return { ok: false, error: 'SĐT không hợp lệ: ' + sdt };
  if (!diaChi && !tinh) return { ok: false, error: 'Cột I (Địa chỉ) + cột L (Tỉnh) đều trống' };
  if (!tinh) return { ok: false, error: `Thiếu Tỉnh/TP — cột I="${diaChi}" + cột L="${String(r0[GHN_COL.TINH - 1] || '')}" đều không có` };
  if (!phuong && !quan) {
    return { ok: false, error: `Thiếu cả Phường lẫn Quận — cột I="${diaChi}" + J="${String(r0[GHN_COL.PHUONG - 1] || '')}" + K="${String(r0[GHN_COL.QUAN - 1] || '')}". Cần KH xác nhận lại địa chỉ.` };
  }

  // Smart lookup: nếu thiếu Quận hoặc Quận sai, tool sẽ tự suy từ Phường + Tỉnh
  const addrLookup = lookupAddressGhn_(phuong, quan, tinh);
  if (!addrLookup) {
    let why = '';
    if (!quan && phuong) why = `phường "${phuong}" không tìm thấy ở bất kỳ quận nào của "${tinh}"`;
    else if (quan && !phuong) why = `phường thiếu, không reverse được`;
    else why = `cả "${quan}" lẫn "${phuong}" đều không match`;
    // Diagnostic: log char codes để truy ký tự ẩn
    Logger.log(`🔬 DIAG đơn ${group.orderId}:`);
    Logger.log(`   phuong="${phuong}" codes=[${charCodesOf_(phuong)}]`);
    Logger.log(`   quan="${quan}"   codes=[${charCodesOf_(quan)}]`);
    Logger.log(`   tinh="${tinh}"   codes=[${charCodesOf_(tinh)}]`);
    Logger.log(`   normalized: P="${normalizeAddrGhn_(phuong)}" Q="${normalizeAddrGhn_(quan)}" T="${normalizeAddrGhn_(tinh)}"`);
    return { ok: false, error: `Không lookup được địa chỉ GHN: ${why}. Parse từ "${diaChi}". Cần KH xác nhận lại địa chỉ.` };
  }
  const toDistrictId = addrLookup.districtId;
  const toWardCode   = addrLookup.wardCode;
  if (addrLookup.source !== 'direct') {
    Logger.log(`Đơn ${group.orderId}: ${addrLookup.source} — phường "${phuong}" → ${addrLookup.districtName}` +
               (quan && quan !== addrLookup.districtName ? ` (KH ghi quận "${quan}" nhưng tool tự fix)` : ''));
  }

  const dim = lookupDimensions_(group, scheme);
  if (!dim.weight || dim.weight < 50) {
    return { ok: false, error: `Không lookup được KL — fallback brand="${dim.brand}" cũng fail` };
  }

  const items = group.rows.map(row => ({
    name: String(row[GHN_COL.TEN_SP - 1] || 'SP').slice(0, 100),
    code: String(row[GHN_COL.MA_SP - 1] || ''),
    quantity: Number(row[GHN_COL.SL - 1] || 1),
    weight: Math.max(50, Math.round(dim.weight / group.rows.length)),
  }));

  const fromName = detectSenderByBrand_(camp, items);
  const isCOD = pttt.includes('cod') || pttt.includes('khi giao') || pttt.includes('cash on');
  const codAmount = isCOD ? Math.round(tongDon) : 0;
  const serviceTypeId = dim.weight > 20000 ? 5 : 2;
  const requiredNote = isCOD ? 'CHOXEMHANGKHONGTHU' : 'KHONGCHOXEMHANG';

  // Customer note → ghi vào GHN note (đoạn KH viết khi đặt đơn). Shipper sẽ thấy.
  const customerNote = String(group.rows[0][GHN_COL.NOTE_DON - 1] || '').trim();

  // Phase 13 — Lookup quà tặng từ Scheme: nếu đơn có gift → ghi note "Kèm quà X" cho kho.
  // Theo video Tefal, gift được setup đầu camp với mã combo CB. Lookup theo SKU CP/CB hoặc Camp.
  const giftInfo = lookupGiftFromScheme_(group, scheme);
  const giftNote = giftInfo ? `🎁 Kèm quà: ${giftInfo.nameGift} (${giftInfo.skuGift})` : '';

  const noteParts = [];
  if (giftNote) noteParts.push(giftNote);
  if (customerNote) noteParts.push(customerNote);
  noteParts.push(`[${group.orderId}]`);
  const ghnNote = noteParts.join(' | ').slice(0, 200);

  const payload = {
    payment_type_id: 1,
    note: ghnNote,
    required_note: requiredNote,
    from_name: fromName,
    from_phone: GHN_CFG.FROM_PHONE,
    from_address: GHN_CFG.FROM_ADDRESS,
    from_ward_name: GHN_CFG.FROM_WARD,
    from_district_name: GHN_CFG.FROM_DISTRICT,
    from_province_name: GHN_CFG.FROM_PROVINCE,
    return_phone: GHN_CFG.FROM_PHONE,
    return_address: GHN_CFG.FROM_ADDRESS,
    return_district_id: fromDistrictId,
    return_ward_code: fromWardCode,
    client_order_code: group.orderId.replace(/[^A-Za-z0-9_-]/g, '').slice(0, 50),
    to_name: tenKH,
    to_phone: sdt,
    to_address: diaChi,
    to_ward_code: String(toWardCode),
    to_district_id: toDistrictId,
    cod_amount: codAmount,
    weight: dim.weight,
    length: dim.length,
    width: dim.width,
    height: dim.height,
    service_type_id: serviceTypeId,
    items: items,
  };

  const res = ghnFetch_('/v2/shipping-order/create', payload);
  if (!res.ok) return { ok: false, error: res.error };

  return {
    ok: true,
    code: res.data.order_code || '',
    fee: res.data.total_fee || 0,
  };
}

// ============================================================
// LOOKUP DIMENSIONS — 3 TẦNG
// ============================================================

function lookupDimensions_(group, scheme) {
  const r0 = group.rows[0];
  const orderType = String(r0[GHN_COL.LOAI_DON - 1] || '');
  const camp      = String(r0[GHN_COL.CAMP - 1] || '').trim();

  // Tầng 1
  if (orderType === 'KOL' && camp) {
    const found = scheme.byCamp.get(camp.toLowerCase());
    if (found) return Object.assign({}, found, { source: 'camp' });
  }
  if (orderType !== 'KOL') {
    let totalW = 0, maxL = 0, maxW = 0, maxH = 0, hit = 0;
    group.rows.forEach(row => {
      const sku = String(row[GHN_COL.MA_SP - 1] || '').trim();
      if (!sku) return;
      const found = scheme.bySku.get(sku);
      if (found) {
        const qty = Number(row[GHN_COL.SL - 1] || 1);
        totalW += found.weight * qty;
        maxL = Math.max(maxL, found.length);
        maxW = Math.max(maxW, found.width);
        maxH = Math.max(maxH, found.height);
        hit++;
      }
    });
    if (hit > 0 && totalW > 0) {
      return { weight: Math.round(totalW), length: maxL, width: maxW, height: maxH, source: 'sku', brand: detectBrandFromItems_(group) };
    }
  }

  // Tầng 2
  let bestScore = 0, bestDim = null;
  group.rows.forEach(row => {
    const name = String(row[GHN_COL.TEN_SP - 1] || '').trim();
    if (!name) return;
    scheme.byNameList.forEach(entry => {
      const s = stringSimilarity_(name, entry.name);
      if (s > bestScore) { bestScore = s; bestDim = entry; }
    });
  });
  if (bestScore >= 0.65 && bestDim) {
    return Object.assign({}, bestDim, { source: `fuzzy(${bestScore.toFixed(2)})` });
  }

  // Tầng 3 — fallback default theo brand. NHÂN với tổng SL của tất cả line item của đơn
  // (mỗi line là 1 dòng trong group.rows). Tránh báo thiếu KL khi KH đặt nhiều SP.
  const brand = detectBrandFromItems_(group);
  const def = GHN_CFG.DEFAULT_BY_BRAND[brand] || GHN_CFG.DEFAULT_BY_BRAND.default;
  let totalQty = 0;
  group.rows.forEach(row => { totalQty += Number(row[GHN_COL.SL - 1] || 1); });
  totalQty = Math.max(1, totalQty);
  return {
    weight: def.weight * totalQty,
    length: def.length, width: def.width, height: def.height,
    source: `default-${brand}×${totalQty}`, brand,
  };
}

function detectBrandFromItems_(group) {
  const txt = group.rows.map(r =>
    (String(r[GHN_COL.TEN_SP - 1] || '') + ' ' + String(r[GHN_COL.CAMP - 1] || '')).toLowerCase()
  ).join(' ');
  if (txt.includes('bosch')) return 'bosch';
  if (txt.includes('tefal')) return 'tefal';
  if (txt.includes('finish')) return 'finish';
  return 'default';
}

function detectSenderByBrand_(camp, items) {
  const txt = (camp + ' ' + items.map(i => i.name).join(' ')).toLowerCase();
  if (txt.includes('tefal')) return GHN_CFG.SENDER_TEFAL;
  return GHN_CFG.SENDER_WELLHOME;
}

// ============================================================
// SCHEME CACHE
// ============================================================

function loadSchemeCache_() {
  // Dual-read: ưu tiên tab local "scheme_master" (Phase 12 sync), fallback file ngoài.
  const local = (typeof loadSchemeFromLocal_ === 'function') ? loadSchemeFromLocal_() : null;
  if (local) {
    Logger.log(`📥 Đọc Scheme từ tab local "scheme_master" (${local.lastRow} rows)`);
    return parseSchemeRows_(local.sheet, local.lastRow, local.lastCol);
  }
  Logger.log(`📥 Tab local chưa có, fallback file ngoài ${GHN_CFG.SCHEME_SHEET_ID}`);
  const ss = SpreadsheetApp.openById(GHN_CFG.SCHEME_SHEET_ID);
  let sheet = null;
  ss.getSheets().forEach(s => { if (s.getSheetId() === GHN_CFG.SCHEME_TAB_GID) sheet = s; });
  if (!sheet) throw new Error('Không tìm thấy tab Scheme với gid=' + GHN_CFG.SCHEME_TAB_GID);
  return parseSchemeRows_(sheet, sheet.getLastRow(), sheet.getLastColumn());
}

function parseSchemeRows_(sheet, lastRow, lastColRaw) {
  const lastCol = Math.max(GHN_SCHEME_COL.CAO, lastColRaw);
  if (lastRow < 2) return { bySku: new Map(), byCamp: new Map(), byNameList: [], giftBySku: new Map(), giftByCamp: new Map() };

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const bySku = new Map();
  const byCamp = new Map();
  const byNameList = [];
  const giftBySku = new Map();      // SKU CP / SKU CB → { skuGift, nameGift }
  const giftByCamp = new Map();     // CMMF HANNAH (lowercase) → { skuGift, nameGift }

  data.forEach(r => {
    const camp     = String(r[GHN_SCHEME_COL.CAMP_KEY - 1] || '').trim();
    const skuCp    = String(r[GHN_SCHEME_COL.SKU_CP - 1]  || '').trim();
    const skuGift  = String(r[GHN_SCHEME_COL.SKU_GIFT - 1] || '').trim();
    const nameGift = String(r[GHN_SCHEME_COL.NAME_GIFT - 1] || '').trim();
    const skuCb    = String(r[GHN_SCHEME_COL.SKU_CB - 1]  || '').trim();
    const nameCb   = String(r[GHN_SCHEME_COL.NAME_CB - 1] || '').trim();
    const kl       = Number(String(r[GHN_SCHEME_COL.KL - 1]   || 0).toString().replace(/[^\d.-]/g, ''));
    const dai      = Number(r[GHN_SCHEME_COL.DAI - 1]  || 0);
    const rong     = Number(r[GHN_SCHEME_COL.RONG - 1] || 0);
    const cao      = Number(r[GHN_SCHEME_COL.CAO - 1]  || 0);

    // Gift map (kể cả khi KL < 50 — gift cũng cần lookup)
    if (skuGift && nameGift) {
      const giftInfo = { skuGift, nameGift };
      if (skuCp && !giftBySku.has(skuCp)) giftBySku.set(skuCp, giftInfo);
      if (skuCb && !giftBySku.has(skuCb)) giftBySku.set(skuCb, giftInfo);
      if (camp && !giftByCamp.has(camp.toLowerCase())) giftByCamp.set(camp.toLowerCase(), giftInfo);
    }

    if (!kl || kl < 50) return;

    const dim = { weight: kl, length: dai || 30, width: rong || 30, height: cao || 30 };
    if (skuCp && !bySku.has(skuCp)) bySku.set(skuCp, dim);
    if (skuCb && !bySku.has(skuCb)) bySku.set(skuCb, dim);
    if (camp  && !byCamp.has(camp.toLowerCase())) byCamp.set(camp.toLowerCase(), dim);
    if (nameCb) byNameList.push(Object.assign({ name: nameCb }, dim));
  });

  return { bySku, byCamp, byNameList, giftBySku, giftByCamp };
}

// ============================================================
// LOOKUP DISTRICT / WARD GHN
// ============================================================

/**
 * SMART address lookup. Trả về { districtId, wardCode, districtName, source } hoặc null.
 * source:
 *   "direct"          — district + ward đều match từ data parse
 *   "ward-only"       — chỉ ward match (district sai/thiếu) → reverse từ ward+province
 *   "reverse-cache"   — đã từng reverse lookup ward+province trước đây, cached
 */
function lookupAddressGhn_(ward, district, province) {
  if (!province) return null;
  const tWard = normalizeAddrGhn_(ward);
  const tDist = normalizeAddrGhn_(district);
  const tProv = normalizeAddrGhn_(province);

  const provCache = JSON.parse(getGhnProp_(GHN_CFG.PROP_PROVINCE_CACHE) || '[]');
  const distCache = JSON.parse(getGhnProp_(GHN_CFG.PROP_DISTRICT_CACHE) || '[]');
  if (provCache.length === 0 || distCache.length === 0) {
    Logger.log('Province/District cache rỗng. Chạy setupDistrictCache() trước.');
    return null;
  }

  const provMatch = provCache.find(p =>
    normalizeAddrGhn_(p.name) === tProv ||
    (p.ext || []).some(e => normalizeAddrGhn_(e) === tProv)
  );
  if (!provMatch) return null;

  const provDistricts = distCache.filter(d => d.pid === provMatch.id);

  // Path 1 — direct: district được parse + match → check ward trong district đó
  if (district && tDist) {
    let distMatch = provDistricts.find(d =>
      normalizeAddrGhn_(d.name) === tDist ||
      (d.ext || []).some(e => normalizeAddrGhn_(e) === tDist)
    );
    if (!distMatch) {
      distMatch = provDistricts.find(d =>
        normalizeAddrGhn_(d.name).includes(tDist) || tDist.includes(normalizeAddrGhn_(d.name))
      );
    }
    if (distMatch && ward) {
      const wards = getWardsOfDistrict_(distMatch.id);
      const wardMatch = findWardMatch_(wards, tWard);
      if (wardMatch) {
        return { districtId: distMatch.id, wardCode: wardMatch.WardCode, districtName: distMatch.name, source: 'direct' };
      }
      // District đúng nhưng ward không khớp → fall through reverse
    }
    // Path "district only" — không có ward, OK lấy district mặc định ward đầu tiên?
    // Bỏ vì không an toàn (GHN từ chối khi ward không match)
  }

  // Path 2 — reverse lookup: có ward + province nhưng district sai/thiếu
  if (!ward) return null;

  // Cached?
  const cacheKey = 'GHN_REV_' + provMatch.id + '_' + tWard.replace(/\s+/g, '_').slice(0, 80);
  const cached = getGhnProp_(cacheKey);
  if (cached) {
    try {
      const obj = JSON.parse(cached);
      return { districtId: obj.did, wardCode: obj.wc, districtName: obj.dn, source: 'reverse-cache' };
    } catch (e) { /* fall through */ }
  }

  // Loop districts of province, find ward
  for (const dist of provDistricts) {
    const wards = getWardsOfDistrict_(dist.id);
    const match = findWardMatch_(wards, tWard);
    if (match) {
      try {
        PropertiesService.getScriptProperties().setProperty(cacheKey, JSON.stringify({
          did: dist.id, wc: match.WardCode, dn: dist.name,
        }));
      } catch (e) { /* properties full — skip cache */ }
      return { districtId: dist.id, wardCode: match.WardCode, districtName: dist.name, source: 'ward-only' };
    }
  }
  return null;
}

function findWardMatch_(wards, normalizedWardName) {
  if (!wards || wards.length === 0 || !normalizedWardName) return null;
  return wards.find(w =>
    normalizeAddrGhn_(w.WardName) === normalizedWardName ||
    (w.NameExtension || []).some(e => normalizeAddrGhn_(e) === normalizedWardName)
  ) || wards.find(w => {
    const wn = normalizeAddrGhn_(w.WardName);
    return wn.includes(normalizedWardName) || normalizedWardName.includes(wn);
  }) || null;
}

/** Cache wards của 1 district (CacheService 6h TTL) — giảm UrlFetch cho đơn cùng district */
function getWardsOfDistrict_(districtId) {
  const cache = CacheService.getScriptCache();
  const key = 'wards_' + districtId;
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* corrupt — refetch */ }
  }
  const res = ghnFetch_('/master-data/ward', { district_id: districtId });
  if (!res.ok || !res.data) return [];
  try { cache.put(key, JSON.stringify(res.data), 21600); }
  catch (e) { /* value too large — skip cache */ }
  return res.data;
}

function lookupDistrictId_(provinceName, districtName) {
  // Đọc 100% từ cache, không call API. Cache nạp 1 lần qua setupDistrictCache().
  const distCache = JSON.parse(getGhnProp_(GHN_CFG.PROP_DISTRICT_CACHE) || '[]');
  const provCache = JSON.parse(getGhnProp_(GHN_CFG.PROP_PROVINCE_CACHE) || '[]');
  if (distCache.length === 0 || provCache.length === 0) {
    Logger.log('Province/District cache rỗng. Chạy setupDistrictCache() trước.');
    return null;
  }

  const tProv = normalizeAddrGhn_(provinceName);
  const tDist = normalizeAddrGhn_(districtName);

  const provMatch = provCache.find(p =>
    normalizeAddrGhn_(p.name) === tProv ||
    (p.ext || []).some(e => normalizeAddrGhn_(e) === tProv)
  );
  if (!provMatch) return null;

  const candidates = distCache.filter(d => d.pid === provMatch.id);
  let best = candidates.find(d =>
    normalizeAddrGhn_(d.name) === tDist ||
    (d.ext || []).some(e => normalizeAddrGhn_(e) === tDist)
  );
  if (!best) {
    best = candidates.find(d =>
      normalizeAddrGhn_(d.name).includes(tDist) || tDist.includes(normalizeAddrGhn_(d.name))
    );
  }
  return best ? best.id : null;
}

/**
 * Debug: log từng bước lookup để biết bước nào fail (province/district/ward).
 * Cách dùng: function _t() { debugAddressLookup("Phường 14", "Quận Gò Vấp", "Hồ Chí Minh"); }
 */
function debugAddressLookup(ward, district, province) {
  Logger.log('===== DEBUG LOOKUP =====');
  Logger.log(`Input: ward="${ward}" | district="${district}" | province="${province}"`);
  const tWard = normalizeAddrGhn_(ward);
  const tDist = normalizeAddrGhn_(district);
  const tProv = normalizeAddrGhn_(province);
  Logger.log(`Normalized: ward="${tWard}" | district="${tDist}" | province="${tProv}"`);

  const provCache = JSON.parse(getGhnProp_(GHN_CFG.PROP_PROVINCE_CACHE) || '[]');
  const distCache = JSON.parse(getGhnProp_(GHN_CFG.PROP_DISTRICT_CACHE) || '[]');
  Logger.log(`Cache: ${provCache.length} provinces, ${distCache.length} districts`);
  if (provCache.length === 0 || distCache.length === 0) {
    Logger.log('❌ Cache rỗng — Run setupDistrictCache trước');
    return;
  }

  // Step 1: match province
  const provMatch = provCache.find(p =>
    normalizeAddrGhn_(p.name) === tProv ||
    (p.ext || []).some(e => normalizeAddrGhn_(e) === tProv)
  );
  if (!provMatch) {
    Logger.log(`❌ Không match province. Sample 10 đầu: ${provCache.slice(0, 10).map(p => p.name).join(' | ')}`);
    return;
  }
  Logger.log(`✓ Province match: "${provMatch.name}" (id=${provMatch.id})`);

  const provDistricts = distCache.filter(d => d.pid === provMatch.id);
  Logger.log(`  ${provDistricts.length} districts trong province này`);

  // Step 2: match district (exact + ext)
  let distMatch = provDistricts.find(d =>
    normalizeAddrGhn_(d.name) === tDist ||
    (d.ext || []).some(e => normalizeAddrGhn_(e) === tDist)
  );
  if (!distMatch) {
    Logger.log(`  Exact match fail. Thử substring...`);
    distMatch = provDistricts.find(d =>
      normalizeAddrGhn_(d.name).includes(tDist) || tDist.includes(normalizeAddrGhn_(d.name))
    );
  }
  if (!distMatch) {
    Logger.log(`❌ Không match district "${district}". Districts trong province:`);
    provDistricts.forEach(d => Logger.log(`    ${d.name} (id=${d.id})${d.ext && d.ext.length ? ' | ext: ' + d.ext.join(', ') : ''}`));
    return;
  }
  Logger.log(`✓ District match: "${distMatch.name}" (id=${distMatch.id})`);

  // Step 3: fetch wards + match
  const wards = getWardsOfDistrict_(distMatch.id);
  Logger.log(`  ${wards.length} wards trong district này`);
  if (wards.length === 0) {
    Logger.log('❌ getWardsOfDistrict_ trả về 0 ward — có thể quota UrlFetch hết / token hỏng');
    return;
  }
  const sampleWards = wards.slice(0, 20).map(w => `${w.WardName}${w.NameExtension && w.NameExtension.length ? '(ext:' + w.NameExtension.join('/') + ')' : ''}`);
  Logger.log(`  Sample: ${sampleWards.join(' | ')}`);

  // Match exact
  let wardMatch = wards.find(w =>
    normalizeAddrGhn_(w.WardName) === tWard ||
    (w.NameExtension || []).some(e => normalizeAddrGhn_(e) === tWard)
  );
  if (!wardMatch) {
    Logger.log(`  Exact ward match fail. Thử substring...`);
    wardMatch = wards.find(w => {
      const wn = normalizeAddrGhn_(w.WardName);
      return wn.includes(tWard) || tWard.includes(wn);
    });
  }
  if (!wardMatch) {
    Logger.log(`❌ Không match ward "${ward}" (norm "${tWard}"). Đã list 20 sample bên trên.`);
    return;
  }
  Logger.log(`✅ FULL MATCH: districtId=${distMatch.id}, wardCode=${wardMatch.WardCode}, wardName="${wardMatch.WardName}"`);
}

function lookupWardCode_(districtId, wardName) {
  const wards = getWardsOfDistrict_(districtId);
  const tWard = normalizeAddrGhn_(wardName);
  if (!tWard) return null;
  const match = findWardMatch_(wards, tWard);
  return match ? match.WardCode : null;
}

// 5 thành phố trực thuộc TW (province-level) — "Thành phố X" với X trong list này là TỈNH, không phải QUẬN
const GHN_PROVINCE_TP = ['hồ chí minh', 'hà nội', 'đà nẵng', 'hải phòng', 'cần thơ', 'huế'];

/**
 * Parse FULL địa chỉ (Phường + Quận + Tỉnh) từ chuỗi cột I.
 * Format chuẩn: "<số nhà + đường>, Phường X, Quận Y, Tỉnh Z[, Việt Nam]"
 * Trả về { ward, district, province } — bất kỳ field nào parse không ra sẽ là ''.
 */
function parseFullAddress_(address) {
  const empty = { ward: '', district: '', province: '' };
  if (!address) return empty;

  // 1. Cleanup: bỏ "Việt Nam" / "Vietnam" / "VN" / dấu chấm/khoảng trắng cuối
  let cleaned = String(address)
    .replace(/[,;\s]*(việt\s*nam|vietnam|\bvn\b)\s*[.\s]*$/i, '')
    .replace(/[.\s,;]+$/, '')
    .trim();

  const segs = cleaned.split(',').map(s => s.trim()).filter(Boolean);
  if (segs.length === 0) return empty;

  let ward = '', district = '', province = '';

  // Chỉ gán province mặc định = seg cuối khi:
  //  (a) có >= 2 segs (single seg mặc định là địa chỉ chi tiết, không phải tỉnh)
  //  (b) seg cuối KHÔNG trông giống tên đường (số nhà đầu / từ khóa đường phố)
  if (segs.length >= 2) {
    const lastSeg = segs[segs.length - 1];
    const looksLikeStreet = /^\d/.test(lastSeg) ||
      /\b(đường|phố|ngõ|hẻm|kiệt|tổ|ấp|thôn|khu phố|kp\.?)\b/i.test(lastSeg);
    if (!looksLikeStreet) province = lastSeg;
  }

  // 2. Scan các segs (trừ cuối) tìm prefix Phường + Quận
  for (let i = 0; i < segs.length - 1; i++) {
    const s = segs[i];

    // WARD
    if (!ward) {
      let mW = s.match(/^(Phường|Xã|Thị trấn)\s+(.+)$/i);
      if (mW) { ward = s; continue; }
      let mP = s.match(/^P\.?\s*(\d+[A-Z]?|[^\d].*)$/i);
      if (mP) { ward = 'Phường ' + mP[1].trim(); continue; }
      let mTT = s.match(/^TT\.?\s+(.+)$/i);
      if (mTT) { ward = 'Thị trấn ' + mTT[1].trim(); continue; }
    }

    // DISTRICT
    if (!district) {
      // "Thành phố X" → district CHỈ KHI X không thuộc 5 TP trực thuộc TW
      let mTP = s.match(/^Thành phố\s+(.+)$/i);
      if (mTP) {
        const cityName = mTP[1].trim().toLowerCase();
        if (GHN_PROVINCE_TP.some(p => cityName.includes(p))) {
          // "Thành phố Hồ Chí Minh" → province (override default)
          province = s;
          continue;
        }
        // "Thành phố Thủ Đức" / "Thành phố Vinh" → district
        district = s;
        continue;
      }
      let mD = s.match(/^(Quận|Huyện|Thị xã)\s+(.+)$/i);
      if (mD) { district = s; continue; }
      let mQ = s.match(/^Q\.?\s*(\d+|.+)$/i);
      if (mQ) { district = 'Quận ' + mQ[1].trim(); continue; }
      let mH = s.match(/^H\.?\s+(.+)$/i);
      if (mH) { district = 'Huyện ' + mH[1].trim(); continue; }
    }
  }

  // 3. Fallback positional CHỈ khi segs >= 4 (đảm bảo đủ slot đường+phường+quận+tỉnh)
  // KH ghi không có prefix: "73 Bùi Dương Lịch, Bình Hưng Hoà A, Bình Tân, Hồ Chí Minh"
  if (segs.length >= 4) {
    if (!district) district = segs[segs.length - 2];
    if (!ward) ward = segs[segs.length - 3];
  }

  return { ward, district, province };
}

function extractWardFromAddress_(address) {
  return parseFullAddress_(address).ward;
}

/**
 * Hợp nhất 3 nguồn địa chỉ: parse cột I (ưu tiên) + fallback cột J/K/L của sheet.
 * Trả về { ward, district, province, source } — source mô tả nguồn từng trường.
 *   source format: "PJ/QK/TI" nghĩa là Phường lấy J, Quận lấy K, Tỉnh lấy I.
 */
function resolveAddress_(row, diaChi) {
  const parsed = parseFullAddress_(diaChi || '');
  const sheetWard     = String(row[GHN_COL.PHUONG - 1] || '').trim();
  const sheetDistrict = String(row[GHN_COL.QUAN   - 1] || '').trim();
  const sheetProvince = String(row[GHN_COL.TINH   - 1] || '').trim();

  let ward = parsed.ward, district = parsed.district, province = parsed.province;
  let pSrc = ward     ? 'I' : '';
  let qSrc = district ? 'I' : '';
  let tSrc = province ? 'I' : '';

  if (!ward     && sheetWard)     { ward = sheetWard; pSrc = 'J'; }
  if (!district && sheetDistrict) { district = sheetDistrict; qSrc = 'K'; }
  if (!province && sheetProvince) { province = sheetProvince; tSrc = 'L'; }

  // Edge: cột J có thể chỉ ghi "Phường 12" mà tên gốc parse cũng ra "Phường 12" — không sao, kết quả như nhau.
  // Edge: cột K có thể ghi "Tân Bình" thay vì "Quận Tân Bình" — normalizeAddrGhn_ trong lookupAddressGhn_ tự strip prefix.

  return {
    ward, district, province,
    source: `P${pSrc || '-'}/Q${qSrc || '-'}/T${tSrc || '-'}`,
  };
}

function normalizeAddrGhn_(s) {
  if (!s) return '';
  return String(s)
    // Strip mọi loại whitespace ẩn: NBSP, zero-width, BOM, narrow no-break, ideographic
    .replace(/[ ­-‏‪-  ⁠　]/g, ' ')
    .toLowerCase()
    .replace(/^(thành phố|tp\.?|tỉnh|quận|huyện|phường|xã|thị trấn|thị xã)\s+/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Trả về chuỗi mô tả char codes — để debug ký tự ẩn */
function charCodesOf_(s) {
  return String(s || '').split('').map(c => c.charCodeAt(0)).join(',');
}

/**
 * Debug 1 đơn cụ thể: log RAW cell values + char codes + normalized + lookup từng bước.
 * Cách dùng:
 *   function _diag() { diagOneOrder("EC325103860"); }
 *   chọn _diag → Run → xem Logs.
 */
function diagOneOrder(orderName) {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, GHN_COL.GHN_AT).getValues();
  const cleanName = String(orderName).replace(/^#/, '');
  const r = data.find(row => String(row[GHN_COL.MA_HARAVAN - 1] || '').replace(/^#/, '') === cleanName);
  if (!r) { Logger.log(`Không tìm thấy đơn ${orderName}`); return; }

  Logger.log(`===== DIAG ĐƠN ${orderName} =====`);
  const cols = [
    { name: 'I (Địa chỉ)', idx: GHN_COL.DIA_CHI },
    { name: 'J (Phường)',  idx: GHN_COL.PHUONG  },
    { name: 'K (Quận)',    idx: GHN_COL.QUAN    },
    { name: 'L (Tỉnh)',    idx: GHN_COL.TINH    },
  ];
  cols.forEach(c => {
    const raw = r[c.idx - 1];
    const s = String(raw || '');
    Logger.log(`  ${c.name}: "${s}" len=${s.length} codes=[${charCodesOf_(s)}]`);
    Logger.log(`     normalized="${normalizeAddrGhn_(s)}"`);
  });

  Logger.log('');
  Logger.log('-- Bước resolveAddress_ --');
  const resolved = resolveAddress_(r, String(r[GHN_COL.DIA_CHI - 1] || ''));
  Logger.log(`  resolved.ward     = "${resolved.ward}"`);
  Logger.log(`  resolved.district = "${resolved.district}"`);
  Logger.log(`  resolved.province = "${resolved.province}"`);
  Logger.log(`  source = ${resolved.source}`);

  Logger.log('');
  Logger.log('-- Bước lookupAddressGhn_ --');
  const lookup = lookupAddressGhn_(resolved.ward, resolved.district, resolved.province);
  if (lookup) {
    Logger.log(`  ✅ OK: districtId=${lookup.districtId}, wardCode=${lookup.wardCode}, source=${lookup.source}`);
  } else {
    Logger.log(`  ❌ FAIL — gọi debugAddressLookup để biết bước nào fail:`);
    debugAddressLookup(resolved.ward, resolved.district, resolved.province);
  }
}

// ============================================================
// HELPERS
// ============================================================

function ghnFetch_(path, payload) {
  const url = GHN_CFG.API_BASE + path;
  const opts = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Token': getGhnProp_(GHN_CFG.PROP_TOKEN),
      'ShopId': String(GHN_CFG.SHOP_ID),
    },
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true,
  };
  let res;
  try { res = UrlFetchApp.fetch(url, opts); }
  catch (e) { return { ok: false, error: 'Network: ' + e.message }; }
  const code = res.getResponseCode();
  const txt = res.getContentText();
  let body;
  try { body = JSON.parse(txt); } catch (e) { return { ok: false, error: `HTTP ${code} non-JSON: ${txt.slice(0,200)}` }; }
  if (code !== 200 || body.code !== 200) {
    const msg = body.message || body.code_message_value || body.code_message || txt.slice(0,200);
    return { ok: false, error: `[${code}/${body.code}] ${msg}` };
  }
  return { ok: true, data: body.data };
}

function openGhnTargetTab_() {
  const ss = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(GHN_CFG.TARGET_TAB);
  if (!sheet) throw new Error('Không tìm thấy tab "' + GHN_CFG.TARGET_TAB + '"');
  return sheet;
}

function getGhnProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function normalizeGhnPhone_(p) {
  let s = String(p || '').replace(/[^\d]/g, '');
  if (s.startsWith('84') && s.length > 9) s = '0' + s.slice(2);
  if (!s.startsWith('0') && s.length === 9) s = '0' + s;
  return s;
}

function formatGhnNow_() {
  return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
}

/**
 * Detect đơn pickup tại showroom — không cần tạo GHN.
 * Heuristic:
 *  - PTTT/shipping_method chứa "tại cửa hàng" / "showroom" / "pickup" / "nhận tại"
 *  - HOẶC địa chỉ trống và PTTT có dấu hiệu pickup khác
 */
function isPickupShipping_(pttt, diaChi) {
  const p = String(pttt || '').toLowerCase();
  if (/tại cửa hàng|tại của hàng|nhận tại|nh[aâ]n tại|showroom|pickup|store pickup|tự đến|tự lấy/i.test(p)) return true;
  // Đơn không có địa chỉ giao + PTTT nói "nhận" → 99% pickup
  if (!String(diaChi || '').trim() && /nhận|pickup|store/i.test(p)) return true;
  return false;
}

/**
 * Lookup quà tặng từ Scheme cho 1 group đơn.
 * Logic: thử match Camp (CMMF HANNAH) → SKU CP → SKU CB.
 * Trả về { skuGift, nameGift } hoặc null nếu đơn không có gift.
 */
function lookupGiftFromScheme_(group, scheme) {
  if (!scheme.giftBySku && !scheme.giftByCamp) return null;
  const r0 = group.rows[0];
  const camp = String(r0[GHN_COL.CAMP - 1] || '').trim().toLowerCase();

  // Match theo Camp (CMMF HANNAH) — đơn KOL ưu tiên
  if (camp && scheme.giftByCamp.has(camp)) return scheme.giftByCamp.get(camp);

  // Match theo SKU của line items (SKU CP hoặc SKU CB)
  for (const row of group.rows) {
    const sku = String(row[GHN_COL.MA_SP - 1] || '').trim();
    if (sku && scheme.giftBySku.has(sku)) return scheme.giftBySku.get(sku);
  }

  return null;
}

/**
 * Pre-flight validate đơn trước khi tạo GHN. Trả array lỗi (empty = OK).
 * Catch sớm các đơn data thiếu/sai để PG fix tay, tránh gọi GHN API rồi nhận 400.
 */
function validateOrderData_(group) {
  const errors = [];
  const firstRow = group.rows[0];
  const rawSdt = String(firstRow[GHN_COL.SDT - 1] || '').trim();
  const sdtClean = rawSdt.replace(/[\s.\-+]/g, '').replace(/^84/, '0');
  const tenKH = String(firstRow[GHN_COL.TEN_KH - 1] || '').trim();
  const tongDon = Number(firstRow[GHN_COL.TONG_DON - 1] || 0);
  const diaChi = String(firstRow[GHN_COL.DIA_CHI - 1] || '').trim();
  const phuong = String(firstRow[GHN_COL.PHUONG - 1] || '').trim();
  const quan = String(firstRow[GHN_COL.QUAN - 1] || '').trim();
  const tinh = String(firstRow[GHN_COL.TINH - 1] || '').trim();

  if (!/^0\d{9,10}$/.test(sdtClean)) errors.push(`SĐT không hợp lệ "${rawSdt}"`);
  if (!tenKH) errors.push('Thiếu tên KH');
  if (tongDon <= 0) errors.push('Tổng đơn = 0');
  if (!diaChi && !phuong) errors.push('Thiếu địa chỉ');
  if (!tinh) errors.push('Thiếu Tỉnh/TP');
  // SP check: phải có ít nhất 1 SKU + SL > 0
  let hasItem = false;
  group.rows.forEach(row => {
    if (String(row[GHN_COL.MA_SP - 1] || '').trim() && Number(row[GHN_COL.SL - 1] || 0) > 0) hasItem = true;
  });
  if (!hasItem) errors.push('Thiếu SKU/SL hợp lệ');

  return errors;
}

/** Email cảnh báo PG list đơn lỗi data */
function sendDataInvalidAlertEmail_(invalidOrders) {
  const ALERT_EMAIL = 'admin@khomes.com.vn';
  const subject = `⛔ [K-Homes] ${invalidOrders.length} đơn LỖI DATA — chặn tạo GHN`;
  const rows = invalidOrders.map((d, i) => {
    const r = d.group.rows[0];
    return `<tr>
      <td>${i + 1}</td><td><b>${d.orderId}</b></td>
      <td>${r[GHN_COL.TEN_KH - 1] || ''}</td>
      <td>${r[GHN_COL.SDT - 1] || ''}</td>
      <td>${r[GHN_COL.TINH - 1] || ''}</td>
      <td style="color:#c0504d;">${d.errors.join('<br>')}</td>
    </tr>`;
  }).join('');
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h2 style="color:#c0504d;">⛔ ${invalidOrders.length} đơn bị chặn tạo GHN (lỗi data)</h2>
      <p>Pre-flight check phát hiện thiếu/sai SĐT, tên KH, địa chỉ, hoặc SKU/SL. PG cần:</p>
      <ol>
        <li>Mở tab "tracking haravan", search Mã đơn → fix data tay</li>
        <li>Set lại cột X = "Đã xác nhận" → cron Phase 2 sẽ retry tự động</li>
      </ol>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
        <tr style="background:#1f4e78;color:#fff;">
          <th>#</th><th>Mã đơn</th><th>Tên KH</th><th>SĐT</th><th>Tỉnh</th><th>Lỗi</th>
        </tr>
        ${rows}
      </table>
    </div>`;
  MailApp.sendEmail({ to: ALERT_EMAIL, subject, htmlBody: html });
}

/**
 * Check tồn kho cho 1 group đơn (1 mã Haravan, có thể có nhiều SKU).
 * Trả về array các SKU thiếu: [{sku, need, have}, ...]. Empty array = đủ hàng.
 * Bỏ qua SKU không có trong stockMap (chưa sync hoặc SP không quản tồn).
 */
function checkStockForGroup_(group, stockMap) {
  if (!stockMap || Object.keys(stockMap).length === 0) return [];   // chưa có data stock → cho qua
  const need = {};
  group.rows.forEach(row => {
    const sku = String(row[GHN_COL.MA_SP - 1] || '').trim();
    if (!sku) return;
    const sl = Number(row[GHN_COL.SL - 1] || 0);
    if (!sl) return;
    need[sku] = (need[sku] || 0) + sl;
  });
  const short = [];
  Object.keys(need).forEach(sku => {
    if (!(sku in stockMap)) return;   // SKU không trong tồn kho → trust kho
    const have = stockMap[sku];
    if (have < need[sku]) short.push({ sku, need: need[sku], have });
  });
  return short;
}

/**
 * Email cảnh báo đơn thiếu hàng — list SKU + mã đơn để PG/kho biết liên hệ NCC.
 */
function sendStockShortAlertEmail_(stockShortOrders) {
  const ALERT_EMAIL = 'admin@khomes.com.vn';
  const subject = `📦 [K-Homes] ${stockShortOrders.length} đơn THIẾU HÀNG — không tạo được GHN`;
  const rows = stockShortOrders.map((s, i) => {
    const itemsHtml = s.shortItems.map(it =>
      `<li><b>${it.sku}</b>: cần ${it.need}, tồn còn ${it.have} <span style="color:#c0504d;">(thiếu ${it.need - it.have})</span></li>`
    ).join('');
    const firstRow = s.group.rows[0];
    const tenKH = firstRow[GHN_COL.TEN_KH - 1] || '';
    const sdt = firstRow[GHN_COL.SDT - 1] || '';
    const tinh = firstRow[GHN_COL.TINH - 1] || '';
    const tongDon = firstRow[GHN_COL.TONG_DON - 1] || 0;
    return `
      <tr>
        <td>${i + 1}</td>
        <td><b>${s.orderId}</b></td>
        <td>${tenKH}</td>
        <td><a href="tel:${sdt}">${sdt}</a></td>
        <td>${tinh}</td>
        <td>${Number(tongDon).toLocaleString('vi-VN')}đ</td>
        <td><ul style="margin:0;padding-left:18px;">${itemsHtml}</ul></td>
      </tr>`;
  }).join('');
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h2 style="color:#c0504d;">📦 ${stockShortOrders.length} đơn THIẾU HÀNG — đã skip tạo GHN</h2>
      <p>Các đơn dưới đây có SKU không đủ tồn kho (theo Haravan inventory cập nhật mỗi sáng 6h). PG/kho cần xử lý:</p>
      <ol>
        <li>Liên hệ NCC nhập thêm hàng → cron sau khi sync stock + Phase 2 sẽ retry tự động</li>
        <li>Hoặc liên hệ KH đổi SKU thay thế / hoãn đơn</li>
        <li>Hoặc set cột X = "Đã xác nhận" (override) nếu kho thực tế còn hàng nhưng Haravan chưa cập nhật</li>
      </ol>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
        <tr style="background:#1f4e78;color:#fff;">
          <th>#</th><th>Mã đơn</th><th>Tên KH</th><th>SĐT</th>
          <th>Tỉnh</th><th>Tổng đơn</th><th>SKU thiếu</th>
        </tr>
        ${rows}
      </table>
      <p style="margin-top:15px;color:#666;">
        ➤ Email tự động từ Phase 2 Wellhome Order Sync. Sau khi nhập thêm hàng, tab "stock haravan" sẽ tự refresh sáng hôm sau và đơn sẽ retry.
      </p>
    </div>`;
  MailApp.sendEmail({ to: ALERT_EMAIL, subject, htmlBody: html });
}

function stringSimilarity_(a, b) {
  a = (a || '').toLowerCase(); b = (b || '').toLowerCase();
  if (!a || !b) return 0;
  if (a === b) return 1;
  const m = a.length, n = b.length;
  if (m > 200 || n > 200) return 0;
  const dp = new Array(n + 1);
  for (let j = 0; j <= n; j++) dp[j] = j;
  for (let i = 1; i <= m; i++) {
    let prev = i - 1; dp[0] = i;
    for (let j = 1; j <= n; j++) {
      const tmp = dp[j];
      dp[j] = a[i-1] === b[j-1] ? prev : 1 + Math.min(prev, dp[j], dp[j-1]);
      prev = tmp;
    }
  }
  const dist = dp[n];
  return 1 - dist / Math.max(m, n);
}

// ============================================================
// UTILITIES
// ============================================================

function showGhnProps() {
  const props = PropertiesService.getScriptProperties().getProperties();
  Object.keys(props).filter(k => k.startsWith('GHN_')).forEach(k => {
    let v = props[k];
    if (k.includes('TOKEN')) v = v.slice(0, 8) + '...' + v.slice(-4);
    if (k.includes('CACHE') && v.length > 80) v = `<JSON ${v.length} chars>`;
    Logger.log(`${k} = ${v}`);
  });
}

function retryFailedOrders() {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const col = sheet.getRange(2, GHN_COL.TT_XU_LY, lastRow - 1, 1).getValues();
  let count = 0;
  col.forEach((r, i) => {
    if (r[0] === 'Lỗi GHN') {
      sheet.getRange(i + 2, GHN_COL.TT_XU_LY).setValue('Đã xác nhận');
      count++;
    }
  });
  Logger.log(`✅ Reset ${count} đơn lỗi → "Đã xác nhận"`);
}

function dryRunOrder(orderName) {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, GHN_COL.GHN_AT).getValues();
  const rows = data.filter(r => String(r[GHN_COL.MA_HARAVAN - 1]) === String(orderName));
  if (rows.length === 0) { Logger.log('Không tìm thấy đơn ' + orderName); return; }
  const scheme = loadSchemeCache_();
  const dim = lookupDimensions_({ rows, rowIndices: [] }, scheme);
  Logger.log(`Đơn ${orderName}: ${rows.length} dòng | KL=${dim.weight}g, ${dim.length}x${dim.width}x${dim.height}cm | source=${dim.source}`);
}

// ============================================================
// 🛑 EMERGENCY — DỪNG/KHỞI ĐỘNG TẤT CẢ 4 CRON
// ============================================================

const PIPELINE_TRIGGERS = ['syncOrders', 'createGhnOrders', 'updateGhnStatuses', 'fulfillHaravanOrders'];

/** List toàn bộ trigger đang có trong project — kiểm tra trước khi pause/resume */
function showAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`Tổng trigger: ${triggers.length}`);
  triggers.forEach((t, i) => {
    const fn = t.getHandlerFunction();
    const src = t.getTriggerSource();
    Logger.log(`  ${i + 1}. ${fn} | source=${src} | id=${t.getUniqueId()}`);
  });
}

/** DỪNG cả 4 cron — đơn KHÔNG bị mất, chỉ tool ngừng tự động chạy */
function pauseAllPipelineTriggers() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (PIPELINE_TRIGGERS.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log(`🛑 Đã PAUSE ${removed} trigger pipeline. Chạy resumeAllPipelineTriggers để khởi động lại.`);
}

/** Khởi động lại 4 cron với schedule chuẩn (Phase 1 + 2: 1h, Phase 3 + 4: 30 phút) */
function resumeAllPipelineTriggers() {
  pauseAllPipelineTriggers();   // xóa cũ trước (idempotent)
  ScriptApp.newTrigger('syncOrders').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('createGhnOrders').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('updateGhnStatuses').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('fulfillHaravanOrders').timeBased().everyMinutes(30).create();
  Logger.log('✅ Đã khởi động 4 cron: syncOrders 1h, createGhnOrders 1h, updateGhnStatuses 30p, fulfillHaravanOrders 30p');
}

// ============================================================
// 🟡 DRY RUN — TEST AN TOÀN, KHÔNG TẠO ĐƠN GHN THẬT
// ============================================================

/**
 * Quét sheet, log toàn bộ payload SẼ gửi GHN, KHÔNG call API tạo đơn.
 * Dùng trước khi chạy createGhnOrders để verify đơn nào sẽ được tạo + KL/Kích thước/Người gửi đúng chưa.
 */
function dryRunCreateGhnOrders() {
  Logger.log('🟡 DRY RUN — KHÔNG gọi GHN API tạo đơn, chỉ log payload sẽ gửi');

  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng, skip'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, GHN_COL.GHN_AT).getValues();
  const groups = {};
  data.forEach((row, idx) => {
    const status = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    if (status !== 'Đã xác nhận') return;
    if (row[GHN_COL.GHN_MA - 1]) return;
    const orderId = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim();
    if (!orderId) return;
    if (!groups[orderId]) groups[orderId] = { orderId, rows: [], rowIndices: [] };
    groups[orderId].rows.push(row);
    groups[orderId].rowIndices.push(idx + 2);
  });
  const orderIds = Object.keys(groups);
  Logger.log(`Số đơn SẼ tạo GHN: ${orderIds.length}`);
  if (orderIds.length === 0) return;

  const scheme = loadSchemeCache_();
  const fromDistrictId = parseInt(getGhnProp_(GHN_CFG.PROP_FROM_DISTRICT_ID), 10);
  const fromWardCode   = String(getGhnProp_(GHN_CFG.PROP_FROM_WARD_CODE) || '');
  if (!fromDistrictId || !fromWardCode) {
    Logger.log(`⚠️ Sender chưa setup. Run setupSenderAddress trước. Vẫn dry-run được nhưng không có district sender.`);
  }

  for (let i = 0; i < Math.min(orderIds.length, 10); i++) {
    const g = groups[orderIds[i]];
    const r0 = g.rows[0];
    Logger.log(`\n--- Đơn ${i + 1}/${orderIds.length}: ${g.orderId} ---`);
    Logger.log(`  KH: ${r0[GHN_COL.TEN_KH - 1]} | SĐT: ${r0[GHN_COL.SDT - 1]}`);
    Logger.log(`  Địa chỉ (cột I): ${r0[GHN_COL.DIA_CHI - 1]}`);
    Logger.log(`  Cột J/K/L: P="${r0[GHN_COL.PHUONG - 1]}" | Q="${r0[GHN_COL.QUAN - 1]}" | T="${r0[GHN_COL.TINH - 1]}"`);
    const resolved = resolveAddress_(r0, String(r0[GHN_COL.DIA_CHI - 1] || ''));
    Logger.log(`  Hợp nhất: P="${resolved.ward}" | Q="${resolved.district}" | T="${resolved.province}" | nguồn=${resolved.source}`);
    if (resolved.province) {
      const addrLookup = lookupAddressGhn_(resolved.ward, resolved.district, resolved.province);
      if (addrLookup) {
        Logger.log(`  ✓ Lookup GHN: districtId=${addrLookup.districtId}, wardCode=${addrLookup.wardCode} (${addrLookup.source})`);
      } else {
        Logger.log(`  ✗ Lookup GHN FAIL — tool sẽ set "Lỗi GHN"`);
      }
    }
    const dim = lookupDimensions_(g, scheme);
    Logger.log(`  KL+Kích thước: ${dim.weight}g | ${dim.length}×${dim.width}×${dim.height}cm | source=${dim.source}`);
    const items = g.rows.map(row => ({
      sku: String(row[GHN_COL.MA_SP - 1] || ''),
      name: String(row[GHN_COL.TEN_SP - 1] || '').slice(0, 60),
      qty: row[GHN_COL.SL - 1],
    }));
    Logger.log(`  Items: ${JSON.stringify(items)}`);
    const camp = String(r0[GHN_COL.CAMP - 1] || '');
    const sender = detectSenderByBrand_(camp, items.map(it => ({ name: it.name })));
    const pttt = String(r0[GHN_COL.PTTT - 1] || '').toLowerCase();
    const isCOD = pttt.includes('cod') || pttt.includes('khi giao') || pttt.includes('cash on');
    const codAmount = isCOD ? Math.round(Number(r0[GHN_COL.TONG_DON - 1] || 0)) : 0;
    Logger.log(`  Người gửi: ${sender} | COD: ${isCOD ? codAmount + 'đ' : 'KHÔNG'}`);
  }
  if (orderIds.length > 10) Logger.log(`\n... (còn ${orderIds.length - 10} đơn không log chi tiết)`);
  Logger.log(`\n✅ DRY RUN OK — chạy createGhnOrders để tạo đơn GHN thật`);
}
