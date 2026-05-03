/**
 * WELLHOME ORDER SYNC — Phase 1 v2 (UX upgraded)
 * Pull đơn mới từ Haravan wellhome về tab "tracking haravan"
 * Cron 1h/lần. Idempotent (không trùng đơn). KHÔNG ghi đè cột X (TT xử lý) và Y (Ghi chú NV).
 */

const CONFIG = {
  HARAVAN_SHOP: 'wellho-me.myharavan.com',
  TARGET_SHEET_ID: '1Bn4C0UdvX2hT82p1pput3VtBDkdJe_jc1-w8ZjKfdGw',
  TARGET_TAB: 'tracking haravan',
  LOOKBACK_DAYS_FIRST_RUN: 7,
  PAGE_SIZE: 50,
  MAX_PAGES_PER_RUN: 5,
  PROP_TOKEN: 'WELLHOME_HARAVAN_TOKEN',
  PROP_LAST_SYNC: 'WELLHOME_LAST_SYNC',
  // Auto-confirm: đơn paid + loại KOL/Organic → set X = "Đã xác nhận"
  // (theo nhân viên video: "đơn đã thanh toán không cần xác nhận với khách").
  // Nội bộ + Khác vẫn giữ "Chưa xử lý" để PG kiểm tra tay.
  AUTO_CONFIRM_PAID: true,
  AUTO_CONFIRM_TYPES: ['KOL', 'Organic'],
};

const HEADERS = [
  'Ngày đặt',           // A
  'Mã đơn Haravan',     // B
  'Mã HO',              // C
  'Loại đơn',           // D
  'Camp',               // E
  'Tên KH',             // F
  'SĐT',                // G
  'Email',              // H
  'Địa chỉ',            // I
  'Phường/Xã',          // J
  'Quận/Huyện',         // K
  'Tỉnh/TP',            // L
  'Vùng miền',          // M
  'PTTT',               // N
  'Mã SP',              // O
  'Tên SP',             // P
  'SL',                 // Q
  'Đơn giá',            // R
  'Tổng đơn',           // S
  'TT thanh toán',      // T
  'TT giao hàng',       // U
  'Ghi chú đơn',        // V
  'Cập nhật lần cuối',  // W
  'Trạng thái xử lý',   // X  ← do nhân viên cập nhật, tool KHÔNG ghi đè
  'Ghi chú nhân viên',  // Y  ← do nhân viên cập nhật, tool KHÔNG ghi đè
];

const STATUS_LIST = ['Chưa xử lý', 'Đang gọi KH', 'Đã xác nhận', 'Đã tạo GHN', 'Lỗi GHN', 'Lỗi data', 'Thiếu hàng', 'Pickup tại SR', 'Hủy', 'Đã giao'];
const ORDER_TYPE_LIST = ['KOL', 'Organic', 'Nội bộ', 'Khác'];
const REGION_LIST = ['Bắc', 'Trung', 'Nam'];
const FINANCIAL_LIST = ['Chờ thanh toán', 'Đã uỷ quyền', 'Đã thanh toán', 'Thanh toán 1 phần', 'Đã hoàn tiền', 'Hoàn 1 phần', 'Đã hủy TT'];
const FULFILLMENT_LIST = ['Đã giao', 'Giao 1 phần', 'Chưa giao', 'Đã trả về kho'];

// ============================================================
// SETUP — chạy LẦN ĐẦU sau khi paste code
// ============================================================

/** Lưu token Haravan vào PropertiesService */
function setupToken() {
  const TOKEN = 'REPLACE_WITH_HARAVAN_TOKEN';
  PropertiesService.getScriptProperties().setProperty(CONFIG.PROP_TOKEN, TOKEN);
  Logger.log('✅ Đã lưu token Haravan');
}

/** Tạo header 25 cột cho tab "tracking haravan" */
function setupHeader() {
  const sheet = openTab_();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  const widths = [110,130,100,80,140,130,100,170,220,110,120,100,80,130,110,230,50,100,110,130,110,200,130,130,200];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange('A:A').setNumberFormat('@');
  // Fix locale en_US → vi_VN để Date không bị parse nhầm dd↔mm.
  // Idempotent — gọi nhiều lần không sao.
  try {
    const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
    if (ss.getSpreadsheetLocale() !== 'vi_VN') {
      ss.setSpreadsheetLocale('vi_VN');
      Logger.log(`✅ Đã đổi locale Sheet sang vi_VN`);
    }
  } catch (e) {
    Logger.log(`⚠️ setSpreadsheetLocale: ${e.message}`);
  }
  Logger.log(`✅ Đã tạo header 25 cột cho tab "${CONFIG.TARGET_TAB}"`);
}

/** Test kết nối Haravan */
function testConnection() {
  const url = `https://${CONFIG.HARAVAN_SHOP}/admin/orders.json?limit=3&order=created_at%20desc`;
  const res = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + getToken_() },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code !== 200) {
    Logger.log(`❌ HTTP ${code}: ${res.getContentText().slice(0, 300)}`);
    return;
  }
  const orders = (JSON.parse(res.getContentText()).orders || []);
  Logger.log(`✅ Kết nối OK — ${orders.length} đơn lấy ra để test`);
  orders.forEach(o => {
    Logger.log(`   ${o.name} | ${o.shipping_address?.name || '?'} | ${formatDate_(o.created_at)} | ${o.total_price}đ`);
  });
}

/** Cài cron 1h/lần */
function setupHourlyTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('syncOrders').timeBased().everyHours(1).create();
  Logger.log('✅ Đã cài trigger: syncOrders chạy mỗi 1h');
}

// ============================================================
// UPGRADE TO V2 — chạy 1 LẦN sau khi paste code v2
// ============================================================

/** Hàm gộp: xóa 2 tab BOSCH cũ + setup header v2 + migrate data + apply formatting */
function setupV2() {
  Logger.log('===== Bắt đầu nâng cấp v2 =====');
  deleteOldBoschTabs();
  setupHeader();
  migrateExistingData();
  applyFormatting();
  Logger.log('===== ✅ Nâng cấp v2 hoàn tất =====');
}

/** Xóa 2 tab BOSCH cũ */
function deleteOldBoschTabs() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const NAMES = ['Đơn hàng', 'Theo dõi vận chuyển'];
  const deleted = [];
  NAMES.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.deleteSheet(sh); deleted.push(name); }
  });
  Logger.log(`✅ Đã xóa tab: ${deleted.join(', ') || '(không tìm thấy tab nào)'}`);
}

/** Fill cột X="Chưa xử lý" và Y="" cho data cũ (chạy 1 lần khi nâng cấp) */
function migrateExistingData() {
  const sheet = openTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Không có data cũ, skip migration'); return; }
  const range = sheet.getRange(2, 24, lastRow - 1, 2); // cột X (24) và Y (25)
  const values = range.getValues();
  let filled = 0;
  values.forEach(r => {
    if (!r[0]) { r[0] = 'Chưa xử lý'; filled++; }
    if (r[1] === undefined || r[1] === null) r[1] = '';
  });
  range.setValues(values);
  Logger.log(`✅ Migration: fill "Chưa xử lý" cho ${filled} dòng cũ`);
}

/** Apply số format + dropdown + conditional formatting */
function applyFormatting() {
  const sheet = openTab_();
  const N = Math.max(sheet.getLastRow(), 1000);

  // Number format số tiền
  sheet.getRange(`Q2:Q${N}`).setNumberFormat('0');
  sheet.getRange(`R2:R${N}`).setNumberFormat('#,##0" đ"');
  sheet.getRange(`S2:S${N}`).setNumberFormat('#,##0" đ"');

  // Dropdown cột X (nhân viên chọn — strict)
  sheet.getRange(`X2:X${N}`).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_LIST, true).setAllowInvalid(false).build()
  );

  // Dropdown các cột tool fill (allow invalid — phòng Haravan thêm giá trị mới)
  setDropdown_(sheet, `D2:D${N}`, ORDER_TYPE_LIST);
  setDropdown_(sheet, `M2:M${N}`, REGION_LIST);
  setDropdown_(sheet, `T2:T${N}`, FINANCIAL_LIST);
  setDropdown_(sheet, `U2:U${N}`, FULFILLMENT_LIST);

  // Conditional formatting
  const rules = [];
  // Cột D: Loại đơn
  rules.push(condEq_(sheet, 'D', N, 'KOL', '#E8DAEF'));
  rules.push(condEq_(sheet, 'D', N, 'Organic', '#D5E8D4'));
  rules.push(condEq_(sheet, 'D', N, 'Nội bộ', '#F2F2F2'));
  rules.push(condEq_(sheet, 'D', N, 'Khác', '#FFE6CC'));
  // Cột T: TT thanh toán
  rules.push(condEq_(sheet, 'T', N, 'Đã thanh toán', '#D5E8D4'));
  rules.push(condEq_(sheet, 'T', N, 'Chờ thanh toán', '#FFF2CC'));
  rules.push(condEq_(sheet, 'T', N, 'Đã hoàn tiền', '#F8CECC'));
  rules.push(condEq_(sheet, 'T', N, 'Đã hủy TT', '#F8CECC'));
  // Cột U: TT giao
  rules.push(condEq_(sheet, 'U', N, 'Đã giao', '#D5E8D4'));
  rules.push(condEq_(sheet, 'U', N, 'Chưa giao', '#FFF2CC'));
  // Cột X: Trạng thái xử lý
  rules.push(condEq_(sheet, 'X', N, 'Chưa xử lý', '#FFF2CC'));
  rules.push(condEq_(sheet, 'X', N, 'Đang gọi KH', '#DAE8FC'));
  rules.push(condEq_(sheet, 'X', N, 'Đã xác nhận', '#D5E8D4'));
  rules.push(condEq_(sheet, 'X', N, 'Đã tạo GHN', '#9BD78A'));
  rules.push(condEq_(sheet, 'X', N, 'Hủy', '#F8CECC'));
  rules.push(condEq_(sheet, 'X', N, 'Đã giao', '#82C26B'));

  // Tô xám nhạt cho dòng tiếp của cùng đơn (nhận diện đơn nhiều SP)
  const dataRange = sheet.getRange(`A2:Y${N}`);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2<>"", $B2=$B1)')
    .setBackground('#F8F8F8')
    .setRanges([dataRange]).build());

  sheet.setConditionalFormatRules(rules);
  Logger.log(`✅ Đã apply formatting (số tiền, dropdown, màu trạng thái)`);
}

function condEq_(sheet, col, lastRow, value, bg) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(value).setBackground(bg)
    .setRanges([sheet.getRange(`${col}2:${col}${lastRow}`)]).build();
}

function setDropdown_(sheet, a1, list) {
  sheet.getRange(a1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(list, true).setAllowInvalid(true).build()
  );
}

// ============================================================
// MAIN — hàm cron chính
// ============================================================

function syncOrders() {
  const sheet = openTab_();
  if (sheet.getLastColumn() < HEADERS.length) setupHeader();

  const props = PropertiesService.getScriptProperties();
  let lastSync = props.getProperty(CONFIG.PROP_LAST_SYNC);
  if (!lastSync) {
    const d = new Date();
    d.setDate(d.getDate() - CONFIG.LOOKBACK_DAYS_FIRST_RUN);
    lastSync = d.toISOString();
    Logger.log(`Lần đầu — lookback ${CONFIG.LOOKBACK_DAYS_FIRST_RUN} ngày từ ${lastSync}`);
  } else {
    Logger.log(`Pull đơn từ ${lastSync}`);
  }

  // Đọc Mã đơn Haravan đã có (cột B) để skip duplicate
  const lastRow = sheet.getLastRow();
  const existing = new Set();
  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, 1).getValues().forEach(r => {
      if (r[0]) existing.add(String(r[0]).trim());
    });
  }

  let page = 1, totalFetched = 0;
  const newRows = [];
  let maxCreatedAt = lastSync;

  while (page <= CONFIG.MAX_PAGES_PER_RUN) {
    const url = `https://${CONFIG.HARAVAN_SHOP}/admin/orders.json`
      + `?limit=${CONFIG.PAGE_SIZE}`
      + `&page=${page}`
      + `&created_at_min=${encodeURIComponent(lastSync)}`
      + `&order=created_at+asc`;
    const res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + getToken_() },
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code !== 200) {
      Logger.log(`❌ Page ${page} HTTP ${code}: ${res.getContentText().slice(0, 200)}`);
      break;
    }
    const orders = JSON.parse(res.getContentText()).orders || [];
    if (orders.length === 0) { Logger.log(`Page ${page}: hết đơn`); break; }
    totalFetched += orders.length;
    Logger.log(`Page ${page}: ${orders.length} đơn`);

    orders.forEach(o => {
      if (o.cancelled_at) return;
      if (['voided', 'refunded'].includes(o.financial_status)) return;
      if (existing.has(o.name)) return;

      const parsed = parseNote_(o.note);
      const orderType = classifyOrderType_(o, parsed);
      const sa = o.shipping_address || {};
      const province = getProvinceWithFallback_(sa);
      const lineItems = o.line_items || [];
      const now = formatDate_(new Date().toISOString());

      const baseRow = [
        formatDate_(o.created_at), o.name, parsed.sourceOrderCode, orderType, parsed.camp,
        sa.name || '', sa.phone || '', o.email || '', sa.address1 || '',
        sa.ward || '', sa.district || '', province,
        detectRegion_(province), o.gateway || '',
      ];
      // Pickup detection sớm: nếu PTTT/shipping_method là "nhận tại cửa hàng"
      // → set "Pickup tại SR" ngay từ Phase 1 (không cần đợi Phase 2 cron 1h sau).
      // Defensive: typeof check vì isPickupShipping_ định nghĩa trong Phase 2.
      const ptttRaw = String(o.gateway || '').toLowerCase();
      const shippingMethod = String((o.shipping_lines && o.shipping_lines[0] && o.shipping_lines[0].title) || '').toLowerCase();
      const ptttCombined = (ptttRaw + ' ' + shippingMethod).trim();
      const isPickup = (typeof isPickupShipping_ === 'function')
        ? isPickupShipping_(ptttCombined, sa.address1 || '')
        : /tại cửa hàng|nhận tại|showroom|pickup|tự đến|tự lấy/i.test(ptttCombined);

      // Auto-confirm: đơn paid + loại KOL/Organic → bỏ qua bước PG xác nhận tay
      const isPaid = (o.financial_status === 'paid');
      const canAutoConfirm = CONFIG.AUTO_CONFIRM_PAID
        && isPaid
        && CONFIG.AUTO_CONFIRM_TYPES.indexOf(orderType) >= 0;

      let initialStatus, initialNote;
      if (isPickup) {
        initialStatus = 'Pickup tại SR';
        initialNote = `Auto-detect ${now}: đơn pickup tại showroom (Phase 1)`;
      } else if (canAutoConfirm) {
        initialStatus = 'Đã xác nhận';
        initialNote = `Auto-confirm ${now}: KH paid (${orderType})`;
      } else {
        initialStatus = 'Chưa xử lý';
        initialNote = '';
      }

      const tailRow = [
        o.total_price || 0,
        translateFinancialStatus_(o.financial_status),
        translateFulfillmentStatus_(o.fulfillment_status),
        parsed.cleanNote, now,
        initialStatus, initialNote,  // X, Y - auto-confirm cho đơn paid KOL/Organic
      ];

      if (lineItems.length === 0) {
        newRows.push([...baseRow, '', '', '', '', ...tailRow]);
      } else {
        lineItems.forEach(li => {
          newRows.push([
            ...baseRow,
            li.sku || '', li.title || '', li.quantity || '', li.price || 0,
            ...tailRow,
          ]);
        });
      }
      if (o.created_at > maxCreatedAt) maxCreatedAt = o.created_at;
    });

    if (orders.length < CONFIG.PAGE_SIZE) break;
    page++;
  }

  Logger.log(`Tổng đơn fetch: ${totalFetched}, dòng append: ${newRows.length}`);
  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, HEADERS.length).setValues(newRows);
    Logger.log(`✅ Đã append ${newRows.length} dòng từ row ${startRow}`);
  }

  if (maxCreatedAt && maxCreatedAt !== lastSync) {
    const next = new Date(new Date(maxCreatedAt).getTime() + 1).toISOString();
    props.setProperty(CONFIG.PROP_LAST_SYNC, next);
    Logger.log(`Cập nhật last sync = ${next}`);
  }
}

// ============================================================
// HELPERS
// ============================================================

function openTab_() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.TARGET_TAB);
  if (!sheet) throw new Error(`Không tìm thấy tab "${CONFIG.TARGET_TAB}"`);
  return sheet;
}

function getToken_() {
  const t = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_TOKEN);
  if (!t) throw new Error('Chưa setup token. Chạy setupToken() trước.');
  return t;
}

function parseNote_(note) {
  if (!note) return { camp: '', brand: '', sourceOrderCode: '', cleanNote: '' };
  const camp = (note.match(/Camp:\s*([^|]+?)\s*(?:\||$)/i) || [])[1] || '';
  const brand = (note.match(/Brand:\s*([^|]+?)\s*(?:\||$)/i) || [])[1] || '';
  const src = (note.match(/Source order code:\s*([^|]+?)\s*(?:\||$)/i) || [])[1] || '';
  let clean = note
    .replace(/Camp:\s*[^|]+\|?/gi, '')
    .replace(/Brand:\s*[^|]+\|?/gi, '')
    .replace(/Source order code:\s*[^|]+\|?/gi, '')
    .replace(/\|\s*\|/g, '|').trim();
  clean = clean.replace(/^\|+\s*|\s*\|+$/g, '').trim();
  return { camp: camp.trim(), brand: brand.trim(), sourceOrderCode: src.trim(), cleanNote: clean };
}

function classifyOrderType_(order, parsed) {
  if (parsed.camp) return 'KOL';
  const note = (order.note || '').toLowerCase();
  if (note.includes('nội bộ') || note.includes('noi bo') || note.includes('nhân viên')) return 'Nội bộ';
  const src = order.source_name || '';
  if (['web', 'pos', 'mobile_app', 'iphone', 'android'].includes(src)) return 'Organic';
  return 'Khác';
}

const REGION_BAC = ['hà nội','hà giang','cao bằng','bắc kạn','tuyên quang','lào cai','điện biên','lai châu','sơn la','yên bái','hoà bình','hòa bình','thái nguyên','lạng sơn','quảng ninh','bắc giang','phú thọ','vĩnh phúc','bắc ninh','hải dương','hải phòng','hưng yên','thái bình','hà nam','nam định','ninh bình'];
const REGION_TRUNG = ['thanh hoá','thanh hóa','nghệ an','hà tĩnh','quảng bình','quảng trị','thừa thiên huế','thừa thiên-huế','huế','đà nẵng','quảng nam','quảng ngãi','bình định','phú yên','khánh hoà','khánh hòa','ninh thuận','bình thuận','kon tum','gia lai','đắk lắk','đắk nông','lâm đồng'];
const REGION_NAM = ['hồ chí minh','tp.hcm','tp hcm','tphcm','sài gòn','bình phước','tây ninh','bình dương','đồng nai','bà rịa','vũng tàu','long an','tiền giang','bến tre','trà vinh','vĩnh long','đồng tháp','an giang','kiên giang','cần thơ','hậu giang','sóc trăng','bạc liêu','cà mau'];

function detectRegion_(province) {
  if (!province) return '';
  const p = province.trim().toLowerCase().replace(/^(thành phố|tp\.?|tỉnh)\s+/i, '');
  if (REGION_BAC.some(x => p.includes(x))) return 'Bắc';
  if (REGION_TRUNG.some(x => p.includes(x))) return 'Trung';
  if (REGION_NAM.some(x => p.includes(x))) return 'Nam';
  return '';
}

/** Nếu province trống → parse từ address1 */
function getProvinceWithFallback_(sa) {
  if (sa && sa.province) return sa.province;
  const addr = (sa && sa.address1 || '').toLowerCase();
  if (!addr) return '';
  const ALL = REGION_BAC.concat(REGION_TRUNG).concat(REGION_NAM);
  // Sort dài → ngắn để khớp "Hà Nội" trước "Hà"
  const sorted = ALL.slice().sort((a, b) => b.length - a.length);
  for (const p of sorted) {
    if (addr.includes(p)) {
      // Capitalize tương đối
      return p.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
    }
  }
  return '';
}

function translateFinancialStatus_(s) {
  return ({
    'pending': 'Chờ thanh toán',
    'authorized': 'Đã uỷ quyền',
    'paid': 'Đã thanh toán',
    'partially_paid': 'Thanh toán 1 phần',
    'refunded': 'Đã hoàn tiền',
    'partially_refunded': 'Hoàn 1 phần',
    'voided': 'Đã hủy TT',
  })[s] || s || '';
}

function translateFulfillmentStatus_(s) {
  return ({
    'fulfilled': 'Đã giao',
    'partial': 'Giao 1 phần',
    'notfulfilled': 'Chưa giao',
    'unfulfilled': 'Chưa giao',
    'restocked': 'Đã trả về kho',
  })[s] || s || '';
}

function formatDate_(iso) {
  if (!iso) return '';
  return Utilities.formatDate(new Date(iso), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
}

// ============================================================
// UTILITIES
// ============================================================

/** Xóa toàn bộ data, giữ header */
function clearAllData() {
  const sheet = openTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  Logger.log('✅ Đã xóa data, giữ header');
}

function resetLastSync() {
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.PROP_LAST_SYNC);
  Logger.log('✅ Đã reset last sync');
}

function forceSyncDaysAgo(n) {
  n = n || 7;
  const d = new Date();
  d.setDate(d.getDate() - n);
  PropertiesService.getScriptProperties().setProperty(CONFIG.PROP_LAST_SYNC, d.toISOString());
  Logger.log(`Set last sync = ${d.toISOString()} (${n} ngày trước). Bắt đầu sync...`);
  syncOrders();
}

// ============================================================
// 🟡 DRY RUN — TEST AN TOÀN, CHỈ ĐỌC HARAVAN, KHÔNG GHI SHEET
// ============================================================

/**
 * Pull đơn từ Haravan như syncOrders nhưng KHÔNG ghi sheet, chỉ log danh sách đơn mới sẽ append.
 * @param {number} daysAgo  (optional) lookback N ngày. Default đọc PROP_LAST_SYNC như syncOrders.
 */
function dryRunSyncOrders(daysAgo) {
  Logger.log('🟡 DRY RUN — KHÔNG ghi sheet, chỉ log đơn sẽ pull');

  const sheet = openTab_();
  let lastSync;
  if (daysAgo && daysAgo > 0) {
    const d = new Date();
    d.setDate(d.getDate() - daysAgo);
    lastSync = d.toISOString();
    Logger.log(`Lookback ${daysAgo} ngày từ ${lastSync}`);
  } else {
    lastSync = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_LAST_SYNC);
    if (!lastSync) {
      const d = new Date();
      d.setDate(d.getDate() - CONFIG.LOOKBACK_DAYS_FIRST_RUN);
      lastSync = d.toISOString();
    }
    Logger.log(`Pull đơn từ ${lastSync}`);
  }

  const lastRow = sheet.getLastRow();
  const existing = new Set();
  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, 1).getValues().forEach(r => {
      if (r[0]) existing.add(String(r[0]).trim());
    });
  }

  let page = 1, totalFetched = 0, newCount = 0, dupCount = 0;
  while (page <= CONFIG.MAX_PAGES_PER_RUN) {
    const url = `https://${CONFIG.HARAVAN_SHOP}/admin/orders.json`
      + `?limit=${CONFIG.PAGE_SIZE}&page=${page}`
      + `&created_at_min=${encodeURIComponent(lastSync)}&order=created_at+asc`;
    const res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + getToken_() },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      Logger.log(`❌ Page ${page}: HTTP ${res.getResponseCode()}`);
      break;
    }
    const orders = JSON.parse(res.getContentText()).orders || [];
    if (orders.length === 0) break;
    totalFetched += orders.length;
    orders.forEach(o => {
      if (o.cancelled_at) return;
      if (['voided', 'refunded'].includes(o.financial_status)) return;
      if (existing.has(o.name)) { dupCount++; return; }
      newCount++;
      const parsed = parseNote_(o.note);
      const orderType = classifyOrderType_(o, parsed);
      Logger.log(`  [MỚI] ${o.name} | ${orderType} | ${formatDate_(o.created_at)} | ${o.total_price}đ | ${(o.shipping_address && o.shipping_address.name) || '?'}`);
    });
    if (orders.length < CONFIG.PAGE_SIZE) break;
    page++;
  }
  Logger.log(`\n✅ DRY RUN OK — Tổng fetch: ${totalFetched}, đơn MỚI sẽ append: ${newCount}, trùng: ${dupCount}`);
  Logger.log(`   Chạy syncOrders để ghi vào sheet thật.`);
}

function showProps() {
  const props = PropertiesService.getScriptProperties().getProperties();
  Object.keys(props).forEach(k => {
    const v = k.includes('TOKEN') ? props[k].slice(0, 8) + '...' + props[k].slice(-4) : props[k];
    Logger.log(`${k} = ${v}`);
  });
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncOrders') ScriptApp.deleteTrigger(t);
  });
  Logger.log('✅ Đã xóa trigger cũ');
}
