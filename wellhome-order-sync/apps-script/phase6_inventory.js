/**
 * PHASE 6 — SYNC TỒN KHO HARAVAN  (paste vào CÙNG project Wellhome Order Sync — file thứ 6)
 * Pull tồn kho từ Haravan API → ghi tab "stock haravan" trong file 1Bn4C0Ud...
 * Cron 1 lần/ngày 6h sáng. Dùng để check trước khi Phase 2 tạo GHN.
 *
 * Nguồn: Haravan /admin/products.json (paginate 50/page) → mỗi product có variants[]
 *   - variant.sku, variant.inventory_quantity, variant.id
 *
 * Output tab "stock haravan" cột A-G:
 *   A. SKU | B. Tên SP | C. Variant title | D. Variant ID
 *   E. Tồn kho | F. Giá bán | G. Cập nhật lúc
 *
 * Public functions:
 *   setupStockHeader            — tạo tab "stock haravan" + header
 *   syncHaravanInventory        — pull tồn kho (cron daily 6h)
 *   setupStockTrigger           — cài cron 6h sáng mỗi ngày
 *   removeStockTrigger          — xóa cron
 *   getStockBySku("SKU123")     — debug lookup 1 SKU
 *   loadStockMap_()             — load Map<SKU, qty> cho Phase 2 dùng
 *   showStockReport             — count đơn thiếu hàng
 */

const STOCK_CFG = {
  TARGET_TAB: 'stock haravan',
  PROP_LAST_SYNC: 'STOCK_LAST_SYNC',
  PROP_STOCK_CACHE: 'STOCK_CACHE_JSON',     // cache toàn bộ stock vào Properties (key=SKU, value=qty)
  PAGE_SIZE: 50,
  MAX_PAGES: 50,                            // 50*50=2500 SP max
  CRON_HOUR: 6,
  LOW_STOCK_THRESHOLD: 0,                   // < threshold = thiếu hàng
};

const STOCK_HEADERS = [
  'SKU', 'Tên SP', 'Variant title', 'Variant ID', 'Tồn kho', 'Giá bán', 'Cập nhật lúc'
];

const STOCK_COL = {
  SKU: 1, TEN_SP: 2, VARIANT_TITLE: 3, VARIANT_ID: 4,
  QTY: 5, PRICE: 6, UPDATED_AT: 7,
};

// ============================================================
// SETUP
// ============================================================

function setupStockHeader() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  let sheet = ss.getSheetByName(STOCK_CFG.TARGET_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(STOCK_CFG.TARGET_TAB);
    Logger.log(`✅ Đã tạo tab mới "${STOCK_CFG.TARGET_TAB}"`);
  }
  sheet.getRange(1, 1, 1, STOCK_HEADERS.length).setValues([STOCK_HEADERS])
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  const widths = [120, 280, 120, 110, 90, 110, 140];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Conditional format: tồn kho <= 0 → đỏ
  const N = Math.max(sheet.getLastRow(), 1000);
  const colE = sheet.getRange(2, STOCK_COL.QTY, N - 1, 1);
  const rules = sheet.getConditionalFormatRules();
  const filtered = rules.filter(r => {
    const ranges = r.getRanges();
    return !ranges.some(rg => rg.getColumn() === STOCK_COL.QTY);
  });
  filtered.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setBackground('#F8CECC').setFontColor('#990000').setBold(true)
    .setRanges([colE]).build());
  filtered.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 5)
    .setBackground('#FFE6CC')
    .setRanges([colE]).build());
  sheet.setConditionalFormatRules(filtered);

  Logger.log(`✅ Setup header tab "${STOCK_CFG.TARGET_TAB}" OK`);
}

function setupStockTrigger() {
  removeStockTrigger();
  ScriptApp.newTrigger('syncHaravanInventory')
    .timeBased().atHour(STOCK_CFG.CRON_HOUR).everyDays(1).create();
  Logger.log(`✅ Cron syncHaravanInventory: mỗi ngày ${STOCK_CFG.CRON_HOUR}h sáng`);
}

function removeStockTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncHaravanInventory') ScriptApp.deleteTrigger(t);
  });
}

// ============================================================
// MAIN — pull Haravan inventory
// ============================================================

function syncHaravanInventory() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  let sheet = ss.getSheetByName(STOCK_CFG.TARGET_TAB);
  if (!sheet) { setupStockHeader(); sheet = ss.getSheetByName(STOCK_CFG.TARGET_TAB); }

  const token = getToken_();   // dùng token chung Phase 1 (CONFIG.PROP_TOKEN)
  if (!token) { Logger.log('❌ Chưa setup token. Chạy setupToken trước.'); return; }

  const rows = [];
  const stockMap = {};
  let totalProducts = 0;
  let totalVariants = 0;

  for (let page = 1; page <= STOCK_CFG.MAX_PAGES; page++) {
    const url = `https://${CONFIG.HARAVAN_SHOP}/admin/products.json` +
      `?limit=${STOCK_CFG.PAGE_SIZE}&page=${page}` +
      `&fields=id,title,variants`;
    const res = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code !== 200) {
      Logger.log(`❌ Page ${page} HTTP ${code}: ${res.getContentText().slice(0, 200)}`);
      break;
    }
    const products = JSON.parse(res.getContentText()).products || [];
    if (products.length === 0) break;
    totalProducts += products.length;

    products.forEach(p => {
      (p.variants || []).forEach(v => {
        const sku = String(v.sku || '').trim();
        if (!sku) return;
        const qty = Number(v.inventory_quantity || 0);
        rows.push([
          sku, p.title || '', v.title || '', String(v.id || ''),
          qty, Number(v.price || 0), formatStockNow_()
        ]);
        if (!stockMap[sku] || qty > stockMap[sku]) stockMap[sku] = qty;   // SKU trùng → max
        totalVariants++;
      });
    });
    Logger.log(`Page ${page}: ${products.length} SP`);
    if (products.length < STOCK_CFG.PAGE_SIZE) break;
    Utilities.sleep(200);   // tránh rate limit Haravan khi paginate nhiều page
  }

  // Ghi đè toàn bộ data (tránh stale rows)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, STOCK_HEADERS.length).clearContent();
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, STOCK_HEADERS.length).setValues(rows);
  }

  // Cache vào Properties (Phase 2 đọc nhanh, không cần mở Sheet mỗi lần)
  // Properties limit 9KB/value → split nếu quá lớn
  try {
    const json = JSON.stringify(stockMap);
    if (json.length < 9000) {
      PropertiesService.getScriptProperties().setProperty(STOCK_CFG.PROP_STOCK_CACHE, json);
    } else {
      Logger.log(`⚠️ Stock cache ${json.length} bytes vượt 9KB Properties → bỏ cache, Phase 2 sẽ đọc Sheet`);
      PropertiesService.getScriptProperties().deleteProperty(STOCK_CFG.PROP_STOCK_CACHE);
    }
  } catch (e) {
    Logger.log(`⚠️ Cache stock: ${e.message}`);
  }

  PropertiesService.getScriptProperties().setProperty(
    STOCK_CFG.PROP_LAST_SYNC, new Date().toISOString());

  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`✅ syncHaravanInventory done ${dur}s. Products: ${totalProducts}, Variants có SKU: ${totalVariants}`);
}

// ============================================================
// PUBLIC HELPERS
// ============================================================

/**
 * Load stock map (SKU → qty). Ưu tiên Properties cache, fallback đọc Sheet.
 * Phase 2 createGhnOrders gọi 1 lần đầu loop.
 */
function loadStockMap_() {
  const cached = PropertiesService.getScriptProperties().getProperty(STOCK_CFG.PROP_STOCK_CACHE);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through */ }
  }
  // Fallback đọc Sheet
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(STOCK_CFG.TARGET_TAB);
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, STOCK_COL.QTY).getValues();
  const map = {};
  data.forEach(r => {
    const sku = String(r[STOCK_COL.SKU - 1] || '').trim();
    if (!sku) return;
    const qty = Number(r[STOCK_COL.QTY - 1] || 0);
    if (!map[sku] || qty > map[sku]) map[sku] = qty;
  });
  return map;
}

function getStockBySku(sku) {
  if (!sku) { Logger.log('Truyền vào SKU. Vd: getStockBySku("SKU123")'); return; }
  const map = loadStockMap_();
  Logger.log(`SKU ${sku}: ${map[sku] !== undefined ? map[sku] + ' tồn' : 'không có trong stock haravan'}`);
}

function showStockReport() {
  const map = loadStockMap_();
  const skus = Object.keys(map);
  let zero = 0, low = 0, ok = 0;
  skus.forEach(s => {
    const q = map[s];
    if (q <= 0) zero++;
    else if (q <= 5) low++;
    else ok++;
  });
  const last = PropertiesService.getScriptProperties().getProperty(STOCK_CFG.PROP_LAST_SYNC);
  Logger.log(`===== Stock report =====`);
  Logger.log(`Tổng SKU: ${skus.length}`);
  Logger.log(`  Hết hàng (≤0): ${zero}`);
  Logger.log(`  Sắp hết (1-5): ${low}`);
  Logger.log(`  OK (>5):       ${ok}`);
  Logger.log(`Last sync: ${last || '(chưa chạy)'}`);
}

function formatStockNow_() {
  return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
}
