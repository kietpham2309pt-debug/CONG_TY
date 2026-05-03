/**
 * PHASE 20 — KOL Performance aggregation
 *
 * Cron daily 19:30 (sau Phase 17). Quét tab "tracking haravan" — group by Camp + Brand,
 * tính KPI: SoDon, GMV, FailRate, AvgGmv, RoiPlaceholder. Ghi tab "kol_performance".
 *
 * Tab kol_performance schema:
 *   A Camp | B Brand | C SoDon | D GMV | E AvgGmv | F FailedOrders | G FailRate%
 *   H DeliveredOrders | I PendingOrders | J UpdatedAt | K Phí KOL (PG fill tay) | L ROI%
 *
 * Phí KOL: chưa có data → để trống cho PG fill tay. Khi K (Phí KOL) có giá trị,
 * công thức cột L tự tính ROI = (GMV - Phí KOL) / Phí KOL * 100.
 */

const KOL_CFG = {
  TAB: 'kol_performance',
  HEADERS: ['Camp', 'Brand', 'SoDon', 'GMV', 'AvgGmv', 'FailedOrders', 'FailRate%',
            'Delivered', 'Pending', 'UpdatedAt', 'PhiKOL', 'ROI%'],
};

function aggregateKolPerformance() {
  const ss = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID);
  const src = ss.getSheetByName(GHN_CFG.TARGET_TAB);
  const lastRow = src.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const data = src.getRange(2, 1, lastRow - 1, GHN_COL.GHN_AT).getValues();
  const groups = {};

  data.forEach(function (row) {
    const loaiDon = String(row[GHN_COL.LOAI_DON - 1] || '').trim();
    if (loaiDon !== 'KOL') return;
    const camp = String(row[GHN_COL.CAMP - 1] || 'Unknown').trim();
    const brand = detectBrand_(row[GHN_COL.TEN_SP - 1] || '');
    const key = camp + '|' + brand;
    if (!groups[key]) {
      groups[key] = { camp, brand, soDon: 0, gmv: 0, failed: 0, delivered: 0, pending: 0 };
    }
    const g = groups[key];
    g.soDon++;
    g.gmv += Number(row[GHN_COL.TONG_DON - 1] || 0);
    const ttXuLy = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    const ttGiao = String(row[GHN_COL.TT_GIAO - 1] || '').trim();
    if (ttXuLy === 'Lỗi GHN' || ttXuLy === 'Lỗi data' || ttXuLy === 'Hủy') g.failed++;
    if (ttGiao === 'Đã giao') g.delivered++;
    else if (ttXuLy === 'Đã xác nhận' || ttXuLy === 'Đã tạo GHN') g.pending++;
  });

  let tab = ss.getSheetByName(KOL_CFG.TAB);
  if (!tab) {
    tab = ss.insertSheet(KOL_CFG.TAB);
    tab.getRange(1, 1, 1, KOL_CFG.HEADERS.length).setValues([KOL_CFG.HEADERS])
      .setFontWeight('bold').setBackground('#fce5cd');
    tab.setFrozenRows(1);
  } else {
    if (tab.getLastRow() > 1) tab.getRange(2, 1, tab.getLastRow() - 1, KOL_CFG.HEADERS.length).clearContent();
  }

  const now = new Date();
  const rows = [];
  Object.keys(groups).forEach(function (k) {
    const g = groups[k];
    const avg = g.soDon > 0 ? Math.round(g.gmv / g.soDon) : 0;
    const failRate = g.soDon > 0 ? Math.round((g.failed / g.soDon) * 1000) / 10 : 0;
    rows.push([g.camp, g.brand, g.soDon, g.gmv, avg, g.failed, failRate, g.delivered, g.pending, now, '', '']);
  });

  if (!rows.length) return { ok: true, count: 0 };
  rows.sort(function (a, b) { return b[3] - a[3] });  // sort GMV desc
  tab.getRange(2, 1, rows.length, KOL_CFG.HEADERS.length).setValues(rows);

  // ROI formula trong cột L (chỉ tính khi cột K có giá trị)
  for (let i = 0; i < rows.length; i++) {
    const r = i + 2;
    tab.getRange(r, 12).setFormula('=IF(K' + r + '>0, ROUND((D' + r + '-K' + r + ')/K' + r + '*100, 1), "")');
  }

  Logger.log('✅ KOL aggregate: ' + rows.length + ' camps');
  return { ok: true, count: rows.length };
}

function detectBrand_(productName) {
  const n = String(productName).toLowerCase();
  if (n.indexOf('bosch') >= 0) return 'Bosch';
  if (n.indexOf('tefal') >= 0) return 'Tefal';
  if (n.indexOf('finish') >= 0) return 'Finish';
  if (n.indexOf('moulinex') >= 0) return 'Moulinex';
  if (n.indexOf('rowenta') >= 0) return 'Rowenta';
  return 'Other';
}

function setupKolTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'aggregateKolPerformance') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('aggregateKolPerformance').timeBased().atHour(19).nearMinute(30).everyDays(1).create();
  Logger.log('✅ Trigger aggregateKolPerformance daily 19:30');
}
