/**
 * PHASE 18 — Cảnh báo phí ship cao bất thường
 *
 * Cron daily 19:15 (sau Phase 17). Quét tab "tracking haravan", tính ratio phí ship/tổng đơn,
 * flag đơn có ratio > THRESHOLD (default 15%). Email top 10 đơn outlier.
 *
 * Mục đích: phát hiện đơn nhập sai DxRxC trong Scheme master → KL quy đổi cao → phí ship đội.
 * Khác Phase 7 (KL quy đổi cảnh báo theo SKU), Phase 18 cảnh báo theo ĐƠN HÀNG cụ thể.
 */

const FEE_ALERT_CFG = {
  THRESHOLD_RATIO: 0.15,    // 15%
  TOP_N: 10,
  ALERT_EMAIL: 'admin@khomes.com.vn',
  MIN_ORDER_VALUE: 200000,  // chỉ flag đơn > 200k (đơn nhỏ ratio dễ cao bất thường, ko quan trọng)
};

function runFeeAlertReport() {
  const ss = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(GHN_CFG.TARGET_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const data = sheet.getRange(2, 1, lastRow - 1, GHN_COL.GHN_AT).getValues();
  const today = new Date();
  const yest = new Date(today.getTime() - 24 * 3600 * 1000);

  const outliers = [];
  data.forEach(function (row) {
    const tongDon = Number(row[GHN_COL.TONG_DON - 1] || 0);
    const phi = Number(row[GHN_COL.GHN_PHI - 1] || 0);
    if (tongDon < FEE_ALERT_CFG.MIN_ORDER_VALUE) return;
    if (!phi) return;
    const ratio = phi / tongDon;
    if (ratio < FEE_ALERT_CFG.THRESHOLD_RATIO) return;
    // Chỉ tính đơn trong 7 ngày gần nhất
    const ngayDat = row[GHN_COL.NGAY_DAT - 1];
    let dayParsed = null;
    if (ngayDat instanceof Date) dayParsed = ngayDat;
    else if (ngayDat) {
      try { dayParsed = new Date(ngayDat); } catch (e) {}
    }
    if (dayParsed && (today.getTime() - dayParsed.getTime()) > 7 * 24 * 3600 * 1000) return;

    outliers.push({
      haravan: row[GHN_COL.MA_HARAVAN - 1],
      maVc: row[GHN_COL.GHN_MA - 1],
      sku: row[GHN_COL.MA_SP - 1],
      tenSp: String(row[GHN_COL.TEN_SP - 1] || '').slice(0, 60),
      tongDon, phi, ratio,
    });
  });

  if (!outliers.length) {
    Logger.log('No fee outliers in last 7 days');
    return { ok: true, count: 0 };
  }

  outliers.sort(function (a, b) { return b.ratio - a.ratio; });
  const top = outliers.slice(0, FEE_ALERT_CFG.TOP_N);

  const dateStr = Utilities.formatDate(today, 'GMT+7', 'dd/MM/yyyy');
  let html = '<h2 style="color:#c0392b">⚠️ Phí ship cao bất thường — ' + dateStr + '</h2>';
  html += '<p>Top ' + top.length + ' đơn có phí ship > <b>' + (FEE_ALERT_CFG.THRESHOLD_RATIO * 100) +
          '%</b> tổng đơn trong 7 ngày qua. Nguyên nhân thường gặp: <b>nhập sai DxRxC trong Scheme master</b> ' +
          '→ KL quy đổi GHN cao → phí ship đội. Xem tab <code>vol_warnings</code> để chỉnh DxRxC.</p>';
  html += '<table border="1" cellpadding="6" style="border-collapse:collapse;font-family:Arial">';
  html += '<tr style="background:#cfe2f3"><th>Mã đơn</th><th>Mã VC</th><th>SKU</th><th>SP</th><th>Tổng đơn</th><th>Phí ship</th><th>Ratio</th></tr>';
  top.forEach(function (o) {
    const pct = Math.round(o.ratio * 1000) / 10;
    const bg = pct > 25 ? '#f4cccc' : pct > 20 ? '#fce5cd' : '#fff2cc';
    html += '<tr style="background:' + bg + '">' +
            '<td>' + (o.haravan || '') + '</td>' +
            '<td>' + (o.maVc || '') + '</td>' +
            '<td>' + (o.sku || '') + '</td>' +
            '<td>' + o.tenSp + '</td>' +
            '<td>' + o.tongDon.toLocaleString('vi-VN') + 'đ</td>' +
            '<td>' + o.phi.toLocaleString('vi-VN') + 'đ</td>' +
            '<td><b>' + pct + '%</b></td></tr>';
  });
  html += '</table>';
  html += '<hr><p style="color:#888;font-size:12px">Phase 18 cron 19:15.</p>';

  MailApp.sendEmail({
    to: FEE_ALERT_CFG.ALERT_EMAIL,
    subject: '[K-Homes] ⚠️ ' + top.length + ' đơn phí ship cao bất thường — ' + dateStr,
    htmlBody: html,
  });

  Logger.log('✅ Fee alert email sent: ' + top.length + ' outliers');
  return { ok: true, count: top.length };
}

function setupFeeAlertTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runFeeAlertReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runFeeAlertReport').timeBased().atHour(19).nearMinute(15).everyDays(1).create();
  Logger.log('✅ Trigger runFeeAlertReport daily 19:15');
}
