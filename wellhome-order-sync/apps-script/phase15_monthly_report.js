/**
 * PHASE 15 — MONTHLY REPORT (file thứ 13)
 *
 * Cron daily 9h sáng nhưng chỉ chạy thực sự khi today.getDate() === 1 (ngày đầu tháng).
 * Lý do: Apps Script triggers không có nthDayOfMonth, workaround daily + check ngày.
 *
 * Đọc tab "đối soát" tháng trước, tổng hợp KPI, gửi email HTML chi tiết.
 *
 * Public:
 *   sendMonthlyReport             — hàm chính (cron daily 9h, skip nếu không phải ngày 1)
 *   forceSendMonthlyReport        — debug — chạy không check ngày
 *   setupMonthlyReportTrigger
 *   removeMonthlyReportTrigger
 */

const MONTHLY_CFG = {
  ALERT_EMAIL: 'admin@khomes.com.vn',
  PROP_LAST_RUN: 'MONTHLY_REPORT_LAST_RUN',
  CRON_HOUR: 9,
};

function sendMonthlyReport() {
  const today = new Date();
  if (today.getDate() !== 1) {
    Logger.log(`Skip — chỉ gửi report ngày 1 mỗi tháng (hôm nay là ngày ${today.getDate()})`);
    return;
  }
  return _sendMonthlyReportInternal_();
}

function forceSendMonthlyReport() {
  Logger.log('🟡 Force sendMonthlyReport (bypass date check)');
  return _sendMonthlyReportInternal_();
}

function _sendMonthlyReportInternal_() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(RECON_CFG.TARGET_TAB);
  if (!sheet) { Logger.log('Tab "đối soát" chưa tạo'); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Tab "đối soát" rỗng'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, RECON_HEADERS.length).getValues();

  // Lấy tháng trước (1st → end of month)
  const today = new Date();
  const firstOfThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const firstOfLastMonth = new Date(firstOfThisMonth.getFullYear(), firstOfThisMonth.getMonth() - 1, 1);
  const lastOfLastMonth = new Date(firstOfThisMonth.getTime() - 86400000);   // ngày cuối tháng trước

  const monthData = [];
  data.forEach(row => {
    const dateStr = String(row[0] || '');
    const m = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!m) return;
    const d = new Date(+m[3], +m[2] - 1, +m[1]);
    if (d >= firstOfLastMonth && d <= lastOfLastMonth) {
      monthData.push({ date: dateStr, raw: row, dateObj: d });
    }
  });

  if (monthData.length === 0) {
    Logger.log(`Không có data tháng ${formatMonth_(firstOfLastMonth)}`);
    return;
  }

  // Tổng hợp
  const totals = aggregateRecon_(monthData);

  // Tuần phân tích — chia 4 tuần (mỗi 7 ngày từ ngày 1)
  const weeks = [
    { label: 'Tuần 1 (1-7)', filter: d => d.getDate() <= 7 },
    { label: 'Tuần 2 (8-14)', filter: d => d.getDate() >= 8 && d.getDate() <= 14 },
    { label: 'Tuần 3 (15-21)', filter: d => d.getDate() >= 15 && d.getDate() <= 21 },
    { label: 'Tuần 4 (22+)', filter: d => d.getDate() >= 22 },
  ];
  const weekStats = weeks.map(w => ({
    label: w.label,
    data: monthData.filter(d => w.filter(d.dateObj)),
  })).map(w => ({
    label: w.label,
    nDays: w.data.length,
    totals: aggregateRecon_(w.data),
  }));

  const monthLabel = formatMonth_(firstOfLastMonth);
  const subject = `📅 [K-Homes] Báo cáo tháng ${monthLabel} — ${totals.donMoi} đơn, GMV ${fmt_(totals.gmv)}đ`;

  const fmt = fmt_;
  const pct = pct_;

  const weekRowsHtml = weekStats.map(w => `
    <tr>
      <td>${w.label}</td>
      <td>${w.nDays} ngày</td>
      <td>${w.totals.donMoi}</td>
      <td>${fmt(w.totals.gmv)}đ</td>
      <td style="color:#c0504d;">${pct(w.totals.ghnLoi, w.totals.donMoi)}</td>
      <td>${pct(w.totals.daGiao, w.totals.donMoi)}</td>
    </tr>`).join('');

  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h1 style="color:#1f4e78;">📅 Báo cáo tháng ${monthLabel}</h1>
      <p style="color:#666;">${formatDateOnly_(firstOfLastMonth)} → ${formatDateOnly_(lastOfLastMonth)} (${monthData.length}/${lastOfLastMonth.getDate()} ngày có data)</p>

      <h2 style="color:#1f4e78;margin-top:25px;">Tổng quan tháng</h2>
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:14px;">
        <tr><td><b>📦 Đơn mới</b></td><td><b>${totals.donMoi}</b></td><td>(TB ${(totals.donMoi/Math.max(monthData.length,1)).toFixed(1)}/ngày)</td></tr>
        <tr><td>💰 Tổng GMV</td><td><b>${fmt(totals.gmv)}đ</b></td><td>(TB ${fmt(Math.round(totals.gmv/Math.max(monthData.length,1)))}đ/ngày)</td></tr>
        <tr><td>🚚 Phí GHN</td><td>${fmt(totals.phiGhn)}đ</td><td>${pct(totals.phiGhn, totals.gmv)} GMV</td></tr>
        <tr><td>✅ Tỷ lệ tạo GHN OK</td><td>${pct(totals.ghnOK, totals.donMoi)}</td><td>(${totals.ghnOK}/${totals.donMoi})</td></tr>
        <tr><td style="color:#c0504d;">⚠️ Tỷ lệ lỗi GHN</td><td>${pct(totals.ghnLoi, totals.donMoi)}</td><td>(${totals.ghnLoi}/${totals.donMoi})</td></tr>
        <tr><td>📤 Fulfill rate</td><td>${pct(totals.fulfilled, totals.donMoi)}</td><td>(${totals.fulfilled}/${totals.donMoi})</td></tr>
        <tr><td>✅ Tỷ lệ giao thành công</td><td>${pct(totals.daGiao, totals.donMoi)}</td><td>(${totals.daGiao}/${totals.donMoi})</td></tr>
      </table>

      <h2 style="color:#1f4e78;margin-top:25px;">Phân loại</h2>
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:14px;">
        <tr><td>KOL</td><td>${totals.kol}</td><td>${pct(totals.kol, totals.donMoi)}</td></tr>
        <tr><td>Organic</td><td>${totals.organic}</td><td>${pct(totals.organic, totals.donMoi)}</td></tr>
        <tr><td>Nội bộ</td><td>${totals.noiBo}</td><td>${pct(totals.noiBo, totals.donMoi)}</td></tr>
        <tr><td>Khác</td><td>${totals.khac}</td><td>${pct(totals.khac, totals.donMoi)}</td></tr>
      </table>

      <h2 style="color:#1f4e78;margin-top:25px;">Phân tích theo tuần</h2>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;width:100%;">
        <tr style="background:#1f4e78;color:#fff;">
          <th>Tuần</th><th>Ngày có data</th><th>Đơn mới</th><th>GMV</th><th>Lỗi GHN %</th><th>Đã giao %</th>
        </tr>
        ${weekRowsHtml}
      </table>

      <p style="margin-top:25px;color:#666;font-size:12px;">
        ➤ Email tự động từ Phase 15 ngày 1 đầu mỗi tháng (cron daily 9h, check date=1).<br>
        ➤ Data nguồn: tab "đối soát" 1Bn4C0Ud.
      </p>
    </div>`;

  MailApp.sendEmail({ to: MONTHLY_CFG.ALERT_EMAIL, subject, htmlBody: html });
  PropertiesService.getScriptProperties().setProperty(
    MONTHLY_CFG.PROP_LAST_RUN, new Date().toISOString());

  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`✅ sendMonthlyReport ${monthLabel} done ${dur}s. ${monthData.length} ngày, GMV ${fmt(totals.gmv)}đ`);
}

function aggregateRecon_(monthData) {
  const totals = { donMoi: 0, kol: 0, organic: 0, noiBo: 0, khac: 0,
    ghnOK: 0, ghnLoi: 0, pickup: 0, thieuHang: 0,
    daGiao: 0, dangGiao: 0, huy: 0, pending: 0,
    fulfilled: 0, gmv: 0, phiGhn: 0 };
  monthData.forEach(d => {
    const r = d.raw;
    totals.donMoi += Number(r[1] || 0);
    totals.kol += Number(r[2] || 0);
    totals.organic += Number(r[3] || 0);
    totals.noiBo += Number(r[4] || 0);
    totals.khac += Number(r[5] || 0);
    totals.ghnOK += Number(r[6] || 0);
    totals.ghnLoi += Number(r[7] || 0);
    totals.pickup += Number(r[8] || 0);
    totals.thieuHang += Number(r[9] || 0);
    totals.daGiao += Number(r[10] || 0);
    totals.dangGiao += Number(r[11] || 0);
    totals.huy += Number(r[12] || 0);
    totals.pending += Number(r[13] || 0);
    totals.fulfilled += Number(r[14] || 0);
    totals.gmv += Number(r[15] || 0);
    totals.phiGhn += Number(r[16] || 0);
  });
  return totals;
}

function fmt_(n) { return Number(n).toLocaleString('vi-VN'); }
function pct_(n, d) { return d > 0 ? Math.round(n / d * 100) + '%' : '0%'; }
function formatMonth_(d) {
  return Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'MM/yyyy');
}

function setupMonthlyReportTrigger() {
  removeMonthlyReportTrigger();
  ScriptApp.newTrigger('sendMonthlyReport')
    .timeBased().atHour(MONTHLY_CFG.CRON_HOUR).everyDays(1).create();
  Logger.log(`✅ Cron sendMonthlyReport daily ${MONTHLY_CFG.CRON_HOUR}h (chỉ gửi khi date=1)`);
}

function removeMonthlyReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendMonthlyReport') ScriptApp.deleteTrigger(t);
  });
}
