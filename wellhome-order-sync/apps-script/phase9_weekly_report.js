/**
 * PHASE 9 — WEEKLY REPORT (file thứ 10)
 *
 * Cron Thứ 2 sáng 9h:
 *   - Đọc tab "đối soát" 7 ngày qua (Mon-Sun của tuần trước)
 *   - Tổng hợp số liệu tuần: GMV, đơn mới, fail rate, fulfill rate, trend
 *   - Email HTML cho admin@khomes.com.vn
 *
 * Public:
 *   sendWeeklyReport             — hàm chính (cron Thứ 2 9h)
 *   manualSendWeeklyReport       — debug (gửi báo cáo tuần này nếu chạy giữa tuần)
 *   setupWeeklyReportTrigger     — cài cron weekly Mon 9h
 *   removeWeeklyReportTrigger
 */

const WEEKLY_CFG = {
  ALERT_EMAIL: 'admin@khomes.com.vn',
  PROP_LAST_RUN: 'WEEKLY_REPORT_LAST_RUN',
};

function sendWeeklyReport() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(RECON_CFG.TARGET_TAB);
  if (!sheet) { Logger.log('Tab "đối soát" chưa tạo'); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Tab "đối soát" rỗng'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, RECON_HEADERS.length).getValues();

  // Lấy tuần trước (Mon → Sun)
  const today = new Date();
  const dayOfWeek = today.getDay();   // 0=Sun, 1=Mon, ..., 6=Sat
  const lastSunday = new Date(today);
  lastSunday.setDate(today.getDate() - (dayOfWeek === 0 ? 7 : dayOfWeek));
  lastSunday.setHours(23, 59, 59, 999);
  const lastMonday = new Date(lastSunday);
  lastMonday.setDate(lastSunday.getDate() - 6);
  lastMonday.setHours(0, 0, 0, 0);

  const weekData = [];
  data.forEach(row => {
    const dateStr = String(row[0] || '');
    const m = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!m) return;
    const d = new Date(+m[3], +m[2] - 1, +m[1]);
    if (d >= lastMonday && d <= lastSunday) {
      weekData.push({ date: dateStr, raw: row });
    }
  });

  if (weekData.length === 0) {
    Logger.log(`Không có data tuần ${formatDateOnly_(lastMonday)} - ${formatDateOnly_(lastSunday)}`);
    return;
  }

  // Tổng hợp
  const totals = { donMoi: 0, kol: 0, organic: 0, noiBo: 0, khac: 0,
    ghnOK: 0, ghnLoi: 0, pickup: 0, thieuHang: 0,
    daGiao: 0, dangGiao: 0, huy: 0, pending: 0,
    fulfilled: 0, gmv: 0, phiGhn: 0, volWarnings: 0 };
  weekData.forEach(d => {
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
    totals.volWarnings = Math.max(totals.volWarnings, Number(r[17] || 0));   // max not sum
  });

  const fmt = (n) => Number(n).toLocaleString('vi-VN');
  const pct = (n, d) => d > 0 ? Math.round(n / d * 100) + '%' : '0%';

  const dailyRowsHtml = weekData.map(d => {
    const r = d.raw;
    return `<tr>
      <td>${d.date}</td>
      <td style="text-align:center;">${r[1]}</td>
      <td style="text-align:center;">${r[6]}</td>
      <td style="text-align:center;color:#c0504d;">${r[7]}</td>
      <td style="text-align:center;">${r[10]}</td>
      <td style="text-align:right;">${fmt(r[15])}đ</td>
    </tr>`;
  }).join('');

  const subject = `📈 [K-Homes] Báo cáo tuần ${formatDateOnly_(lastMonday)} - ${formatDateOnly_(lastSunday)}`;
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h1 style="color:#1f4e78;">📈 Báo cáo tuần</h1>
      <p style="color:#666;">${formatDateOnly_(lastMonday)} → ${formatDateOnly_(lastSunday)} (${weekData.length}/7 ngày có data)</p>

      <h2 style="color:#1f4e78;margin-top:25px;">Tổng quan</h2>
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:14px;">
        <tr><td><b>📦 Đơn mới</b></td><td><b>${totals.donMoi}</b></td><td>(TB ${(totals.donMoi/weekData.length).toFixed(1)}/ngày)</td></tr>
        <tr><td>💰 Tổng GMV</td><td><b>${fmt(totals.gmv)}đ</b></td><td>(TB ${fmt(Math.round(totals.gmv/weekData.length))}đ/ngày)</td></tr>
        <tr><td>🚚 Phí GHN</td><td>${fmt(totals.phiGhn)}đ</td><td>${pct(totals.phiGhn, totals.gmv)} GMV</td></tr>
        <tr><td>✅ Tỷ lệ tạo GHN OK</td><td>${pct(totals.ghnOK, totals.donMoi)}</td><td>(${totals.ghnOK}/${totals.donMoi})</td></tr>
        <tr><td style="color:#c0504d;">⚠️ Tỷ lệ lỗi GHN</td><td>${pct(totals.ghnLoi, totals.donMoi)}</td><td>(${totals.ghnLoi}/${totals.donMoi})</td></tr>
        <tr><td>📤 Tỷ lệ Fulfill Haravan</td><td>${pct(totals.fulfilled, totals.donMoi)}</td><td>(${totals.fulfilled}/${totals.donMoi})</td></tr>
        <tr><td>✅ Đã giao thành công</td><td>${pct(totals.daGiao, totals.donMoi)}</td><td>(${totals.daGiao}/${totals.donMoi})</td></tr>
      </table>

      <h2 style="color:#1f4e78;margin-top:25px;">Phân loại đơn</h2>
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:14px;">
        <tr><td>KOL</td><td>${totals.kol}</td><td>${pct(totals.kol, totals.donMoi)}</td></tr>
        <tr><td>Organic</td><td>${totals.organic}</td><td>${pct(totals.organic, totals.donMoi)}</td></tr>
        <tr><td>Nội bộ</td><td>${totals.noiBo}</td><td>${pct(totals.noiBo, totals.donMoi)}</td></tr>
        <tr><td>Khác</td><td>${totals.khac}</td><td>${pct(totals.khac, totals.donMoi)}</td></tr>
      </table>

      <h2 style="color:#1f4e78;margin-top:25px;">Chi tiết từng ngày</h2>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;width:100%;">
        <tr style="background:#1f4e78;color:#fff;">
          <th>Ngày</th><th>Đơn mới</th><th>GHN OK</th><th>Lỗi GHN</th><th>Đã giao</th><th>GMV</th>
        </tr>
        ${dailyRowsHtml}
      </table>

      ${totals.volWarnings > 0 ? `
      <p style="margin-top:20px;padding:12px;background:#fff8e1;border-left:4px solid #e58a00;">
        ⚖️ Tuần qua có <b>${totals.volWarnings}</b> SKU có KL quy đổi cao bất thường. Xem email "Báo cáo cuối ngày" hằng ngày để có gợi ý điều chỉnh DxRxC.
      </p>` : ''}

      <p style="margin-top:25px;color:#666;font-size:12px;">
        ➤ Email tự động từ Phase 9 Wellhome Order Sync gửi mỗi Thứ 2 sáng 9h.<br>
        ➤ Data nguồn: tab "đối soát" trong Sheet K-Homès Bosch.
      </p>
    </div>`;

  MailApp.sendEmail({ to: WEEKLY_CFG.ALERT_EMAIL, subject, htmlBody: html });
  PropertiesService.getScriptProperties().setProperty(
    WEEKLY_CFG.PROP_LAST_RUN, new Date().toISOString());

  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`✅ sendWeeklyReport ${dur}s. ${weekData.length} ngày data, GMV ${fmt(totals.gmv)}đ`);
}

function manualSendWeeklyReport() {
  Logger.log('🟡 Manual sendWeeklyReport');
  sendWeeklyReport();
}

function setupWeeklyReportTrigger() {
  removeWeeklyReportTrigger();
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  Logger.log('✅ Cron sendWeeklyReport mỗi Thứ 2 9h sáng');
}

function removeWeeklyReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t);
  });
}
