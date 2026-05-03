/**
 * PHASE 8 — AUTO-RETRY đơn LỖI GHN có địa chỉ đầy đủ  (file thứ 9)
 *
 * Cron daily 9h sáng:
 *   - Quét đơn cột X = "Lỗi GHN" tồn ≥ 6h
 *   - Check địa chỉ giao có đầy đủ Phường/Quận/Tỉnh (cột J/K/L)
 *   - Nếu đầy đủ → reset cột X = "Đã xác nhận" để Phase 2 cron 1h sau retry
 *   - Ghi log retry vào cột Y "Auto-retry <lần thứ N> @<timestamp>"
 *   - Limit MAX 3 lần retry/đơn (tránh loop vô hạn)
 *
 * Public:
 *   autoRetryFailedGhnOrders     — hàm chính (cron 9h)
 *   manualAutoRetry              — debug
 *   setupAutoRetryTrigger        — cài cron
 *   removeAutoRetryTrigger
 */

const AUTO_RETRY_CFG = {
  CRON_HOUR: 9,
  MIN_AGE_HOURS: 6,             // chỉ retry đơn lỗi tồn ≥ 6h (PG có thời gian sửa)
  MAX_RETRIES: 3,               // mỗi đơn retry max 3 lần
  PROP_LAST_RUN: 'AUTO_RETRY_LAST_RUN',
  ALERT_EMAIL: 'admin@khomes.com.vn',
};

function autoRetryFailedGhnOrders() {
  const t0 = Date.now();
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng, skip'); return; }

  const lastCol = Math.max(33, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const retried = [];
  const skipReasons = { tooNew: 0, missingAddress: 0, tooManyRetries: 0, wrongStatus: 0 };
  const now = formatGhnNow_();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const status = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    if (status !== 'Lỗi GHN') { skipReasons.wrongStatus++; continue; }

    const ngayDat = parseDate_(row[GHN_COL.NGAY_DAT - 1]);
    const ageHrs = ngayDat ? (Date.now() - ngayDat.getTime()) / 3600000 : 0;
    if (ageHrs < AUTO_RETRY_CFG.MIN_AGE_HOURS) { skipReasons.tooNew++; continue; }

    const phuong = String(row[GHN_COL.PHUONG - 1] || '').trim();
    const quan = String(row[GHN_COL.QUAN - 1] || '').trim();
    const tinh = String(row[GHN_COL.TINH - 1] || '').trim();
    if (!phuong || !quan || !tinh) { skipReasons.missingAddress++; continue; }

    // Đếm số lần retry trước đó từ cột Y note
    const note = String(row[GHN_COL.GHI_CHU_NV - 1] || '');
    const retryMatches = note.match(/Auto-retry (\d+)/g) || [];
    if (retryMatches.length >= AUTO_RETRY_CFG.MAX_RETRIES) {
      skipReasons.tooManyRetries++; continue;
    }
    const nextRetryNo = retryMatches.length + 1;

    const orderId = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim();
    const rowIdx = i + 2;
    try {
      sheet.getRange(rowIdx, GHN_COL.TT_XU_LY).setValue('Đã xác nhận');
      const newNote = note + (note ? ' | ' : '') + `Auto-retry ${nextRetryNo} @${now}`;
      sheet.getRange(rowIdx, GHN_COL.GHI_CHU_NV).setValue(newNote);
      retried.push({ orderId, retryNo: nextRetryNo, tenKH: row[GHN_COL.TEN_KH - 1] });
    } catch (e) {
      Logger.log(`⚠️ Retry row ${rowIdx} fail: ${e.message}`);
    }
  }

  PropertiesService.getScriptProperties().setProperty(
    AUTO_RETRY_CFG.PROP_LAST_RUN, new Date().toISOString());

  if (retried.length > 0) {
    try { sendAutoRetryEmail_(retried); }
    catch (e) { Logger.log(`⚠️ sendAutoRetryEmail: ${e.message}`); }
  }

  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`✅ autoRetryFailedGhnOrders ${dur}s. Retried: ${retried.length}. Skip: ${JSON.stringify(skipReasons)}`);
}

function sendAutoRetryEmail_(retried) {
  const subject = `🔄 [K-Homes] Auto-retry ${retried.length} đơn lỗi GHN`;
  const rowsHtml = retried.map((r, i) => `
    <tr><td>${i + 1}</td><td><b>${r.orderId}</b></td><td>${r.tenKH || ''}</td>
        <td style="text-align:center;">${r.retryNo}</td></tr>`).join('');
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h2 style="color:#1f4e78;">🔄 Auto-retry đơn lỗi GHN — ${retried.length} đơn</h2>
      <p>Các đơn dưới đây đã có địa chỉ đầy đủ (cột J/K/L) và đã tồn lỗi GHN ≥ ${AUTO_RETRY_CFG.MIN_AGE_HOURS}h.
      Tool tự reset cột X = "Đã xác nhận" → cron Phase 2 (1h/lần) sẽ tự thử tạo GHN lần nữa.</p>
      <p>Nếu lần retry này cũng fail (max ${AUTO_RETRY_CFG.MAX_RETRIES} lần), PG cần fallback sang Fbox/FFM tay.</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
        <tr style="background:#1f4e78;color:#fff;">
          <th>#</th><th>Mã đơn</th><th>Tên KH</th><th>Lần retry</th>
        </tr>
        ${rowsHtml}
      </table>
      <p style="margin-top:15px;color:#666;font-size:12px;">
        ➤ Email tự động từ Phase 8 mỗi sáng 9h. Sau retry, kết quả sẽ được Phase 2 (10h) ghi vào cột Z + email alert nếu vẫn lỗi.
      </p>
    </div>`;
  MailApp.sendEmail({ to: AUTO_RETRY_CFG.ALERT_EMAIL, subject, htmlBody: html });
}

function manualAutoRetry() {
  Logger.log('🟡 Manual autoRetryFailedGhnOrders');
  autoRetryFailedGhnOrders();
}

function setupAutoRetryTrigger() {
  removeAutoRetryTrigger();
  ScriptApp.newTrigger('autoRetryFailedGhnOrders')
    .timeBased().atHour(AUTO_RETRY_CFG.CRON_HOUR).everyDays(1).create();
  Logger.log(`✅ Cron autoRetryFailedGhnOrders mỗi ngày ${AUTO_RETRY_CFG.CRON_HOUR}h sáng`);
}

function removeAutoRetryTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoRetryFailedGhnOrders') ScriptApp.deleteTrigger(t);
  });
}
