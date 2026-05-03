/**
 * PHASE 2.5 — DAILY FAILED ORDERS SUMMARY (paste vào CÙNG project Wellhome Order Sync — file thứ 7)
 *
 * Theo nhân viên video Tefal: "đơn lỗi GHN → bắt buộc làm thủ công bằng FFM (Fbox)".
 * Vì K-Homes Wellhome chưa có Fbox API, hàm này gửi email tổng hợp mỗi sáng 8h:
 *   - Đơn cột X = "Lỗi GHN" còn pending (không phải mới fail trong 1h trước, vì email Phase 2 đã alert)
 *   - Đơn cột X = "Đã tạo GHN" + cột AE (Số lần giao fail GHN) >= 3
 *   - Đơn cột X = "Thiếu hàng" còn tồn > 24h
 *
 * Email gửi PG để PG quyết định:
 *   - Tạo đơn FFM tay trên https://asia.qianyierp.com/
 *   - Hoặc fix đơn (sửa địa chỉ) + retry GHN
 *
 * Public:
 *   sendDailyGhnFailedSummary    — hàm chính (cron daily 8h)
 *   manualSendDailySummary       — gọi tay debug
 *   setupDailySummaryTrigger     — cài cron 8h sáng
 *   removeDailySummaryTrigger    — xóa
 */

const DAILY_SUMMARY_CFG = {
  ALERT_EMAIL: 'admin@khomes.com.vn',
  CRON_HOUR: 8,
  GHN_FAIL_THRESHOLD: 3,         // đếm AE (số lần giao fail) ≥ 3 → đề xuất FFM
  STUCK_HOURS: 24,               // "Thiếu hàng" / "Lỗi GHN" tồn quá X giờ → đưa vào summary
  PROP_LAST_RUN: 'DAILY_SUMMARY_LAST_RUN',
};

function sendDailyGhnFailedSummary() {
  const t0 = Date.now();
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng, skip'); return; }

  const lastCol = Math.max(33, sheet.getLastColumn());   // tới cột AG
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const groups = { ghnFailed: [], stuckGhn: [], stuckStock: [], deliveryFailHigh: [], duplicates: [] };
  const seenOrders = new Set();
  const bySdt24h = {};   // SĐT → list orders trong 24h gần đây — phát hiện duplicate
  const oneDayMs = 86400000;

  data.forEach((row, idx) => {
    const orderId = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim();
    if (!orderId || seenOrders.has(orderId)) return;
    seenOrders.add(orderId);

    const status = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    const ngayDat = row[GHN_COL.NGAY_DAT - 1];
    const ageHrs = computeAgeHours_(ngayDat);

    // Detect đơn duplicate: cùng SĐT, đơn trong 24h gần đây
    const sdt = String(row[GHN_COL.SDT - 1] || '').trim().replace(/[\s.\-+]/g, '').replace(/^84/, '0');
    if (sdt && ageHrs <= 24) {
      if (!bySdt24h[sdt]) bySdt24h[sdt] = [];
      bySdt24h[sdt].push({
        orderId, ageHrs,
        tenKH: row[GHN_COL.TEN_KH - 1] || '',
        tongDon: Number(row[GHN_COL.TONG_DON - 1] || 0),
        tinh: row[GHN_COL.TINH - 1] || '',
        status,
      });
    }

    const baseInfo = {
      orderId,
      tenKH: row[GHN_COL.TEN_KH - 1] || '',
      sdt: row[GHN_COL.SDT - 1] || '',
      tinh: row[GHN_COL.TINH - 1] || '',
      diaChi: row[GHN_COL.DIA_CHI - 1] || '',
      tongDon: Number(row[GHN_COL.TONG_DON - 1] || 0),
      tenSP: row[GHN_COL.TEN_SP - 1] || '',
      camp: row[GHN_COL.CAMP - 1] || '',
      ghiChu: row[GHN_COL.GHI_CHU_NV - 1] || '',
      ageHrs,
    };

    if (status === 'Lỗi GHN') {
      if (ageHrs >= DAILY_SUMMARY_CFG.STUCK_HOURS) {
        groups.stuckGhn.push(baseInfo);
      } else {
        groups.ghnFailed.push(baseInfo);
      }
    } else if (status === 'Thiếu hàng' && ageHrs >= DAILY_SUMMARY_CFG.STUCK_HOURS) {
      groups.stuckStock.push(baseInfo);
    } else if (status === 'Đã tạo GHN') {
      // Đơn đã tạo GHN nhưng giao fail >= 3 lần → đề xuất FFM hoặc liên hệ KH
      const failCount = Number(row[30] || 0);   // cột AE = 31, idx 30
      if (failCount >= DAILY_SUMMARY_CFG.GHN_FAIL_THRESHOLD) {
        groups.deliveryFailHigh.push({
          ...baseInfo,
          failCount,
          ghnCode: row[GHN_COL.GHN_MA - 1] || '',
        });
      }
    }
  });

  // Phase 14 — group đơn duplicate: SĐT có ≥ 2 đơn trong 24h
  Object.keys(bySdt24h).forEach(sdt => {
    const orders = bySdt24h[sdt];
    if (orders.length >= 2) {
      groups.duplicates.push({ sdt, orders });
    }
  });

  const total = groups.ghnFailed.length + groups.stuckGhn.length +
                groups.stuckStock.length + groups.deliveryFailHigh.length +
                groups.duplicates.length;

  if (total === 0) {
    Logger.log(`✅ Không có đơn nào cần xử lý sáng nay`);
    PropertiesService.getScriptProperties().setProperty(
      DAILY_SUMMARY_CFG.PROP_LAST_RUN, new Date().toISOString());
    return;
  }

  sendDailySummaryEmail_(groups);
  PropertiesService.getScriptProperties().setProperty(
    DAILY_SUMMARY_CFG.PROP_LAST_RUN, new Date().toISOString());
  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`✅ sendDailyGhnFailedSummary done ${dur}s. Total ${total} đơn cần xử lý.`);
}

function sendDailySummaryEmail_(groups) {
  const total = groups.ghnFailed.length + groups.stuckGhn.length +
                groups.stuckStock.length + groups.deliveryFailHigh.length;
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  const subject = `📋 [K-Homes] Tổng hợp ${total} đơn cần xử lý — ${dateStr}`;

  const sections = [];

  if (groups.stuckGhn.length > 0) {
    sections.push(buildSection_(
      `🔴 ${groups.stuckGhn.length} đơn LỖI GHN tồn ≥ ${DAILY_SUMMARY_CFG.STUCK_HOURS}h`,
      'Cần PG xử lý GẤP — chuyển sang Fbox (FFM) hoặc fix lỗi và retry GHN.',
      'https://asia.qianyierp.com/',
      groups.stuckGhn, '#c0504d'
    ));
  }
  if (groups.deliveryFailHigh.length > 0) {
    sections.push(buildSection_(
      `🟠 ${groups.deliveryFailHigh.length} đơn GHN giao thất bại ≥ ${DAILY_SUMMARY_CFG.GHN_FAIL_THRESHOLD} lần`,
      'PG cần gọi điện trực tiếp KH để đảm bảo lần giao tới thành công, hoặc đổi địa chỉ.',
      null, groups.deliveryFailHigh, '#e58a00'
    ));
  }
  if (groups.stuckStock.length > 0) {
    sections.push(buildSection_(
      `🟡 ${groups.stuckStock.length} đơn THIẾU HÀNG tồn ≥ ${DAILY_SUMMARY_CFG.STUCK_HOURS}h`,
      'Liên hệ NCC nhập thêm hàng. Sau khi sync stock + Phase 2 sẽ retry tự động.',
      null, groups.stuckStock, '#bf9000'
    ));
  }
  if (groups.ghnFailed.length > 0) {
    sections.push(buildSection_(
      `🔵 ${groups.ghnFailed.length} đơn LỖI GHN mới (< ${DAILY_SUMMARY_CFG.STUCK_HOURS}h)`,
      'Chờ PG xử lý trong ngày — fix lỗi hoặc fallback FFM.',
      null, groups.ghnFailed, '#1f4e78'
    ));
  }
  if (groups.duplicates.length > 0) {
    sections.push(buildDuplicateSection_(groups.duplicates));
  }

  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h1 style="color:#1f4e78;margin-bottom:5px;">📋 K-Homes Wellhome — Tổng hợp đơn cần xử lý</h1>
      <p style="color:#666;margin-top:0;">Ngày ${dateStr} — Tổng <b>${total}</b> đơn</p>
      ${sections.join('<br>')}
      <p style="margin-top:25px;color:#666;font-size:12px;">
        ➤ Email tự động gửi mỗi sáng 8h từ Phase 2.5 Wellhome Order Sync.<br>
        ➤ Sau khi xử lý xong, cập nhật cột X trong Sheet → đơn sẽ tự rời khỏi summary lần sau.
      </p>
    </div>`;

  MailApp.sendEmail({ to: DAILY_SUMMARY_CFG.ALERT_EMAIL, subject, htmlBody: html });
}

function buildDuplicateSection_(duplicates) {
  const dupRows = duplicates.map((d, i) => {
    const orderRows = d.orders.map(o =>
      `<tr style="background:#fff;">
        <td style="font-size:11px;">└─</td>
        <td><b>${o.orderId}</b></td>
        <td>${o.tenKH}</td>
        <td>${o.status}</td>
        <td>${Number(o.tongDon).toLocaleString('vi-VN')}đ</td>
        <td>${Math.round(o.ageHrs)}h trước</td>
      </tr>`).join('');
    return `
      <tr style="background:#e8eaf6;">
        <td colspan="6"><b>SĐT ${d.sdt}</b> — ${d.orders.length} đơn trong 24h</td>
      </tr>
      ${orderRows}`;
  }).join('');
  return `
    <div style="margin-top:20px;padding:15px;background:#f8f8f8;border-left:4px solid #5c6bc0;">
      <h3 style="margin:0 0 8px 0;color:#5c6bc0;">🔁 ${duplicates.length} KH đặt nhiều đơn trong 24h</h3>
      <p style="margin:0 0 10px 0;font-size:13px;">PG nên gọi xác nhận: KH muốn nhập đơn lặp hay double-click nhầm? Nếu là double-click → hủy bớt đơn.</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:12px;width:100%;">
        <tr style="background:#5c6bc0;color:#fff;">
          <th></th><th>Mã đơn</th><th>Tên KH</th><th>Trạng thái</th><th>Tổng</th><th>Tuổi</th>
        </tr>
        ${dupRows}
      </table>
    </div>`;
}

function buildSection_(title, hint, ctaUrl, items, color) {
  const ctaHtml = ctaUrl
    ? `<a href="${ctaUrl}" target="_blank" style="display:inline-block;padding:6px 12px;background:${color};color:#fff;text-decoration:none;border-radius:4px;font-size:13px;">Mở Fbox</a>`
    : '';
  const rowsHtml = items.map((it, i) => {
    const failCol = it.failCount !== undefined ? `<td style="text-align:center;color:${color};font-weight:bold;">${it.failCount}</td>` : '';
    const ghnCol = it.ghnCode !== undefined ? `<td>${it.ghnCode}</td>` : '';
    return `
      <tr>
        <td>${i + 1}</td>
        <td><b>${it.orderId}</b></td>
        ${ghnCol}
        ${failCol}
        <td>${it.tenKH}</td>
        <td><a href="tel:${it.sdt}">${it.sdt}</a></td>
        <td>${it.tinh}</td>
        <td>${Number(it.tongDon).toLocaleString('vi-VN')}đ</td>
        <td>${Math.round(it.ageHrs)}h</td>
        <td style="font-size:12px;color:#666;">${String(it.ghiChu).slice(0, 80)}</td>
      </tr>`;
  }).join('');
  const headerCells = items[0] && items[0].failCount !== undefined
    ? '<th>#</th><th>Mã đơn</th><th>Mã VC</th><th>Fail</th><th>KH</th><th>SĐT</th><th>Tỉnh</th><th>Tổng</th><th>Tuổi</th><th>Ghi chú</th>'
    : '<th>#</th><th>Mã đơn</th><th>KH</th><th>SĐT</th><th>Tỉnh</th><th>Tổng</th><th>Tuổi</th><th>Ghi chú</th>';

  return `
    <div style="margin-top:20px;padding:15px;background:#f8f8f8;border-left:4px solid ${color};">
      <h3 style="margin:0 0 8px 0;color:${color};">${title}</h3>
      <p style="margin:0 0 10px 0;font-size:13px;">${hint} ${ctaHtml}</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:12px;width:100%;">
        <tr style="background:${color};color:#fff;">${headerCells}</tr>
        ${rowsHtml}
      </table>
    </div>`;
}

function computeAgeHours_(ngayDat) {
  if (!ngayDat) return 0;
  let d;
  if (ngayDat instanceof Date) d = ngayDat;
  else if (typeof ngayDat === 'string') {
    // Handle "dd/MM/yyyy HH:mm" format
    const m = ngayDat.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
    if (m) d = new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]);
    else d = new Date(ngayDat);
  } else return 0;
  if (isNaN(d.getTime())) return 0;
  return (Date.now() - d.getTime()) / 3600000;
}

function manualSendDailySummary() {
  Logger.log('🟡 Manual run sendDailyGhnFailedSummary');
  sendDailyGhnFailedSummary();
}

function setupDailySummaryTrigger() {
  removeDailySummaryTrigger();
  ScriptApp.newTrigger('sendDailyGhnFailedSummary')
    .timeBased().atHour(DAILY_SUMMARY_CFG.CRON_HOUR).everyDays(1).create();
  Logger.log(`✅ Cron sendDailyGhnFailedSummary mỗi ngày ${DAILY_SUMMARY_CFG.CRON_HOUR}h sáng`);
}

function removeDailySummaryTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyGhnFailedSummary') ScriptApp.deleteTrigger(t);
  });
}
