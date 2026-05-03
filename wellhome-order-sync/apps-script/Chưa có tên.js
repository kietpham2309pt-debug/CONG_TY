/**
 * BOUND SCRIPT — gắn vào Sheet K-Homès Bosch chính (1Bn4C0Ud...)
 * Mục đích: hiển thị menu "🚀 K-Homes" cho TẤT CẢ user mở Sheet (kể cả PG/nhân viên).
 * Click menu → gọi Web App của project Wellhome standalone → tool chạy as admin@khomes.com.vn.
 *
 * CÁCH PASTE:
 *   1. Mở https://docs.google.com/spreadsheets/d/1Bn4C0UdvX2hT82p1pput3VtBDkdJe_jc1-w8ZjKfdGw/edit
 *   2. Menu Extensions → Apps Script → mở project bound (tự tạo nếu chưa có)
 *   3. Xóa code mặc định → paste toàn bộ file này → Save
 *   4. Update WEBAPP_URL bên dưới (lấy từ project Wellhome — xem hướng dẫn ở dưới)
 *   5. Refresh tab Sheet (F5) → menu "🚀 K-Homes" xuất hiện cho mọi user mở
 */

// ⚠️ PHẢI cập nhật URL này sau khi deploy Web App của project Wellhome
const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxM3XIGhIYK2Cpvj65h9bGuxMjvEIw_arQNbjj2RJSFtC40DjK5qeVbiv2oByXV8AkkUQ/exec';

// ⚠️ Phải KHỚP với GHN_WEBAPP_SECRET trong file phase2_ghn của project Wellhome
const WEBAPP_SECRET = 'REPLACE_WITH_WEBAPP_SECRET';

/** Simple onOpen trigger — TẤT CẢ user mở Sheet đều thấy menu này */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 K-Homes')
    .addItem('Tạo đơn GHN ngay', 'createGhnOrdersManual')
    .addItem('Retry đơn lỗi GHN', 'retryFailedOrdersManual')
    .addSeparator()
    .addItem('Cập nhật trạng thái GHN ngay', 'updateGhnStatusManual')
    .addItem('Fulfill Haravan ngay', 'fulfillHaravanManual')
    .addSeparator()
    .addItem('📋 In vận đơn GHN ngay', 'printGhnLabelsManual')
    .addItem('📦 Pull tồn kho Haravan ngay', 'syncInventoryManual')
    .addSeparator()
    .addItem('📊 Báo cáo cuối ngày ngay', 'dailyReconManual')
    .addItem('📨 Tổng hợp đơn cần xử lý', 'dailySummaryManual')
    .addItem('🩺 Health check', 'healthCheckManual')
    .addToUi();
}

function dailyReconManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '📊 Chạy báo cáo cuối ngày ngay',
    'Tool sẽ:\n' +
    '   • Đếm thống kê đơn hôm nay (đơn mới, GHN OK/lỗi, đã giao, ...)\n' +
    '   • Phân tích KL quy đổi GHN — list SKU cần review Scheme\n' +
    '   • Append 1 dòng vào tab "đối soát"\n' +
    '   • Gửi email tổng hợp HTML\n\n' +
    'Mặc định cron tự chạy 18h chiều. Click YES để chạy ngay.',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  callWebApp_('daily_recon',
    '📊 Đang tổng hợp báo cáo, đợi 5-15 giây...',
    '✅ Đã gửi email + cập nhật tab "đối soát"');
}

function dailySummaryManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '📨 Tổng hợp đơn cần xử lý',
    'Tool sẽ gửi email tổng hợp:\n' +
    '   • Đơn LỖI GHN tồn (cần fallback Fbox/FFM)\n' +
    '   • Đơn giao thất bại ≥ 3 lần (cần gọi KH)\n' +
    '   • Đơn THIẾU HÀNG tồn quá 24h\n\n' +
    'Mặc định cron tự chạy 8h sáng. Click YES để chạy ngay.',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  callWebApp_('daily_summary',
    '📨 Đang quét đơn cần xử lý, đợi 5-10 giây...',
    '✅ Đã gửi email tổng hợp (nếu có đơn cần xử lý)');
}

function healthCheckManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('🩺 Đang check health...', '⏳', -1);
  try {
    const res = UrlFetchApp.fetch(WEBAPP_URL, {
      method: 'post', muteHttpExceptions: true,
      payload: { secret: WEBAPP_SECRET, action: 'health_check' },
    });
    const r = JSON.parse(res.getContentText());
    if (!r.ok) { ss.toast('❌ ' + r.error, 'Lỗi', 30); return; }
    const h = r.health;
    const html = HtmlService.createHtmlOutput(
      `<div style="font-family:Arial;padding:15px;font-size:13px;">
        <h3 style="color:#1f4e78;margin-top:0;">🩺 K-Homes System Health</h3>
        <p><b>Locale:</b> ${h.spreadsheetLocale}</p>
        <p><b>Triggers (${h.triggerCount}):</b></p>
        <ul style="margin-top:0;">${h.triggers.map(t => `<li>${t}</li>`).join('')}</ul>
        <p><b>Tabs:</b></p>
        <ul style="margin-top:0;">${Object.keys(h.tabs).map(name => {
          const t = h.tabs[name];
          return `<li><b>${name}:</b> ${t.exists ? `${t.rows} rows × ${t.cols} cols` : '<span style="color:red;">KHÔNG TỒN TẠI</span>'}</li>`;
        }).join('')}</ul>
        <p><b>Last runs:</b></p>
        <ul style="margin-top:0;font-size:12px;">${Object.keys(h.lastRuns).map(k =>
          `<li><b>${k}:</b> ${h.lastRuns[k]}</li>`).join('')}</ul>
      </div>`).setWidth(550).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, '🩺 Health Check');
    ss.toast('✅ Health check OK', '', 3);
  } catch (e) {
    ss.toast('❌ ' + e.message, 'Lỗi', 30);
  }
}

function printGhnLabelsManual() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!WEBAPP_URL || WEBAPP_URL === 'PASTE_WEB_APP_URL_HERE') {
    ss.toast('❌ Chưa setup WEBAPP_URL', 'Error', 30); return;
  }
  ss.toast('📋 Đang gen URL in tem GHN...', '⏳', -1);
  try {
    const res = UrlFetchApp.fetch(WEBAPP_URL, {
      method: 'post', muteHttpExceptions: true,
      payload: { secret: WEBAPP_SECRET, action: 'print_labels' },
    });
    const r = JSON.parse(res.getContentText());
    if (!r.ok) { ss.toast('❌ ' + r.error, 'Lỗi', 30); return; }
    if (r.count === 0) { ss.toast('Không có đơn nào chưa in tem', 'Info', 10); return; }
    // Mở dialog có nút mở URL print
    const html = HtmlService.createHtmlOutput(
      `<div style="font-family:Arial;padding:20px;text-align:center;">
        <h2 style="color:#1f4e78;">📋 ${r.count} đơn sẵn sàng in</h2>
        <p>Click nút bên dưới để mở PDF tem GHN trong tab mới:</p>
        <a href="${r.url}" target="_blank" rel="noopener"
           style="display:inline-block;padding:12px 24px;background:#1f4e78;color:#fff;
                  text-decoration:none;border-radius:6px;font-weight:bold;font-size:16px;">
           Mở PDF in tem (${r.count} đơn)
        </a>
        <p style="margin-top:20px;color:#666;font-size:12px;">Đơn đã in được mark cột AG. Lần sau click sẽ chỉ in đơn mới.</p>
      </div>`).setWidth(420).setHeight(220);
    ui.showModalDialog(html, '📋 In vận đơn GHN');
    ss.toast(`✅ Đã gen tem cho ${r.count} đơn`, '', 5);
  } catch (e) {
    ss.toast('❌ ' + e.message, 'Lỗi', 30);
  }
}

function syncInventoryManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '📦 Pull tồn kho Haravan ngay',
    'Tool sẽ pull TẤT CẢ tồn kho từ Haravan API → ghi tab "stock haravan".\n' +
    'Mặc định cron tự chạy 6h sáng mỗi ngày. Click YES nếu cần update ngay.\n\n' +
    'Đợi 1-3 phút tuỳ số SP. Tiếp tục?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  callWebApp_('sync_inventory',
    '📦 Đang pull tồn kho Haravan, đợi 1-3 phút...',
    '✅ Đã cập nhật tab "stock haravan". Cron Phase 2 lần sau sẽ check tồn theo data này.');
}

function createGhnOrdersManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '🚀 Tạo đơn GHN ngay',
    'Tool sẽ quét tab "tracking haravan" và tạo đơn GHN cho TẤT CẢ dòng có:\n\n' +
    '   • Cột X = "Đã xác nhận"\n' +
    '   • Cột Z (Mã VC GHN) trống\n\n' +
    'Sau đó cập nhật X, Z, AA, AB. Đơn lỗi sẽ ghi cột Y.\n\nTiếp tục?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  callWebApp_('create_ghn_orders',
    '🚀 Đang tạo đơn GHN, đợi 10-60 giây tuỳ số đơn...',
    '✅ Hoàn tất! Kiểm tra cột Z (Mã VC), AA (Phí), Y (lỗi nếu có).');
}

function retryFailedOrdersManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '🔄 Retry đơn lỗi GHN',
    'Reset các đơn cột X = "Lỗi GHN" → "Đã xác nhận" để retry.\n\n' +
    'Sau đó click "Tạo đơn GHN ngay" để chạy ngay (hoặc đợi cron 1h tự chạy).\n\nTiếp tục?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  callWebApp_('retry_failed',
    '🔄 Đang reset đơn lỗi...',
    '✅ Đã reset. Click "Tạo đơn GHN ngay" để retry.');
}

function updateGhnStatusManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '📦 Cập nhật trạng thái GHN ngay',
    'Tool sẽ poll API GHN cho TẤT CẢ đơn có Mã VC (cột Z) chưa ở trạng thái cuối\n' +
    '(Đã giao / Hủy / Đã trả / Mất / Hỏng), cập nhật:\n\n' +
    '   • Cột U "TT giao hàng" — label tiếng Việt\n' +
    '   • Cột AC "GHN cập nhật lúc" — timestamp\n' +
    '   • Cột X "Trạng thái xử lý" — tự đổi sang "Đã giao" / "Hủy" / "Lỗi GHN" khi đơn vào trạng thái cuối\n\n' +
    'Tiếp tục?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  callWebApp_('update_ghn_status',
    '📦 Đang poll GHN, đợi 30 giây - 2 phút tuỳ số đơn...',
    '✅ Hoàn tất! Kiểm tra cột U, AC, X.');
}

function fulfillHaravanManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '📤 Fulfill Haravan ngay',
    'Tool sẽ đẩy Mã VC GHN ngược về Haravan (đánh dấu đơn "Đã giao cho ĐVVC")\n' +
    'cho TẤT CẢ dòng có:\n\n' +
    '   • Cột Z (Mã VC GHN) đã có\n' +
    '   • Cột AD (Fulfill lúc) trống\n' +
    '   • Cột X KHÁC "Hủy" và "Lỗi GHN"\n\n' +
    '⚠️ Sau khi fulfill, đơn trên Haravan sẽ chuyển sang trạng thái "Đã giao cho ĐVVC".\n' +
    'Tool có set notify_customer=false nên KH KHÔNG nhận email từ Haravan.\n\nTiếp tục?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  callWebApp_('fulfill_haravan',
    '📤 Đang fulfill Haravan, đợi 30 giây - 2 phút tuỳ số đơn...',
    '✅ Hoàn tất! Kiểm tra cột AD và trên Haravan admin.');
}

/** Helper: gọi Web App với toast progress */
function callWebApp_(action, runningMsg, successMsg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!WEBAPP_URL || WEBAPP_URL === 'PASTE_WEB_APP_URL_HERE') {
    ss.toast('❌ Chưa setup WEBAPP_URL trong bound script. Liên hệ admin.', 'Error', 30);
    return;
  }
  ss.toast(runningMsg, '⏳', -1);
  try {
    const res = UrlFetchApp.fetch(WEBAPP_URL, {
      method: 'post',
      payload: { action: action, secret: WEBAPP_SECRET },
      muteHttpExceptions: true,
      followRedirects: true,
    });
    const code = res.getResponseCode();
    if (code !== 200) {
      ss.toast('❌ HTTP ' + code + ': ' + res.getContentText().slice(0, 100), 'Error', 30);
      return;
    }
    let body;
    try { body = JSON.parse(res.getContentText()); }
    catch (e) {
      ss.toast('❌ Web App trả về không phải JSON. Có thể chưa deploy đúng.', 'Error', 30);
      return;
    }
    if (body.ok) ss.toast(successMsg, 'Done', 10);
    else ss.toast('❌ ' + (body.error || 'Unknown error'), 'Error', 30);
  } catch (e) {
    ss.toast('❌ ' + e.message, 'Error', 30);
  }
}
