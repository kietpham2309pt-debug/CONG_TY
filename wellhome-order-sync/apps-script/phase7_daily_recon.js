/**
 * PHASE 7 — DAILY RECON (Đối soát cuối ngày)  (paste vào CÙNG project — file thứ 8)
 *
 * Cron 18h chiều mỗi ngày:
 *   - Quét tab "tracking haravan", tính các metrics ngày hôm nay
 *   - Append 1 dòng vào tab "đối soát" với data thống kê
 *   - Phân tích KL quy đổi GHN — list các SKU có volumetric ratio > 1.5 (phí ship cao)
 *   - Email tổng hợp HTML cho admin@khomes.com.vn
 *
 * Theo nhân viên video Tefal: KL quy đổi GHN = (D × R × C) / 6000 (gram). Nếu KL quy đổi cao hơn
 * KL thực > 2kg → điều chỉnh DxRxC trong Scheme để giảm phí ship.
 *
 * Public:
 *   setupReconHeader            — tạo tab "đối soát" + header
 *   runDailyRecon               — hàm chính (cron 18h)
 *   manualRunDailyRecon         — debug
 *   setupReconTrigger           — cài cron 18h
 *   removeReconTrigger
 *   showReconReport             — log thống kê 7 ngày gần nhất
 */

const RECON_CFG = {
  TARGET_TAB: 'đối soát',
  ALERT_EMAIL: 'admin@khomes.com.vn',
  CRON_HOUR: 18,
  VOL_RATIO_WARN: 1.5,           // KL quy đổi / KL thực > 1.5 = đề xuất review Scheme
  VOL_DIVISOR: 6000,             // GHN dùng /6000 cho hàng tiêu chuẩn
  PROP_LAST_RUN: 'RECON_LAST_RUN',
};

const RECON_HEADERS = [
  'Ngày', 'Đơn mới', 'KOL', 'Organic', 'Nội bộ', 'Khác',
  'Tạo GHN OK', 'Lỗi GHN', 'Pickup SR', 'Thiếu hàng',
  'Đã giao', 'Đang giao', 'Hủy', 'Pending',
  'Fulfill Haravan', 'Tổng GMV (đ)', 'Phí GHN (đ)',
  'Vol ratio cảnh báo', 'Cập nhật lúc'
];

// ============================================================
// SETUP
// ============================================================

function setupReconHeader() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  let sheet = ss.getSheetByName(RECON_CFG.TARGET_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(RECON_CFG.TARGET_TAB);
    Logger.log(`✅ Đã tạo tab mới "${RECON_CFG.TARGET_TAB}"`);
  }
  sheet.getRange(1, 1, 1, RECON_HEADERS.length).setValues([RECON_HEADERS])
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  const widths = [100, 80, 60, 70, 70, 60,
                  90, 80, 80, 90,
                  80, 90, 60, 80,
                  120, 130, 110, 100, 140];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Format số tiền cho cột P (16) Tổng GMV và Q (17) Phí GHN
  const N = Math.max(sheet.getLastRow(), 1000);
  sheet.getRange(`P2:P${N}`).setNumberFormat('#,##0" đ"');
  sheet.getRange(`Q2:Q${N}`).setNumberFormat('#,##0" đ"');

  Logger.log(`✅ Setup header tab "${RECON_CFG.TARGET_TAB}" OK`);
}

function setupReconTrigger() {
  removeReconTrigger();
  ScriptApp.newTrigger('runDailyRecon')
    .timeBased().atHour(RECON_CFG.CRON_HOUR).everyDays(1).create();
  Logger.log(`✅ Cron runDailyRecon mỗi ngày ${RECON_CFG.CRON_HOUR}h chiều`);
}

function removeReconTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runDailyRecon') ScriptApp.deleteTrigger(t);
  });
}

// ============================================================
// MAIN
// ============================================================

function runDailyRecon() {
  const t0 = Date.now();
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng, skip'); return; }

  const lastCol = Math.max(33, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const today = todayDateOnly_();
  const stats = {
    donMoi: 0, kol: 0, organic: 0, noiBo: 0, khac: 0,
    ghnOK: 0, ghnLoi: 0, pickup: 0, thieuHang: 0,
    daGiao: 0, dangGiao: 0, huy: 0, pending: 0,
    fulfilled: 0, gmv: 0, phiGhn: 0,
  };

  const seenOrders = new Set();
  data.forEach(row => {
    const ngayDat = parseDate_(row[GHN_COL.NGAY_DAT - 1]);
    if (!ngayDat) return;
    if (!isSameDate_(ngayDat, today)) return;

    const orderId = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim();
    if (!orderId) return;

    const isFirstRow = !seenOrders.has(orderId);
    if (isFirstRow) seenOrders.add(orderId);

    if (isFirstRow) {
      stats.donMoi++;
      const loaiDon = String(row[GHN_COL.LOAI_DON - 1] || '').trim();
      if (loaiDon === 'KOL') stats.kol++;
      else if (loaiDon === 'Organic') stats.organic++;
      else if (loaiDon === 'Nội bộ') stats.noiBo++;
      else stats.khac++;

      stats.gmv += Number(row[GHN_COL.TONG_DON - 1] || 0);

      const status = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
      if (status === 'Đã tạo GHN' || status === 'Đã giao' || status === 'Đang giao') stats.ghnOK++;
      else if (status === 'Lỗi GHN') stats.ghnLoi++;
      else if (status === 'Pickup tại SR') stats.pickup++;
      else if (status === 'Thiếu hàng') stats.thieuHang++;

      if (status === 'Đã giao') stats.daGiao++;
      else if (status === 'Đã tạo GHN' || status === 'Đang gọi KH' || status === 'Đã xác nhận') stats.dangGiao++;
      else if (status === 'Hủy') stats.huy++;
      else stats.pending++;

      const fulfillAt = row[29];   // cột AD = 30, idx 29
      if (fulfillAt) stats.fulfilled++;

      stats.phiGhn += Number(row[GHN_COL.GHN_PHI - 1] || 0);
    }
  });

  // Phân tích KL quy đổi từ Scheme
  const volWarnings = analyzeVolumetricWeights_();

  // Ghi snapshot vào tab "vol_warnings" trong file 1Bn4C0Ud (đối chứng tại chỗ)
  try { writeVolWarningsToTab_(volWarnings); }
  catch (e) { Logger.log(`⚠️ writeVolWarningsToTab_: ${e.message}`); }

  const dateStr = formatDateOnly_(today);
  const updateAt = formatGhnNow_();
  const newRow = [
    dateStr,
    stats.donMoi, stats.kol, stats.organic, stats.noiBo, stats.khac,
    stats.ghnOK, stats.ghnLoi, stats.pickup, stats.thieuHang,
    stats.daGiao, stats.dangGiao, stats.huy, stats.pending,
    stats.fulfilled, stats.gmv, stats.phiGhn,
    volWarnings.length, updateAt,
  ];

  // Append vào tab "đối soát" (idempotent: nếu hôm nay đã có row → update)
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  let reconSheet = ss.getSheetByName(RECON_CFG.TARGET_TAB);
  if (!reconSheet) { setupReconHeader(); reconSheet = ss.getSheetByName(RECON_CFG.TARGET_TAB); }

  const reconLastRow = reconSheet.getLastRow();
  let targetRow = reconLastRow + 1;
  if (reconLastRow >= 2) {
    const lastDateCell = String(reconSheet.getRange(reconLastRow, 1).getValue() || '');
    if (lastDateCell === dateStr) targetRow = reconLastRow;   // overwrite same-day row
  }
  reconSheet.getRange(targetRow, 1, 1, RECON_HEADERS.length).setValues([newRow]);

  // Email summary
  try {
    sendReconEmail_(stats, volWarnings, dateStr);
  } catch (e) {
    Logger.log(`⚠️ sendReconEmail: ${e.message}`);
  }

  PropertiesService.getScriptProperties().setProperty(
    RECON_CFG.PROP_LAST_RUN, new Date().toISOString());

  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`✅ runDailyRecon ${dateStr} done ${dur}s. ${stats.donMoi} đơn mới, GMV ${stats.gmv.toLocaleString('vi-VN')}đ`);
}

/**
 * Đọc Scheme master (ưu tiên tab local "scheme_master", fallback file ngoài),
 * tính ratio = vol_weight / actual_weight cho mỗi SKU.
 * Trả về list SKU có ratio > VOL_RATIO_WARN.
 */
function analyzeVolumetricWeights_() {
  const warnings = [];
  try {
    let schemeSheet = null;
    const local = (typeof loadSchemeFromLocal_ === 'function') ? loadSchemeFromLocal_() : null;
    if (local) {
      schemeSheet = local.sheet;
    } else {
      schemeSheet = SpreadsheetApp.openById(GHN_CFG.SCHEME_SHEET_ID).getSheets()
        .find(s => s.getSheetId() === GHN_CFG.SCHEME_TAB_GID);
    }
    if (!schemeSheet) return warnings;
    const lastRow = schemeSheet.getLastRow();
    if (lastRow < 2) return warnings;
    const data = schemeSheet.getRange(2, 1, lastRow - 1,
      Math.max(GHN_SCHEME_COL.CAO, schemeSheet.getLastColumn())).getValues();

    data.forEach(row => {
      const sku = String(row[GHN_SCHEME_COL.SKU_CB - 1] || '').trim() ||
                  String(row[GHN_SCHEME_COL.SKU_CP - 1] || '').trim();
      if (!sku) return;
      const name = String(row[GHN_SCHEME_COL.NAME_CB - 1] || '').trim();
      const w = Number(row[GHN_SCHEME_COL.KL - 1] || 0);
      const l = Number(row[GHN_SCHEME_COL.DAI - 1] || 0);
      const wd = Number(row[GHN_SCHEME_COL.RONG - 1] || 0);
      const h = Number(row[GHN_SCHEME_COL.CAO - 1] || 0);
      if (w <= 0 || l <= 0 || wd <= 0 || h <= 0) return;

      const volW = (l * wd * h) / RECON_CFG.VOL_DIVISOR * 1000;   // gram
      const ratio = volW / w;
      if (ratio > RECON_CFG.VOL_RATIO_WARN) {
        // Đề xuất giảm DxRxC để ratio = 1: scale_factor = (volW_target / volW)^(1/3)
        const scale = Math.cbrt(1 / ratio);
        warnings.push({
          sku, name,
          actualW: w, volW: Math.round(volW), ratio: ratio.toFixed(2),
          suggestion: `Hiện ${l}×${wd}×${h}cm → ${Math.round(l * scale)}×${Math.round(wd * scale)}×${Math.round(h * scale)}cm`
        });
      }
    });
  } catch (e) {
    Logger.log(`⚠️ analyzeVolumetricWeights_: ${e.message}`);
  }
  return warnings;
}

/**
 * Ghi snapshot vol warnings vào tab "vol_warnings" trong file 1Bn4C0Ud (K-Homès Bosch chính).
 * Mỗi lần Phase 7 chạy → clear data + rewrite. Header giữ nguyên.
 * Cấu trúc cột: SKU | Tên SP | KL thực (g) | KL quy đổi (g) | Ratio | DxRxC hiện tại | DxRxC đề xuất | Detect lúc
 */
function writeVolWarningsToTab_(warnings) {
  const TAB = 'vol_warnings';
  const HEADERS = ['SKU', 'Tên SP', 'KL thực (g)', 'KL quy đổi (g)', 'Ratio',
                   'DxRxC hiện tại', 'DxRxC đề xuất', 'Detect lúc'];
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  let sheet = ss.getSheetByName(TAB);
  if (!sheet) {
    sheet = ss.insertSheet(TAB);
    Logger.log(`✅ Tạo tab mới "${TAB}"`);
  }

  // Re-setup header (idempotent)
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#e58a00').setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  const widths = [120, 320, 90, 110, 70, 130, 150, 140];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Conditional format ratio col E (5): ratio >= 2 → đỏ, 1.5-2 → cam
  const N = Math.max(sheet.getLastRow(), 1000);
  const colE = sheet.getRange(2, 5, N - 1, 1);
  const rules = sheet.getConditionalFormatRules().filter(r => {
    const ranges = r.getRanges();
    return !ranges.some(rg => rg.getColumn() === 5);
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(2.0)
    .setBackground('#F8CECC').setFontColor('#990000').setBold(true)
    .setRanges([colE]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1.5, 1.99)
    .setBackground('#FFE6CC')
    .setRanges([colE]).build());
  sheet.setConditionalFormatRules(rules);

  // Clear data cũ (giữ row 1 header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();

  if (!warnings || warnings.length === 0) {
    sheet.getRange(2, 1).setValue('(Không có SKU nào có KL quy đổi cao bất thường)')
      .setFontStyle('italic').setFontColor('#888888');
    return;
  }

  const now = formatGhnNow_();
  // Sort theo ratio descending — ưu tiên SKU lệch nhiều nhất lên đầu
  const sorted = warnings.slice().sort((a, b) => parseFloat(b.ratio) - parseFloat(a.ratio));
  const rows = sorted.map(w => {
    const cur = w.suggestion.match(/Hiện ([\d×]+cm)/);
    const sug = w.suggestion.match(/→ ([\d×]+cm)/);
    return [
      w.sku, w.name, w.actualW, w.volW, parseFloat(w.ratio),
      cur ? cur[1] : '', sug ? sug[1] : '', now
    ];
  });
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);

  Logger.log(`✅ writeVolWarningsToTab_: ghi ${rows.length} SKU vào tab "${TAB}"`);
}

function sendReconEmail_(stats, volWarnings, dateStr) {
  const subject = `📊 [K-Homes] Đối soát ngày ${dateStr} — ${stats.donMoi} đơn mới, GMV ${stats.gmv.toLocaleString('vi-VN')}đ`;

  const ratioFulfill = stats.donMoi > 0 ? Math.round(stats.fulfilled / stats.donMoi * 100) : 0;
  const ratioGhnFail = stats.donMoi > 0 ? Math.round(stats.ghnLoi / stats.donMoi * 100) : 0;

  const volSection = volWarnings.length > 0 ? `
    <div style="margin-top:20px;padding:15px;background:#fff8e1;border-left:4px solid #e58a00;">
      <h3 style="margin:0 0 8px 0;color:#e58a00;">⚖️ ${volWarnings.length} SKU có KL quy đổi cao — gợi ý review Scheme</h3>
      <p style="margin:0 0 10px 0;font-size:13px;">
        Theo công thức GHN: KL quy đổi = (D×R×C)/6000. Khi KL quy đổi > KL thực × ${RECON_CFG.VOL_RATIO_WARN}
        → phí ship đang cao hơn cần thiết. Gợi ý điều chỉnh DxRxC trong file Scheme để giảm chi phí.
      </p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:12px;width:100%;">
        <tr style="background:#e58a00;color:#fff;">
          <th>SKU</th><th>Tên SP</th><th>KL thực</th><th>KL quy đổi</th><th>Ratio</th><th>Đề xuất kích thước mới</th>
        </tr>
        ${volWarnings.slice(0, 30).map(w => `
          <tr>
            <td><b>${w.sku}</b></td>
            <td style="font-size:11px;">${String(w.name).slice(0, 60)}</td>
            <td>${w.actualW}g</td>
            <td>${w.volW}g</td>
            <td style="color:#e58a00;font-weight:bold;">${w.ratio}×</td>
            <td style="font-size:12px;">${w.suggestion}</td>
          </tr>`).join('')}
      </table>
      ${volWarnings.length > 30 ? `<p style="font-size:12px;color:#666;">...còn ${volWarnings.length - 30} SKU nữa</p>` : ''}
    </div>` : '';

  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h1 style="color:#1f4e78;margin-bottom:5px;">📊 Báo cáo cuối ngày — ${dateStr}</h1>
      <p style="color:#666;margin-top:0;">K-Homes Wellhome Order Pipeline</p>

      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:13px;margin-top:15px;">
        <tr style="background:#1f4e78;color:#fff;">
          <th colspan="2">Tổng quan</th>
        </tr>
        <tr><td>📦 Đơn mới</td><td><b>${stats.donMoi}</b></td></tr>
        <tr><td>💰 Tổng GMV</td><td><b>${stats.gmv.toLocaleString('vi-VN')}đ</b></td></tr>
        <tr><td>🚚 Phí GHN</td><td>${stats.phiGhn.toLocaleString('vi-VN')}đ ${stats.gmv > 0 ? `(${(stats.phiGhn/stats.gmv*100).toFixed(1)}%)` : ''}</td></tr>
      </table>

      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:13px;margin-top:15px;">
        <tr style="background:#1f4e78;color:#fff;"><th colspan="2">Phân loại đơn</th></tr>
        <tr><td>KOL</td><td>${stats.kol}</td></tr>
        <tr><td>Organic</td><td>${stats.organic}</td></tr>
        <tr><td>Nội bộ</td><td>${stats.noiBo}</td></tr>
        <tr><td>Khác</td><td>${stats.khac}</td></tr>
      </table>

      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:13px;margin-top:15px;">
        <tr style="background:#1f4e78;color:#fff;"><th colspan="2">Pipeline status</th></tr>
        <tr><td>✅ Tạo GHN OK</td><td>${stats.ghnOK}</td></tr>
        <tr><td style="color:#c0504d;">⚠️ Lỗi GHN</td><td><b>${stats.ghnLoi}</b> (${ratioGhnFail}%)</td></tr>
        <tr><td>📦 Pickup tại SR</td><td>${stats.pickup}</td></tr>
        <tr><td>📦 Thiếu hàng</td><td>${stats.thieuHang}</td></tr>
        <tr><td>📤 Fulfill Haravan</td><td>${stats.fulfilled} (${ratioFulfill}%)</td></tr>
      </table>

      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:13px;margin-top:15px;">
        <tr style="background:#1f4e78;color:#fff;"><th colspan="2">Trạng thái giao</th></tr>
        <tr><td>✅ Đã giao</td><td>${stats.daGiao}</td></tr>
        <tr><td>🚚 Đang giao</td><td>${stats.dangGiao}</td></tr>
        <tr><td>❌ Hủy</td><td>${stats.huy}</td></tr>
        <tr><td>⏳ Pending khác</td><td>${stats.pending}</td></tr>
      </table>

      ${volSection}

      <p style="margin-top:25px;color:#666;font-size:12px;">
        ➤ Email tự động gửi 18h mỗi ngày từ Phase 7 Wellhome Order Sync.<br>
        ➤ Tab "đối soát" trong Sheet K-Homès Bosch lưu lại lịch sử thống kê các ngày.
      </p>
    </div>`;

  MailApp.sendEmail({ to: RECON_CFG.ALERT_EMAIL, subject, htmlBody: html });
}

// ============================================================
// HELPERS
// ============================================================

function todayDateOnly_() {
  const d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function formatDateOnly_(d) {
  return Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
}

function isSameDate_(a, b) {
  return a.getFullYear() === b.getFullYear() &&
         a.getMonth() === b.getMonth() &&
         a.getDate() === b.getDate();
}

function parseDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'string') {
    // Format "dd/MM/yyyy HH:mm"
    const m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function manualRunDailyRecon() {
  Logger.log('🟡 Manual run runDailyRecon');
  runDailyRecon();
}

function showReconReport() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(RECON_CFG.TARGET_TAB);
  if (!sheet) { Logger.log('Tab "đối soát" chưa tạo, chạy setupReconHeader trước'); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Tab "đối soát" rỗng'); return; }
  const startRow = Math.max(2, lastRow - 6);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, RECON_HEADERS.length).getValues();
  Logger.log(`===== Đối soát 7 ngày gần nhất =====`);
  data.forEach(r => {
    Logger.log(`${r[0]} | Đơn mới: ${r[1]} | GHN OK: ${r[6]} | Lỗi: ${r[7]} | Đã giao: ${r[10]} | GMV: ${Number(r[15]).toLocaleString('vi-VN')}đ`);
  });
}
