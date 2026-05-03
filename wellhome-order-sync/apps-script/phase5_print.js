/**
 * PHASE 5 — IN VẬN ĐƠN GHN  (paste vào CÙNG project Wellhome Order Sync — file thứ 5)
 * Tận dụng GHN_CFG, GHN_COL, ghnFetch_, openGhnTargetTab_, formatGhnNow_ của Phase 2.
 *
 * Workflow:
 *   - Quét tab "tracking haravan" → lấy mã VC GHN (cột Z) của đơn:
 *     • có mã VC
 *     • cột AG (Đã in lúc) trống
 *     • cột X NOT in [Hủy, Đã giao, Lỗi GHN, Pickup tại SR]
 *   - Call POST /v2/a5/gen-token với batch order_codes → nhận token
 *   - URL print: https://online-gateway.ghn.vn/a5/public-api/printA5?token={token}
 *   - Web app trả URL về bound script → bound script mở dialog → user click in
 *   - Sau in xong, mark cột AG = timestamp để cron sau không in lại
 *
 * Public functions:
 *   setupPrintExtHeader            — add cột AG "Đã in lúc"
 *   printAllUnprintedLabels        — gen URL in cho TẤT CẢ đơn chưa in (gộp 1 PDF)
 *   markPrintedNow(orderCodes)     — mark các đơn đã in vào cột AG
 *   manualPrintOne("GYWxxx")       — debug: gen URL cho 1 đơn
 *   showPrintReport                — count đơn chưa in / đã in
 */

const GHN_PRINT_CFG = {
  GEN_TOKEN_PATH: '/v2/a5/gen-token',
  PRINT_URL_BASE: 'https://online-gateway.ghn.vn/a5/public-api/printA5',
  PRINT_URL_A6_BASE: 'https://online-gateway.ghn.vn/a5/public-api/printA6',
  MAX_LABELS_PER_BATCH: 100,
  // Đơn KHÔNG cần in (đã giao xong, hủy, hoặc lỗi không tạo được)
  NO_PRINT_STATUSES: ['Hủy', 'Đã giao', 'Lỗi GHN', 'Pickup tại SR'],
};

const GHN_PRINT_COL = {
  PRINTED_AT: 33,   // AG — Đã in vận đơn lúc (cột AC=29 sync, AD=30 fulfill, AE=31 fail count, AF=32 alerted_at)
};

// ============================================================
// SETUP
// ============================================================

function setupPrintExtHeader() {
  const sheet = openGhnTargetTab_();
  sheet.getRange(1, GHN_PRINT_COL.PRINTED_AT)
    .setValue('Đã in vận đơn lúc')
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setColumnWidth(GHN_PRINT_COL.PRINTED_AT, 140);
  Logger.log('✅ Đã thêm cột AG "Đã in vận đơn lúc"');
}

// ============================================================
// CORE — gen token + return URL
// ============================================================

/**
 * Lấy danh sách mã VC GHN cần in: có mã VC, cột AG trống, cột X non-terminal.
 * Trả về { codes: [...], rows: [...rowIndices...] }
 */
function listUnprintedGhnCodes_() {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { codes: [], rows: [] };

  const lastCol = Math.max(GHN_PRINT_COL.PRINTED_AT, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const codes = [];
  const rows = [];
  const seen = new Set();   // dedupe cùng mã VC nhiều dòng

  data.forEach((row, idx) => {
    const code = String(row[GHN_COL.GHN_MA - 1] || '').trim();
    if (!code) return;
    const printedAt = row[GHN_PRINT_COL.PRINTED_AT - 1];
    if (printedAt) return;   // đã in
    const status = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    if (GHN_PRINT_CFG.NO_PRINT_STATUSES.indexOf(status) >= 0) return;

    if (!seen.has(code)) {
      seen.add(code);
      codes.push(code);
    }
    rows.push(idx + 2);
  });

  return { codes, rows };
}

/**
 * Web app handler — gen token cho tất cả đơn chưa in, return URL.
 * Bound script gọi action này, nhận URL, mở dialog cho user in.
 */
function generatePrintUrlForUnprinted_() {
  const { codes, rows } = listUnprintedGhnCodes_();
  if (codes.length === 0) {
    return { ok: true, url: null, count: 0, message: 'Không có đơn nào chưa in' };
  }
  if (codes.length > GHN_PRINT_CFG.MAX_LABELS_PER_BATCH) {
    Logger.log(`⚠️ ${codes.length} đơn vượt giới hạn ${GHN_PRINT_CFG.MAX_LABELS_PER_BATCH}, chỉ in ${GHN_PRINT_CFG.MAX_LABELS_PER_BATCH} đơn đầu`);
    codes.length = GHN_PRINT_CFG.MAX_LABELS_PER_BATCH;
  }

  const token = genGhnPrintToken_(codes);
  if (!token) {
    return { ok: false, error: 'Không gen được token GHN — check log Phase 5' };
  }

  // Mark printed_at cho tất cả các row liên quan
  const sheet = openGhnTargetTab_();
  const now = formatGhnNow_();
  const rowsToMark = [];
  const lastCol = Math.max(GHN_PRINT_COL.PRINTED_AT, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  data.forEach((row, idx) => {
    const code = String(row[GHN_COL.GHN_MA - 1] || '').trim();
    if (codes.indexOf(code) >= 0) rowsToMark.push(idx + 2);
  });
  rowsToMark.forEach(r => {
    try { sheet.getRange(r, GHN_PRINT_COL.PRINTED_AT).setValue(now); }
    catch (e) { Logger.log(`⚠️ markPrinted row ${r}: ${e.message}`); }
  });

  const url = `${GHN_PRINT_CFG.PRINT_URL_BASE}?token=${encodeURIComponent(token)}`;
  Logger.log(`✅ Gen URL in tem cho ${codes.length} đơn: ${url}`);
  return { ok: true, url, count: codes.length, codes };
}

/** Gọi GHN gen-token API */
function genGhnPrintToken_(orderCodes) {
  if (!orderCodes || orderCodes.length === 0) return null;
  const res = ghnFetch_(GHN_PRINT_CFG.GEN_TOKEN_PATH, { order_codes: orderCodes });
  if (!res.ok) {
    Logger.log(`❌ gen-token GHN: ${res.error}`);
    return null;
  }
  // GHN response: { code:200, data:{token:"..."} } hoặc trả token trực tiếp
  const token = (res.data && res.data.token) || res.token;
  if (!token) {
    Logger.log(`❌ gen-token GHN response không có token: ${JSON.stringify(res).slice(0, 300)}`);
    return null;
  }
  return token;
}

// ============================================================
// PUBLIC — gọi tay từ menu
// ============================================================

/**
 * In tất cả đơn chưa in. Khi gọi tay, log URL ra console (user copy đi mở browser).
 * Khi gọi qua web app, return URL cho bound script.
 */
function printAllUnprintedLabels() {
  const r = generatePrintUrlForUnprinted_();
  if (!r.ok) { Logger.log(`❌ ${r.error}`); return; }
  if (r.count === 0) { Logger.log('Không có đơn nào chưa in'); return; }
  Logger.log(`📋 Gen tem cho ${r.count} đơn: ${r.codes.join(', ')}`);
  Logger.log(`👉 URL in tem: ${r.url}`);
  Logger.log(`👉 Copy URL trên paste vào browser để in PDF (gộp ${r.count} đơn).`);
}

function manualPrintOne(orderCode) {
  if (!orderCode) { Logger.log('Truyền vào Mã VC GHN. Vd: manualPrintOne("GYWW9VAP")'); return; }
  const token = genGhnPrintToken_([orderCode]);
  if (!token) return;
  const url = `${GHN_PRINT_CFG.PRINT_URL_BASE}?token=${encodeURIComponent(token)}`;
  Logger.log(`👉 URL in tem ${orderCode}: ${url}`);
}

/** Reset cột AG cho 1 đơn — dùng khi user muốn in lại */
function resetPrintedFlag(orderCode) {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, GHN_PRINT_COL.PRINTED_AT).getValues();
  let count = 0;
  data.forEach((row, idx) => {
    if (String(row[GHN_COL.GHN_MA - 1] || '').trim() === orderCode) {
      sheet.getRange(idx + 2, GHN_PRINT_COL.PRINTED_AT).setValue('');
      count++;
    }
  });
  Logger.log(`✅ Reset cột AG cho ${count} dòng có Mã VC ${orderCode}`);
}

function showPrintReport() {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng'); return; }
  const data = sheet.getRange(2, 1, lastRow - 1, GHN_PRINT_COL.PRINTED_AT).getValues();
  let total = 0, printed = 0, unprinted = 0, noCode = 0, terminal = 0;
  const seen = new Set();
  data.forEach(row => {
    const code = String(row[GHN_COL.GHN_MA - 1] || '').trim();
    if (!code) { noCode++; return; }
    if (seen.has(code)) return;
    seen.add(code);
    total++;
    const status = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    if (GHN_PRINT_CFG.NO_PRINT_STATUSES.indexOf(status) >= 0) { terminal++; return; }
    if (row[GHN_PRINT_COL.PRINTED_AT - 1]) printed++; else unprinted++;
  });
  Logger.log(`===== Báo cáo in tem =====`);
  Logger.log(`Tổng đơn có Mã VC (unique): ${total}`);
  Logger.log(`  Đã in:        ${printed}`);
  Logger.log(`  Chưa in:      ${unprinted}`);
  Logger.log(`  Terminal:     ${terminal} (Hủy/Đã giao/Lỗi/Pickup — không cần in)`);
  Logger.log(`Dòng không có Mã VC: ${noCode}`);
}
