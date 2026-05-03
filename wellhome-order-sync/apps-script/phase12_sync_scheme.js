/**
 * PHASE 12 — SYNC SCHEME MASTER + Tab vol_warnings (file thứ 12)
 *
 * Mục đích: gom tất cả data đối chứng vào file 1Bn4C0Ud (K-Homès Bosch chính)
 *   - Tab `scheme_master` — clone từ file ngoài 1BWUAMv6 (Scheme campaign master)
 *   - Tab `vol_warnings` — snapshot SKU có KL quy đổi cao bất thường (Phase 7 ghi)
 *
 * Cron daily 5h sáng (trước Phase 6 sync stock 6h, trước Phase 2 cron 1h):
 *   - Pull tab Scheme (gid=1684440583) từ file 1BWUAMv6
 *   - Ghi đè vào tab `scheme_master` trong 1Bn4C0Ud
 *
 * Phase 2 + Phase 7 sau khi sync sẽ ưu tiên đọc tab local (nhanh hơn + offline-resilient).
 *
 * Public:
 *   syncSchemeToLocal              — hàm chính (cron 5h)
 *   manualSyncScheme               — debug
 *   setupSyncSchemeTrigger
 *   removeSyncSchemeTrigger
 *   loadSchemeFromLocal_           — helper: đọc tab scheme_master local (Phase 2/7 dùng)
 *   showSchemeReport               — log info tab local
 */

const SCHEME_SYNC_CFG = {
  LOCAL_TAB: 'scheme_master',
  CRON_HOUR: 5,
  PROP_LAST_SYNC: 'SCHEME_LAST_SYNC',
};

// ============================================================
// SYNC
// ============================================================

function syncSchemeToLocal() {
  const t0 = Date.now();
  const srcSs = SpreadsheetApp.openById(GHN_CFG.SCHEME_SHEET_ID);
  let srcSheet = null;
  srcSs.getSheets().forEach(s => {
    if (s.getSheetId() === GHN_CFG.SCHEME_TAB_GID) srcSheet = s;
  });
  if (!srcSheet) {
    Logger.log(`❌ Không tìm thấy tab Scheme gid=${GHN_CFG.SCHEME_TAB_GID} trong file ${GHN_CFG.SCHEME_SHEET_ID}`);
    return;
  }
  const srcName = srcSheet.getName();
  const srcLastRow = srcSheet.getLastRow();
  const srcLastCol = srcSheet.getLastColumn();
  if (srcLastRow < 1 || srcLastCol < 1) {
    Logger.log('⚠️ Source Scheme tab rỗng, skip sync');
    return;
  }
  const data = srcSheet.getRange(1, 1, srcLastRow, srcLastCol).getValues();
  const widths = [];
  for (let c = 1; c <= srcLastCol; c++) widths.push(srcSheet.getColumnWidth(c));

  const targetSs = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  let targetSheet = targetSs.getSheetByName(SCHEME_SYNC_CFG.LOCAL_TAB);
  if (!targetSheet) {
    targetSheet = targetSs.insertSheet(SCHEME_SYNC_CFG.LOCAL_TAB);
    Logger.log(`✅ Tạo tab mới "${SCHEME_SYNC_CFG.LOCAL_TAB}"`);
  }
  // Ensure target có đủ rows/cols
  const curRows = Math.max(targetSheet.getMaxRows(), srcLastRow);
  const curCols = Math.max(targetSheet.getMaxColumns(), srcLastCol);
  if (targetSheet.getMaxRows() < curRows) targetSheet.insertRowsAfter(targetSheet.getMaxRows(), curRows - targetSheet.getMaxRows());
  if (targetSheet.getMaxColumns() < curCols) targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), curCols - targetSheet.getMaxColumns());

  targetSheet.clear();
  targetSheet.getRange(1, 1, srcLastRow, srcLastCol).setValues(data);
  // Highlight header row 1+2 giống Scheme gốc (style cứng vì copy values mất format)
  targetSheet.getRange(1, 1, Math.min(2, srcLastRow), srcLastCol)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');
  targetSheet.setFrozenRows(Math.min(2, srcLastRow));
  // Set column widths
  widths.forEach((w, i) => targetSheet.setColumnWidth(i + 1, w));

  // Note metadata cell A row last+2: timestamp last sync để user biết tab fresh
  const noteRow = srcLastRow + 2;
  targetSheet.getRange(noteRow, 1).setValue(`Last sync from "${srcName}" (file 1BWUAMv6...): ${formatGhnNow_()}`)
    .setFontStyle('italic').setFontColor('#888888');

  PropertiesService.getScriptProperties().setProperty(
    SCHEME_SYNC_CFG.PROP_LAST_SYNC, new Date().toISOString());

  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`✅ syncSchemeToLocal ${dur}s. Cloned ${srcLastRow} rows × ${srcLastCol} cols từ "${srcName}"`);
}

function manualSyncScheme() {
  Logger.log('🟡 Manual syncSchemeToLocal');
  syncSchemeToLocal();
}

function setupSyncSchemeTrigger() {
  removeSyncSchemeTrigger();
  ScriptApp.newTrigger('syncSchemeToLocal')
    .timeBased().atHour(SCHEME_SYNC_CFG.CRON_HOUR).everyDays(1).create();
  Logger.log(`✅ Cron syncSchemeToLocal mỗi ngày ${SCHEME_SYNC_CFG.CRON_HOUR}h sáng`);
}

function removeSyncSchemeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncSchemeToLocal') ScriptApp.deleteTrigger(t);
  });
}

// ============================================================
// LOCAL READ HELPER — Phase 2 + Phase 7 dùng (dual-read)
// ============================================================

/**
 * Đọc tab `scheme_master` local. Trả về { sheet, hasData } hoặc null nếu chưa sync.
 * Phase 2/7 fallback file ngoài nếu return null.
 */
function loadSchemeFromLocal_() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(SCHEME_SYNC_CFG.LOCAL_TAB);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return null;
  return { sheet, lastRow, lastCol: sheet.getLastColumn() };
}

function showSchemeReport() {
  const local = loadSchemeFromLocal_();
  const last = PropertiesService.getScriptProperties().getProperty(SCHEME_SYNC_CFG.PROP_LAST_SYNC);
  Logger.log(`===== Scheme Master Report =====`);
  if (!local) {
    Logger.log(`⚠️ Tab "${SCHEME_SYNC_CFG.LOCAL_TAB}" chưa có data. Phase 2/7 đang đọc file ngoài.`);
  } else {
    Logger.log(`✅ Tab local "${SCHEME_SYNC_CFG.LOCAL_TAB}": ${local.lastRow} rows × ${local.lastCol} cols`);
  }
  Logger.log(`Last sync: ${last || '(chưa)'}`);
}
