/**
 * SETUP RUNNER — gộp tất cả hàm setup của 6 phase để chạy 1 lệnh duy nhất.
 * Dùng khi paste code mới hoặc anh muốn re-init.
 *
 * Public:
 *   runAllSetups        — chạy setup toàn bộ (Phase 1+2+3+4+5+6)
 *   runIncrementalSetup — chỉ chạy setup các file mới (Phase 5+6 + locale fix)
 */

function runAllSetups() {
  Logger.log('===== Run All Setups (toàn bộ 6 phase) =====');
  const steps = [
    ['Phase 1 — setupHeader (locale + cột A plain text)', setupHeader],
    ['Phase 2 — setupGhnAll', setupGhnAll],
    ['Phase 3 — setupGhnStatusAll', setupGhnStatusAll],
    ['Phase 4 — setupHv4ExtHeader', setupHv4ExtHeader],
    ['Phase 5 — setupPrintExtHeader', setupPrintExtHeader],
    ['Phase 6 — setupStockHeader', setupStockHeader],
  ];
  steps.forEach(([label, fn]) => {
    try {
      Logger.log(`▶ ${label}`);
      fn();
    } catch (e) {
      Logger.log(`❌ ${label}: ${e.message}`);
    }
  });
  Logger.log('===== ✅ Done all setups =====');
}

/** Chỉ chạy setup các thay đổi mới (Phase 5 in tem + Phase 6 stock + locale). */
function runIncrementalSetup() {
  Logger.log('===== Run Incremental Setup =====');
  const steps = [
    ['Phase 1 — setupHeader (locale + cột A plain text)', setupHeader],
    ['Phase 4 — setupHv4ExtHeader (cột AD)', setupHv4ExtHeader],
    ['Phase 3 — setupGhnStatusExtHeader (cột AC, AE, AF)', setupGhnStatusExtHeader],
    ['Phase 5 — setupPrintExtHeader (cột AG print)', setupPrintExtHeader],
    ['Phase 6 — setupStockHeader (tạo tab stock haravan)', setupStockHeader],
    ['Phase 7 — setupReconHeader (tạo tab đối soát)', setupReconHeader],
    ['Phase 12 — syncSchemeToLocal (clone Scheme master vào tab local)', syncSchemeToLocal],
    ['Phase 6 — syncHaravanInventory (pull tồn kho lần đầu)', syncHaravanInventory],
  ];
  steps.forEach(([label, fn]) => {
    try {
      Logger.log(`▶ ${label}`);
      fn();
    } catch (e) {
      Logger.log(`❌ ${label}: ${e.message}`);
    }
  });
  Logger.log('===== ✅ Done incremental setup =====');
}

// ============================================================
// AUDIT + REPAIR — fix bug Phase 3 column conflict (03/05/2026)
// ============================================================

/**
 * Inspect current state cột AC-AG. Trả về header + sample data row 2-6.
 */
function auditTrackingColumns() {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  const cols = [
    { idx: 29, label: 'AC', expected: 'GHN cập nhật lúc' },
    { idx: 30, label: 'AD', expected: 'Fulfill Haravan lúc' },
    { idx: 31, label: 'AE', expected: 'Số lần giao fail' },
    { idx: 32, label: 'AF', expected: 'Đã alert lúc' },
    { idx: 33, label: 'AG', expected: 'Đã in vận đơn lúc' },
  ];
  const result = { headerStatus: [], sampleData: [] };
  cols.forEach(c => {
    const header = String(sheet.getRange(1, c.idx).getValue() || '');
    const ok = header === c.expected;
    result.headerStatus.push(`${c.label}(${c.idx}): "${header}" ${ok ? '✅' : '❌ EXPECTED: "' + c.expected + '"'}`);
  });
  if (lastRow > 1) {
    const n = Math.min(5, lastRow - 1);
    const data = sheet.getRange(2, 29, n, 5).getValues();
    data.forEach((row, i) => {
      const types = row.map(v => `${typeof v}/${v instanceof Date ? 'Date' : ''}/${String(v).slice(0, 30)}`);
      result.sampleData.push(`Row ${i + 2}: ${types.join(' | ')}`);
    });
  }
  Logger.log('===== AUDIT cột AC-AG =====');
  result.headerStatus.forEach(s => Logger.log(s));
  Logger.log('--- Sample data (5 rows đầu) ---');
  result.sampleData.forEach(s => Logger.log(s));
  return result;
}

/**
 * REPAIR sau bug Phase 3 column conflict:
 *   Trước fix: Phase 3 ghi GHN_FAIL_COUNT vào cột AD (30) thay vì AE (31)
 *              → đè timestamp Phase 4 FULFILL_AT
 *   Sau fix: code đã đổi GHN_FAIL_COUNT=31, GHN_ALERTED_AT=32
 *
 * Repair logic per row:
 *   - AD (30): nếu là Number (fail count Phase 3 ghi sai) → MOVE sang AE → clear AD
 *              nếu là Date / string "dd/MM/yyyy HH:mm" → giữ (Phase 4 timestamp đúng)
 *   - AE (31): nếu là Date / string timestamp (Phase 3 ghi alerted_at sai) → MOVE sang AF → clear AE
 *              (sau migration AD, AE có thể đã có Number — keep)
 *   - AF (32): chỉ giữ Date timestamp.
 *
 * Sau migration: re-setup headers + data validation.
 */
function repairColumnsAfterPhase3Fix() {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng, không cần repair'); return; }

  const N = lastRow - 1;
  const adData = sheet.getRange(2, 30, N, 1).getValues();
  const aeData = sheet.getRange(2, 31, N, 1).getValues();

  const newAd = [];
  const newAe = [];
  const newAf = [];
  let stats = { adKeepDate: 0, adKeepStr: 0, adNumberToAe: 0, adEmpty: 0,
                aeDateToAf: 0, aeStrToAf: 0, aeNumberKeep: 0, aeEmpty: 0 };
  const isTimestampStr = s => typeof s === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}/.test(s);

  for (let i = 0; i < N; i++) {
    const ad = adData[i][0];
    const ae = aeData[i][0];
    let outAd = '', outAe = '', outAf = '';

    // Cột AD (30) — phải là FULFILL_AT của Phase 4
    if (ad instanceof Date) { outAd = ad; stats.adKeepDate++; }
    else if (isTimestampStr(ad)) { outAd = ad; stats.adKeepStr++; }
    else if (typeof ad === 'number') {
      // Number = fail count Phase 3 ghi sai → move sang AE
      outAe = ad; stats.adNumberToAe++;
    }
    else { stats.adEmpty++; }

    // Cột AE (31) — sau fix phải là GHN_FAIL_COUNT (Number)
    if (ae instanceof Date) {
      // Date trong AE = alerted_at Phase 3 ghi sai → move sang AF
      outAf = ae; stats.aeDateToAf++;
    }
    else if (isTimestampStr(ae)) {
      outAf = ae; stats.aeStrToAf++;
    }
    else if (typeof ae === 'number') {
      // Number trong AE đã đúng (fail count) → giữ. Conflict với migration AD?
      // Nếu adNumberToAe đã set outAe → giữ outAe đó; nếu không, lấy ae current.
      if (outAe === '') outAe = ae;
      stats.aeNumberKeep++;
    }
    else { stats.aeEmpty++; }

    newAd.push([outAd]); newAe.push([outAe]); newAf.push([outAf]);
  }

  // Ghi đồng loạt (3 batch writes thay vì N×3 calls)
  sheet.getRange(2, 30, N, 1).setValues(newAd);
  sheet.getRange(2, 31, N, 1).setValues(newAe);
  sheet.getRange(2, 32, N, 1).setValues(newAf);

  Logger.log(`===== REPAIR DONE (${N} rows) =====`);
  Logger.log(`AD: keep_date=${stats.adKeepDate}, keep_str=${stats.adKeepStr}, number→AE=${stats.adNumberToAe}, empty=${stats.adEmpty}`);
  Logger.log(`AE: date→AF=${stats.aeDateToAf}, str→AF=${stats.aeStrToAf}, number_keep=${stats.aeNumberKeep}, empty=${stats.aeEmpty}`);

  // Re-setup headers
  setupHv4ExtHeader();          // AD
  setupGhnStatusExtHeader();    // AC + AE + AF
  setupPrintExtHeader();        // AG (idempotent)

  Logger.log('✅ Repair + re-setup headers done');
  return stats;
}

/**
 * Tổng hợp state hệ thống — verify deployment health.
 * Trả về số trigger, last sync timestamps, tab counts.
 */
function healthCheck() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  const triggers = ScriptApp.getProjectTriggers();
  const triggerNames = triggers.map(t => `${t.getHandlerFunction()} (every ${t.getTriggerSource()})`);

  const props = PropertiesService.getScriptProperties();
  const tabs = {
    'tracking haravan': ss.getSheetByName('tracking haravan'),
    'stock haravan': ss.getSheetByName('stock haravan'),
    'đối soát': ss.getSheetByName('đối soát'),
    'scheme_master': ss.getSheetByName('scheme_master'),
    'vol_warnings': ss.getSheetByName('vol_warnings'),
  };
  const tabStats = {};
  Object.keys(tabs).forEach(name => {
    const s = tabs[name];
    tabStats[name] = s ? { exists: true, rows: s.getLastRow(), cols: s.getLastColumn() } : { exists: false };
  });

  const result = {
    spreadsheetLocale: ss.getSpreadsheetLocale(),
    triggers: triggerNames,
    triggerCount: triggers.length,
    tabs: tabStats,
    lastRuns: {
      phase1_sync: props.getProperty('WELLHOME_LAST_SYNC') || '(chưa)',
      phase3_status: props.getProperty('GHN_STATUS_LAST_RUN') || '(chưa)',
      phase3_alert: props.getProperty('GHN_STATUS_LAST_ALERT') || '(chưa)',
      phase4_fulfill: props.getProperty('HV4_LAST_RUN') || '(chưa)',
      phase6_inventory: props.getProperty('STOCK_LAST_SYNC') || '(chưa)',
      phase2_5_summary: props.getProperty('DAILY_SUMMARY_LAST_RUN') || '(chưa)',
      phase7_recon: props.getProperty('RECON_LAST_RUN') || '(chưa)',
      phase12_scheme: props.getProperty('SCHEME_LAST_SYNC') || '(chưa)',
    },
  };
  Logger.log('===== HEALTH CHECK =====');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/** Cài tất cả cron triggers. */
function setupAllTriggers() {
  Logger.log('===== Setup all triggers =====');
  const steps = [
    ['Phase 1 — setupHourlyTrigger (syncOrders 1h)', setupHourlyTrigger],
    ['Phase 2 — setupGhnTrigger (createGhnOrders 1h)', setupGhnTrigger],
    ['Phase 2.5 — setupDailySummaryTrigger (daily 8h)', setupDailySummaryTrigger],
    ['Phase 3 — setupGhnStatusTrigger (updateGhnStatuses 30p)', setupGhnStatusTrigger],
    ['Phase 4 — setupHv4Trigger (fulfillHaravanOrders 30p)', setupHv4Trigger],
    ['Phase 6 — setupStockTrigger (syncHaravanInventory daily 6h)', setupStockTrigger],
    ['Phase 7 — setupReconTrigger (runDailyRecon daily 18h)', setupReconTrigger],
    ['Phase 8 — setupAutoRetryTrigger (autoRetryFailedGhnOrders daily 9h)', setupAutoRetryTrigger],
    ['Phase 9 — setupWeeklyReportTrigger (sendWeeklyReport Mon 9h)', setupWeeklyReportTrigger],
    ['Phase 11 — setupHealthMonitorTrigger (runHealthMonitor Sun 23h)', setupHealthMonitorTrigger],
    ['Phase 12 — setupSyncSchemeTrigger (syncSchemeToLocal daily 5h)', setupSyncSchemeTrigger],
    ['Phase 15 — setupMonthlyReportTrigger (sendMonthlyReport daily 9h, check date=1)', setupMonthlyReportTrigger],
    ['Phase 17 — setupCodReconTrigger (runCodRecon daily 19h)', typeof setupCodReconTrigger === 'function' ? setupCodReconTrigger : function () {}],
    ['Phase 18 — setupFeeAlertTrigger (runFeeAlertReport daily 19:15)', typeof setupFeeAlertTrigger === 'function' ? setupFeeAlertTrigger : function () {}],
    ['Phase 19 — setupPickupTriggers (notify 9h + fulfill 1h)', typeof setupPickupTriggers === 'function' ? setupPickupTriggers : function () {}],
    ['Phase 20 — setupKolTrigger (aggregateKolPerformance daily 19:30)', typeof setupKolTrigger === 'function' ? setupKolTrigger : function () {}],
  ];
  steps.forEach(([label, fn]) => {
    try {
      Logger.log(`▶ ${label}`);
      fn();
    } catch (e) {
      Logger.log(`❌ ${label}: ${e.message}`);
    }
  });
  Logger.log('===== ✅ Done all triggers =====');
}
