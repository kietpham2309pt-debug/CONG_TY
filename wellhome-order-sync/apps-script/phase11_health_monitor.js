/**
 * PHASE 11 — AUTO HEALTH MONITOR (file thứ 11)
 *
 * Cron weekly Sunday 23h:
 *   - Verify hệ thống đang khoẻ:
 *     • 9 cron triggers expected
 *     • 3 tabs (tracking haravan, stock haravan, đối soát) tồn tại
 *     • Headers AC-AG đúng vị trí (sau bug Phase 3 column conflict)
 *     • Locale = vi_VN
 *     • Last runs không stale (Phase 1 < 24h, Phase 6 < 25h, etc)
 *   - Nếu phát hiện issue → email alert PG để fix
 *
 * Public:
 *   runHealthMonitor             — hàm chính (cron Sun 23h)
 *   manualHealthMonitor          — debug
 *   setupHealthMonitorTrigger
 *   removeHealthMonitorTrigger
 */

const HEALTH_CFG = {
  ALERT_EMAIL: 'admin@khomes.com.vn',
  EXPECTED_TRIGGERS: [
    'syncOrders', 'createGhnOrders', 'sendDailyGhnFailedSummary',
    'updateGhnStatuses', 'fulfillHaravanOrders', 'syncHaravanInventory',
    'runDailyRecon', 'autoRetryFailedGhnOrders', 'sendWeeklyReport',
    'runHealthMonitor', 'syncSchemeToLocal', 'sendMonthlyReport',
  ],
  EXPECTED_TABS: ['tracking haravan', 'stock haravan', 'đối soát'],
  EXPECTED_HEADERS: {
    29: 'GHN cập nhật lúc',
    30: 'Fulfill Haravan lúc',
    31: 'Số lần giao fail',
    32: 'Đã alert lúc',
    33: 'Đã in vận đơn lúc',
  },
  STALE_THRESHOLDS_HRS: {
    phase1: 25,        // syncOrders 1h → tối đa 25h là đã có run gần đây
    phase6: 26,        // inventory daily 6h
    phase3: 1,         // status 30p → 1h là quá lâu
    phase4: 1,
  },
};

function runHealthMonitor() {
  const t0 = Date.now();
  const issues = [];
  const ok = [];

  // 1. Triggers
  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  HEALTH_CFG.EXPECTED_TRIGGERS.forEach(name => {
    if (triggers.indexOf(name) >= 0) ok.push(`Trigger ${name}`);
    else issues.push(`❌ Thiếu trigger: ${name}`);
  });

  // 2. Tabs
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID);
  HEALTH_CFG.EXPECTED_TABS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) ok.push(`Tab "${name}" (${sh.getLastRow()} rows)`);
    else issues.push(`❌ Thiếu tab: "${name}"`);
  });

  // 3. Headers AC-AG trên tracking haravan
  const tracking = ss.getSheetByName('tracking haravan');
  if (tracking) {
    Object.keys(HEALTH_CFG.EXPECTED_HEADERS).forEach(col => {
      const expected = HEALTH_CFG.EXPECTED_HEADERS[col];
      const actual = String(tracking.getRange(1, +col).getValue() || '');
      if (actual === expected) ok.push(`Header col ${col}: ${actual}`);
      else issues.push(`❌ Header col ${col}: "${actual}" (expected "${expected}")`);
    });
  }

  // 4. Locale
  if (ss.getSpreadsheetLocale() === 'vi_VN') ok.push('Locale vi_VN');
  else issues.push(`⚠️ Locale = ${ss.getSpreadsheetLocale()} (expected vi_VN)`);

  // 5. Stale check last runs
  const props = PropertiesService.getScriptProperties();
  const checkStale = (label, propKey, thresholdHrs) => {
    const ts = props.getProperty(propKey);
    if (!ts) { issues.push(`⚠️ ${label}: chưa có last run`); return; }
    const ageHrs = (Date.now() - new Date(ts).getTime()) / 3600000;
    if (ageHrs > thresholdHrs) {
      issues.push(`⚠️ ${label}: lần chạy cuối ${ageHrs.toFixed(1)}h trước (> ${thresholdHrs}h threshold)`);
    } else {
      ok.push(`${label}: last run ${ageHrs.toFixed(1)}h trước`);
    }
  };
  checkStale('Phase 1 syncOrders', 'WELLHOME_LAST_SYNC', HEALTH_CFG.STALE_THRESHOLDS_HRS.phase1);
  checkStale('Phase 6 syncHaravanInventory', 'STOCK_LAST_SYNC', HEALTH_CFG.STALE_THRESHOLDS_HRS.phase6);
  checkStale('Phase 3 updateGhnStatuses', 'GHN_STATUS_LAST_RUN', HEALTH_CFG.STALE_THRESHOLDS_HRS.phase3);
  checkStale('Phase 4 fulfillHaravanOrders', 'HV4_LAST_RUN', HEALTH_CFG.STALE_THRESHOLDS_HRS.phase4);

  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`===== Health Monitor (${dur}s) =====`);
  Logger.log(`✅ ${ok.length} OK | ❌ ${issues.length} issues`);
  ok.forEach(s => Logger.log('  ✅ ' + s));
  issues.forEach(s => Logger.log('  ' + s));

  // Email chỉ khi có issues hoặc force mode
  if (issues.length > 0) {
    try { sendHealthAlertEmail_(issues, ok); }
    catch (e) { Logger.log(`⚠️ sendHealthAlertEmail: ${e.message}`); }
  }

  return { issues, okCount: ok.length };
}

function sendHealthAlertEmail_(issues, ok) {
  const subject = `🩺 [K-Homes] Health monitor — ${issues.length} vấn đề phát hiện`;
  const issueRows = issues.map(s => `<li style="color:#c0504d;">${s}</li>`).join('');
  const okRows = ok.slice(0, 20).map(s => `<li style="color:#5b9b3f;">✅ ${s}</li>`).join('');
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h1 style="color:#c0504d;">🩺 Health Monitor — ${issues.length} vấn đề</h1>
      <p>Auto-check phát hiện một số bất thường trong hệ thống Wellhome Order Sync. PG/admin cần kiểm tra:</p>

      <h2 style="color:#c0504d;">🔴 Vấn đề (${issues.length})</h2>
      <ul>${issueRows}</ul>

      <h3 style="color:#666;">Khắc phục thường gặp</h3>
      <ul style="font-size:13px;">
        <li><b>Thiếu trigger:</b> chạy hàm <code>setupAllTriggers</code> trong Apps Script editor (hoặc POST action <code>setup_all_triggers</code>)</li>
        <li><b>Thiếu tab/Header sai:</b> chạy <code>runIncrementalSetup</code></li>
        <li><b>Locale sai:</b> chạy <code>setupHeader</code> trong Phase 1 (auto fix locale)</li>
        <li><b>Last run stale:</b> trigger có thể bị disable do quota Google. Vào <code>script.google.com</code> → Triggers → check execution history</li>
      </ul>

      <h2 style="color:#5b9b3f;">✅ ${ok.length} mục OK</h2>
      <ul style="font-size:12px;color:#666;">${okRows}</ul>

      <p style="margin-top:15px;color:#666;font-size:12px;">
        ➤ Email tự động từ Phase 11 mỗi Chủ Nhật 23h. Chỉ gửi khi có ≥ 1 issue.
      </p>
    </div>`;
  MailApp.sendEmail({ to: HEALTH_CFG.ALERT_EMAIL, subject, htmlBody: html });
}

function manualHealthMonitor() {
  Logger.log('🟡 Manual runHealthMonitor');
  return runHealthMonitor();
}

function setupHealthMonitorTrigger() {
  removeHealthMonitorTrigger();
  ScriptApp.newTrigger('runHealthMonitor')
    .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).create();
  Logger.log('✅ Cron runHealthMonitor mỗi CN 23h');
}

function removeHealthMonitorTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runHealthMonitor') ScriptApp.deleteTrigger(t);
  });
}
