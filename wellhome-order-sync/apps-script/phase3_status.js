/**
 * PHASE 3 — UPDATE GHN STATUS  (paste vào CÙNG project Wellhome Order Sync — file thứ 3)
 * Cron 30 phút quét tab "tracking haravan", với mỗi đơn có Mã VC GHN (cột Z) chưa ở
 * trạng thái terminal (Đã giao / Hủy / Đã trả / Mất / Hỏng) → call /v2/shipping-order/detail
 * cập nhật cột U "TT giao hàng" + cột AC "GHN cập nhật lúc".
 *
 * Khi đơn delivered  → set cột X = "Đã giao"
 * Khi đơn cancel    → set cột X = "Hủy"
 * Khi đơn lost/damage→ set cột X = "Lỗi GHN" + note cột Y
 *
 * Tận dụng GHN_CFG, GHN_COL, ghnFetch_, openGhnTargetTab_, getGhnProp_, formatGhnNow_
 * đã khai báo trong Phase 2 (cùng project).
 *
 * Public functions:
 *   setupGhnStatusExtHeader   — thêm cột AC "GHN cập nhật lúc"
 *   setupGhnStatusAll         — gộp setup
 *   updateGhnStatuses         — hàm chính (cron 30 phút)
 *   manualUpdateOneOrder("GYWW9VAP") — debug 1 đơn
 *   showGhnStatusReport       — count theo từng status hiện có
 *   setupGhnStatusTrigger     — cài cron 30 phút
 *   removeGhnStatusTrigger    — xóa trigger
 */

const GHN_STATUS_CFG = {
  PROP_LAST_RUN: 'GHN_STATUS_LAST_RUN',
  PROP_LAST_ALERT: 'GHN_STATUS_LAST_ALERT',
  MAX_ORDERS_PER_RUN: 100,
  REFRESH_TERMINAL_AFTER_DAYS: 0,   // 0 = không refresh đơn đã terminal
  SLEEP_MS: 250,                     // delay giữa các call GHN tránh rate limit
  FAIL_THRESHOLD: 3,                 // ≥ 3 lần giao fail → cảnh báo PG
  ALERT_EMAIL: 'admin@khomes.com.vn',
  ALERT_COOLDOWN_HRS: 6,             // Cảnh báo lại sau 6h (tránh spam mail)
};

const GHN_STATUS_COL = {
  // CRITICAL — column indices MUST NOT conflict với Phase 4 (AD=30 FULFILL_AT) hay Phase 5 (AG=33 PRINTED_AT):
  GHN_SYNC_AT: 29,    // AC — GHN cập nhật lúc
  GHN_FAIL_COUNT: 31, // AE — Số lần giao fail
  GHN_ALERTED_AT: 32, // AF — Đã alert cảnh báo lúc (tránh spam)
};

// Các trạng thái không cần re-poll (đơn đã chốt cuối cùng)
const GHN_TERMINAL_LABELS = ['Đã giao', 'Hủy', 'Đã trả', 'Mất', 'Hỏng'];

// Mapping GHN status → label tiếng Việt thân thiện
const GHN_STATUS_MAP = {
  ready_to_pick:           'Chờ lấy hàng',
  picking:                 'Đang lấy hàng',
  money_collect_picking:   'Thu tiền (lấy)',
  picked:                  'Đã lấy hàng',
  storing:                 'Đang ở kho',
  transporting:            'Đang luân chuyển',
  sorting:                 'Đang phân loại',
  delivering:              'Đang giao',
  money_collect_delivering:'Đang giao + COD',
  delivered:               'Đã giao',
  delivery_fail:           'Giao thất bại',
  waiting_to_return:       'Chờ trả hàng',
  return:                  'Đang trả',
  return_transporting:     'Luân chuyển trả',
  return_sorting:          'Phân loại trả',
  returning:               'Đang đi trả',
  return_fail:             'Trả thất bại',
  returned:                'Đã trả',
  cancel:                  'Hủy',
  exception:               'Ngoại lệ',
  damage:                  'Hỏng',
  lost:                    'Mất',
};

// ============================================================
// SETUP
// ============================================================

function setupGhnStatusExtHeader() {
  const sheet = openGhnTargetTab_();
  // Cột AC — GHN cập nhật lúc
  sheet.getRange(1, GHN_STATUS_COL.GHN_SYNC_AT)
    .setValue('GHN cập nhật lúc')
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setColumnWidth(GHN_STATUS_COL.GHN_SYNC_AT, 140);

  // Cột AE — Số lần giao fail
  sheet.getRange(1, GHN_STATUS_COL.GHN_FAIL_COUNT)
    .setValue('Số lần giao fail')
    .setFontWeight('bold')
    .setBackground('#c0504d')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setColumnWidth(GHN_STATUS_COL.GHN_FAIL_COUNT, 110);

  // Cột AF — Đã alert lúc
  sheet.getRange(1, GHN_STATUS_COL.GHN_ALERTED_AT)
    .setValue('Đã alert lúc')
    .setFontWeight('bold')
    .setBackground('#c0504d')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setColumnWidth(GHN_STATUS_COL.GHN_ALERTED_AT, 140);

  // Conditional format cột AE: ô đỏ khi count >= FAIL_THRESHOLD
  const N = Math.max(sheet.getLastRow(), 1000);
  const colE = sheet.getRange(2, GHN_STATUS_COL.GHN_FAIL_COUNT, N - 1, 1);
  const rules = sheet.getConditionalFormatRules();
  const filtered = rules.filter(r => {
    const ranges = r.getRanges();
    return !ranges.some(rg => rg.getColumn() === GHN_STATUS_COL.GHN_FAIL_COUNT);
  });
  filtered.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(GHN_STATUS_CFG.FAIL_THRESHOLD)
    .setBackground('#F8CECC')
    .setFontColor('#990000')
    .setBold(true)
    .setRanges([colE]).build());
  sheet.setConditionalFormatRules(filtered);

  Logger.log('✅ Đã thêm cột AC "GHN cập nhật lúc" + AE "Số lần giao fail" + AF "Đã alert lúc"');
}

function updateGhnDeliveryStatusDropdown() {
  // Cập nhật dropdown cột U (TT giao hàng) với toàn bộ label GHN có thể trả về
  const sheet = openGhnTargetTab_();
  const N = Math.max(sheet.getLastRow(), 1000);
  const labels = Array.from(new Set(Object.values(GHN_STATUS_MAP)));
  sheet.getRange(2, GHN_COL.TT_GIAO, N - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(labels, true)
      .setAllowInvalid(true)
      .build()
  );

  // Conditional format: terminal xanh, fail đỏ
  const rules = sheet.getConditionalFormatRules();
  const colU = sheet.getRange(`U2:U${N}`);
  // Bỏ rule cũ trên cột U (giữ rule cột khác)
  const filtered = rules.filter(r => {
    const ranges = r.getRanges();
    return !ranges.some(rg => rg.getColumn() === GHN_COL.TT_GIAO);
  });
  filtered.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Đã giao')
    .setBackground('#D5E8D4')
    .setRanges([colU]).build());
  filtered.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Hủy')
    .setBackground('#F8CECC')
    .setRanges([colU]).build());
  filtered.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Giao thất bại')
    .setBackground('#FFE6CC')
    .setRanges([colU]).build());
  filtered.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Đã trả')
    .setBackground('#FFF2CC')
    .setRanges([colU]).build());
  sheet.setConditionalFormatRules(filtered);

  Logger.log(`✅ Dropdown cột U cập nhật ${labels.length} giá trị + conditional format`);
}

function setupGhnStatusAll() {
  Logger.log('===== Bắt đầu setup Phase 3 GHN Status =====');
  setupGhnStatusExtHeader();
  updateGhnDeliveryStatusDropdown();
  Logger.log('===== ✅ Setup xong. Test bằng updateGhnStatuses, sau đó setupGhnStatusTrigger =====');
}

function setupGhnStatusTrigger() {
  removeGhnStatusTrigger();
  ScriptApp.newTrigger('updateGhnStatuses').timeBased().everyMinutes(30).create();
  Logger.log('✅ Đã cài trigger: updateGhnStatuses chạy mỗi 30 phút');
}

function removeGhnStatusTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'updateGhnStatuses') ScriptApp.deleteTrigger(t);
  });
  Logger.log('✅ Đã xóa trigger cũ của updateGhnStatuses');
}

// ============================================================
// MAIN — hàm cron chính
// ============================================================

function updateGhnStatuses() {
  const t0 = Date.now();
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng, skip'); return; }

  // Range mở rộng tới cột AC để đọc cả Mã VC + sync timestamp
  const lastCol = Math.max(GHN_STATUS_COL.GHN_SYNC_AT, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Group theo Mã VC (1 đơn Haravan có thể có nhiều dòng nhưng cùng Mã VC)
  const seen = new Set();
  const targets = [];   // { ghnCode, rows: [rowIndex...] }

  data.forEach((row, idx) => {
    const ghnCode = String(row[GHN_COL.GHN_MA - 1] || '').trim();
    if (!ghnCode) return;
    const curStatus = String(row[GHN_COL.TT_GIAO - 1] || '').trim();
    if (GHN_TERMINAL_LABELS.indexOf(curStatus) >= 0) return;

    if (!seen.has(ghnCode)) {
      seen.add(ghnCode);
      targets.push({ ghnCode, rows: [idx + 2] });
    } else {
      // Tìm trong targets, append rowIndex
      const t = targets.find(x => x.ghnCode === ghnCode);
      if (t) t.rows.push(idx + 2);
    }
  });

  Logger.log(`Tổng đơn cần update: ${targets.length} (đơn có Mã VC, chưa terminal)`);
  if (targets.length === 0) {
    PropertiesService.getScriptProperties().setProperty(
      GHN_STATUS_CFG.PROP_LAST_RUN, new Date().toISOString());
    return;
  }

  let okCount = 0, failCount = 0;
  const toAlert = [];   // đơn cần email cảnh báo (fail >= threshold, chưa alert trong cooldown)
  const now = formatGhnNow_();
  const nowMs = Date.now();
  const cooldownMs = GHN_STATUS_CFG.ALERT_COOLDOWN_HRS * 3600 * 1000;

  for (let i = 0; i < Math.min(targets.length, GHN_STATUS_CFG.MAX_ORDERS_PER_RUN); i++) {
    const t = targets[i];
    try {
      const res = ghnFetch_('/v2/shipping-order/detail', { order_code: t.ghnCode });
      if (!res.ok) {
        failCount++;
        Logger.log(`❌ ${t.ghnCode}: ${res.error}`);
        // Đơn không tồn tại bên GHN → ghi note 1 lần, không update label
        const isNotFound = /not found|không tồn tại|404/i.test(res.error);
        if (isNotFound) {
          t.rows.forEach(r => {
            sheet.getRange(r, GHN_COL.GHI_CHU_NV).setValue(
              appendNote_(sheet.getRange(r, GHN_COL.GHI_CHU_NV).getValue(),
                `GHN không tìm thấy ${t.ghnCode}`));
            sheet.getRange(r, GHN_STATUS_COL.GHN_SYNC_AT).setValue(now);
          });
        }
        Utilities.sleep(GHN_STATUS_CFG.SLEEP_MS);
        continue;
      }
      const ghnStatus = String(res.data.status || '').toLowerCase();
      const friendly = GHN_STATUS_MAP[ghnStatus] || ghnStatus || '(unknown)';
      const deliveryFails = countDeliveryFails_(res.data.log);

      t.rows.forEach(r => {
        sheet.getRange(r, GHN_COL.TT_GIAO).setValue(friendly);
        sheet.getRange(r, GHN_STATUS_COL.GHN_SYNC_AT).setValue(now);
        sheet.getRange(r, GHN_STATUS_COL.GHN_FAIL_COUNT).setValue(deliveryFails);

        // Đồng bộ cột X (Trạng thái xử lý) khi đơn ở terminal state
        if (ghnStatus === 'delivered') {
          trySetCellValue_(sheet, r, GHN_COL.TT_XU_LY, 'Đã giao');
        } else if (ghnStatus === 'cancel' || ghnStatus === 'returned') {
          trySetCellValue_(sheet, r, GHN_COL.TT_XU_LY, 'Hủy');
        } else if (ghnStatus === 'lost' || ghnStatus === 'damage') {
          trySetCellValue_(sheet, r, GHN_COL.TT_XU_LY, 'Lỗi GHN');
          sheet.getRange(r, GHN_COL.GHI_CHU_NV).setValue(
            appendNote_(sheet.getRange(r, GHN_COL.GHI_CHU_NV).getValue(),
              `GHN báo ${friendly} (${t.ghnCode})`));
        }
      });

      // Collect đơn cần alert: fail >= threshold + chưa terminal + cooldown OK
      if (deliveryFails >= GHN_STATUS_CFG.FAIL_THRESHOLD &&
          GHN_TERMINAL_LABELS.indexOf(friendly) < 0) {
        const firstRow = t.rows[0];
        const lastAlert = sheet.getRange(firstRow, GHN_STATUS_COL.GHN_ALERTED_AT).getValue();
        const lastAlertMs = lastAlert instanceof Date ? lastAlert.getTime()
          : (lastAlert ? Date.parse(lastAlert) || 0 : 0);
        if (nowMs - lastAlertMs > cooldownMs) {
          toAlert.push({
            ghnCode: t.ghnCode,
            row: firstRow,
            failCount: deliveryFails,
            status: friendly,
            tenKH: sheet.getRange(firstRow, GHN_COL.TEN_KH).getValue(),
            sdt: sheet.getRange(firstRow, GHN_COL.SDT).getValue(),
            tinh: sheet.getRange(firstRow, GHN_COL.TINH).getValue(),
            tenSP: sheet.getRange(firstRow, GHN_COL.TEN_SP).getValue(),
            maHaravan: sheet.getRange(firstRow, GHN_COL.MA_HARAVAN).getValue(),
          });
        }
      }
      okCount++;
    } catch (e) {
      failCount++;
      Logger.log(`❌ ${t.ghnCode} exception: ${e.message}`);
    }
    Utilities.sleep(GHN_STATUS_CFG.SLEEP_MS);
  }

  PropertiesService.getScriptProperties().setProperty(
    GHN_STATUS_CFG.PROP_LAST_RUN, new Date().toISOString());

  // Gửi email cảnh báo nếu có đơn fail >= threshold
  if (toAlert.length > 0) {
    try {
      sendFailureAlertEmail_(toAlert);
      // Mark alerted_at để cooldown
      toAlert.forEach(a => {
        sheet.getRange(a.row, GHN_STATUS_COL.GHN_ALERTED_AT).setValue(now);
      });
      Logger.log(`📧 Đã gửi alert ${toAlert.length} đơn giao fail ≥ ${GHN_STATUS_CFG.FAIL_THRESHOLD} lần`);
    } catch (e) {
      Logger.log(`⚠️ Lỗi gửi mail alert: ${e.message}`);
    }
  }

  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`✅ Done ${dur}s. OK: ${okCount}, Lỗi: ${failCount}, Alert: ${toAlert.length}, Skip terminal: ${data.length - targets.length - countNoCode_(data)}`);
}

/**
 * Đếm số lần delivery_fail trong log history GHN.
 * res.data.log là array các status change. Mỗi entry { status, updated_date }.
 */
function countDeliveryFails_(log) {
  if (!log || !Array.isArray(log)) return 0;
  let count = 0;
  log.forEach(entry => {
    const s = String(entry.status || '').toLowerCase();
    if (s === 'delivery_fail') count++;
  });
  return count;
}

/**
 * Gửi email tổng hợp cảnh báo PG về các đơn giao thất bại nhiều lần.
 * Nội dung HTML table — dễ đọc trên mobile.
 */
function sendFailureAlertEmail_(orders) {
  const subject = `🚨 [K-Homes] ${orders.length} đơn giao thất bại ≥ ${GHN_STATUS_CFG.FAIL_THRESHOLD} lần — cần liên hệ KH`;
  const rows = orders.map((o, i) => `
    <tr>
      <td>${i + 1}</td>
      <td><b>${o.ghnCode}</b></td>
      <td>${o.maHaravan || ''}</td>
      <td>${o.tenKH || ''}</td>
      <td><a href="tel:${o.sdt}">${o.sdt || ''}</a></td>
      <td>${o.tinh || ''}</td>
      <td>${o.tenSP || ''}</td>
      <td style="color:#c0504d;font-weight:bold;text-align:center;">${o.failCount}</td>
      <td>${o.status}</td>
    </tr>`).join('');

  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
      <h2 style="color:#c0504d;">🚨 Cảnh báo đơn giao thất bại ≥ ${GHN_STATUS_CFG.FAIL_THRESHOLD} lần</h2>
      <p>Các đơn dưới đây đã <b>giao thất bại từ ${GHN_STATUS_CFG.FAIL_THRESHOLD} lần trở lên</b>. PG cần <b>gọi điện trực tiếp KH</b> để xác nhận lại địa chỉ / hẹn giao lại / hỏi vướng mắc, tránh đơn bị hoàn về kho.</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;">
        <tr style="background:#1f4e78;color:#fff;">
          <th>#</th><th>Mã VC GHN</th><th>Mã Haravan</th><th>Tên KH</th><th>SĐT</th>
          <th>Tỉnh</th><th>Sản phẩm</th><th>Số lần fail</th><th>Trạng thái GHN</th>
        </tr>
        ${rows}
      </table>
      <p style="margin-top:15px;color:#666;">
        ➤ Email tự động gửi từ Phase 3 Wellhome Order Sync. Cảnh báo lại sau ${GHN_STATUS_CFG.ALERT_COOLDOWN_HRS}h nếu đơn vẫn chưa được xử lý.<br>
        ➤ Sau khi gọi KH, cập nhật cột Y "Ghi chú" trong sheet để team biết.
      </p>
    </div>`;

  MailApp.sendEmail({
    to: GHN_STATUS_CFG.ALERT_EMAIL,
    subject: subject,
    htmlBody: html,
  });
  PropertiesService.getScriptProperties().setProperty(
    GHN_STATUS_CFG.PROP_LAST_ALERT, new Date().toISOString());
}

/**
 * Test gửi mail alert tay (debug). Gọi tay từ Apps Script editor.
 * Quét toàn bộ tab tracking haravan, tìm đơn có cột AE >= threshold + chưa terminal,
 * gửi mail bất kể cooldown.
 */
function manualSendFailureAlert() {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng'); return; }
  const lastCol = Math.max(GHN_STATUS_COL.GHN_FAIL_COUNT, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const orders = [];
  data.forEach((row, idx) => {
    const failCount = Number(row[GHN_STATUS_COL.GHN_FAIL_COUNT - 1] || 0);
    if (failCount < GHN_STATUS_CFG.FAIL_THRESHOLD) return;
    const status = String(row[GHN_COL.TT_GIAO - 1] || '').trim();
    if (GHN_TERMINAL_LABELS.indexOf(status) >= 0) return;
    orders.push({
      ghnCode:   row[GHN_COL.GHN_MA - 1],
      row:       idx + 2,
      failCount: failCount,
      status:    status,
      tenKH:     row[GHN_COL.TEN_KH - 1],
      sdt:       row[GHN_COL.SDT - 1],
      tinh:      row[GHN_COL.TINH - 1],
      tenSP:     row[GHN_COL.TEN_SP - 1],
      maHaravan: row[GHN_COL.MA_HARAVAN - 1],
    });
  });
  if (orders.length === 0) { Logger.log('Không có đơn nào fail ≥ threshold'); return; }
  Logger.log(`📧 Test gửi alert ${orders.length} đơn`);
  sendFailureAlertEmail_(orders);
}

function countNoCode_(data) {
  let n = 0;
  data.forEach(r => { if (!String(r[GHN_COL.GHN_MA - 1] || '').trim()) n++; });
  return n;
}

function trySetCellValue_(sheet, row, col, value) {
  try { sheet.getRange(row, col).setValue(value); }
  catch (e) {
    Logger.log(`⚠️ Không set được [${row},${col}]="${value}": ${e.message}`);
  }
}

function appendNote_(cur, add) {
  const s = String(cur || '').trim();
  if (!s) return add;
  if (s.indexOf(add) >= 0) return s;   // tránh duplicate note
  return s + ' | ' + add;
}

// ============================================================
// DEBUG / UTILITY
// ============================================================

function manualUpdateOneOrder(ghnOrderCode) {
  if (!ghnOrderCode) { Logger.log('Truyền vào Mã VC GHN. Vd: manualUpdateOneOrder("GYWW9VAP")'); return; }
  const res = ghnFetch_('/v2/shipping-order/detail', { order_code: ghnOrderCode });
  if (!res.ok) { Logger.log(`❌ ${res.error}`); return; }
  const ghnStatus = String(res.data.status || '');
  const friendly = GHN_STATUS_MAP[ghnStatus.toLowerCase()] || ghnStatus;
  Logger.log(`Mã VC: ${ghnOrderCode}`);
  Logger.log(`Trạng thái GHN: ${ghnStatus} → "${friendly}"`);
  Logger.log(`COD: ${res.data.cod_amount || 0} đ — đã thu: ${res.data.cod_collect_date || 'chưa'}`);
  Logger.log(`Phí ship: ${res.data.total_fee || 0} đ`);
  Logger.log(`Cập nhật cuối: ${res.data.updated_date || 'N/A'}`);
  if (res.data.log && res.data.log.length) {
    Logger.log(`History (${res.data.log.length} bước):`);
    res.data.log.slice(-5).forEach(l => {
      Logger.log(`  ${l.updated_date} — ${l.status}`);
    });
  }
}

// ============================================================
// 🟡 DRY RUN — chỉ đếm đơn sẽ poll, KHÔNG call API GHN
// ============================================================

function dryRunUpdateGhnStatuses() {
  Logger.log('🟡 DRY RUN — KHÔNG call GHN API, chỉ thống kê');
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng'); return; }
  const lastCol = Math.max(GHN_STATUS_COL.GHN_SYNC_AT, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const seen = new Set();
  const targets = [];
  let noCode = 0, terminal = 0;
  data.forEach(row => {
    const ghnCode = String(row[GHN_COL.GHN_MA - 1] || '').trim();
    if (!ghnCode) { noCode++; return; }
    const curStatus = String(row[GHN_COL.TT_GIAO - 1] || '').trim();
    if (GHN_TERMINAL_LABELS.indexOf(curStatus) >= 0) { terminal++; return; }
    if (!seen.has(ghnCode)) { seen.add(ghnCode); targets.push({ ghnCode, status: curStatus || '(trống)' }); }
  });
  Logger.log(`Tổng dòng: ${data.length} | Không có Mã VC: ${noCode} | Terminal (skip): ${terminal} | Cần poll: ${targets.length}`);
  targets.slice(0, 30).forEach((t, i) => {
    Logger.log(`  ${i + 1}. ${t.ghnCode} (cột U hiện tại: ${t.status})`);
  });
  if (targets.length > 30) Logger.log(`  ... còn ${targets.length - 30} đơn nữa`);
  Logger.log(`\n✅ DRY RUN OK — chạy updateGhnStatuses để poll thật. Mỗi đơn 1 API call GHN.`);
}

function showGhnStatusReport() {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng'); return; }
  const data = sheet.getRange(2, GHN_COL.TT_GIAO, lastRow - 1, 1).getValues();
  const counts = {};
  data.forEach(r => {
    const v = String(r[0] || '').trim() || '(trống)';
    counts[v] = (counts[v] || 0) + 1;
  });
  Logger.log('===== Phân bố trạng thái cột U (TT giao hàng) =====');
  Object.keys(counts).sort((a, b) => counts[b] - counts[a]).forEach(k => {
    Logger.log(`${k}: ${counts[k]}`);
  });
  const last = getGhnProp_(GHN_STATUS_CFG.PROP_LAST_RUN);
  Logger.log(`Lần chạy cuối: ${last || '(chưa chạy)'}`);
}
