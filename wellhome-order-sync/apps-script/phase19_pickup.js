/**
 * PHASE 19 — Pickup tại showroom workflow
 *
 * Trước Phase 19: đơn "Pickup tại SR" skip GHN, không ai báo PG showroom HN.
 *
 * Phase 19:
 *   - notifyNewPickupOrders cron 9h sáng: email PG list đơn pickup mới chưa noti
 *   - fulfillPickupOrders cron 1h cùng Phase 4: đơn Pickup + TT giao "Đã giao" + AD trống
 *     → fulfill Haravan với note "Pickup tại showroom" (KHÔNG cần Mã VC GHN)
 *
 * Cột mới:
 *   AO (41) — Pickup notified at
 *
 * PG mark "Đã giao" cho đơn pickup: đổi cột U (TT giao) trong sheet thành "Đã giao"
 * (dropdown đã có sẵn).
 */

const PICKUP_CFG = {
  COL_PICKUP_NOTIFIED: 41,   // AO
  EXT_HEADER: 'Pickup notified',
  PG_SR_EMAIL: 'admin@khomes.com.vn',  // TODO: đổi sang email PG SR HN khi có
};

function setupPickupHeader_() {
  const sheet = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID).getSheetByName(GHN_CFG.TARGET_TAB);
  const curCols = sheet.getMaxColumns();
  if (curCols < PICKUP_CFG.COL_PICKUP_NOTIFIED) {
    sheet.insertColumnsAfter(curCols, PICKUP_CFG.COL_PICKUP_NOTIFIED - curCols);
  }
  sheet.getRange(1, PICKUP_CFG.COL_PICKUP_NOTIFIED).setValue(PICKUP_CFG.EXT_HEADER)
    .setFontWeight('bold').setBackground('#ead1dc');
  Logger.log('✅ Setup AO header Phase 19');
}

function notifyNewPickupOrders() {
  setupPickupHeader_();
  const ss = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(GHN_CFG.TARGET_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const data = sheet.getRange(2, 1, lastRow - 1, PICKUP_CFG.COL_PICKUP_NOTIFIED).getValues();
  const newPickups = [];
  data.forEach(function (row, idx) {
    const ttXuLy = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    const notified = row[PICKUP_CFG.COL_PICKUP_NOTIFIED - 1];
    if (ttXuLy === 'Pickup tại SR' && !notified) {
      newPickups.push({
        rowIdx: idx + 2,
        haravan: row[GHN_COL.MA_HARAVAN - 1],
        tenKh: row[GHN_COL.TEN_KH - 1],
        sdt: row[GHN_COL.SDT - 1],
        sku: row[GHN_COL.MA_SP - 1],
        tenSp: row[GHN_COL.TEN_SP - 1],
        sl: row[GHN_COL.SL - 1],
        tong: row[GHN_COL.TONG_DON - 1],
      });
    }
  });

  if (!newPickups.length) {
    Logger.log('No new pickup orders');
    return { ok: true, count: 0 };
  }

  // Email PG SR
  const today = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy');
  let html = '<h2 style="color:#9900ff">🏪 Đơn Pickup tại showroom — ' + today + '</h2>';
  html += '<p>Có <b>' + newPickups.length + '</b> đơn KH chọn nhận tại showroom. Vui lòng:';
  html += '<ol><li>Chuẩn bị hàng theo SKU + SL</li>';
  html += '<li>Khi KH đến lấy → mở Sheet → đổi cột "TT giao" thành <b>"Đã giao"</b></li>';
  html += '<li>Tool sẽ tự fulfill Haravan trong 30 phút sau đó</li></ol></p>';
  html += '<table border="1" cellpadding="6" style="border-collapse:collapse;font-family:Arial">';
  html += '<tr style="background:#cfe2f3"><th>Mã đơn</th><th>KH</th><th>SĐT</th><th>SKU</th><th>SP</th><th>SL</th><th>Tổng</th></tr>';
  newPickups.forEach(function (p) {
    html += '<tr><td>' + p.haravan + '</td><td>' + p.tenKh + '</td><td>' + p.sdt + '</td><td>' +
            p.sku + '</td><td>' + (p.tenSp || '').slice(0, 50) + '</td><td>' + p.sl + '</td><td>' +
            (p.tong || 0).toLocaleString('vi-VN') + 'đ</td></tr>';
  });
  html += '</table>';

  MailApp.sendEmail({
    to: PICKUP_CFG.PG_SR_EMAIL,
    subject: '[K-Homes] 🏪 ' + newPickups.length + ' đơn Pickup cần chuẩn bị — ' + today,
    htmlBody: html,
  });

  // Mark notified
  newPickups.forEach(function (p) {
    sheet.getRange(p.rowIdx, PICKUP_CFG.COL_PICKUP_NOTIFIED).setValue(new Date());
  });

  Logger.log('✅ Notified ' + newPickups.length + ' pickup orders');
  return { ok: true, count: newPickups.length };
}

function fulfillPickupOrders() {
  const ss = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(GHN_CFG.TARGET_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const data = sheet.getRange(2, 1, lastRow - 1, GHN_COL.GHN_AT).getValues();
  // HV4_FULFILL_AT cột 30 (AD) — defined in phase4
  const HV4_FULFILL_AT_COL = 30;
  const allCols = sheet.getRange(2, 1, lastRow - 1, HV4_FULFILL_AT_COL).getValues();
  const candidates = [];
  allCols.forEach(function (row, idx) {
    const ttXuLy = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    const ttGiao = String(row[GHN_COL.TT_GIAO - 1] || '').trim();
    const fulfilled = row[HV4_FULFILL_AT_COL - 1];
    if (ttXuLy === 'Pickup tại SR' && ttGiao === 'Đã giao' && !fulfilled) {
      candidates.push({
        rowIdx: idx + 2,
        haravan: String(row[GHN_COL.MA_HARAVAN - 1] || '').trim(),
      });
    }
  });

  if (!candidates.length) return { ok: true, count: 0 };

  let okCount = 0, failCount = 0;
  candidates.forEach(function (c) {
    if (!c.haravan) return;
    try {
      // Tận dụng logic Phase 4: hv4LookupOrderByName_ + post fulfillment với tracking_number "PICKUP-SR"
      const lookup = hv4LookupOrderByName_(c.haravan);
      if (!lookup || !lookup.order) {
        sheet.getRange(c.rowIdx, HV4_FULFILL_AT_COL).setValue('LỖI: lookup fail');
        failCount++;
        return;
      }
      if (lookup.order.fulfillment_status === 'fulfilled') {
        sheet.getRange(c.rowIdx, HV4_FULFILL_AT_COL).setValue(new Date());
        appendNote_(sheet, c.rowIdx, 'Pickup: đã fulfilled trên Haravan trước');
        okCount++;
        return;
      }
      const lineItems = (lookup.order.line_items || []).map(function (li) {
        return { id: li.id, quantity: li.quantity };
      });
      const url = 'https://' + CONFIG.HARAVAN_SHOP + '/admin/orders/' + lookup.order.id + '/fulfillments.json';
      const tokenH = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_TOKEN);
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + tokenH },
        payload: JSON.stringify({
          fulfillment: {
            tracking_number: 'PICKUP-SR-' + c.haravan,
            tracking_company: 'Pickup tại showroom',
            notify_customer: false,
            line_items: lineItems,
          },
        }),
        muteHttpExceptions: true,
      });
      const code = res.getResponseCode();
      if (code >= 200 && code < 300) {
        sheet.getRange(c.rowIdx, HV4_FULFILL_AT_COL).setValue(new Date());
        okCount++;
      } else {
        const body = res.getContentText().slice(0, 500);
        sheet.getRange(c.rowIdx, HV4_FULFILL_AT_COL).setValue('LỖI ' + code);
        appendNote_(sheet, c.rowIdx, 'Pickup fulfill fail: ' + body.slice(0, 100));
        failCount++;
      }
      Utilities.sleep(350);
    } catch (err) {
      sheet.getRange(c.rowIdx, HV4_FULFILL_AT_COL).setValue('LỖI: ' + err.message.slice(0, 80));
      failCount++;
    }
  });

  Logger.log('Pickup fulfill: ok=' + okCount + ' fail=' + failCount);
  return { ok: true, count: candidates.length, fulfilled: okCount, failed: failCount };
}

function setupPickupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const name = t.getHandlerFunction();
    if (name === 'notifyNewPickupOrders' || name === 'fulfillPickupOrders') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('notifyNewPickupOrders').timeBased().atHour(9).everyDays(1).create();
  ScriptApp.newTrigger('fulfillPickupOrders').timeBased().everyHours(1).create();
  Logger.log('✅ Trigger Phase 19: notify 9h + fulfill 1h');
}
