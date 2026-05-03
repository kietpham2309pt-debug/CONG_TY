/**
 * PHASE 4 — FULFILL HARAVAN  (paste vào CÙNG project Wellhome Order Sync — file thứ 4)
 * Cron 30 phút quét tab "tracking haravan", với mỗi đơn có Mã VC GHN (cột Z) + chưa fulfill
 * (cột AD trống) + chưa Hủy → call Haravan POST /admin/orders/<id>/fulfillments.json
 * gắn tracking_number = Mã VC GHN, tracking_company = "GHN" (viết tắt, tránh auto-link Haravan-GHN integration).
 *
 * Lý do: KH mua trên wellhome.asia → cần thấy mã vận đơn trong tài khoản của họ.
 *
 * Tận dụng:
 *   CONFIG.HARAVAN_SHOP, CONFIG.PROP_TOKEN  — Phase 1
 *   GHN_COL                                  — Phase 2
 *   openGhnTargetTab_, formatGhnNow_        — Phase 2
 *
 * Public functions:
 *   setupHv4ExtHeader        — thêm cột AD "Fulfill Haravan lúc"
 *   testHv4Connection        — verify token có quyền write_fulfillments
 *   fulfillHaravanOrders     — hàm chính (cron 30 phút)
 *   manualFulfillOne("RT12510680") — fulfill tay 1 đơn (debug)
 *   resetHv4FulfillRow(rowNumber)   — xóa cột AD của 1 dòng để retry
 *   showHv4Report            — đếm số đơn đã/chưa fulfill
 *   setupHv4Trigger          — cài cron 30 phút
 *   removeHv4Trigger         — xóa trigger
 */

const HV4_CFG = {
  // Để 'GHN' (viết tắt) thay vì 'Giao Hàng Nhanh' để Haravan KHÔNG auto-link sang GHN
  // shop integration của Tefal (5146557) — gây 422 "chưa được cấp quyền". Chỉ ghi text.
  TRACKING_COMPANY: 'GHN',
  NOTIFY_CUSTOMER: false,
  MAX_ORDERS_PER_RUN: 50,
  SLEEP_MS: 350,
  PROP_LAST_RUN: 'HV4_LAST_RUN',
};

const HV4_COL = {
  FULFILL_AT: 30,  // AD — Fulfill Haravan lúc (timestamp)
};

const HV4_TERMINAL_X = ['Hủy', 'Lỗi GHN', 'Pickup tại SR', 'Thiếu hàng'];   // không fulfill các đơn này

// ============================================================
// SETUP
// ============================================================

function setupHv4ExtHeader() {
  const sheet = openGhnTargetTab_();
  sheet.getRange(1, HV4_COL.FULFILL_AT)
    .setValue('Fulfill Haravan lúc')
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setColumnWidth(HV4_COL.FULFILL_AT, 140);
  Logger.log('✅ Đã thêm cột AD "Fulfill Haravan lúc"');
}

function testHv4Connection() {
  // Lookup 1 đơn bất kỳ + dry-fetch fulfillments endpoint (GET) — không write
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng — không có đơn để test'); return; }
  // Đọc cột A (Ngày đặt) và cột B (Mã đơn) cùng lúc
  const sample = sheet.getRange(2, 1, Math.min(20, lastRow - 1), 2).getValues()
    .map(r => ({ createdAt: r[0], name: String(r[1] || '').trim() }))
    .filter(x => x.name);
  if (sample.length === 0) { Logger.log('Không có Mã đơn Haravan'); return; }

  const { name, createdAt } = sample[0];
  const lookup = hv4LookupOrderByName_(name, createdAt);
  if (!lookup.ok) { Logger.log(`❌ Lookup ${name}: ${lookup.error}`); return; }
  Logger.log(`✅ Lookup OK — đơn ${name}: id=${lookup.order.id}, items=${(lookup.order.line_items || []).length}`);

  // (1) Test GET fulfillments — verify scope read_fulfillments
  const urlGet = `https://${CONFIG.HARAVAN_SHOP}/admin/orders/${lookup.order.id}/fulfillments.json`;
  const resGet = UrlFetchApp.fetch(urlGet, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + getHv4Token_() },
    muteHttpExceptions: true,
  });
  const codeGet = resGet.getResponseCode();
  if (codeGet === 200) {
    const arr = (JSON.parse(resGet.getContentText()).fulfillments || []);
    Logger.log(`✅ GET fulfillments OK — read_fulfillments có quyền (${arr.length} fulfillment hiện có)`);
  } else {
    Logger.log(`❌ GET HTTP ${codeGet} — thiếu scope read_fulfillments. Body: ${resGet.getContentText().slice(0, 200)}`);
    return;
  }

  // (2) KHÔNG còn test POST fulfillments với body trống (đã từng vô tình tạo fulfillment trên
  // đơn sample, gây nhầm lẫn cho KH). Cách verify scope write_orders an toàn nhất là chạy
  // `manualFulfillOne("<orderName>")` trên 1 đơn thật cụ thể — nếu HTTP 200/201 = scope OK.
  Logger.log(`ℹ️ Để verify scope com.write_orders, chạy 'manualFulfillOne("<Mã đơn>")' với 1 đơn thật.`);
  Logger.log(`   Nếu trả về "✅ Đã fulfill" → scope OK. Nếu HTTP 403 → vào Haravan Partners bật "Đơn hàng = Đọc và ghi" + re-install app.`);
}

function setupHv4Trigger() {
  removeHv4Trigger();
  ScriptApp.newTrigger('fulfillHaravanOrders').timeBased().everyMinutes(30).create();
  Logger.log('✅ Đã cài trigger: fulfillHaravanOrders chạy mỗi 30 phút');
}

function removeHv4Trigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'fulfillHaravanOrders') ScriptApp.deleteTrigger(t);
  });
  Logger.log('✅ Đã xóa trigger cũ của fulfillHaravanOrders');
}

// ============================================================
// MAIN — hàm cron chính
// ============================================================

function fulfillHaravanOrders() {
  const t0 = Date.now();
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng, skip'); return; }

  const lastCol = Math.max(HV4_COL.FULFILL_AT, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Group theo Mã đơn Haravan (cột B). Một đơn có thể nhiều dòng line item, fulfill 1 lần thôi.
  const groups = {};
  data.forEach((row, idx) => {
    const orderName = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim();
    if (!orderName) return;
    const ghnCode = String(row[GHN_COL.GHN_MA - 1] || '').trim();
    if (!ghnCode) return;
    const fulfilledAt = String(row[HV4_COL.FULFILL_AT - 1] || '').trim();
    if (fulfilledAt) return;
    const xStatus = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    if (HV4_TERMINAL_X.indexOf(xStatus) >= 0) return;

    if (!groups[orderName]) {
      groups[orderName] = { orderName, ghnCode, createdAt: row[GHN_COL.NGAY_DAT - 1], rows: [] };
    }
    groups[orderName].rows.push(idx + 2);
  });

  const orderNames = Object.keys(groups);
  Logger.log(`Tổng đơn cần fulfill: ${orderNames.length}`);
  if (orderNames.length === 0) {
    PropertiesService.getScriptProperties().setProperty(
      HV4_CFG.PROP_LAST_RUN, new Date().toISOString());
    return;
  }

  let okCount = 0, failCount = 0, skipCount = 0;
  const now = formatGhnNow_();

  for (let i = 0; i < Math.min(orderNames.length, HV4_CFG.MAX_ORDERS_PER_RUN); i++) {
    const g = groups[orderNames[i]];
    try {
      // Step 1: lookup order by name + date → get id + line_items
      const lookup = hv4LookupOrderByName_(g.orderName, g.createdAt);
      if (!lookup.ok) {
        failCount++;
        Logger.log(`❌ Lookup ${g.orderName}: ${lookup.error}`);
        appendNoteAllRows_(sheet, g.rows, `Fulfill fail (lookup): ${lookup.error}`);
        Utilities.sleep(HV4_CFG.SLEEP_MS);
        continue;
      }

      // Step 2a: nếu đơn đã có fulfillment với tracking_number trùng → coi như done
      const existing = (lookup.order.fulfillments || []);
      const alreadyMatch = existing.some(f =>
        String(f.tracking_number || '').toUpperCase() === g.ghnCode.toUpperCase());
      if (alreadyMatch) {
        skipCount++;
        markFulfilledRows_(sheet, g.rows, now);
        appendNoteAllRows_(sheet, g.rows, `Đã fulfill trước với cùng tracking ${g.ghnCode}`);
        Utilities.sleep(HV4_CFG.SLEEP_MS);
        continue;
      }

      // Step 2b: đơn đã ở trạng thái fulfilled (toàn bộ line_items đã fulfill bằng tracking khác)
      // → tránh POST → 422 "Đơn hàng đã hoàn thành". Mark AD + note để nhân viên biết.
      const fStatus = String(lookup.order.fulfillment_status || '').toLowerCase();
      if (fStatus === 'fulfilled') {
        skipCount++;
        markFulfilledRows_(sheet, g.rows, now);
        const otherTrackings = existing.map(f => f.tracking_number).filter(Boolean).join(', ');
        appendNoteAllRows_(sheet, g.rows,
          `Đơn Haravan đã fulfill bằng tracking khác (${otherTrackings || 'không có tracking'}) — KHÔNG POST lại`);
        Logger.log(`⏭️ ${g.orderName}: đã fulfilled trước (tracking khác). Mark AD, không POST.`);
        Utilities.sleep(HV4_CFG.SLEEP_MS);
        continue;
      }

      // Step 3: build line_items array để gửi
      const lineItems = (lookup.order.line_items || []).map(li => ({
        id: li.id, quantity: li.quantity,
      }));

      // Step 4: POST fulfillment
      const res = hv4CreateFulfillment_(lookup.order.id, g.ghnCode, lineItems);
      if (!res.ok) {
        // Nếu Haravan trả lỗi "đã hoàn thành" / "fulfilled" → đơn đã fulfill bằng cách khác,
        // mark AD để skip, tránh cron retry vô tận.
        const isAlreadyDone = /đã hoàn thành|đã giao|fulfilled|already/i.test(res.error);
        if (isAlreadyDone) {
          skipCount++;
          markFulfilledRows_(sheet, g.rows, now);
          appendNoteAllRows_(sheet, g.rows,
            `Haravan báo đơn đã hoàn thành — mark AD để skip cron tới: ${res.error}`);
          Logger.log(`⏭️ ${g.orderName}: Haravan báo đã hoàn thành. Mark AD, không retry.`);
          Utilities.sleep(HV4_CFG.SLEEP_MS);
          continue;
        }
        failCount++;
        Logger.log(`❌ Fulfill ${g.orderName} (id=${lookup.order.id}): ${res.error}`);
        appendNoteAllRows_(sheet, g.rows, `Fulfill fail: ${res.error}`);
        Utilities.sleep(HV4_CFG.SLEEP_MS);
        continue;
      }
      okCount++;
      markFulfilledRows_(sheet, g.rows, now);
    } catch (e) {
      failCount++;
      Logger.log(`❌ ${g.orderName} exception: ${e.message}`);
      appendNoteAllRows_(sheet, g.rows, `Fulfill exception: ${e.message}`);
    }
    Utilities.sleep(HV4_CFG.SLEEP_MS);
  }

  PropertiesService.getScriptProperties().setProperty(
    HV4_CFG.PROP_LAST_RUN, new Date().toISOString());

  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  Logger.log(`✅ Done ${dur}s. OK: ${okCount}, Skip(đã fulfill): ${skipCount}, Lỗi: ${failCount}`);
}

function markFulfilledRows_(sheet, rows, timestamp) {
  rows.forEach(r => sheet.getRange(r, HV4_COL.FULFILL_AT).setValue(timestamp));
}

function appendNoteAllRows_(sheet, rows, note) {
  rows.forEach(r => {
    const cur = sheet.getRange(r, GHN_COL.GHI_CHU_NV).getValue();
    const merged = appendNote_(cur, note);
    sheet.getRange(r, GHN_COL.GHI_CHU_NV).setValue(merged);
  });
}

// ============================================================
// HARAVAN API CALLS
// ============================================================

/**
 * Lookup đơn Haravan bằng khoảng thời gian (Haravan API không filter ?name= reliable).
 * @param {string} orderName  Mã đơn (vd "#EC325103846")
 * @param {Date|string} createdAt  Ngày đặt từ cột A. Có thể là Date object hoặc string "dd/MM/yyyy HH:mm".
 */
function hv4LookupOrderByName_(orderName, createdAt) {
  const cleanName = String(orderName).replace(/^#/, '').trim();
  const target = cleanName.toUpperCase();

  const refDate = hv4ParseDate_(createdAt);
  if (!refDate) {
    return { ok: false, error: `Không parse được ngày đặt cho ${orderName} (giá trị="${createdAt}")` };
  }

  // Cửa sổ -2h tới +26h quanh thời điểm đặt — đảm bảo bao gồm đơn dù có lệch timezone
  const minDate = new Date(refDate.getTime() - 2 * 3600 * 1000);
  const maxDate = new Date(refDate.getTime() + 26 * 3600 * 1000);

  let allOrders = [];
  let lastError = '';
  for (let page = 1; page <= 4; page++) {   // tối đa 4 page × 50 = 200 đơn / cửa sổ
    const url = `https://${CONFIG.HARAVAN_SHOP}/admin/orders.json` +
      `?created_at_min=${encodeURIComponent(minDate.toISOString())}` +
      `&created_at_max=${encodeURIComponent(maxDate.toISOString())}` +
      `&status=any&limit=50&page=${page}&order=created_at+asc`;
    let res;
    try {
      res = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + getHv4Token_() },
        muteHttpExceptions: true,
      });
    } catch (e) { lastError = 'Network: ' + e.message; break; }
    const code = res.getResponseCode();
    const txt = res.getContentText();
    if (code !== 200) { lastError = `HTTP ${code}: ${txt.slice(0, 150)}`; break; }
    let body;
    try { body = JSON.parse(txt); } catch (e) { lastError = 'Non-JSON'; break; }
    const orders = body.orders || [];
    if (orders.length === 0) break;

    const exact = orders.find(o => {
      const n = String(o.name || '').replace(/^#/, '').toUpperCase();
      return n === target;
    });
    if (exact) return { ok: true, order: exact };

    allOrders = allOrders.concat(orders);
    if (orders.length < 50) break;   // hết page
  }

  const namesPreview = allOrders.slice(0, 8).map(o => o.name).join(', ') || '(rỗng)';
  const winStr = `${minDate.toISOString().slice(0,16)} → ${maxDate.toISOString().slice(0,16)}`;
  return {
    ok: false,
    error: `Không tìm thấy "${orderName}" trong cửa sổ ${winStr} (${allOrders.length} đơn quét). Mẫu: [${namesPreview}]${lastError ? ' | ' + lastError : ''}`,
  };
}

function hv4ParseDate_(v) {
  if (!v) return null;

  // Nếu là Date object (Sheet đã auto-parse khi ghi cell), có khả năng locale Sheet
  // là en_US đã hiểu "01/05/2026" thành Jan 5 thay vì May 1. Áp dụng heuristic swap
  // dd↔mm khi Date "kỳ lạ" (cách hôm nay > 30 ngày) và swap sẽ gần hôm nay hơn.
  if (v instanceof Date && !isNaN(v.getTime())) {
    const now = Date.now();
    const diffDays = Math.abs((now - v.getTime()) / 86400000);
    const dd = v.getDate();
    const mm = v.getMonth() + 1;
    if (diffDays > 30 && dd <= 12 && mm <= 12 && dd !== mm) {
      const swapped = new Date(
        v.getFullYear(), dd - 1, mm,
        v.getHours(), v.getMinutes(), v.getSeconds()
      );
      const swappedDiff = Math.abs((now - swapped.getTime()) / 86400000);
      if (swappedDiff < diffDays && swappedDiff <= 30) {
        Logger.log(`   ⚠️ Auto-swap dd↔mm: ${v.toISOString().slice(0,16)} → ${swapped.toISOString().slice(0,16)} (Sheet locale en_US)`);
        return swapped;
      }
    }
    return v;
  }

  const s = String(v).trim();
  // Format Phase 1 lưu: "dd/MM/yyyy HH:mm" (Asia/Ho_Chi_Minh)
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) {
    const dd = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10) - 1;
    const yy = parseInt(m[3], 10);
    const hh = m[4] ? parseInt(m[4], 10) : 0;
    const mi = m[5] ? parseInt(m[5], 10) : 0;
    // Đặt giờ theo +07:00 → quy về UTC bằng cách trừ 7h
    const utcMs = Date.UTC(yy, mm, dd, hh, mi) - 7 * 3600 * 1000;
    return new Date(utcMs);
  }
  // Thử ISO
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function hv4CreateFulfillment_(orderId, trackingNumber, lineItems) {
  const url = `https://${CONFIG.HARAVAN_SHOP}/admin/orders/${orderId}/fulfillments.json`;
  const payload = {
    fulfillment: {
      tracking_number: trackingNumber,
      tracking_company: HV4_CFG.TRACKING_COMPANY,
      notify_customer: HV4_CFG.NOTIFY_CUSTOMER,
      line_items: lineItems,
    },
  };
  let res;
  try {
    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + getHv4Token_() },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
  } catch (e) {
    return { ok: false, error: 'Network: ' + e.message };
  }
  const code = res.getResponseCode();
  const txt = res.getContentText();
  if (code === 201 || code === 200) {
    let body; try { body = JSON.parse(txt); } catch (e) { body = {}; }
    return { ok: true, fulfillment: body.fulfillment || null };
  }
  return { ok: false, error: `HTTP ${code}: ${txt.slice(0, 250)}` };
}

function getHv4Token_() {
  const t = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_TOKEN);
  if (!t) throw new Error('Chưa setup token Haravan. Chạy setupToken() (Phase 1) trước.');
  return t;
}

// ============================================================
// DEBUG / UTILITY
// ============================================================

function manualFulfillOne(orderName) {
  if (!orderName) {
    Logger.log('Truyền vào Mã đơn Haravan. Vd: manualFulfillOne("RT12510680")');
    return;
  }
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, HV4_COL.FULFILL_AT).getValues();
  const cleanName = String(orderName).replace(/^#/, '');

  let ghnCode = '', createdAt = null;
  const rows = [];
  data.forEach((row, idx) => {
    const n = String(row[GHN_COL.MA_HARAVAN - 1] || '').replace(/^#/, '');
    if (n === cleanName) {
      const code = String(row[GHN_COL.GHN_MA - 1] || '').trim();
      if (code) ghnCode = code;
      if (!createdAt) createdAt = row[GHN_COL.NGAY_DAT - 1];
      rows.push(idx + 2);
    }
  });
  if (rows.length === 0) { Logger.log(`Không tìm thấy đơn ${orderName} trong sheet`); return; }
  if (!ghnCode) { Logger.log(`Đơn ${orderName} chưa có Mã VC GHN, không fulfill được`); return; }

  Logger.log(`Đơn ${orderName}: ${rows.length} dòng, Mã VC = ${ghnCode}, Ngày đặt = ${createdAt}`);
  const lookup = hv4LookupOrderByName_(orderName, createdAt);
  if (!lookup.ok) { Logger.log(`❌ Lookup: ${lookup.error}`); return; }
  Logger.log(`Haravan order id=${lookup.order.id}, line_items=${(lookup.order.line_items || []).length}, fulfillments=${(lookup.order.fulfillments || []).length}`);

  const lineItems = (lookup.order.line_items || []).map(li => ({ id: li.id, quantity: li.quantity }));
  const res = hv4CreateFulfillment_(lookup.order.id, ghnCode, lineItems);
  if (!res.ok) { Logger.log(`❌ Fulfill: ${res.error}`); return; }
  Logger.log(`✅ Fulfill OK — fulfillment id=${res.fulfillment ? res.fulfillment.id : '?'}`);
  const now = formatGhnNow_();
  rows.forEach(r => sheet.getRange(r, HV4_COL.FULFILL_AT).setValue(now));
}

function resetHv4FulfillRow(rowNumber) {
  if (!rowNumber || rowNumber < 2) { Logger.log('Truyền row >= 2'); return; }
  const sheet = openGhnTargetTab_();
  sheet.getRange(rowNumber, HV4_COL.FULFILL_AT).clearContent();
  Logger.log(`✅ Đã clear cột AD row ${rowNumber} — lần cron tới sẽ retry fulfill`);
}

// ============================================================
// 🟡 DRY RUN — TEST AN TOÀN, KHÔNG GỌI API WRITE
// ============================================================

/**
 * Dry-run 1 đơn: lookup + log payload SẼ POST, KHÔNG call POST fulfillment.
 * Dùng khi muốn xem tool sẽ làm gì với 1 đơn cụ thể trước khi fulfill thật.
 * Vd: function _t() { dryRunFulfillOne("EC325103846"); }
 */
function dryRunFulfillOne(orderName) {
  if (!orderName) { Logger.log('Truyền vào Mã đơn. Vd: dryRunFulfillOne("EC325103846")'); return; }
  Logger.log('🟡 DRY RUN — KHÔNG gọi API write, chỉ log payload sẽ gửi');

  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, HV4_COL.FULFILL_AT).getValues();
  const cleanName = String(orderName).replace(/^#/, '');

  let ghnCode = '', createdAt = null;
  const rows = [];
  data.forEach((row, idx) => {
    const n = String(row[GHN_COL.MA_HARAVAN - 1] || '').replace(/^#/, '');
    if (n === cleanName) {
      const code = String(row[GHN_COL.GHN_MA - 1] || '').trim();
      if (code) ghnCode = code;
      if (!createdAt) createdAt = row[GHN_COL.NGAY_DAT - 1];
      rows.push(idx + 2);
    }
  });
  if (rows.length === 0) { Logger.log(`❌ Không tìm thấy đơn ${orderName} trong sheet`); return; }
  if (!ghnCode) { Logger.log(`❌ Đơn ${orderName} chưa có Mã VC GHN — không fulfill được`); return; }

  Logger.log(`Đơn ${orderName}: ${rows.length} dòng, Mã VC = ${ghnCode}, Ngày đặt = ${createdAt}`);
  const lookup = hv4LookupOrderByName_(orderName, createdAt);
  if (!lookup.ok) { Logger.log(`❌ Lookup: ${lookup.error}`); return; }

  const o = lookup.order;
  const existingFf = (o.fulfillments || []);
  Logger.log(`Haravan id=${o.id}, line_items=${(o.line_items || []).length}, fulfillments hiện có=${existingFf.length}`);
  existingFf.forEach((f, i) => {
    Logger.log(`  Fulfillment cũ #${i + 1}: tracking=${f.tracking_number}, company=${f.tracking_company}, status=${f.status}`);
  });

  const alreadyMatch = existingFf.some(f =>
    String(f.tracking_number || '').toUpperCase() === ghnCode.toUpperCase());
  if (alreadyMatch) {
    Logger.log(`⏭️ SKIP — đơn đã có fulfillment với tracking_number = ${ghnCode}, tool sẽ chỉ mark cột AD`);
    return;
  }

  const lineItems = (o.line_items || []).map(li => ({ id: li.id, quantity: li.quantity }));
  const payload = {
    fulfillment: {
      tracking_number: ghnCode,
      tracking_company: HV4_CFG.TRACKING_COMPANY,
      notify_customer: HV4_CFG.NOTIFY_CUSTOMER,
      line_items: lineItems,
    },
  };
  Logger.log(`📋 Payload SẼ POST tới /admin/orders/${o.id}/fulfillments.json:`);
  Logger.log(JSON.stringify(payload, null, 2));
  Logger.log(`✅ DRY RUN OK — gọi manualFulfillOne("${cleanName}") để fulfill thật`);
}

/**
 * Dry-run TOÀN BỘ: quét sheet, log danh sách đơn SẼ fulfill, KHÔNG call API.
 * Hữu ích trước khi run fulfillHaravanOrders lần đầu để biết sẽ đụng bao nhiêu đơn.
 */
function dryRunFulfillAll() {
  Logger.log('🟡 DRY RUN — KHÔNG gọi Haravan API write, chỉ thống kê');

  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng'); return; }
  const lastCol = Math.max(HV4_COL.FULFILL_AT, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const groups = {};
  data.forEach((row, idx) => {
    const orderName = String(row[GHN_COL.MA_HARAVAN - 1] || '').trim();
    if (!orderName) return;
    const ghnCode = String(row[GHN_COL.GHN_MA - 1] || '').trim();
    if (!ghnCode) return;
    const fulfilledAt = String(row[HV4_COL.FULFILL_AT - 1] || '').trim();
    if (fulfilledAt) return;
    const xStatus = String(row[GHN_COL.TT_XU_LY - 1] || '').trim();
    if (HV4_TERMINAL_X.indexOf(xStatus) >= 0) return;

    if (!groups[orderName]) {
      groups[orderName] = { orderName, ghnCode, createdAt: row[GHN_COL.NGAY_DAT - 1], rows: [] };
    }
    groups[orderName].rows.push(idx + 2);
  });
  const orderNames = Object.keys(groups);
  Logger.log(`Tổng đơn SẼ fulfill: ${orderNames.length}`);
  orderNames.slice(0, 30).forEach((n, i) => {
    const g = groups[n];
    Logger.log(`  ${i + 1}. ${n} | Mã VC = ${g.ghnCode} | ${g.rows.length} dòng | Ngày đặt = ${g.createdAt}`);
  });
  if (orderNames.length > 30) Logger.log(`  ... (còn ${orderNames.length - 30} đơn nữa)`);
  Logger.log(`✅ DRY RUN OK — chạy fulfillHaravanOrders để fulfill thật`);
}

function showHv4Report() {
  const sheet = openGhnTargetTab_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet rỗng'); return; }
  const data = sheet.getRange(2, 1, lastRow - 1, HV4_COL.FULFILL_AT).getValues();
  let hasGhn = 0, fulfilled = 0, pending = 0, terminal = 0;
  data.forEach(r => {
    const ghn = String(r[GHN_COL.GHN_MA - 1] || '').trim();
    if (!ghn) return;
    hasGhn++;
    const ad = String(r[HV4_COL.FULFILL_AT - 1] || '').trim();
    const x = String(r[GHN_COL.TT_XU_LY - 1] || '').trim();
    if (ad) fulfilled++;
    else if (HV4_TERMINAL_X.indexOf(x) >= 0) terminal++;
    else pending++;
  });
  Logger.log('===== Phase 4 — Báo cáo fulfill Haravan =====');
  Logger.log(`Tổng đơn có Mã VC GHN: ${hasGhn} (đếm theo dòng)`);
  Logger.log(`  Đã fulfill: ${fulfilled}`);
  Logger.log(`  Chưa fulfill (chờ cron): ${pending}`);
  Logger.log(`  Skip (Hủy/Lỗi GHN): ${terminal}`);
  const last = PropertiesService.getScriptProperties().getProperty(HV4_CFG.PROP_LAST_RUN);
  Logger.log(`Lần chạy cuối: ${last || '(chưa chạy)'}`);
}
