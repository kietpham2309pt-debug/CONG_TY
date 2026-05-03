/**
 * PHASE 21 — recoverOrderByName helper
 *
 * Khi PG xóa nhầm row trong sheet, Phase 1 không pull lại (vì PROP_LAST_SYNC
 * đã đi qua thời điểm đó). Hàm này pull đơn cụ thể từ Haravan API về sheet.
 *
 * Action qua Web App POST: action=recover_order&order_name=EC325103860
 */

function recoverOrderByName(orderName) {
  if (!orderName) throw new Error('orderName required');
  const orderClean = String(orderName).trim();
  Logger.log('Recover đơn: ' + orderClean);

  const url = 'https://' + CONFIG.HARAVAN_SHOP + '/admin/orders.json'
            + '?name=' + encodeURIComponent(orderClean) + '&status=any&limit=5';
  const token = PropertiesService.getScriptProperties().getProperty(CONFIG.PROP_TOKEN);
  const res = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    return { ok: false, error: 'Haravan HTTP ' + res.getResponseCode() };
  }

  let orders = JSON.parse(res.getContentText()).orders || [];
  // Haravan ?name= filter có thể không exact, lọc lại
  orders = orders.filter(function (o) { return o.name === orderClean; });

  if (!orders.length) {
    // Fallback: search 14 ngày gần nhất
    const since = new Date(); since.setDate(since.getDate() - 14);
    const url2 = 'https://' + CONFIG.HARAVAN_SHOP + '/admin/orders.json'
               + '?status=any&limit=250&created_at_min=' + encodeURIComponent(since.toISOString());
    const res2 = UrlFetchApp.fetch(url2, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true,
    });
    if (res2.getResponseCode() === 200) {
      const all = JSON.parse(res2.getContentText()).orders || [];
      orders = all.filter(function (o) { return o.name === orderClean; });
    }
  }

  if (!orders.length) {
    return { ok: false, error: 'Không tìm thấy đơn ' + orderClean + ' trong Haravan' };
  }

  const order = orders[0];
  const sheet = SpreadsheetApp.openById(CONFIG.TARGET_SHEET_ID).getSheetByName(CONFIG.TARGET_TAB);
  const lastRow = sheet.getLastRow();

  // Check existing
  let existingRowIdx = -1;
  if (lastRow > 1) {
    const codes = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < codes.length; i++) {
      if (String(codes[i][0]).trim() === orderClean) { existingRowIdx = i + 2; break; }
    }
  }

  // Build row tận dụng logic Phase 1 nhưng inline ở đây để tránh refactor lớn
  const parsed = parseNote_(order.note);
  const orderType = classifyOrderType_(order, parsed);
  const sa = order.shipping_address || {};
  const province = getProvinceWithFallback_(sa);
  const lineItems = order.line_items || [];
  const now = formatDate_(new Date().toISOString());
  const ptttRaw = String(order.gateway || '').toLowerCase();
  const shippingMethod = String((order.shipping_lines && order.shipping_lines[0] && order.shipping_lines[0].title) || '').toLowerCase();
  const ptttCombined = (ptttRaw + ' ' + shippingMethod).trim();
  const isPickup = (typeof isPickupShipping_ === 'function')
    ? isPickupShipping_(ptttCombined, sa.address1 || '')
    : /tại cửa hàng|nhận tại|showroom|pickup|tự đến|tự lấy/i.test(ptttCombined);

  // Recovery: KHÔNG auto-confirm, để PG kiểm tra
  let initialStatus = isPickup ? 'Pickup tại SR' : 'Chưa xử lý';
  const initialNote = `RECOVERED ${now}: pulled lại từ Haravan via recoverOrderByName`;

  const baseRow = [
    formatDate_(order.created_at), order.name, parsed.sourceOrderCode, orderType, parsed.camp,
    sa.name || '', sa.phone || '', order.email || '', sa.address1 || '',
    sa.ward || '', sa.district || '', province,
    detectRegion_(province), order.gateway || '',
  ];

  const writes = [];
  if (lineItems.length === 0) {
    writes.push(baseRow.concat(['', '', 0, 0, order.total_price || 0,
      translateFinancialStatus_(order.financial_status), translateFulfillmentStatus_(order.fulfillment_status),
      parsed.cleanNote, now, initialStatus, initialNote]));
  } else {
    lineItems.forEach(function (li) {
      writes.push(baseRow.concat([
        li.sku || '', li.name || '',
        li.quantity || 1, li.price || 0,
        order.total_price || 0,
        translateFinancialStatus_(order.financial_status),
        translateFulfillmentStatus_(order.fulfillment_status),
        parsed.cleanNote, now,
        initialStatus, initialNote,
      ]));
    });
  }

  // Append (nếu existing thì xóa rồi thêm lại)
  if (existingRowIdx > 0) {
    sheet.deleteRow(existingRowIdx);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, writes.length, writes[0].length).setValues(writes);

  Logger.log('✅ Recovered ' + writes.length + ' rows for order ' + orderClean);
  return { ok: true, order: orderClean, rows_added: writes.length, was_existing: existingRowIdx > 0 };
}
