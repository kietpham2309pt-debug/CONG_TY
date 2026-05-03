/**
 * PHASE 17 — Đối soát COD GHN daily
 *
 * Cron 19h sau Phase 7. Quét đơn có Mã VC + TT giao = "Đã giao" + AN trống → gọi GHN
 * /v2/shipping-order/detail lấy:
 *   - cod_amount: COD GHN thu hộ (về account ~3 ngày sau)
 *   - total_fee: phí ship final
 * Ghi vào AL (COD GHN thu) / AM (Phí ship final) / AN (Đối soát lúc).
 *
 * Email tổng kết:
 *   - Tổng COD chờ về tài khoản
 *   - Top 5 đơn phí ship lệch ước tính > 1000đ
 *   - Đơn lỗi gọi API
 */

const COD_RECON_CFG = {
  COL_COD_AMOUNT: 38,    // AL
  COL_TOTAL_FEE: 39,     // AM
  COL_RECON_AT: 40,      // AN

  EXT_HEADERS: ['COD GHN thu', 'Phí ship final', 'Đối soát lúc'],

  MAX_PER_RUN: 100,
  TERMINAL_TT_GIAO: ['Đã giao'],
  ALERT_EMAIL: 'admin@khomes.com.vn',
};

function setupCodReconHeader_() {
  const sheet = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID).getSheetByName(GHN_CFG.TARGET_TAB);
  const curCols = sheet.getMaxColumns();
  if (curCols < COD_RECON_CFG.COL_RECON_AT) {
    sheet.insertColumnsAfter(curCols, COD_RECON_CFG.COL_RECON_AT - curCols);
  }
  sheet.getRange(1, COD_RECON_CFG.COL_COD_AMOUNT, 1, 3).setValues([COD_RECON_CFG.EXT_HEADERS]);
  sheet.getRange(1, COD_RECON_CFG.COL_COD_AMOUNT, 1, 3)
    .setFontWeight('bold').setBackground('#fff2cc');
  Logger.log('✅ Setup AL-AN headers Phase 17');
}

function runCodRecon() {
  setupCodReconHeader_();
  const ss = SpreadsheetApp.openById(GHN_CFG.TARGET_SHEET_ID);
  const sheet = ss.getSheetByName(GHN_CFG.TARGET_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Empty sheet'); return { ok: true, processed: 0 }; }

  const data = sheet.getRange(2, 1, lastRow - 1, COD_RECON_CFG.COL_RECON_AT).getValues();
  const candidates = [];
  data.forEach(function (row, idx) {
    const ttGiao = String(row[GHN_COL.TT_GIAO - 1] || '').trim();
    const maVc   = String(row[GHN_COL.GHN_MA - 1] || '').trim();
    const reconAt = row[COD_RECON_CFG.COL_RECON_AT - 1];
    if (maVc && COD_RECON_CFG.TERMINAL_TT_GIAO.indexOf(ttGiao) >= 0 && !reconAt) {
      candidates.push({
        rowIdx: idx + 2,
        maVc,
        haravan: row[GHN_COL.MA_HARAVAN - 1],
        tongDon: Number(row[GHN_COL.TONG_DON - 1] || 0),
        phiUocTinh: Number(row[GHN_COL.GHN_PHI - 1] || 0),
      });
    }
  });

  if (!candidates.length) {
    Logger.log('No candidates for COD recon');
    return { ok: true, processed: 0 };
  }

  const limited = candidates.slice(0, COD_RECON_CFG.MAX_PER_RUN);
  const stats = { ok: 0, fail: 0, totalCod: 0, totalFee: 0, feeDiffs: [] };
  const errors = [];

  limited.forEach(function (c) {
    try {
      const detail = ghnFetch_('/v2/shipping-order/detail', { order_code: c.maVc });
      if (!detail || detail.code !== 200 || !detail.data) {
        errors.push({ maVc: c.maVc, error: (detail && detail.message) || 'No data' });
        sheet.getRange(c.rowIdx, COD_RECON_CFG.COL_RECON_AT).setValue('LỖI: ' + ((detail && detail.message) || 'no data'));
        stats.fail++;
        return;
      }
      const cod = Number(detail.data.cod_amount || 0);
      const fee = Number(detail.data.total_fee || detail.data.fee || 0);
      sheet.getRange(c.rowIdx, COD_RECON_CFG.COL_COD_AMOUNT).setValue(cod);
      sheet.getRange(c.rowIdx, COD_RECON_CFG.COL_TOTAL_FEE).setValue(fee);
      sheet.getRange(c.rowIdx, COD_RECON_CFG.COL_RECON_AT).setValue(new Date());
      stats.ok++;
      stats.totalCod += cod;
      stats.totalFee += fee;
      const diff = fee - c.phiUocTinh;
      if (Math.abs(diff) > 1000) {
        stats.feeDiffs.push({ maVc: c.maVc, haravan: c.haravan, est: c.phiUocTinh, real: fee, diff });
      }
      Utilities.sleep(250);
    } catch (err) {
      errors.push({ maVc: c.maVc, error: err.message });
      stats.fail++;
    }
  });

  stats.feeDiffs.sort(function (a, b) { return Math.abs(b.diff) - Math.abs(a.diff); });
  sendCodReconEmail_(stats, errors, limited.length);
  return { ok: true, processed: limited.length, stats };
}

function sendCodReconEmail_(stats, errors, total) {
  const today = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy');
  const subject = '[K-Homes] Đối soát COD GHN ' + today;
  let html = '<h2 style="color:#1a73e8">💰 Đối soát COD GHN ngày ' + today + '</h2>';
  html += '<table border="1" cellpadding="8" style="border-collapse:collapse;font-family:Arial">';
  html += '<tr><td><b>Đơn xử lý</b></td><td>' + total + '</td></tr>';
  html += '<tr style="background:#d9ead3"><td><b>OK</b></td><td>' + stats.ok + '</td></tr>';
  html += '<tr style="background:#f4cccc"><td><b>Lỗi</b></td><td>' + stats.fail + '</td></tr>';
  html += '<tr style="background:#fff2cc"><td><b>💰 COD chờ về account</b></td><td>' + stats.totalCod.toLocaleString('vi-VN') + ' đ</td></tr>';
  html += '<tr><td><b>🚚 Tổng phí ship final</b></td><td>' + stats.totalFee.toLocaleString('vi-VN') + ' đ</td></tr>';
  html += '</table>';

  if (stats.feeDiffs.length) {
    html += '<h3 style="color:#e8710a">⚠️ Phí ship lệch ước tính (top 5)</h3>';
    html += '<table border="1" cellpadding="6" style="border-collapse:collapse;font-family:Arial">';
    html += '<tr style="background:#cfe2f3"><th>Mã VC</th><th>Mã đơn</th><th>Ước tính</th><th>Thực tế</th><th>Lệch</th></tr>';
    stats.feeDiffs.slice(0, 5).forEach(function (d) {
      const bg = d.diff > 0 ? '#fce5cd' : '#d9ead3';
      html += '<tr style="background:' + bg + '"><td>' + d.maVc + '</td><td>' + (d.haravan || '') +
              '</td><td>' + d.est.toLocaleString('vi-VN') + '</td><td>' + d.real.toLocaleString('vi-VN') +
              '</td><td><b>' + (d.diff > 0 ? '+' : '') + d.diff.toLocaleString('vi-VN') + '</b></td></tr>';
    });
    html += '</table>';
  }

  if (errors.length) {
    html += '<h3 style="color:#c00">❌ Đơn lỗi đối soát</h3><ul>';
    errors.slice(0, 10).forEach(function (e) { html += '<li>' + e.maVc + ': ' + e.error + '</li>'; });
    html += '</ul>';
  }

  html += '<hr><p style="color:#888;font-size:12px">Phase 17 cron 19h. COD chờ về sẽ vào account ngân hàng ~3-7 ngày sau ngày giao.</p>';

  MailApp.sendEmail({ to: COD_RECON_CFG.ALERT_EMAIL, subject, htmlBody: html });
}

function setupCodReconTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runCodRecon') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runCodRecon').timeBased().atHour(19).everyDays(1).create();
  Logger.log('✅ Trigger runCodRecon daily 19h');
}
