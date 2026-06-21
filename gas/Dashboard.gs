/**
 * ダッシュボード（メニューシート）のサマリー情報を更新する。
 */
function updateDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(CONFIG.SHEET_NAMES.MENU);
  if (!dashboard) return;

  // --- 請求書サマリー ---
  var invoiceSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVOICES);
  var total = 0, confirmed = 0, review = 0, aiRead = 0;
  if (invoiceSheet && invoiceSheet.getLastRow() > 1) {
    var data = invoiceSheet.getDataRange().getValues();
    var statusCol = CONFIG.INVOICE_HEADERS.indexOf('ステータス');
    for (var i = 1; i < data.length; i++) {
      total++;
      var status = data[i][statusCol];
      if (status === CONFIG.STATUSES.CONFIRMED) confirmed++;
      else if (status === CONFIG.STATUSES.REVIEW) review++;
      else if (status === CONFIG.STATUSES.AI_READ) aiRead++;
    }
  }
  dashboard.getRange('B7').setValue(total);
  dashboard.getRange('B8').setValue(confirmed).setFontColor('#137333');
  dashboard.getRange('B9').setValue(review).setFontColor('#e37400');
  dashboard.getRange('B10').setValue(aiRead).setFontColor('#1967d2');

  // --- 入金消込サマリー ---
  var reconSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RECONCILE);
  var rec = 0, partial = 0, unpaid = 0, unmatched = 0;
  if (reconSheet && reconSheet.getLastRow() > 1) {
    var rdata = reconSheet.getDataRange().getValues();
    var rStatusCol = CONFIG.RECONCILE_HEADERS.indexOf('ステータス');
    for (var k = 1; k < rdata.length; k++) {
      var rs = rdata[k][rStatusCol];
      if (rs === CONFIG.RECON_STATUSES.RECONCILED) rec++;
      else if (rs === CONFIG.RECON_STATUSES.PARTIAL) partial++;
      else if (rs === CONFIG.RECON_STATUSES.UNPAID) unpaid++;
      else if (rs === CONFIG.RECON_STATUSES.UNMATCHED || rs === CONFIG.RECON_STATUSES.OVERPAID) unmatched++;
    }
  }
  dashboard.getRange('B16').setValue(rec).setFontColor('#137333');
  dashboard.getRange('B17').setValue(partial).setFontColor('#e37400');
  dashboard.getRange('B18').setValue(unpaid).setFontColor('#c5221f');
  dashboard.getRange('B19').setValue(unmatched).setFontColor('#9334e6');

  // 最終更新日時
  dashboard.getRange('A12').setValue(
    '最終更新: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
  ).setFontColor('#80868b').setFontSize(10);
}
