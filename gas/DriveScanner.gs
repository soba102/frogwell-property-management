/**
 * Google Drive の請求書フォルダをスキャンし、
 * まだ処理されていない新規ファイルのリストを返す。
 *
 * 「処理済み」の判定は、スプレッドシートの請求書データシートに
 * 元ファイルURLが既に存在するかどうかで行う。
 */
function scanNewInvoiceFiles() {
  var folderId = getInvoiceFolderId();
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var processedUrls = getProcessedFileUrls_();
  var newFiles = [];

  while (files.hasNext()) {
    var file = files.next();
    var mimeType = file.getMimeType();

    var isSupported = (
      mimeType === 'application/pdf' ||
      mimeType === 'image/jpeg' ||
      mimeType === 'image/png' ||
      mimeType === 'image/webp'
    );
    if (!isSupported) continue;

    var url = file.getUrl();
    if (processedUrls[url]) continue;

    newFiles.push({
      id: file.getId(),
      name: file.getName(),
      url: url,
      mimeType: mimeType,
      blob: file.getBlob(),
    });
  }

  Logger.log('新規ファイル: ' + newFiles.length + ' 件');
  return newFiles;
}

/** 既に処理済みのファイルURLセットを取得する */
function getProcessedFileUrls_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVOICES);
  if (!sheet || sheet.getLastRow() <= 1) return {};

  var urlColIndex = CONFIG.INVOICE_HEADERS.indexOf('元ファイルURL');
  var data = sheet.getRange(2, urlColIndex + 1, sheet.getLastRow() - 1, 1).getValues();

  var urls = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) urls[data[i][0]] = true;
  }
  return urls;
}
