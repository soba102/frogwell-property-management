/**
 * オーナーにレポート PDF をメール送信する。
 */

/**
 * 全オーナー分のレポートを生成してメール送信する。
 * @returns {number} 送信したメールの件数
 */
function sendAllOwnerReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ownerSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OWNERS);
  var owners = ownerSheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < owners.length; i++) {
    var ownerName = String(owners[i][0] || '').trim();
    var email = owners[i][1];
    if (!ownerName) continue;

    if (!email || email.indexOf('@') === -1) {
      Logger.log('メールアドレスなし、スキップ: ' + ownerName);
      continue;
    }

    try {
      var pdfFile = generateOwnerReport(ownerName);
      var pdfBlob = pdfFile.getBlob();
      sendReportEmail_(email, ownerName, pdfBlob);
      count++;
      Logger.log('送信完了: ' + ownerName + ' (' + email + ')');
    } catch (e) {
      Logger.log('送信エラー (' + ownerName + '): ' + e.message);
    }

    if (i < owners.length - 1) Utilities.sleep(1000);
  }
  return count;
}

/** レポートメールを送信する */
function sendReportEmail_(toEmail, ownerName, pdfBlob) {
  var now = new Date();
  var yearMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月');

  var companyName = 'Property Management';
  var companyEmail = '';
  var companyTel = '';
  try { companyName = getSettingValue('COMPANY_NAME'); } catch (e) {}
  try { companyEmail = getSettingValue('COMPANY_EMAIL'); } catch (e) {}
  try { companyTel = getSettingValue('COMPANY_TEL'); } catch (e) {}

  var subject = '【月次収支レポート】' + yearMonth + ' / ' + ownerName + '様';

  var body = [
    ownerName + ' 様',
    '',
    'いつもお世話になっております。',
    yearMonth + '分の収支レポートを添付にてお送りいたします。',
    '',
    'ご不明点がございましたらお気軽にお問い合わせください。',
    '',
    '──────────────',
    companyName,
    companyTel ? 'Tel: ' + companyTel : '',
    companyEmail ? 'Email: ' + companyEmail : '',
  ].filter(function(line) { return line !== ''; }).join('\n');

  GmailApp.sendEmail(toEmail, subject, body, {
    attachments: [pdfBlob],
    name: companyName,
  });
}
