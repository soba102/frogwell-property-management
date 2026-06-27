/**
 * スプレッドシートの全シートを作成し、ヘッダーとサンプルデータを投入する。
 * メニューの「初期セットアップ」から実行される。
 */
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.INVOICES,   CONFIG.INVOICE_HEADERS);
  createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.PROPERTIES, CONFIG.PROPERTY_HEADERS);
  createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.OWNERS,     CONFIG.OWNER_HEADERS);
  createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.TENANTS,    CONFIG.TENANT_HEADERS);
  createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.BUILDING_ALIAS, CONFIG.BUILDING_ALIAS_HEADERS);
  createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.RECONCILE,  CONFIG.RECONCILE_HEADERS);
  createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.SETTINGS,   CONFIG.SETTINGS_HEADERS);
  createReportTemplateSheet_(ss);
  createDashboardSheet_(ss);

  insertSampleMasterData_(ss);
  insertDefaultSettings_(ss);

  // 問い合わせ対応（Phase 2）のシート・設定キーも作成する
  if (typeof setupInquirySheets === 'function') setupInquirySheets();

  // デフォルトの「シート1」があれば削除
  const sheet1 = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
}

/** シートを作成してヘッダー行をセットする。既存なら何もしない */
function createSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  const existing = sheet.getRange(1, 1, 1, Math.max(headers.length, 1)).getValues()[0];
  if (existing[0] === headers[0]) return sheet;

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f1f3f4');
  sheet.setFrozenRows(1);
  return sheet;
}

/** レポートテンプレート用のシートを作成 */
function createReportTemplateSheet_(ss) {
  const name = CONFIG.SHEET_NAMES.REPORT_TPL;
  let sheet = ss.getSheetByName(name);
  if (sheet) return;

  sheet = ss.insertSheet(name);
  sheet.hideSheet();

  // レポートレイアウト
  const layout = [
    ['月次収支レポート / Monthly Property Statement'],
    [''],
    ['レポート番号:', '', '作成日:', ''],
    ['オーナー名:', '', '対象月:', ''],
    ['物件:', '', '物件数:', ''],
    [''],
    ['■ 収入の部'],
    ['項目', '詳細', '金額 (MYR)'],
    ['', '', ''],
    ['', '', ''],
    ['', '収入 合計', '=SUM(C9:C10)'],
    [''],
    ['■ 支出の部'],
    ['項目', '詳細', '金額 (MYR)'],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '', ''],
    ['', '支出 合計', '=SUM(C15:C19)'],
    [''],
    ['', '差引 オーナー様お受取額', '=C11-C20'],
  ];
  sheet.getRange(1, 1, layout.length, 3).setValues(layout);
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sheet.getRange(7, 1).setFontWeight('bold');
  sheet.getRange(13, 1).setFontWeight('bold');
  sheet.getRange(22, 2, 1, 2).setFontWeight('bold').setFontSize(12);
}

/** ダッシュボード（メニュー）シートを作成 */
function createDashboardSheet_(ss) {
  const name = CONFIG.SHEET_NAMES.MENU;
  let sheet = ss.getSheetByName(name);
  if (sheet) return;

  sheet = ss.insertSheet(name, 0);
  sheet.getRange('A1').setValue('不動産管理 AI自動化ダッシュボード')
    .setFontSize(18).setFontWeight('bold');
  sheet.getRange('A3').setValue('操作は上部メニュー「AI管理ツール」から実行してください。')
    .setFontColor('#5f6368');
  sheet.getRange('A5').setValue('■ ステータスサマリー').setFontWeight('bold');

  // サマリー表のヘッダー
  const summaryHeaders = ['指標', '件数'];
  sheet.getRange('A6:B6').setValues([summaryHeaders])
    .setFontWeight('bold').setBackground('#f1f3f4');
  sheet.getRange('A7').setValue('今月の請求書');
  sheet.getRange('A8').setValue('確認済');
  sheet.getRange('A9').setValue('要確認');
  sheet.getRange('A10').setValue('AI読取済');

  sheet.getRange('A14').setValue('■ 入金消込サマリー').setFontWeight('bold');
  sheet.getRange('A15:B15').setValues([['指標', '件数']])
    .setFontWeight('bold').setBackground('#f1f3f4');
  sheet.getRange('A16').setValue('消込済');
  sheet.getRange('A17').setValue('一部入金');
  sheet.getRange('A18').setValue('未入金');
  sheet.getRange('A19').setValue('要確認（未照合）');

  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 150);
}

/** サンプルのマスターデータを投入する */
function insertSampleMasterData_(ss) {
  var result = { properties: 0, owners: 0, tenants: 0 };

  // 物件マスター（先方のオーナーリスト準拠：物件名・部屋番号・オーナー名・住所）
  var propSheet = createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.PROPERTIES, CONFIG.PROPERTY_HEADERS);
  if (propSheet.getLastRow() <= 1) {
    var properties = [
      ['St Regis',             '46-5',    'Beslife Co., Ltd',          ''],
      ['Solaris Shop Lot',     '20-G',    'Hideo Ito Holdings Sdn Bhd', ''],
      ['Santuari Pantai Park', '29',      'Hideo Ito Holdings Sdn Bhd', ''],
      ['The Mews',             'A-32-2',  'Ito',                       ''],
      ['The Mews',             'A-36-1',  'Ito',                       ''],
      ['The Mews',             'A-26-3A', 'Sekizawa',                  ''],
      ['Vivo Residence',       'B-30-3A', 'Hidaka',                    ''],
      ['Vivo Residence',       'B-29-3A', 'Enomoto',                   ''],
      ['Vivo Residence',       'C-32-3',  'Kariya',                    ''],
      ['St Mary Residence',    'A2-7-3A', 'Joinbest International',     ''],
      ['St Mary Residence',    'A3-18-3', 'Yamaguchi',                 ''],
      ['St Mary Residence',    'A3-19-2', 'Itao',                      ''],
      ['M City',               '2-29-17', 'Oka',                       ''],
      ['Lucentia Residence',   'T1-19-3a','OG Proptech Sdn Bhd',       ''],
      ['Sentral Suites',       '3-28-C',  'Landex Project',            ''],
      ['Tropicana Gardens',    'B1-35-6', 'Murakami',                  ''],
      ['Face Platinum Suites', 'D-37-1',  'Nakano',                    ''],
      ['Residensi 22',         'A-33-3A', 'Tanaka',                    ''],
      ['Lucentia Residence',   'T1-38-8', 'Lee',                       ''],
      ['Lucentia Residence',   'T1-13-3A','Akiyama',                   ''],
    ];
    propSheet.getRange(2, 1, properties.length, properties[0].length).setValues(properties);
    result.properties = properties.length;
  }

  // 建物別名（名寄せ用）のサンプル。検出名 → 正規 建物名
  var aliasSheet = createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.BUILDING_ALIAS, CONFIG.BUILDING_ALIAS_HEADERS);
  if (aliasSheet.getLastRow() <= 1) {
    var aliases = [
      ['The Residences at St Regis', 'St Regis', 'サンプル：表記ゆれの名寄せ例'],
      ['Mews Suites',                'The Mews', 'サンプル：表記ゆれの名寄せ例'],
      ['PANGSAPURI SERVIS MEWS',     'The Mews', 'TNB等の請求書は住所にこの表記で出る'],
    ];
    aliasSheet.getRange(2, 1, aliases.length, aliases[0].length).setValues(aliases);
  }

  // オーナー連絡先マスター（オーナー名キー。メールは先方提供後に入力）
  var ownerSheet = createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.OWNERS, CONFIG.OWNER_HEADERS);
  if (ownerSheet.getLastRow() <= 1) {
    // 物件マスターから重複を除いたオーナー名を自動投入
    var seen = {};
    var owners = [];
    for (var op = 0; op < (typeof properties !== 'undefined' ? properties.length : 0); op++) {
      var on = String(properties[op][2] || '').trim();
      if (!on || seen[on]) continue;
      seen[on] = true;
      owners.push([on, '', '', 'レポート送付先メールを入力してください']);
    }
    if (owners.length > 0) {
      ownerSheet.getRange(2, 1, owners.length, owners[0].length).setValues(owners);
      result.owners = owners.length;
    }
  }

  // テナントマスター（先方フォーマット：物件名＋区画番号がキー。物件マスターと整合）
  var tenantSheet = createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.TENANTS, CONFIG.TENANT_HEADERS);
  if (tenantSheet.getLastRow() <= 1) {
    // 物件名, 区画番号, テナント名, 契約開始, 契約終了, 更新期間, 解約予告(月), 月間家賃, 敷金S, 敷金U, 備考, メール, WhatsApp番号
    var tenants = [
      ['The Mews',          'A-32-2',  'Tanaka Yuki',    '2025/01/01', '2026/12/31', '1年更新', 2, 3500, 3500, 0,   '', 'tanaka.yuki@example.com', '+60123456701'],
      ['St Mary Residence', 'A3-19-2', 'Lim Wei Ming',   '2025/03/01', '2026/02/28', '1年更新', 2, 2800, 2800, 0,   '', 'lim.weiming@example.com', '+60123456702'],
      ['M City',            '2-29-17', 'Siti Nurhaliza', '2024/06/01', '2026/05/31', '2年更新', 3, 4200, 4200, 0,   '', 'siti.n@example.com',      '+60123456703'],
      ['Vivo Residence',    'B-29-3A', 'Raj Kumar',      '2025/07/01', '2026/06/30', '1年更新', 2, 2500, 2500, 500, '', 'raj.kumar@example.com',   '+60123456704'],
    ];
    tenantSheet.getRange(2, 1, tenants.length, tenants[0].length).setValues(tenants);
    result.tenants = tenants.length;
  }

  return result;
}

/** デフォルトの設定値を投入する */
/** デフォルトの設定キー一覧（キー名・初期値） */
function defaultSettingsList_() {
  return [
    ['GEMINI_API_KEY',     '（ここに Google AI Studio の API キーを貼り付け）'],
    ['INVOICE_FOLDER_ID',  '（ここに Google Drive の請求書フォルダIDを貼り付け）'],
    ['REPORT_FOLDER_ID',   '（ここに Google Drive のレポート出力フォルダIDを貼り付け）'],
    ['DEPOSIT_FOLDER_ID',  '（ここに Google Drive の入金CSVフォルダIDを貼り付け）'],
    ['COMPANY_NAME',       'Sample Property Management Sdn Bhd'],
    ['COMPANY_EMAIL',      'info@example.com'],
    ['COMPANY_TEL',        '+60-xx-xxxx-xxxx'],
    ['MANAGEMENT_FEE_PCT', '8'],
  ];
}

function insertDefaultSettings_(ss) {
  var settingsSheet = createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.SETTINGS, CONFIG.SETTINGS_HEADERS);
  settingsSheet.setColumnWidth(1, 200);
  settingsSheet.setColumnWidth(2, 500);
  ensureSettingsKeys_(ss);
}

/** デフォルト設定キーのうち、未登録のものを「設定」シートに追記する */
function ensureSettingsKeys_(ss) {
  var settingsSheet = createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.SETTINGS, CONFIG.SETTINGS_HEADERS);
  var existing = {};
  if (settingsSheet.getLastRow() > 1) {
    var data = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) existing[data[i][0]] = true;
    }
  }

  var defaults = defaultSettingsList_();
  var toAdd = [];
  for (var j = 0; j < defaults.length; j++) {
    if (!existing[defaults[j][0]]) toAdd.push(defaults[j]);
  }
  if (toAdd.length > 0) {
    settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
  }
  return toAdd.length;
}
