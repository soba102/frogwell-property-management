/* =============================================================
 * 不動産管理 AI自動化モック ─ 統合コード
 * -------------------------------------------------------------
 * GAS オンラインエディタに貼り付ける用の1ファイル版です。
 * このファイルの中身を全てコピーして、Apps Script エディタの
 * 「コード.gs」（最初からあるファイル）に貼り付けてください。
 * ============================================================= */


/* =============================================================
 * Config ─ 設定値・定数・ヘルパー
 * ============================================================= */

const CONFIG = {
  SHEET_NAMES: {
    MENU:       'メニュー',
    INVOICES:   '請求書データ',
    PROPERTIES: 'マスター_物件',
    OWNERS:     'マスター_オーナー',
    TENANTS:    'マスター_テナント',
    BUILDING_ALIAS: 'マスター_建物別名',
    RECONCILE:  '入金消込',
    SETTINGS:   '設定',
    REPORT_TPL: '_レポートテンプレート',
  },

  // 末尾の「ユニット番号」「住所」は後から追加した列。
  // 既存コードが先頭からの列位置に依存しているため、必ず末尾に追加すること。
  INVOICE_HEADERS: [
    'ID', '請求日', '物件名', 'オーナー', '費目',
    '金額 (MYR)', '支払先', '元ファイルURL', 'ステータス', '備考',
    'ユニット番号', '住所',
  ],

  // 先方の「オーナーリスト」フォーマットに準拠。物件名＋部屋番号でオーナー名に紐づく（ID廃止）。
  PROPERTY_HEADERS: ['物件名', '部屋番号', 'オーナー名', '住所'],

  BUILDING_ALIAS_HEADERS: ['検出された建物名', '正規 建物名', '備考'],

  // オーナー連絡先マスター。オーナー名をキーに、レポート送付先メール等を持つ。
  OWNER_HEADERS: ['オーナー名', 'メール', '管理手数料率(%)', '備考'],

  // 先方のテナント契約情報一覧フォーマットに準拠。
  // 物件特定キーは「物件名＋区画番号」（請求書マッチングと同一キー）。
  // オーナーは「マスター_物件（物件名＋区画番号→オーナー）」経由で解決するため列を持たない。
  TENANT_HEADERS: [
    '物件名', '区画番号', 'テナント名', '契約開始', '契約終了',
    '更新期間', '解約予告(月)', '月間家賃 (RM)', '敷金Security (RM)', '敷金Utility (RM)', '備考',
    'メール', 'WhatsApp番号',
  ],

  RECONCILE_HEADERS: [
    '対象月', 'オーナー', '物件', 'テナント',
    '請求額 (MYR)', '入金額 (MYR)', '入金日', 'ステータス', '備考',
  ],

  SETTINGS_HEADERS: ['キー名', '値'],

  DRIVE_FOLDERS: {
    INVOICES_KEY: 'INVOICE_FOLDER_ID',
    REPORTS_KEY:  'REPORT_FOLDER_ID',
    DEPOSITS_KEY: 'DEPOSIT_FOLDER_ID',
  },

  GEMINI: {
    MODEL: 'gemini-2.5-flash-lite',
    API_KEY_SETTING: 'GEMINI_API_KEY',
  },

  STATUSES: {
    AI_READ:   'AI読取済',
    REVIEW:    '要確認',
    CONFIRMED: '確認済',
  },

  RECON_STATUSES: {
    RECONCILED: '消込済',
    UNPAID:     '未入金',
    PARTIAL:    '一部入金',
    OVERPAID:   '過入金',
    UNMATCHED:  '要確認',
  },
};

/** 「設定」シートから値を取得する */
function getSettingValue(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  if (!sheet) throw new Error('「設定」シートが見つかりません。Setup を実行してください。');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  throw new Error('設定キー "' + key + '" が見つかりません。設定シートに追加してください。');
}

/** Gemini API キーを取得 */
function getGeminiApiKey() {
  return getSettingValue(CONFIG.GEMINI.API_KEY_SETTING);
}

/** 請求書フォルダIDを取得 */
function getInvoiceFolderId() {
  return getSettingValue(CONFIG.DRIVE_FOLDERS.INVOICES_KEY);
}

/** レポート出力フォルダIDを取得 */
function getReportFolderId() {
  return getSettingValue(CONFIG.DRIVE_FOLDERS.REPORTS_KEY);
}

/** 入金CSVフォルダIDを取得 */
function getDepositFolderId() {
  return getSettingValue(CONFIG.DRIVE_FOLDERS.DEPOSITS_KEY);
}


/* =============================================================
 * Main ─ メニュー登録・エントリーポイント
 * ============================================================= */

/**
 * スプレッドシート起動時にカスタムメニューを追加する。
 * GAS の予約関数名 onOpen で自動実行される。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI管理ツール')
    .addItem('初期セットアップ', 'runSetup')
    .addItem('サンプルマスター投入', 'runInsertSampleMaster')
    .addSeparator()
    .addItem('OCR バッチ実行', 'runOcrBatch')
    .addItem('オーナー再紐付け', 'runReassignOwners')
    .addSeparator()
    .addItem('入金CSV取込（消込）', 'runImportDeposits')
    .addSeparator()
    .addItem('レポート生成（全オーナー）', 'runGenerateAllReports')
    .addItem('レポート生成（オーナー選択）', 'runGenerateSelectedReport')
    .addSeparator()
    .addItem('レポート一括メール送信', 'runSendAllReports')
    .addSeparator()
    .addItem('問い合わせ機能セットアップ', 'runSetupInquiry')
    .addItem('問い合わせ受信チェック（Gmail）', 'runProcessEmailInbox')
    .addItem('選択チケットに返信送信', 'runReplyActiveTicket')
    .addItem('選択チケットを完了にする', 'runMarkActiveTicketDone')
    .addSeparator()
    .addItem('ダッシュボード更新', 'runUpdateDashboard')
    .addToUi();
}

function runSetup() {
  setupAllSheets();
  SpreadsheetApp.getUi().alert(
    'セットアップ完了',
    'シート構造とサンプルデータを作成しました。\n「設定」シートに API キーとフォルダ ID を入力してください。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function runInsertSampleMaster() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = insertSampleMasterData_(ss);
    ensureSettingsKeys_(ss);
    ui.alert(
      'サンプルマスター投入完了',
      '物件: ' + result.properties + ' 件 / オーナー: ' + result.owners + ' 件 / テナント: ' + result.tenants + ' 件 を投入しました。\n'
        + '（既にデータがある場合はスキップされます）',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runOcrBatch() {
  const ui = SpreadsheetApp.getUi();
  try {
    const count = processNewInvoices();
    updateDashboard();
    ui.alert('OCR 完了', count + ' 件の請求書を処理しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runReassignOwners() {
  const ui = SpreadsheetApp.getUi();
  try {
    const count = reassignOwners();
    updateDashboard();
    ui.alert('オーナー再紐付け完了', count + ' 行のオーナーを更新しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runImportDeposits() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = importBankDeposits();
    updateDashboard();
    ui.alert(
      '入金消込 完了',
      '消込済: ' + result.reconciled + ' 件\n'
        + '一部入金: ' + result.partial + ' 件\n'
        + '未入金: ' + result.unpaid + ' 件\n'
        + '要確認（未照合）: ' + result.unmatched + ' 件\n\n'
        + '「入金消込」シートをご確認ください。',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runGenerateAllReports() {
  const ui = SpreadsheetApp.getUi();
  try {
    const count = generateAllOwnerReports();
    ui.alert('レポート生成完了', count + ' 件のレポートを生成しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runGenerateSelectedReport() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'オーナー選択',
    'レポートを生成するオーナー名を入力してください（例: Ito）:',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  try {
    const pdf = generateOwnerReport(result.getResponseText().trim());
    ui.alert('完了', 'レポートを生成しました: ' + pdf.getName(), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runSendAllReports() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'メール送信確認',
    '全オーナーにレポートをメール送信します。よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    const count = sendAllOwnerReports();
    ui.alert('送信完了', count + ' 件のメールを送信しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runUpdateDashboard() {
  updateDashboard();
  SpreadsheetApp.getUi().alert('ダッシュボードを更新しました。');
}

/* ===== 問い合わせ対応（Phase 2） ===== */

function runSetupInquiry() {
  const ui = SpreadsheetApp.getUi();
  try {
    setupInquirySheets();
    ui.alert('問い合わせ機能セットアップ完了',
      '「問い合わせ管理」「問い合わせ履歴」「ナレッジ_FAQ」シートと設定キーを用意しました。\n'
        + '「設定」シートで通知先・自動返信・Twilio（WhatsApp利用時）を入力してください。',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runProcessEmailInbox() {
  const ui = SpreadsheetApp.getUi();
  try {
    const count = processEmailInbox();
    ui.alert('問い合わせ受信チェック完了', count + ' 件の問い合わせを取り込みました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runReplyActiveTicket() {
  const ui = SpreadsheetApp.getUi();
  try {
    const id = replyActiveTicket();
    ui.alert('送信完了', id + ' に返信を送信しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function runMarkActiveTicketDone() {
  const ui = SpreadsheetApp.getUi();
  try {
    const id = markActiveTicketDone();
    ui.alert('完了', id + ' を完了にしました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}


/* =============================================================
 * Setup ─ シート作成・サンプルデータ投入
 * ============================================================= */

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

  const sheet1 = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
}

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

function createReportTemplateSheet_(ss) {
  const name = CONFIG.SHEET_NAMES.REPORT_TPL;
  let sheet = ss.getSheetByName(name);
  if (sheet) return;

  sheet = ss.insertSheet(name);
  sheet.hideSheet();

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
  // 既存の設定は保持しつつ、不足しているキーだけ追記する
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


/* =============================================================
 * DriveScanner ─ Drive フォルダの新規ファイル検知
 * ============================================================= */

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


/* =============================================================
 * OCR ─ Gemini 2.0 Flash API 呼出し
 * ============================================================= */

function extractInvoiceData(fileBlob, mimeType, fileName) {
  var apiKey = getGeminiApiKey();
  var base64Data = Utilities.base64Encode(fileBlob.getBytes());

  var prompt = buildOcrPrompt_(fileName);
  var payload = buildGeminiPayload_(prompt, base64Data, mimeType);

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + CONFIG.GEMINI.MODEL + ':generateContent?key=' + apiKey;

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  var response = fetchGeminiWithRetry_(url, options);
  var json = JSON.parse(response.getContentText());
  return parseGeminiResponse_(json);
}

/**
 * Gemini API を呼び出す。混雑(503)・レート制限(429)・一時障害(500)の場合は
 * 指数バックオフで自動リトライする。
 */
function fetchGeminiWithRetry_(url, options) {
  var maxAttempts = 4;
  var waitMs = 2000;
  var lastStatus = 0;
  var lastBody = '';

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    var response = UrlFetchApp.fetch(url, options);
    var status = response.getResponseCode();
    if (status === 200) return response;

    lastStatus = status;
    lastBody = response.getContentText();

    var retryable = (status === 503 || status === 429 || status === 500);
    if (!retryable || attempt === maxAttempts) break;

    Logger.log('Gemini API ' + status + ' のため ' + (waitMs / 1000) + '秒後にリトライ (' + attempt + '/' + (maxAttempts - 1) + ')');
    Utilities.sleep(waitMs);
    waitMs *= 2; // 2s → 4s → 8s
  }

  throw new Error('Gemini API エラー (' + lastStatus + '): ' + String(lastBody).substring(0, 200));
}

function buildOcrPrompt_(fileName) {
  return [
    'あなたは不動産管理会社の請求書を読み取るAIアシスタントです。',
    '添付された請求書の画像/PDFから、以下の情報を抽出してJSON形式で返してください。',
    '',
    '必ず以下のJSON形式で返してください（他のテキストは不要）:',
    '{',
    '  "invoice_date": "YYYY/MM/DD形式の請求日（見つからなければ空文字）",',
    '  "property_name": "建物名・物件名（例: The Residences at The St. Regis, The Mews。見つからなければ空文字）",',
    '  "unit_no": "ユニット番号・部屋番号・ロット番号（例: 46-5, A-36-1, A-32-2。見つからなければ空文字）",',
    '  "address": "物件の所在地住所（請求先や支払先の住所ではなく、対象物件の住所。見つからなければ空文字）",',
    '  "category": "費目カテゴリ（水道代/電気代/管理費/修繕費/保険料/その他）",',
    '  "amount": 数値（MYR、税込金額。通貨記号やカンマは除去）,',
    '  "payee": "支払先（請求元の会社名）",',
    '  "confidence": "high/medium/low（全体的な読取自信度）",',
    '  "notes": "特記事項（読取が不確かな箇所の説明等）"',
    '}',
    '',
    'ファイル名の参考情報: ' + fileName,
    '',
    '注意:',
    '- 金額は数値のみ（RM記号やカンマを含めない）',
    '- 日付が見つからない場合は空文字にする',
    '- property_name は建物名（部屋番号は含めない）、unit_no は部屋番号のみを入れる',
    '- unit_no は "Unit No."、"No. Lot"、"ALAMAT PREMIS" 等の近くにあることが多い',
    '- 物件名・部屋番号がファイル名に含まれていればそれも参考にする',
    '- マレー語の請求書もある（ALAMAT=住所, PANGSAPURI=アパート 等）',
    '- 読取に自信がない項目がある場合は confidence を low にし notes で説明する',
  ].join('\n');
}

function buildGeminiPayload_(prompt, base64Data, mimeType) {
  return {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType, data: base64Data } },
      ],
    }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
    },
  };
}

function parseGeminiResponse_(json) {
  try {
    var text = json.candidates[0].content.parts[0].text;
    var cleaned = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
    var data = JSON.parse(cleaned);

    return {
      invoice_date:  data.invoice_date  || '',
      property_name: data.property_name || '',
      unit_no:       data.unit_no       || '',
      address:       data.address       || '',
      category:      data.category      || 'その他',
      amount:        parseFloat(data.amount) || 0,
      payee:         data.payee         || '',
      confidence:    data.confidence    || 'medium',
      notes:         data.notes         || '',
    };
  } catch (e) {
    Logger.log('Gemini レスポンスのパースに失敗: ' + e.message);
    return {
      invoice_date: '', property_name: '', unit_no: '', address: '', category: 'その他',
      amount: 0, payee: '', confidence: 'low',
      notes: 'AI応答のパースに失敗しました。手動で入力してください。',
    };
  }
}


/* =============================================================
 * SheetWriter ─ OCR 結果のスプレッドシート書込み
 * ============================================================= */

function processNewInvoices() {
  var files = scanNewInvoiceFiles();
  if (files.length === 0) {
    Logger.log('新規ファイルなし');
    return 0;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVOICES);
  if (!sheet) throw new Error('「請求書データ」シートが見つかりません。');

  var nextId = getNextInvoiceId_(sheet);
  var count = 0;

  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    Logger.log('処理中: ' + file.name);

    try {
      var data = extractInvoiceData(file.blob, file.mimeType, file.name);
      // 物件名・ユニット番号はファイル名も補助に使う
      var unitNo = data.unit_no || extractUnitFromText_(file.name);
      var match = resolveOwner_(ss, data.property_name, unitNo, data.address);
      var status = mapConfidenceToStatus_(data.confidence);
      if (!match.ownerName) status = CONFIG.STATUSES.REVIEW;

      var row = [
        'INV-' + String(nextId).padStart(4, '0'),
        data.invoice_date,
        data.property_name,
        match.ownerName,
        data.category,
        data.amount,
        data.payee,
        file.url,
        status,
        match.note ? (data.notes ? data.notes + ' / ' + match.note : match.note) : data.notes,
        unitNo,
        data.address,
      ];

      sheet.appendRow(row);
      applyRowFormatting_(sheet, sheet.getLastRow(), status);
      nextId++;
      count++;
    } catch (e) {
      Logger.log('ファイル処理エラー (' + file.name + '): ' + e.message);
      sheet.appendRow([
        'INV-' + String(nextId).padStart(4, '0'),
        '', '', '', '', 0, '',
        file.url,
        CONFIG.STATUSES.REVIEW,
        'OCR処理エラー: ' + e.message,
        '', '',
      ]);
      nextId++;
      count++;
    }

    if (i < files.length - 1) Utilities.sleep(1000);
  }

  Logger.log(count + ' 件処理完了');
  return count;
}

function getNextInvoiceId_(sheet) {
  if (sheet.getLastRow() <= 1) return 1;

  var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var maxNum = 0;
  for (var i = 0; i < ids.length; i++) {
    var match = String(ids[i][0]).match(/INV-(\d+)/);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  return maxNum + 1;
}

/**
 * 物件名（建物名）＋部屋番号からオーナー名を特定する。
 * 建物名はエイリアス（別名）シートで名寄せし、一致しなければ別名シートに自動追記する。
 * @returns {Object} { ownerName, propertyName, note }
 */
function resolveOwner_(ss, buildingName, unitNo, address) {
  var result = { ownerName: '', propertyName: '', note: '' };

  var propSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROPERTIES);
  if (!propSheet) return result;

  var rawBuilding = buildingName || '';
  var canonical = canonicalBuildingName_(ss, rawBuilding, address);

  var props = propSheet.getDataRange().getValues();
  // ヘッダー: 物件名(0) 部屋番号(1) オーナー名(2) 住所(3)
  var bAndU = [];           // 建物＋ユニット 一致
  var buildingMatches = []; // 建物のみ一致
  var unitMatches = [];     // ユニットのみ一致（建物名は不一致）
  for (var i = 1; i < props.length; i++) {
    var masterBuilding = props[i][0];
    var bMatch = buildingNamesMatch_(canonical, masterBuilding);
    var uMatch = unitNo && props[i][1] && unitEquals_(unitNo, props[i][1]);
    if (bMatch) buildingMatches.push(props[i]);
    if (uMatch) unitMatches.push(props[i]);
    if (bMatch && uMatch) bAndU.push(props[i]);
  }

  var chosen = null;
  if (bAndU.length === 1) {
    chosen = bAndU[0];
  } else if (bAndU.length > 1) {
    result.note = '建物＋ユニットで複数候補あり。要確認';
  } else if (buildingMatches.length === 1) {
    chosen = buildingMatches[0];
    if (unitNo) result.note = 'ユニット未一致のため建物のみで紐付け';
  } else if (unitMatches.length === 1) {
    // 建物名が一致しなくても、ユニット番号がマスター全体で一意なら特定する。
    // （TNB等の請求書は建物名が住所表記でしか出ず、部屋番号はファイル名にしか無いケースに対応）
    chosen = unitMatches[0];
    result.note = '建物名は不一致だがユニット番号が一意のため特定';
  } else if (buildingMatches.length > 1) {
    result.note = '同一建物に複数物件あり、ユニット番号で特定できず。要確認';
  } else if (unitMatches.length > 1) {
    result.note = 'ユニット番号が複数物件に存在。建物名で特定できず。要確認';
  } else {
    if (rawBuilding) {
      registerBuildingAlias_(ss, rawBuilding);
      result.note = '建物名がマスター未登録。「マスター_建物別名」に追記しました';
    } else {
      result.note = '建物名・ユニット番号を特定できませんでした';
    }
    return result;
  }

  if (!chosen) return result;

  result.propertyName = chosen[0];
  result.ownerName = String(chosen[2] || '').trim();
  return result;
}

/** 旧シグネチャ互換（物件名のみ）。reassignOwners 等から呼ばれる */
function resolveOwnerName_(ss, propertyName, unitNo, address) {
  return resolveOwner_(ss, propertyName, unitNo || extractUnitFromText_(propertyName), address || '').ownerName;
}

/** 比較用に文字列を正規化（小文字化・記号と空白を除去） */
function normalizeForMatch_(s) {
  return String(s).toLowerCase().replace(/[\s,.\-_、。'"]/g, '');
}

/**
 * テキストからユニット番号らしき文字列を抽出する。
 * 例: "A-36-1", "46-5", "Unit 46-05", "A-32-2 Mews" → "A-36-1" 等
 */
function extractUnitFromText_(s) {
  if (!s) return '';
  // ハイフン連結トークンを抽出し、数字を含む最初のものを採用。
  // A-301, B-102, C-05, D-205, 46-5, A-36-1, A3-19-2, 2-29-17, T1-19-03A 等に対応。
  var tokens = String(s).match(/[A-Za-z0-9]+(?:-[A-Za-z0-9]+)+/g);
  if (!tokens) return '';
  for (var i = 0; i < tokens.length; i++) {
    if (/[0-9]/.test(tokens[i])) return tokens[i];
  }
  return '';
}

/**
 * ユニット番号が実質同じか判定する。
 * 大文字化・各セグメントのゼロ詰めを吸収（46-5 == 46-05, A-36-1 == a-36-01）。
 */
function unitEquals_(a, b) {
  return normalizeUnit_(a) === normalizeUnit_(b);
}

/** ユニット番号を正規化（区切りで分割し、数字の先頭ゼロを除去して再結合） */
function normalizeUnit_(s) {
  if (!s) return '';
  var segs = String(s).toUpperCase().replace(/\s/g, '').split(/[\-\/]/);
  var out = [];
  for (var i = 0; i < segs.length; i++) {
    var seg = segs[i];
    if (!seg) continue;
    seg = seg.replace(/^0+(\d)/, '$1').replace(/\b0+(\d)/g, '$1');
    out.push(seg);
  }
  return out.join('-');
}

/** 2つの建物名が一致するか（正規化して部分一致も許容） */
function buildingNamesMatch_(a, b) {
  var na = normalizeForMatch_(a);
  var nb = normalizeForMatch_(b);
  if (!na || !nb) return false;
  return na === nb || na.indexOf(nb) !== -1 || nb.indexOf(na) !== -1;
}

/**
 * 建物名をエイリアスシートで名寄せして正規名を返す。
 * 一致が無ければ入力された建物名をそのまま返す。
 */
function canonicalBuildingName_(ss, buildingName, address) {
  var name = buildingName || '';
  var aliasSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.BUILDING_ALIAS);
  if (aliasSheet && aliasSheet.getLastRow() > 1) {
    var data = aliasSheet.getDataRange().getValues();
    // ヘッダー: 検出された建物名(0) 正規 建物名(1) 備考(2)
    for (var i = 1; i < data.length; i++) {
      var detected = data[i][0];
      var canonical = data[i][1];
      if (!detected || !canonical) continue;
      if (buildingNamesMatch_(name, detected) ||
          (address && normalizeForMatch_(address).indexOf(normalizeForMatch_(detected)) !== -1)) {
        return canonical;
      }
    }
  }
  return name;
}

/**
 * 未登録の建物名を「マスター_建物別名」シートへ追記する（重複は追記しない）。
 * 正規名は空欄で起票し、人が後から正規建物名を入力する運用。
 */
function registerBuildingAlias_(ss, buildingName) {
  if (!buildingName) return;
  var aliasSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.BUILDING_ALIAS);
  if (!aliasSheet) {
    aliasSheet = ss.insertSheet(CONFIG.SHEET_NAMES.BUILDING_ALIAS);
    aliasSheet.getRange(1, 1, 1, CONFIG.BUILDING_ALIAS_HEADERS.length)
      .setValues([CONFIG.BUILDING_ALIAS_HEADERS]).setFontWeight('bold').setBackground('#f1f3f4');
    aliasSheet.setFrozenRows(1);
  }
  if (aliasSheet.getLastRow() > 1) {
    var existing = aliasSheet.getRange(2, 1, aliasSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (normalizeForMatch_(existing[i][0]) === normalizeForMatch_(buildingName)) return;
    }
  }
  aliasSheet.appendRow([buildingName, '', '未割当 - 正規の建物名を入力してください']);
}

/**
 * 既存の「請求書データ」行について、物件名からオーナーを再計算して埋め直す。
 * OCR をやり直さずにオーナー列だけを修正したいときに使う。
 * @returns {number} オーナーを更新した行数
 */
function reassignOwners() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVOICES);
  if (!sheet || sheet.getLastRow() <= 1) return 0;

  var propCol  = CONFIG.INVOICE_HEADERS.indexOf('物件名') + 1;
  var unitCol  = CONFIG.INVOICE_HEADERS.indexOf('ユニット番号') + 1;
  var addrCol  = CONFIG.INVOICE_HEADERS.indexOf('住所') + 1;
  var ownerCol = CONFIG.INVOICE_HEADERS.indexOf('オーナー') + 1;
  var lastRow = sheet.getLastRow();

  var propNames = sheet.getRange(2, propCol, lastRow - 1, 1).getValues();
  var unitNos   = sheet.getRange(2, unitCol, lastRow - 1, 1).getValues();
  var addresses = sheet.getRange(2, addrCol, lastRow - 1, 1).getValues();
  var updated = 0;

  for (var i = 0; i < propNames.length; i++) {
    var propertyName = propNames[i][0];
    if (!propertyName) continue;
    var unitNo = unitNos[i][0] || extractUnitFromText_(propertyName);
    var match = resolveOwner_(ss, propertyName, unitNo, addresses[i][0]);
    if (match.ownerName) {
      sheet.getRange(i + 2, ownerCol).setValue(match.ownerName);
      updated++;
    }
  }
  Logger.log('オーナー再紐付け: ' + updated + ' 行更新');
  return updated;
}

function mapConfidenceToStatus_(confidence) {
  switch (confidence) {
    case 'high':   return CONFIG.STATUSES.AI_READ;
    case 'medium': return CONFIG.STATUSES.AI_READ;
    case 'low':    return CONFIG.STATUSES.REVIEW;
    default:       return CONFIG.STATUSES.REVIEW;
  }
}

function applyRowFormatting_(sheet, rowNum, status) {
  var range = sheet.getRange(rowNum, 1, 1, CONFIG.INVOICE_HEADERS.length);
  switch (status) {
    case CONFIG.STATUSES.AI_READ:
      range.setBackground('#e8f0fe');
      break;
    case CONFIG.STATUSES.REVIEW:
      range.setBackground('#fef7e0');
      break;
    case CONFIG.STATUSES.CONFIRMED:
      range.setBackground('#e6f4ea');
      break;
  }
}


/* =============================================================
 * ReportGenerator ─ オーナー別レポート PDF 生成
 * ============================================================= */

function generateAllOwnerReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ownerSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OWNERS);
  var owners = ownerSheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < owners.length; i++) {
    var ownerName = String(owners[i][0] || '').trim();
    if (!ownerName) continue;
    try {
      generateOwnerReport(ownerName);
      count++;
    } catch (e) {
      Logger.log('レポート生成エラー (' + ownerName + '): ' + e.message);
    }
    if (i < owners.length - 1) Utilities.sleep(500);
  }
  return count;
}

function generateOwnerReport(ownerName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ownerName = String(ownerName || '').trim();
  var ownerInfo = getOwnerInfo_(ss, ownerName);
  var properties = getOwnerProperties_(ss, ownerName);
  var invoices = getOwnerInvoices_(ss, ownerName);

  var now = new Date();
  var yearMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月');
  var slug = ownerName.replace(/[^A-Za-z0-9]/g, '').substr(0, 10) || 'OWNER';
  var reportId = 'RPT-' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM') + '-' + slug;

  var tempSheet = createTempReportSheet_(ss, ownerInfo, properties, invoices, yearMonth, reportId);
  SpreadsheetApp.flush();

  var pdf = exportSheetAsPdf_(ss, tempSheet, ownerInfo.name + '様_' + yearMonth);

  ss.deleteSheet(tempSheet);

  var folderId = getReportFolderId();
  var folder = DriveApp.getFolderById(folderId);
  var savedFile = folder.createFile(pdf);

  Logger.log('レポート生成完了: ' + savedFile.getName());
  return savedFile;
}

function getOwnerInfo_(ss, ownerName) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OWNERS);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === ownerName) {
        return {
          name: ownerName,
          email: data[i][1],
          feePct: data[i][2] !== '' && data[i][2] != null ? parseFloat(data[i][2]) : null,
        };
      }
    }
  }
  return { name: ownerName, email: '', feePct: null };
}

function getOwnerProperties_(ss, ownerName) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROPERTIES);
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2] || '').trim() === ownerName) {
      results.push({ name: data[i][0], unit: data[i][1], address: data[i][3] || '' });
    }
  }
  return results;
}

function getOwnerInvoices_(ss, ownerName) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVOICES);
  if (sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var results = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[3] === ownerName) {
      results.push({
        date: row[1],
        property: row[2],
        category: row[4],
        amount: row[5],
        payee: row[6],
        status: row[8],
      });
    }
  }
  return results;
}

/**
 * オーナーの収入明細行を組み立てる。
 * テナントマスターの賃料を基準に、入金消込シートに実入金があれば
 * 入金状況（消込済/一部入金/未入金）も詳細に併記する。
 * @returns {Array} [項目, 詳細, 金額] の配列
 */
function buildOwnerIncomeLines_(ss, ownerName, yearMonth) {
  var tenantSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TENANTS);
  if (!tenantSheet || tenantSheet.getLastRow() <= 1) return [];

  // オーナーが保有する物件（建物名・部屋番号）を収集
  var propSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROPERTIES);
  var ownerUnits = [];
  if (propSheet && propSheet.getLastRow() > 1) {
    var pdata = propSheet.getDataRange().getValues();
    // 物件名(0) 部屋番号(1) オーナー名(2) 住所(3)
    for (var p = 1; p < pdata.length; p++) {
      if (String(pdata[p][2] || '').trim() === ownerName) {
        ownerUnits.push({ building: pdata[p][0], unit: pdata[p][1] });
      }
    }
  }
  if (ownerUnits.length === 0) return [];

  // テナント（物件名＋区画番号）をオーナーの物件に突合
  var tdata = tenantSheet.getDataRange().getValues();
  var ownerTenants = [];
  for (var i = 1; i < tdata.length; i++) {
    var tb = tdata[i][0], tu = tdata[i][1];
    if (!tb && !tu) continue;
    for (var u = 0; u < ownerUnits.length; u++) {
      var uniMatch = !ownerUnits[u].unit || !tu || unitEquals_(ownerUnits[u].unit, tu);
      if (buildingNamesMatch_(ownerUnits[u].building, tb) && uniMatch) {
        ownerTenants.push({ name: tdata[i][2], rent: parseFloat(tdata[i][7]) || 0 });
        break;
      }
    }
  }
  if (ownerTenants.length === 0) return [];

  // 消込シートからテナント名→入金状況のマップを作る
  var paidMap = {};
  var reconSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RECONCILE);
  if (reconSheet && reconSheet.getLastRow() > 1) {
    var rdata = reconSheet.getDataRange().getValues();
    var tenantCol = CONFIG.RECONCILE_HEADERS.indexOf('テナント');
    var paidCol   = CONFIG.RECONCILE_HEADERS.indexOf('入金額 (MYR)');
    var statusCol = CONFIG.RECONCILE_HEADERS.indexOf('ステータス');
    for (var r = 1; r < rdata.length; r++) {
      var tn = rdata[r][tenantCol];
      if (tn) paidMap[tn] = { paid: parseFloat(rdata[r][paidCol]) || 0, status: rdata[r][statusCol] };
    }
  }

  var lines = [];
  for (var t = 0; t < ownerTenants.length; t++) {
    var ten = ownerTenants[t];
    var info = paidMap[ten.name];
    var detail = yearMonth + '分賃料';
    if (info) detail += '（' + info.status + '）';
    lines.push(['賃料収入: ' + ten.name, detail, ten.rent]);
  }
  return lines;
}

function createTempReportSheet_(ss, ownerInfo, properties, invoices, yearMonth, reportId) {
  var sheet = ss.insertSheet('_tmp_report');
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日');
  var propertyLabels = properties.map(function(p) {
    return p.unit ? (p.name + ' ' + p.unit) : p.name;
  });
  var propertyName = propertyLabels.length > 0 ? propertyLabels.join(' / ') : '─';
  var propertyCount = properties.length;

  var feePct = 8;
  try { feePct = parseFloat(getSettingValue('MANAGEMENT_FEE_PCT')); } catch (e) {}
  if (ownerInfo.feePct != null && !isNaN(ownerInfo.feePct)) feePct = ownerInfo.feePct;

  // 収入: テナントマスター（このオーナーの賃料）＋消込の実入金状況
  var incomeLines = buildOwnerIncomeLines_(ss, ownerInfo.name, yearMonth);
  var totalIncome = 0;
  for (var n = 0; n < incomeLines.length; n++) totalIncome += parseFloat(incomeLines[n][2]) || 0;
  if (incomeLines.length === 0) {
    incomeLines.push(['賃料収入', yearMonth + '分家賃（テナント未登録）', 0]);
  }

  var expenses = [];
  var totalExpense = 0;
  for (var i = 0; i < invoices.length; i++) {
    expenses.push([invoices[i].category, invoices[i].payee + '（' + invoices[i].date + '）', invoices[i].amount]);
    totalExpense += parseFloat(invoices[i].amount) || 0;
  }
  var managementFee = Math.round(totalIncome * feePct / 100);
  expenses.push(['管理手数料', '賃料の' + feePct + '%', managementFee]);
  totalExpense += managementFee;

  var netAmount = totalIncome - totalExpense;

  var rows = [];
  rows.push(['月次収支レポート / Monthly Property Statement', '', '']);
  rows.push(['', '', '']);
  rows.push(['レポート番号:', reportId, '作成日: ' + dateStr]);
  rows.push(['オーナー名:', ownerInfo.name + ' 様', '対象月: ' + yearMonth]);
  rows.push(['物件:', propertyName, '物件数: ' + propertyCount]);
  rows.push(['', '', '']);

  rows.push(['■ 収入の部', '', '']);
  rows.push(['項目', '詳細', '金額 (MYR)']);
  for (var m = 0; m < incomeLines.length; m++) rows.push(incomeLines[m]);
  rows.push(['', '収入 合計', totalIncome]);
  rows.push(['', '', '']);

  var incomeHeaderRow = 7;
  var expenseHeaderRow = rows.length + 1;

  rows.push(['■ 支出の部', '', '']);
  rows.push(['項目', '詳細', '金額 (MYR)']);
  for (var j = 0; j < expenses.length; j++) {
    rows.push(expenses[j]);
  }
  rows.push(['', '支出 合計', totalExpense]);
  rows.push(['', '', '']);
  rows.push(['', '差引 オーナー様お受取額', 'MYR ' + netAmount.toLocaleString()]);

  rows.push(['', '', '']);
  rows.push(['※ 本レポートは自動生成されています。', '', '']);

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);

  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sheet.getRange(incomeHeaderRow, 1).setFontWeight('bold');
  sheet.getRange(expenseHeaderRow, 1).setFontWeight('bold');
  sheet.getRange(rows.length - 2, 2, 1, 2).setFontWeight('bold').setFontSize(12);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 150);

  return sheet;
}

function exportSheetAsPdf_(ss, sheet, fileNameBase) {
  var ssId = ss.getId();
  var sheetId = sheet.getSheetId();

  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export'
    + '?format=pdf'
    + '&gid=' + sheetId
    + '&size=A4'
    + '&portrait=true'
    + '&fitw=true'
    + '&gridlines=false'
    + '&printtitle=false'
    + '&sheetnames=false'
    + '&pagenum=UNDEFINED'
    + '&fzr=false';

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF エクスポートに失敗しました: ' + response.getResponseCode());
  }

  return response.getBlob().setName('レポート_' + fileNameBase + '.pdf');
}


/* =============================================================
 * EmailSender ─ レポートメール送信
 * ============================================================= */

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


/* =============================================================
 * Dashboard ─ ダッシュボード（サマリー）更新
 * ============================================================= */

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

  dashboard.getRange('A12').setValue(
    '最終更新: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
  ).setFontColor('#80868b').setFontSize(10);
}


/* =============================================================
 * Reconciliation ─ 銀行CSV取込・名寄せ・入金消込
 * ============================================================= */

/**
 * 入金CSVフォルダの未処理CSVを読み込み、テナント賃料と突合して
 * 「入金消込」シートを作り直す。
 * @returns {Object} 件数サマリー {reconciled, partial, unpaid, overpaid, unmatched}
 */
function importBankDeposits() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var deposits = readDepositsFromFolder_();
  if (deposits.length === 0) {
    throw new Error('入金CSVが見つかりません。入金CSVフォルダにCSVファイルを置いてください。');
  }

  var tenants = readTenants_(ss);
  if (tenants.length === 0) {
    throw new Error('テナントマスターが空です。先に「サンプルマスター投入」を実行してください。');
  }

  // 対象月: 最初の入金日の年月（なければ今月）
  var targetMonth = deposits.length > 0 && deposits[0].date
    ? Utilities.formatDate(deposits[0].date, 'Asia/Tokyo', 'yyyy/MM')
    : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');

  // 各テナントへ入金をマッチング（摘要から区画番号を抽出して突合）
  var paidByTenant = {};      // tenant index -> {amount, date}
  var unmatchedDeposits = []; // 区画番号は読めたがマスター該当無しの入金
  for (var d = 0; d < deposits.length; d++) {
    var dep = deposits[d];
    var depUnit = extractUnitFromText_(dep.name);
    var idx = matchDepositByUnit_(depUnit, dep.name, tenants);
    if (idx >= 0) {
      if (!paidByTenant[idx]) paidByTenant[idx] = { amount: 0, date: dep.date };
      paidByTenant[idx].amount += dep.amount;
      if (dep.date) paidByTenant[idx].date = dep.date;
    } else if (depUnit) {
      unmatchedDeposits.push(dep);
    }
    // 区画番号が読めない入金は家賃以外とみなして無視（手数料・JOMPAY等のノイズ除外）
  }

  // 消込結果を組み立て
  var rows = [];
  var summary = { reconciled: 0, partial: 0, unpaid: 0, overpaid: 0, unmatched: 0 };

  for (var t = 0; t < tenants.length; t++) {
    var ten = tenants[t];
    var paid = paidByTenant[t] ? paidByTenant[t].amount : 0;
    var paidDate = paidByTenant[t] && paidByTenant[t].date
      ? Utilities.formatDate(paidByTenant[t].date, 'Asia/Tokyo', 'yyyy/MM/dd') : '';
    var expected = ten.rent;
    var status, note;

    if (paid === 0) {
      status = CONFIG.RECON_STATUSES.UNPAID;
      note = '入金が確認できません';
      summary.unpaid++;
    } else if (paid === expected) {
      status = CONFIG.RECON_STATUSES.RECONCILED;
      note = '自動マッチ';
      summary.reconciled++;
    } else if (paid < expected) {
      status = CONFIG.RECON_STATUSES.PARTIAL;
      note = '差額 RM ' + (expected - paid).toLocaleString();
      summary.partial++;
    } else {
      status = CONFIG.RECON_STATUSES.OVERPAID;
      note = '過入金 RM ' + (paid - expected).toLocaleString();
      summary.overpaid++;
    }

    rows.push([
      targetMonth, ten.ownerName, ten.propertyName, ten.name,
      expected, paid, paidDate, status, note,
    ]);
  }

  // 名義不一致の入金（誰の入金か特定できないもの）
  for (var u = 0; u < unmatchedDeposits.length; u++) {
    var ud = unmatchedDeposits[u];
    rows.push([
      targetMonth, '', '', ud.name,
      '', ud.amount,
      ud.date ? Utilities.formatDate(ud.date, 'Asia/Tokyo', 'yyyy/MM/dd') : '',
      CONFIG.RECON_STATUSES.UNMATCHED, '区画番号がマスターと一致しません',
    ]);
    summary.unmatched++;
  }

  writeReconciliation_(ss, rows);
  Logger.log('消込完了: ' + JSON.stringify(summary));
  return summary;
}

/** 入金CSVフォルダ内の全CSVを読み込み、入金明細の配列を返す */
function readDepositsFromFolder_() {
  var folderId = getDepositFolderId();
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var deposits = [];

  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName().toLowerCase();
    if (name.indexOf('.csv') === -1) continue;

    var content = file.getBlob().getDataAsString('UTF-8');
    var parsed = parseDepositCsv_(content);
    for (var i = 0; i < parsed.length; i++) deposits.push(parsed[i]);
  }
  Logger.log('入金明細: ' + deposits.length + ' 件');
  return deposits;
}

/**
 * 銀行CSVのテキストを解析して入金明細を返す。
 * ヘッダー行から「日付」「摘要(名義)」「入金額」の列を推測する。
 * 入金額(クレジット)が正の行のみ対象。
 */
function parseDepositCsv_(content) {
  var table = Utilities.parseCsv(content);
  if (!table || table.length < 2) return [];

  var header = table[0].map(function(h) { return String(h).toLowerCase().trim(); });

  var dateIdx = findColumnIndex_(header, ['date', '日付', '取引日', '取引日付', '年月日']);
  var descIdx = findColumnIndex_(header, ['description', '摘要', '内容', 'お取扱内容', '振込名義', '振込人', 'narrative']);
  var creditIdx = findColumnIndex_(header, ['credit', '入金', '入金額', 'お預入れ', 'お預入れ金額', 'deposit', '預り金']);
  var amountIdx = findColumnIndex_(header, ['amount', '金額']);

  var results = [];
  for (var r = 1; r < table.length; r++) {
    var row = table[r];
    if (!row || row.length === 0) continue;

    var name = descIdx >= 0 ? String(row[descIdx] || '').trim() : '';
    var rawAmount = '';
    if (creditIdx >= 0) rawAmount = row[creditIdx];
    else if (amountIdx >= 0) rawAmount = row[amountIdx];

    var amount = parseAmount_(rawAmount);
    if (!amount || amount <= 0) continue; // 入金(正)の行のみ

    var dateVal = dateIdx >= 0 ? parseDate_(row[dateIdx]) : null;

    results.push({ date: dateVal, name: name, amount: amount });
  }
  return results;
}

/** ヘッダー候補から最初に一致する列のインデックスを返す */
function findColumnIndex_(header, candidates) {
  for (var c = 0; c < candidates.length; c++) {
    for (var h = 0; h < header.length; h++) {
      if (header[h].indexOf(candidates[c]) !== -1) return h;
    }
  }
  return -1;
}

/** 金額文字列を数値化（通貨記号・カンマ・空白を除去） */
function parseAmount_(s) {
  if (s === null || s === undefined) return 0;
  var cleaned = String(s).replace(/[^0-9.\-]/g, '');
  var num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

/** 日付文字列を Date 化（yyyy/mm/dd, yyyy-mm-dd 等に対応） */
function parseDate_(s) {
  if (!s) return null;
  var str = String(s).trim().replace(/[.\-]/g, '/');
  var m = str.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  var d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * テナントマスター（先方フォーマット）を読み込む。
 * 物件名＋区画番号から「マスター_物件」経由でオーナーを解決して付与する。
 * ヘッダー: 物件名(0) 区画番号(1) テナント名(2) 契約開始(3) 契約終了(4)
 *          更新期間(5) 解約予告(6) 月間家賃(7) 敷金S(8) 敷金U(9) 備考(10)
 */
function readTenants_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TENANTS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var tenants = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var building = row[0];
    var unit = row[1];
    if (!building && !unit) continue;
    var resolved = resolveOwner_(ss, building, unit, '');
    tenants.push({
      building: building,
      unit: unit,
      name: row[2],
      rent: parseFloat(row[7]) || 0,
      ownerName: resolved.ownerName || '',
      propertyName: building,
    });
  }
  return tenants;
}

/**
 * 入金（銀行明細の摘要）から抽出した区画番号でテナントを特定する。
 * 同一区画番号が複数物件に存在する場合は、摘要内の建物名で絞り込む。
 * @returns {number} 一致したテナントの配列インデックス（無ければ -1）
 */
function matchDepositByUnit_(depUnit, depDesc, tenants) {
  if (!depUnit) return -1;
  var candidates = [];
  for (var i = 0; i < tenants.length; i++) {
    if (tenants[i].unit && unitEquals_(depUnit, tenants[i].unit)) candidates.push(i);
  }
  if (candidates.length === 1) return candidates[0];
  if (candidates.length > 1) {
    var desc = normalizeForMatch_(depDesc);
    for (var c = 0; c < candidates.length; c++) {
      var b = tenants[candidates[c]].building;
      if (b && desc.indexOf(normalizeForMatch_(b)) !== -1) return candidates[c];
    }
  }
  return -1;
}

/** 消込結果を「入金消込」シートに書き出す（ヘッダー以下を全置換） */
function writeReconciliation_(ss, rows) {
  var sheet = createSheetWithHeaders_(ss, CONFIG.SHEET_NAMES.RECONCILE, CONFIG.RECONCILE_HEADERS);

  // 既存データをクリア（ヘッダーは残す）
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.RECONCILE_HEADERS.length).clearContent();
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, CONFIG.RECONCILE_HEADERS.length).setBackground(null);
  }
  if (rows.length === 0) return;

  sheet.getRange(2, 1, rows.length, CONFIG.RECONCILE_HEADERS.length).setValues(rows);

  // ステータスごとに色付け
  var statusCol = CONFIG.RECONCILE_HEADERS.indexOf('ステータス');
  for (var i = 0; i < rows.length; i++) {
    var color = null;
    switch (rows[i][statusCol]) {
      case CONFIG.RECON_STATUSES.RECONCILED: color = '#e6f4ea'; break;
      case CONFIG.RECON_STATUSES.PARTIAL:    color = '#fef7e0'; break;
      case CONFIG.RECON_STATUSES.UNPAID:     color = '#fce8e6'; break;
      case CONFIG.RECON_STATUSES.OVERPAID:   color = '#fef7e0'; break;
      case CONFIG.RECON_STATUSES.UNMATCHED:  color = '#f3e8fd'; break;
    }
    if (color) {
      sheet.getRange(i + 2, 1, 1, CONFIG.RECONCILE_HEADERS.length).setBackground(color);
    }
  }
}


/* =============================================================
 * Inquiry ─ 問い合わせ対応自動化（Phase 2）
 * -------------------------------------------------------------
 * チャネル非依存の中核（受信 → 分類 → 照合 → チケット → 通知 → 返信）と、
 * メール（Gmail）／WhatsApp（Twilio）の入出力アダプタ。
 * 設計の正は docs/05_問い合わせ自動化_企画.md。
 * ============================================================= */

var INQUIRY = {
  SHEETS: {
    TICKETS: '問い合わせ管理',
    HISTORY: '問い合わせ履歴',
    FAQ:     'ナレッジ_FAQ',
  },
  TICKET_HEADERS: [
    '受付ID', '受付日時', 'チャネル', '送信者名', '送信者連絡先', '件名', '本文',
    '分類', '緊急度', '物件名', '区画番号', 'テナント名', 'オーナー名', '物件候補(AI)',
    '担当', 'ステータス', '返信モード', 'スレッドID', '要約', '返信文(編集可)', '備考',
  ],
  HISTORY_HEADERS: ['履歴ID', '受付ID', '日時', '方向', '送信者', '本文', '添付URL'],
  FAQ_HEADERS: ['カテゴリ', 'キーワード', '模範回答', '自動送信可', '最終更新'],
  SETTINGS: {
    GMAIL_QUERY:  'INQUIRY_GMAIL_QUERY',
    NOTIFY_EMAIL: 'INQUIRY_NOTIFY_EMAIL',
    AUTO_REPLY:   'INQUIRY_AUTO_REPLY',
    TW_SID:       'TWILIO_ACCOUNT_SID',
    TW_TOKEN:     'TWILIO_AUTH_TOKEN',
    TW_FROM:      'TWILIO_WHATSAPP_FROM',
  },
  STATUS: {
    NEW: '新規', IN_PROGRESS: '対応中', WAIT_APPROVAL: '承認待ち', DONE: '完了', REVIEW: '要確認',
  },
  REPLY_MODE: { AUTO: '自動送信済', DRAFT: '下書き待ち', SENT: '送信済' },
  CATEGORIES: ['緊急', '修繕', '契約', '支払い', '近隣', 'FAQ', '空き照会', 'その他'],
  AUTO_REPLY_CATEGORIES: ['FAQ', '空き照会'],
  GMAIL_MAX_THREADS: 20,
  BODY_MAX: 5000,
};

function setupInquirySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  createSheetWithHeaders_(ss, INQUIRY.SHEETS.TICKETS, INQUIRY.TICKET_HEADERS);
  createSheetWithHeaders_(ss, INQUIRY.SHEETS.HISTORY, INQUIRY.HISTORY_HEADERS);
  var faq = createSheetWithHeaders_(ss, INQUIRY.SHEETS.FAQ, INQUIRY.FAQ_HEADERS);
  if (faq.getLastRow() <= 1) {
    var now = new Date();
    faq.getRange(2, 1, 4, INQUIRY.FAQ_HEADERS.length).setValues([
      ['FAQ',    'ゴミ 出し 回収',  'ゴミは可燃が火・木、リサイクルが土曜の朝8時までに所定の集積所へお出しください。', 'YES', now],
      ['FAQ',    '駐車 場 来客',    '来客用駐車場は管理事務所で受付後、B1Fをご利用いただけます（1日まで）。',          'YES', now],
      ['空き照会', '空室 空き 募集',  '現在の空室状況は「マスター_空き物件」をご確認のうえ、最新情報をご案内します。',     'NO',  now],
      ['内見',    '内見 鍵 visit',  '内見は平日10-17時に受付。鍵は管理事務所保管のため、前日までにご予約ください。',     'NO',  now],
    ]);
  }
  ensureInquirySettings_(ss);
  ensureTenantContactColumns_(ss);
}

function ensureInquirySettings_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  if (!sheet) return;
  var existing = {};
  if (sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) if (data[i][0]) existing[data[i][0]] = true;
  }
  var defaults = [
    [INQUIRY.SETTINGS.GMAIL_QUERY,  'label:inquiry is:unread'],
    [INQUIRY.SETTINGS.NOTIFY_EMAIL, '（担当者の通知先メールを入力）'],
    [INQUIRY.SETTINGS.AUTO_REPLY,   'false'],
    [INQUIRY.SETTINGS.TW_SID,       '（Twilio Account SID。WhatsApp利用時に入力）'],
    [INQUIRY.SETTINGS.TW_TOKEN,     '（Twilio Auth Token）'],
    [INQUIRY.SETTINGS.TW_FROM,      'whatsapp:+14155238886'],
  ];
  var toAdd = [];
  for (var j = 0; j < defaults.length; j++) if (!existing[defaults[j][0]]) toAdd.push(defaults[j]);
  if (toAdd.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
}

function ensureTenantContactColumns_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TENANTS);
  if (!sheet) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var needed = ['メール', 'WhatsApp番号'];
  for (var i = 0; i < needed.length; i++) {
    if (headers.indexOf(needed[i]) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1)
        .setValue(needed[i]).setFontWeight('bold').setBackground('#f1f3f4');
    }
  }
}

function handleInbound_(ss, inbound) {
  var cls = classifyInquiry_(inbound);
  var match = matchInquiry_(ss, inbound, cls);
  var status = match.confident ? INQUIRY.STATUS.NEW : INQUIRY.STATUS.REVIEW;
  var ticket = createTicket_(ss, inbound, cls, match, status);
  appendHistory_(ss, ticket.id, '受信', inbound.fromName || inbound.fromAddress, inbound.body,
    (inbound.attachments && inbound.attachments[0] ? inbound.attachments[0].url : ''));
  notifyStaff_(ss, ticket, inbound, cls, match);
  var autoOn = String(getSettingSafe_(INQUIRY.SETTINGS.AUTO_REPLY)).toLowerCase() === 'true';
  var canAuto = autoOn && match.confident && cls.auto_reply_ok
    && INQUIRY.AUTO_REPLY_CATEGORIES.indexOf(cls.category) >= 0;
  if (canAuto) {
    var sent = dispatchReply_(ss, ticket.row, cls.suggested_reply_ja);
    if (sent) {
      appendHistory_(ss, ticket.id, '送信', '(自動返信)', cls.suggested_reply_ja, '');
      setTicketField_(ss, ticket.row, '返信モード', INQUIRY.REPLY_MODE.AUTO);
      setTicketField_(ss, ticket.row, 'ステータス', INQUIRY.STATUS.DONE);
    }
  } else {
    var ack = '【自動受付】お問い合わせを受け付けました。担当者が内容を確認し、追ってご連絡いたします。';
    var ackSent = dispatchReply_(ss, ticket.row, ack);
    if (ackSent) appendHistory_(ss, ticket.id, '送信', '(自動受付)', ack, '');
    setTicketField_(ss, ticket.row, '返信モード', INQUIRY.REPLY_MODE.DRAFT);
  }
  return ticket;
}

function classifyInquiry_(inbound) {
  var prompt = [
    'あなたは不動産管理会社の問い合わせ対応AIです。',
    '次の問い合わせメッセージを読み、JSONのみで返してください（前後に文章を付けない）。',
    '',
    '分類カテゴリは次から1つ: ' + INQUIRY.CATEGORIES.join(' / '),
    '緊急度は 高 / 中 / 低 のいずれか。水漏れ・停電・締め出し等は「緊急」「高」。',
    '',
    '出力JSON:',
    '{',
    '  "category": "カテゴリ",',
    '  "urgency": "高/中/低",',
    '  "property_hint": "本文から推測できる建物名（無ければ空文字）",',
    '  "unit_hint": "本文から推測できる部屋番号（例 A-36-1。無ければ空文字）",',
    '  "tenant_hint": "差出人名や本文中の名前（無ければ空文字）",',
    '  "summary_ja": "30字程度の日本語要約",',
    '  "suggested_reply_ja": "丁寧な日本語の返信文の下書き",',
    '  "auto_reply_ok": true または false（定型のFAQ/空き照会で自信がある場合のみtrue）',
    '}',
    '',
    '--- 件名 ---',
    String(inbound.subject || ''),
    '--- 本文 ---',
    String(inbound.body || '').substring(0, INQUIRY.BODY_MAX),
  ].join('\n');
  var fallback = {
    category: 'その他', urgency: '中', property_hint: '', unit_hint: '', tenant_hint: '',
    summary_ja: String(inbound.subject || inbound.body || '').substring(0, 30),
    suggested_reply_ja: 'お問い合わせありがとうございます。担当者より追ってご連絡いたします。',
    auto_reply_ok: false,
  };
  try {
    var data = callGeminiText_(prompt);
    return {
      category: data.category || 'その他',
      urgency: data.urgency || '中',
      property_hint: data.property_hint || '',
      unit_hint: data.unit_hint || '',
      tenant_hint: data.tenant_hint || '',
      summary_ja: data.summary_ja || fallback.summary_ja,
      suggested_reply_ja: data.suggested_reply_ja || fallback.suggested_reply_ja,
      auto_reply_ok: data.auto_reply_ok === true,
    };
  } catch (e) {
    Logger.log('分類失敗（フォールバック使用）: ' + e.message);
    return fallback;
  }
}

function matchInquiry_(ss, inbound, cls) {
  var res = { property: '', unit: '', tenant: '', owner: '', candidate: '', note: '', confident: false };
  var tenant = findTenantByContact_(ss, inbound.fromAddress);
  if (tenant) {
    res.property = tenant.property; res.unit = tenant.unit; res.tenant = tenant.name;
    res.owner = resolveOwner_(ss, tenant.property, tenant.unit, '').ownerName;
    res.confident = true; res.note = '差出人の連絡先からテナント特定';
    return res;
  }
  res.candidate = [cls.property_hint, cls.unit_hint].filter(function (x) { return x; }).join(' / ');
  if (cls.property_hint || cls.unit_hint) {
    var m = resolveOwner_(ss, cls.property_hint || '', cls.unit_hint || '', '');
    if (m.ownerName) {
      res.property = m.propertyName; res.unit = cls.unit_hint || ''; res.owner = m.ownerName;
      res.confident = true; res.note = '本文の手がかりから特定' + (m.note ? '（' + m.note + '）' : '');
      return res;
    }
    res.note = '本文の手がかりからは一意に特定できず（候補のみ）';
  } else {
    res.note = '物件の手がかりなし';
  }
  return res;
}

function createTicket_(ss, inbound, cls, match, status) {
  var sheet = ss.getSheetByName(INQUIRY.SHEETS.TICKETS)
    || createSheetWithHeaders_(ss, INQUIRY.SHEETS.TICKETS, INQUIRY.TICKET_HEADERS);
  var id = nextSeqId_(sheet, 'INQ-', '受付ID');
  var row = {};
  row['受付ID'] = id;
  row['受付日時'] = inbound.receivedAt || new Date();
  row['チャネル'] = inbound.channel;
  row['送信者名'] = inbound.fromName || '';
  row['送信者連絡先'] = inbound.fromAddress || '';
  row['件名'] = inbound.subject || '';
  row['本文'] = String(inbound.body || '').substring(0, INQUIRY.BODY_MAX);
  row['分類'] = cls.category;
  row['緊急度'] = cls.urgency;
  row['物件名'] = match.property;
  row['区画番号'] = match.unit;
  row['テナント名'] = match.tenant;
  row['オーナー名'] = match.owner;
  row['物件候補(AI)'] = match.candidate;
  row['担当'] = '';
  row['ステータス'] = status;
  row['返信モード'] = '';
  row['スレッドID'] = inbound.threadId || '';
  row['要約'] = cls.summary_ja;
  row['返信文(編集可)'] = cls.suggested_reply_ja;
  row['備考'] = match.note;
  var values = INQUIRY.TICKET_HEADERS.map(function (h) { return row[h] !== undefined ? row[h] : ''; });
  sheet.appendRow(values);
  var rowNum = sheet.getLastRow();
  colorTicketRow_(sheet, rowNum, cls.urgency, status);
  return { id: id, row: rowNum };
}

function appendHistory_(ss, ticketId, direction, sender, body, attachmentUrl) {
  var sheet = ss.getSheetByName(INQUIRY.SHEETS.HISTORY)
    || createSheetWithHeaders_(ss, INQUIRY.SHEETS.HISTORY, INQUIRY.HISTORY_HEADERS);
  var hid = nextSeqId_(sheet, 'H-', '履歴ID');
  sheet.appendRow([hid, ticketId, new Date(), direction, sender || '',
    String(body || '').substring(0, INQUIRY.BODY_MAX), attachmentUrl || '']);
}

function notifyStaff_(ss, ticket, inbound, cls, match) {
  var to = getSettingSafe_(INQUIRY.SETTINGS.NOTIFY_EMAIL);
  if (!to || to.indexOf('@') === -1) { Logger.log('通知先未設定のためスタッフ通知スキップ'); return; }
  var urgent = (cls.urgency === '高' || cls.category === '緊急');
  var subject = (urgent ? '【緊急】' : '【問い合わせ】') + cls.category + ' / ' + ticket.id;
  var lines = [
    '新しい問い合わせを受け付けました。', '',
    '受付ID: ' + ticket.id,
    'チャネル: ' + inbound.channel,
    '分類 / 緊急度: ' + cls.category + ' / ' + cls.urgency,
    '物件: ' + (match.property || '（要確認）') + ' ' + (match.unit || ''),
    'オーナー: ' + (match.owner || '（要確認）'),
    'テナント: ' + (match.tenant || inbound.fromName || ''),
    '送信者: ' + (inbound.fromName || '') + ' <' + inbound.fromAddress + '>',
    '要約: ' + cls.summary_ja,
    match.confident ? '' : '※ 物件/テナントを自動特定できませんでした。要確認です。', '',
    '--- 本文 ---',
    String(inbound.body || '').substring(0, 1000), '',
    '「' + INQUIRY.SHEETS.TICKETS + '」シートで対応してください。',
  ];
  try {
    GmailApp.sendEmail(to, subject, lines.filter(function (l) { return l !== ''; }).join('\n'));
  } catch (e) { Logger.log('スタッフ通知メール失敗: ' + e.message); }
}

function dispatchReply_(ss, rowNum, body) {
  var channel = getTicketField_(ss, rowNum, 'チャネル');
  var contact = getTicketField_(ss, rowNum, '送信者連絡先');
  if (!body) return false;
  if (channel === 'whatsapp') return sendWhatsApp_(contact, body);
  var threadId = getTicketField_(ss, rowNum, 'スレッドID');
  var subject = getTicketField_(ss, rowNum, '件名');
  return sendEmailReply_(threadId, contact, 'Re: ' + (subject || 'お問い合わせの件'), body);
}

function processEmailInbox() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var query = getSettingSafe_(INQUIRY.SETTINGS.GMAIL_QUERY) || 'label:inquiry is:unread';
  var threads = GmailApp.search(query, 0, INQUIRY.GMAIL_MAX_THREADS);
  var count = 0;
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var msgs = thread.getMessages();
    var msg = msgs[msgs.length - 1];
    try {
      var inbound = buildInboundFromEmail_(thread, msg);
      handleInbound_(ss, inbound);
      thread.markRead();
      count++;
    } catch (e) { Logger.log('メール処理エラー: ' + e.message); }
    Utilities.sleep(1000);
  }
  return count;
}

function buildInboundFromEmail_(thread, msg) {
  var atts = msg.getAttachments() || [];
  var attMeta = [];
  for (var i = 0; i < atts.length; i++) {
    attMeta.push({ name: atts[i].getName(), mimeType: atts[i].getContentType(), url: '' });
  }
  return {
    channel: 'email',
    receivedAt: msg.getDate(),
    fromAddress: extractEmail_(msg.getFrom()),
    fromName: extractName_(msg.getFrom()),
    subject: msg.getSubject(),
    body: msg.getPlainBody(),
    attachments: attMeta,
    threadId: thread.getId(),
  };
}

function sendEmailReply_(threadId, to, subject, body) {
  try {
    if (threadId) {
      var th = GmailApp.getThreadById(threadId);
      if (th) { th.reply(body); return true; }
    }
  } catch (e) { Logger.log('スレッド返信失敗、新規送信に切替: ' + e.message); }
  if (to && to.indexOf('@') !== -1) {
    GmailApp.sendEmail(to, subject || 'お問い合わせの件', body);
    return true;
  }
  Logger.log('返信先メール不明のため送信スキップ');
  return false;
}

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    if (p.From && p.Body !== undefined) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var inbound = {
        channel: 'whatsapp', receivedAt: new Date(),
        fromAddress: String(p.From).replace('whatsapp:', ''),
        fromName: p.ProfileName || '', subject: '',
        body: p.Body || '', attachments: [], threadId: String(p.From),
      };
      handleInbound_(ss, inbound);
    }
  } catch (err) { Logger.log('doPost エラー: ' + err.message); }
  return ContentService.createTextOutput('<Response></Response>')
    .setMimeType(ContentService.MimeType.XML);
}

function sendWhatsApp_(to, body) {
  var sid = getSettingSafe_(INQUIRY.SETTINGS.TW_SID);
  var token = getSettingSafe_(INQUIRY.SETTINGS.TW_TOKEN);
  var from = getSettingSafe_(INQUIRY.SETTINGS.TW_FROM);
  if (!sid || sid.charAt(0) === '（' || !token || token.charAt(0) === '（' || !from) {
    Logger.log('Twilio 未設定のため WhatsApp 送信をスキップ');
    return false;
  }
  var url = 'https://api.twilio.com/2010-04-01/Accounts/' + sid + '/Messages.json';
  var toAddr = (String(to).indexOf('whatsapp:') === 0) ? to : ('whatsapp:' + to);
  var fromAddr = (String(from).indexOf('whatsapp:') === 0) ? from : ('whatsapp:' + from);
  var options = {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(sid + ':' + token) },
    payload: { From: fromAddr, To: toAddr, Body: body },
    muteHttpExceptions: true,
  };
  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  if (code >= 200 && code < 300) return true;
  Logger.log('Twilio 送信失敗 (' + code + '): ' + resp.getContentText().substring(0, 200));
  return false;
}

function replyActiveTicket() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== INQUIRY.SHEETS.TICKETS) {
    throw new Error('「' + INQUIRY.SHEETS.TICKETS + '」シートで対象行を選択してください。');
  }
  var row = sheet.getActiveCell().getRow();
  if (row < 2) throw new Error('データ行を選択してください。');
  var body = getTicketField_(ss, row, '返信文(編集可)');
  if (!body) throw new Error('「返信文(編集可)」が空です。文面を入力してから送信してください。');
  var sent = dispatchReply_(ss, row, body);
  if (!sent) throw new Error('送信に失敗しました（チャネル設定・宛先をご確認ください）。');
  var id = getTicketField_(ss, row, '受付ID');
  appendHistory_(ss, id, '送信', '(担当)', body, '');
  setTicketField_(ss, row, '返信モード', INQUIRY.REPLY_MODE.SENT);
  setTicketField_(ss, row, 'ステータス', INQUIRY.STATUS.IN_PROGRESS);
  return id;
}

function markActiveTicketDone() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== INQUIRY.SHEETS.TICKETS) {
    throw new Error('「' + INQUIRY.SHEETS.TICKETS + '」シートで対象行を選択してください。');
  }
  var row = sheet.getActiveCell().getRow();
  if (row < 2) throw new Error('データ行を選択してください。');
  setTicketField_(ss, row, 'ステータス', INQUIRY.STATUS.DONE);
  colorTicketRow_(sheet, row, getTicketField_(ss, row, '緊急度'), INQUIRY.STATUS.DONE);
  return getTicketField_(ss, row, '受付ID');
}

function callGeminiText_(prompt) {
  var apiKey = (typeof getGeminiApiKey === 'function')
    ? getGeminiApiKey() : getSettingValue(CONFIG.GEMINI.API_KEY_SETTING);
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + CONFIG.GEMINI.MODEL + ':generateContent?key=' + apiKey;
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1, responseMimeType: 'application/json' },
  };
  var options = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true,
  };
  var response = fetchGeminiWithRetry_(url, options);
  var json = JSON.parse(response.getContentText());
  var text = json.candidates[0].content.parts[0].text;
  var cleaned = String(text).replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
  return JSON.parse(cleaned);
}

function findTenantByContact_(ss, contact) {
  if (!contact) return null;
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TENANTS);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var iProp = headers.indexOf('物件名');
  var iUnit = headers.indexOf('区画番号');
  var iName = headers.indexOf('テナント名');
  var iMail = headers.indexOf('メール');
  var iWa = headers.indexOf('WhatsApp番号');
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var emailKey = normalizeEmail_(contact);
  var phoneKey = normalizePhone_(contact);
  for (var i = 0; i < data.length; i++) {
    if (iMail >= 0 && data[i][iMail] && normalizeEmail_(data[i][iMail]) === emailKey && emailKey) {
      return { property: data[i][iProp], unit: data[i][iUnit], name: data[i][iName] };
    }
    if (iWa >= 0 && data[i][iWa] && normalizePhone_(data[i][iWa]) === phoneKey && phoneKey) {
      return { property: data[i][iProp], unit: data[i][iUnit], name: data[i][iName] };
    }
  }
  return null;
}

function getTicketField_(ss, rowNum, headerName) {
  var sheet = ss.getSheetByName(INQUIRY.SHEETS.TICKETS);
  var col = INQUIRY.TICKET_HEADERS.indexOf(headerName) + 1;
  if (col <= 0) return '';
  return sheet.getRange(rowNum, col).getValue();
}

function setTicketField_(ss, rowNum, headerName, value) {
  var sheet = ss.getSheetByName(INQUIRY.SHEETS.TICKETS);
  var col = INQUIRY.TICKET_HEADERS.indexOf(headerName) + 1;
  if (col <= 0) return;
  sheet.getRange(rowNum, col).setValue(value);
}

function nextSeqId_(sheet, prefix, idHeader) {
  var max = 0;
  if (sheet.getLastRow() > 1) {
    var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var re = new RegExp(prefix.replace(/[-]/g, '\\-') + '(\\d+)');
    for (var i = 0; i < ids.length; i++) {
      var m = String(ids[i][0]).match(re);
      if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
    }
  }
  return prefix + String(max + 1).padStart(4, '0');
}

function colorTicketRow_(sheet, rowNum, urgency, status) {
  var range = sheet.getRange(rowNum, 1, 1, INQUIRY.TICKET_HEADERS.length);
  if (status === INQUIRY.STATUS.DONE) { range.setBackground('#e6f4ea'); return; }
  if (status === INQUIRY.STATUS.REVIEW) { range.setBackground('#fef7e0'); return; }
  if (urgency === '高') { range.setBackground('#fce8e6'); return; }
  range.setBackground(null);
}

function getSettingSafe_(key) {
  try { return String(getSettingValue(key) || ''); } catch (e) { return ''; }
}

function extractEmail_(from) {
  var m = String(from || '').match(/<([^>]+)>/);
  if (m) return m[1].trim();
  var m2 = String(from || '').match(/[^\s<>]+@[^\s<>]+/);
  return m2 ? m2[0].trim() : String(from || '').trim();
}

function extractName_(from) {
  var m = String(from || '').match(/^\s*"?([^"<]+?)"?\s*</);
  return m ? m[1].trim() : '';
}

function normalizeEmail_(s) {
  return String(s || '').toLowerCase().trim();
}

function normalizePhone_(s) {
  return String(s || '').replace(/[^0-9]/g, '');
}
