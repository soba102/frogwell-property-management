/**
 * 全体設定値。スプレッドシートの「設定」シートから動的に読む値と、
 * コード内の固定定数をここに集約する。
 */

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

  // 先方の「オーナーリスト」フォーマットに準拠。物件名＋部屋番号でオーナー名に紐づく。
  // オーナーは「名前」で識別する（ID廃止）。
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
