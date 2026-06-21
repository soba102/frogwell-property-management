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
    .addItem('ダッシュボード更新', 'runUpdateDashboard')
    .addToUi();
}

/* ===== メニューから呼ばれるエントリーポイント ===== */

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

function runOcrBatch() {
  const ui = SpreadsheetApp.getUi();
  try {
    const count = processNewInvoices();
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
