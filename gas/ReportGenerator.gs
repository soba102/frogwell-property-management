/**
 * オーナー別の月次収支レポートを PDF として生成する。
 * テンプレートシートにデータを流し込み → PDF エクスポート → Drive 保存。
 */

/**
 * 全オーナー分のレポートを一括生成する。
 * @returns {number} 生成したレポートの件数
 */
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

/**
 * 指定オーナーの月次レポートを PDF で生成する。
 * @param {string} ownerName - オーナー名
 * @returns {File} 生成された PDF ファイル
 */
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

  // 一時レポートシートを作成
  var tempSheet = createTempReportSheet_(ss, ownerInfo, properties, invoices, yearMonth, reportId);
  SpreadsheetApp.flush();

  // PDF エクスポート
  var pdf = exportSheetAsPdf_(ss, tempSheet, ownerInfo.name + '様_' + yearMonth);

  // 一時シートを削除
  ss.deleteSheet(tempSheet);

  // レポートフォルダに保存
  var folderId = getReportFolderId();
  var folder = DriveApp.getFolderById(folderId);
  var savedFile = folder.createFile(pdf);

  Logger.log('レポート生成完了: ' + savedFile.getName());
  return savedFile;
}

/** オーナー連絡先情報を取得する（連絡先マスター未登録でも氏名で続行） */
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

/** オーナーの物件一覧を取得する（物件名＋部屋番号） */
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

/** オーナーの請求書データを取得する（確認済のもののみ） */
function getOwnerInvoices_(ss, ownerName) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVOICES);
  if (sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
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

/** 一時レポートシートを作成しデータを流し込む */
function createTempReportSheet_(ss, ownerInfo, properties, invoices, yearMonth, reportId) {
  var sheet = ss.insertSheet('_tmp_report');
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日');
  var propertyLabels = properties.map(function(p) {
    return p.unit ? (p.name + ' ' + p.unit) : p.name;
  });
  var propertyName = propertyLabels.length > 0 ? propertyLabels.join(' / ') : '─';
  var propertyCount = properties.length;

  // 管理手数料率の取得（オーナー個別設定 > グローバル設定 > 既定8%）
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

  // 支出の集計
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

  // シートに書き込み
  var rows = [];
  rows.push(['月次収支レポート / Monthly Property Statement', '', '']);
  rows.push(['', '', '']);
  rows.push(['レポート番号:', reportId, '作成日: ' + dateStr]);
  rows.push(['オーナー名:', ownerInfo.name + ' 様', '対象月: ' + yearMonth]);
  rows.push(['物件:', propertyName, '物件数: ' + propertyCount]);
  rows.push(['', '', '']);

  var incomeHeaderRow = 7;
  rows.push(['■ 収入の部', '', '']);
  rows.push(['項目', '詳細', '金額 (MYR)']);
  for (var m = 0; m < incomeLines.length; m++) rows.push(incomeLines[m]);
  rows.push(['', '収入 合計', totalIncome]);
  rows.push(['', '', '']);

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

  // 書式設定
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sheet.getRange(incomeHeaderRow, 1).setFontWeight('bold');
  sheet.getRange(expenseHeaderRow, 1).setFontWeight('bold');
  sheet.getRange(rows.length - 2, 2, 1, 2).setFontWeight('bold').setFontSize(12);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 150);

  return sheet;
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

/** スプレッドシートの特定シートを PDF としてエクスポートする */
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
