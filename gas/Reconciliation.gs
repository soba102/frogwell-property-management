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
  var paidByTenant = {};
  var unmatchedDeposits = [];
  for (var d = 0; d < deposits.length; d++) {
    var dep = deposits[d];
    var depUnit = extractUnitFromText_(dep.name);
    var idx = matchDepositByUnit_(depUnit, dep.name, tenants);
    if (idx >= 0) {
      if (!paidByTenant[idx]) paidByTenant[idx] = { amount: 0, date: dep.date };
      paidByTenant[idx].amount += dep.amount;
      if (dep.date) paidByTenant[idx].date = dep.date;
    } else if (depUnit) {
      // 区画番号は読めたがマスターに該当無し → 要確認
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
    if (!amount || amount <= 0) continue;

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

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.RECONCILE_HEADERS.length).clearContent();
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, CONFIG.RECONCILE_HEADERS.length).setBackground(null);
  }
  if (rows.length === 0) return;

  sheet.getRange(2, 1, rows.length, CONFIG.RECONCILE_HEADERS.length).setValues(rows);

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
