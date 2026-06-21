/**
 * OCR で抽出したデータをスプレッドシートの「請求書データ」シートに書き込む。
 */

/**
 * Drive スキャン → OCR → シート書込みの一連のフローを実行する。
 * Main.gs の runOcrBatch() から呼ばれる。
 * @returns {number} 処理した請求書の件数
 */
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

    // API レート制限対策
    if (i < files.length - 1) Utilities.sleep(1000);
  }

  Logger.log(count + ' 件処理完了');
  return count;
}

/** 次の請求書IDの連番を取得する */
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

  // 建物名が空なら住所から推測（住所にも建物名が含まれることが多い）
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
    // 建物もユニットも見つからない → 別名シートへ未割当として記録
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
    // 数字部分の先頭ゼロを除去（05 → 5、43A → 43A）
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
      // 建物名 or 住所のいずれかが検出名を内包すれば正規名へ
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
  // 既出チェック
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

/** AI の confidence をステータスにマッピングする */
function mapConfidenceToStatus_(confidence) {
  switch (confidence) {
    case 'high':   return CONFIG.STATUSES.AI_READ;
    case 'medium': return CONFIG.STATUSES.AI_READ;
    case 'low':    return CONFIG.STATUSES.REVIEW;
    default:       return CONFIG.STATUSES.REVIEW;
  }
}

/** ステータスに応じて行の背景色を設定する */
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
