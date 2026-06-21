/**
 * Gemini 2.0 Flash API を呼び出し、請求書画像/PDFから構造化データを抽出する。
 *
 * @param {Blob} fileBlob - Drive から取得したファイルの Blob
 * @param {string} mimeType - ファイルの MIME タイプ
 * @param {string} fileName - ファイル名（コンテキスト情報として使用）
 * @returns {Object} 抽出された請求書データ
 */
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

    // リトライ対象（一時的なエラー）か判定
    var retryable = (status === 503 || status === 429 || status === 500);
    if (!retryable || attempt === maxAttempts) break;

    Logger.log('Gemini API ' + status + ' のため ' + (waitMs / 1000) + '秒後にリトライ (' + attempt + '/' + (maxAttempts - 1) + ')');
    Utilities.sleep(waitMs);
    waitMs *= 2; // 2s → 4s → 8s
  }

  throw new Error('Gemini API エラー (' + lastStatus + '): ' + String(lastBody).substring(0, 200));
}

/** OCR用プロンプトを組み立てる */
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

/** Gemini API リクエストのペイロードを組み立てる */
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

/** Gemini API レスポンスをパースして構造化データを返す */
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
