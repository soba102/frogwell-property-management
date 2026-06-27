/* =============================================================
 * Inquiry ─ 問い合わせ対応自動化（Phase 2）
 * -------------------------------------------------------------
 * チャネル非依存の中核（受信 → 分類 → 照合 → チケット → 通知 → 返信）と、
 * メール（Gmail）／WhatsApp（Twilio）の入出力アダプタを1ファイルにまとめる。
 *
 * 設計の正は docs/05_問い合わせ自動化_企画.md。
 * Phase 1（請求書OCR・レポート・消込）のシート/関数（resolveOwner_, getSettingValue,
 * fetchGeminiWithRetry_, CONFIG 等）を再利用する。
 * ============================================================= */

var INQUIRY = {
  SHEETS: {
    TICKETS: '問い合わせ管理',
    HISTORY: '問い合わせ履歴',
    FAQ:     'ナレッジ_FAQ',
  },

  // 末尾に列を足す場合は getTicketField_ がヘッダー名で引くため順序非依存。
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
    NEW:           '新規',
    IN_PROGRESS:   '対応中',
    WAIT_APPROVAL: '承認待ち',
    DONE:          '完了',
    REVIEW:        '要確認',
  },

  REPLY_MODE: {
    AUTO:  '自動送信済',
    DRAFT: '下書き待ち',
    SENT:  '送信済',
  },

  CATEGORIES: ['緊急', '修繕', '契約', '支払い', '近隣', 'FAQ', '空き照会', 'その他'],

  // 自動返信を許可するカテゴリ（定型のみ）。詳細は企画書 §7。
  AUTO_REPLY_CATEGORIES: ['FAQ', '空き照会'],

  GMAIL_MAX_THREADS: 20,
  BODY_MAX: 5000,
};


/* =============================================================
 * セットアップ
 * ============================================================= */

/** 問い合わせ機能のシートを作成し、設定キー・サンプルFAQを投入する */
function setupInquirySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheetWithHeaders_(ss, INQUIRY.SHEETS.TICKETS, INQUIRY.TICKET_HEADERS);
  createSheetWithHeaders_(ss, INQUIRY.SHEETS.HISTORY, INQUIRY.HISTORY_HEADERS);
  var faq = createSheetWithHeaders_(ss, INQUIRY.SHEETS.FAQ, INQUIRY.FAQ_HEADERS);

  if (faq.getLastRow() <= 1) {
    var now = new Date();
    faq.getRange(2, 1, 4, INQUIRY.FAQ_HEADERS.length).setValues([
      ['FAQ',    'ゴミ 出し 回収',     'ゴミは可燃が火・木、リサイクルが土曜の朝8時までに所定の集積所へお出しください。', 'YES', now],
      ['FAQ',    '駐車 場 来客',       '来客用駐車場は管理事務所で受付後、B1Fをご利用いただけます（1日まで）。',          'YES', now],
      ['空き照会', '空室 空き 募集',     '現在の空室状況は「マスター_空き物件」をご確認のうえ、最新情報をご案内します。',     'NO',  now],
      ['内見',    '内見 鍵 visit',     '内見は平日10-17時に受付。鍵は管理事務所保管のため、前日までにご予約ください。',     'NO',  now],
    ]);
  }

  ensureInquirySettings_(ss);
  ensureTenantContactColumns_(ss);
}

/** 設定シートに問い合わせ用キーが無ければ追記する */
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
  for (var j = 0; j < defaults.length; j++) {
    if (!existing[defaults[j][0]]) toAdd.push(defaults[j]);
  }
  if (toAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
  }
}

/** マスター_テナントに「メール」「WhatsApp番号」列が無ければ追加する */
function ensureTenantContactColumns_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TENANTS);
  if (!sheet) return;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var needed = ['メール', 'WhatsApp番号'];
  for (var i = 0; i < needed.length; i++) {
    if (headers.indexOf(needed[i]) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1)
        .setValue(needed[i]).setFontWeight('bold').setBackground('#f1f3f4');
    }
  }
}


/* =============================================================
 * 中核パイプライン
 * ============================================================= */

/**
 * 1件の受信メッセージ（チャネル非依存の InboundMessage）を処理する。
 * メールアダプタ・WhatsApp(doPost) の両方からここに集約される。
 * @param {Object} inbound { channel, receivedAt, fromAddress, fromName, subject, body, attachments, threadId }
 * @returns {Object} ticket { id, row }
 */
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
    // 定型外・特定不可は受付の一次返信のみ。本対応は人が行う（§7 ポリシー）。
    var ack = '【自動受付】お問い合わせを受け付けました。担当者が内容を確認し、追ってご連絡いたします。';
    var ackSent = dispatchReply_(ss, ticket.row, ack);
    if (ackSent) appendHistory_(ss, ticket.id, '送信', '(自動受付)', ack, '');
    setTicketField_(ss, ticket.row, '返信モード', INQUIRY.REPLY_MODE.DRAFT);
  }

  return ticket;
}

/**
 * Gemini でメッセージを分類・要約・下書き生成する（テキスト入力）。
 * @returns {Object} { category, urgency, property_hint, unit_hint, tenant_hint, summary_ja, suggested_reply_ja, auto_reply_ok }
 */
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
    category: 'その他', urgency: '中',
    property_hint: '', unit_hint: '', tenant_hint: '',
    summary_ja: String(inbound.subject || inbound.body || '').substring(0, 30),
    suggested_reply_ja: 'お問い合わせありがとうございます。担当者より追ってご連絡いたします。',
    auto_reply_ok: false,
  };

  try {
    var data = callGeminiText_(prompt);
    return {
      category:           data.category || 'その他',
      urgency:            data.urgency || '中',
      property_hint:      data.property_hint || '',
      unit_hint:          data.unit_hint || '',
      tenant_hint:        data.tenant_hint || '',
      summary_ja:         data.summary_ja || fallback.summary_ja,
      suggested_reply_ja: data.suggested_reply_ja || fallback.suggested_reply_ja,
      auto_reply_ok:      data.auto_reply_ok === true,
    };
  } catch (e) {
    Logger.log('分類失敗（フォールバック使用）: ' + e.message);
    return fallback;
  }
}

/**
 * 送信者・物件・オーナーを特定する（企画書 §5.7 の優先順）。
 * ①差出人の連絡先 → ②本文の手がかり → ③不明なら要確認。
 */
function matchInquiry_(ss, inbound, cls) {
  var res = { property: '', unit: '', tenant: '', owner: '', candidate: '', note: '', confident: false };

  // ① 差出人の連絡先で一意に引く（最も確実）
  var tenant = findTenantByContact_(ss, inbound.fromAddress);
  if (tenant) {
    res.property = tenant.property;
    res.unit = tenant.unit;
    res.tenant = tenant.name;
    res.owner = resolveOwner_(ss, tenant.property, tenant.unit, '').ownerName;
    res.confident = true;
    res.note = '差出人の連絡先からテナント特定';
    return res;
  }

  // ② 本文の手がかり（あくまで候補）
  res.candidate = [cls.property_hint, cls.unit_hint].filter(function (x) { return x; }).join(' / ');
  if (cls.property_hint || cls.unit_hint) {
    var m = resolveOwner_(ss, cls.property_hint || '', cls.unit_hint || '', '');
    if (m.ownerName) {
      res.property = m.propertyName;
      res.unit = cls.unit_hint || '';
      res.owner = m.ownerName;
      res.confident = true;
      res.note = '本文の手がかりから特定' + (m.note ? '（' + m.note + '）' : '');
      return res;
    }
    res.note = '本文の手がかりからは一意に特定できず（候補のみ）';
  } else {
    res.note = '物件の手がかりなし';
  }

  // ③ 不明 → 要確認（推測で確定しない）
  return res;
}

/** チケットを起票して { id, row } を返す */
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

/** やりとりを履歴シートに1行追記する */
function appendHistory_(ss, ticketId, direction, sender, body, attachmentUrl) {
  var sheet = ss.getSheetByName(INQUIRY.SHEETS.HISTORY)
    || createSheetWithHeaders_(ss, INQUIRY.SHEETS.HISTORY, INQUIRY.HISTORY_HEADERS);
  var hid = nextSeqId_(sheet, 'H-', '履歴ID');
  sheet.appendRow([hid, ticketId, new Date(), direction, sender || '',
    String(body || '').substring(0, INQUIRY.BODY_MAX), attachmentUrl || '']);
}

/** 担当者へメール通知する */
function notifyStaff_(ss, ticket, inbound, cls, match) {
  var to = getSettingSafe_(INQUIRY.SETTINGS.NOTIFY_EMAIL);
  if (!to || to.indexOf('@') === -1) {
    Logger.log('通知先未設定のためスタッフ通知スキップ');
    return;
  }
  var urgent = (cls.urgency === '高' || cls.category === '緊急');
  var subject = (urgent ? '【緊急】' : '【問い合わせ】') + cls.category + ' / ' + ticket.id;
  var lines = [
    '新しい問い合わせを受け付けました。',
    '',
    '受付ID: ' + ticket.id,
    'チャネル: ' + inbound.channel,
    '分類 / 緊急度: ' + cls.category + ' / ' + cls.urgency,
    '物件: ' + (match.property || '（要確認）') + ' ' + (match.unit || ''),
    'オーナー: ' + (match.owner || '（要確認）'),
    'テナント: ' + (match.tenant || inbound.fromName || ''),
    '送信者: ' + (inbound.fromName || '') + ' <' + inbound.fromAddress + '>',
    '要約: ' + cls.summary_ja,
    match.confident ? '' : '※ 物件/テナントを自動特定できませんでした。要確認です。',
    '',
    '--- 本文 ---',
    String(inbound.body || '').substring(0, 1000),
    '',
    '「' + INQUIRY.SHEETS.TICKETS + '」シートで対応してください。',
  ];
  try {
    GmailApp.sendEmail(to, subject, lines.filter(function (l) { return l !== ''; }).join('\n'));
  } catch (e) {
    Logger.log('スタッフ通知メール失敗: ' + e.message);
  }
}


/* =============================================================
 * 返信ディスパッチ（チャネル別）
 * ============================================================= */

/** チケット行のチャネルに応じて返信を送る */
function dispatchReply_(ss, rowNum, body) {
  var channel = getTicketField_(ss, rowNum, 'チャネル');
  var contact = getTicketField_(ss, rowNum, '送信者連絡先');
  if (!body) return false;

  if (channel === 'whatsapp') {
    return sendWhatsApp_(contact, body);
  }
  var threadId = getTicketField_(ss, rowNum, 'スレッドID');
  var subject = getTicketField_(ss, rowNum, '件名');
  return sendEmailReply_(threadId, contact, 'Re: ' + (subject || 'お問い合わせの件'), body);
}


/* =============================================================
 * メールアダプタ（Gmail）
 * ============================================================= */

/** Gmail を検索して未処理の問い合わせを取り込む。@returns {number} 処理件数 */
function processEmailInbox() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var query = getSettingSafe_(INQUIRY.SETTINGS.GMAIL_QUERY) || 'label:inquiry is:unread';
  var threads = GmailApp.search(query, 0, INQUIRY.GMAIL_MAX_THREADS);
  var count = 0;

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var msgs = thread.getMessages();
    var msg = msgs[msgs.length - 1]; // スレッド最新を対象
    try {
      var inbound = buildInboundFromEmail_(thread, msg);
      handleInbound_(ss, inbound);
      thread.markRead();
      count++;
    } catch (e) {
      Logger.log('メール処理エラー: ' + e.message);
    }
    Utilities.sleep(1000); // API レート対策
  }
  return count;
}

/** Gmail メッセージを InboundMessage に変換する */
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

/** スレッドが取れれば同一スレッドに返信。取れなければ新規送信 */
function sendEmailReply_(threadId, to, subject, body) {
  try {
    if (threadId) {
      var th = GmailApp.getThreadById(threadId);
      if (th) { th.reply(body); return true; }
    }
  } catch (e) {
    Logger.log('スレッド返信失敗、新規送信に切替: ' + e.message);
  }
  if (to && to.indexOf('@') !== -1) {
    GmailApp.sendEmail(to, subject || 'お問い合わせの件', body);
    return true;
  }
  Logger.log('返信先メール不明のため送信スキップ');
  return false;
}


/* =============================================================
 * WhatsApp アダプタ（Twilio）
 * ============================================================= */

/**
 * Twilio Webhook 受信。Apps Script を「ウェブアプリ」として公開し、
 * Twilio の WhatsApp 受信 Webhook(POST) に公開URLを設定する。
 * Twilio は application/x-www-form-urlencoded で From / Body / ProfileName 等を送る。
 */
function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    if (p.From && p.Body !== undefined) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var inbound = {
        channel: 'whatsapp',
        receivedAt: new Date(),
        fromAddress: String(p.From).replace('whatsapp:', ''),
        fromName: p.ProfileName || '',
        subject: '',
        body: p.Body || '',
        attachments: [],
        threadId: String(p.From),
      };
      handleInbound_(ss, inbound);
    }
  } catch (err) {
    Logger.log('doPost エラー: ' + err.message);
  }
  // 空の TwiML を返す（Twilio が自動返信を二重送信しないように）
  return ContentService.createTextOutput('<Response></Response>')
    .setMimeType(ContentService.MimeType.XML);
}

/** Twilio 経由で WhatsApp メッセージを送る。未設定ならスキップ。@returns {boolean} */
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


/* =============================================================
 * メニュー操作（選択行ベース）
 * ============================================================= */

/** アクティブ行のチケットに「返信文(編集可)」の内容を送信する */
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

/** アクティブ行のチケットを完了にする */
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


/* =============================================================
 * ヘルパー
 * ============================================================= */

/** Gemini をテキスト入力で呼び、JSON をパースして返す */
function callGeminiText_(prompt) {
  var apiKey = (typeof getGeminiApiKey === 'function')
    ? getGeminiApiKey()
    : getSettingValue(CONFIG.GEMINI.API_KEY_SETTING);

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + CONFIG.GEMINI.MODEL + ':generateContent?key=' + apiKey;
  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1, responseMimeType: 'application/json' },
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  var response = fetchGeminiWithRetry_(url, options);
  var json = JSON.parse(response.getContentText());
  var text = json.candidates[0].content.parts[0].text;
  var cleaned = String(text).replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
  return JSON.parse(cleaned);
}

/** マスター_テナントを連絡先（メール/WhatsApp番号）で検索する */
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

/** チケット行から指定ヘッダーの値を取得する */
function getTicketField_(ss, rowNum, headerName) {
  var sheet = ss.getSheetByName(INQUIRY.SHEETS.TICKETS);
  var col = INQUIRY.TICKET_HEADERS.indexOf(headerName) + 1;
  if (col <= 0) return '';
  return sheet.getRange(rowNum, col).getValue();
}

/** チケット行の指定ヘッダーに値を設定する */
function setTicketField_(ss, rowNum, headerName, value) {
  var sheet = ss.getSheetByName(INQUIRY.SHEETS.TICKETS);
  var col = INQUIRY.TICKET_HEADERS.indexOf(headerName) + 1;
  if (col <= 0) return;
  sheet.getRange(rowNum, col).setValue(value);
}

/** "PREFIX0001" 形式の連番IDを発番する */
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

/** 緊急度・ステータスで行に色を付ける */
function colorTicketRow_(sheet, rowNum, urgency, status) {
  var range = sheet.getRange(rowNum, 1, 1, INQUIRY.TICKET_HEADERS.length);
  if (status === INQUIRY.STATUS.DONE) { range.setBackground('#e6f4ea'); return; }
  if (status === INQUIRY.STATUS.REVIEW) { range.setBackground('#fef7e0'); return; }
  if (urgency === '高') { range.setBackground('#fce8e6'); return; }
  range.setBackground(null);
}

/** 設定値を安全に取得する（無ければ空文字） */
function getSettingSafe_(key) {
  try { return String(getSettingValue(key) || ''); } catch (e) { return ''; }
}

/** "Name <a@b.com>" からメールアドレスを抽出する */
function extractEmail_(from) {
  var m = String(from || '').match(/<([^>]+)>/);
  if (m) return m[1].trim();
  var m2 = String(from || '').match(/[^\s<>]+@[^\s<>]+/);
  return m2 ? m2[0].trim() : String(from || '').trim();
}

/** "Name <a@b.com>" から表示名を抽出する */
function extractName_(from) {
  var m = String(from || '').match(/^\s*"?([^"<]+?)"?\s*</);
  return m ? m[1].trim() : '';
}

/** メール比較用に正規化（小文字・トリム） */
function normalizeEmail_(s) {
  return String(s || '').toLowerCase().trim();
}

/** 電話番号比較用に正規化（数字のみ。whatsapp:プレフィックスや記号を除去） */
function normalizePhone_(s) {
  return String(s || '').replace(/[^0-9]/g, '');
}
