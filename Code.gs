/**
 * ============================================================
 * Portfolio #3: Gmail問い合わせ自動振り分け → スプレッドシート転記
 * ------------------------------------------------------------
 * 1. Gmailから未処理の問い合わせメールを検索
 * 2. 件名・本文をルールで判定してカテゴリ分け（Gmailラベルも付与）
 * 3. 本文から 氏名 / 会社名 / メール / 電話番号 を抽出
 * 4. スプレッドシートに1件1行で記録
 * 5. カテゴリごとにラウンドロビンで担当者を自動アサイン → 通知メール
 *
 * ※ スプレッドシートにバインドしたスクリプト（拡張機能 → Apps Script）として使用
 * ============================================================
 */

// ===== 設定 =====
const CONFIG = {
  // 取り込み対象の検索条件（Gmailの検索窓で事前に結果を確認しておくと安全）
  BASE_QUERY: 'subject:(お問い合わせ OR 問い合わせ OR 見積) newer_than:7d',
  PROCESSED_LABEL: 'GAS処理済み',        // 取り込み済みスレッドに付けるラベル
  CATEGORY_LABEL_PARENT: '問い合わせ',   // カテゴリラベルの親（例: 問い合わせ/見積）
  NOTIFY_KEYWORD: '担当アサイン通知',     // 通知メールの件名キーワード（再取り込み防止に使用）
  MAX_THREADS_PER_RUN: 20,               // 1回の実行で処理する最大スレッド数
  BODY_EXCERPT_LENGTH: 300,              // シートに残す本文抜粋の文字数
  NOTIFY_ASSIGNEE: true,                 // 担当者へ通知メールを送るか
  DRY_RUN: false,                        // true: シート・ラベル・通知を一切変更せずログ出力のみ
  TRIGGER_MINUTES: 10,                   // 定期実行の間隔（1 / 5 / 10 / 15 / 30）
  TIMEZONE: 'Asia/Tokyo',
  DEFAULT_CATEGORY: 'その他',
  UNASSIGNED: '未割当',
  SHEETS: {
    INQUIRIES: '問い合わせ一覧',
    RULES: '振り分けルール',
    MEMBERS: '担当者マスタ',
    LOG: '実行ログ',
  },
};

const INQUIRY_HEADERS = [
  '受付ID', '受信日時', 'カテゴリ', 'ステータス', '担当者', '担当者メール',
  '件名', '差出人名', '差出人メール', '会社名', '電話番号', '本文抜粋',
  'Gmailリンク', 'メッセージID', '取込日時',
];
const RULE_HEADERS = ['優先度', 'カテゴリ', 'キーワード（カンマ区切り）', '判定対象（件名/本文/両方）', '有効'];
const MEMBER_HEADERS = ['担当者名', 'メールアドレス', '担当カテゴリ（カンマ区切り / *=全て）', '有効'];
const LOG_HEADERS = ['日時', 'レベル', '内容'];
const STATUS_OPTIONS = ['未対応', '対応中', '完了', '対応不要'];

// 本文から項目を拾うためのラベル候補（フォーム送信メールの「お名前：xxx」「【会社名】xxx」形式を想定）
const FIELD_LABELS = {
  name: ['お名前', '氏名', 'ご担当者名', '担当者名', '名前'],
  company: ['会社名', '企業名', '法人名', '御社名', '貴社名', '団体名'],
  email: ['メールアドレス', 'E-?mail', 'Eメール', 'メール'],
  phone: ['電話番号', '連絡先電話番号', 'TEL', '電話'],
};

// ===== メイン処理（トリガーから呼ばれる） =====
function processInquiries() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) {
    writeLog_('WARN', '前回の実行がまだ終わっていないためスキップしました');
    return;
  }
  try {
    run_();
  } catch (e) {
    writeLog_('ERROR', `処理全体でエラー: ${e.message}\n${e.stack}`);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function run_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetOrThrow_(ss, CONFIG.SHEETS.INQUIRIES);
  const threads = GmailApp.search(buildQuery_(), 0, CONFIG.MAX_THREADS_PER_RUN);
  if (threads.length === 0) {
    console.log('新着の問い合わせはありません');
    return;
  }

  const rules = loadRules_(ss);
  const members = loadMembers_(ss);
  const existingIds = loadExistingMessageIds_(sheet);
  const state = loadState_();

  // 1) 解析
  const drafts = [];
  const duplicateThreads = [];
  threads.forEach((thread) => {
    try {
      const message = thread.getMessages()[0]; // スレッド先頭 = 問い合わせ本体
      if (existingIds.has(message.getId())) {
        duplicateThreads.push(thread);
        return;
      }
      drafts.push({ thread, record: parseMessage_(message, thread, rules) });
    } catch (e) {
      writeLog_('ERROR', `メール解析に失敗 threadId=${thread.getId()}: ${e.message}`);
    }
  });

  // 2) 受信順に並べて受付ID・担当者を採番（Gmail検索は新しい順で返るため）
  drafts.sort((a, b) => a.record.receivedAt - b.record.receivedAt);
  drafts.forEach(({ record }) => {
    state.seq = (state.seq || 0) + 1;
    record.id = 'INQ-' + String(state.seq).padStart(5, '0');
    const member = assign_(record.category, members, state);
    record.assignee = member ? member.name : CONFIG.UNASSIGNED;
    record.assigneeEmail = member ? member.email : '';
  });

  if (CONFIG.DRY_RUN) {
    drafts.forEach(({ record }) => console.log('[DRY_RUN]', JSON.stringify(record, null, 2)));
    console.log(`[DRY_RUN] 取込予定 ${drafts.length}件 / 重複 ${duplicateThreads.length}件（シート・ラベル・通知は変更していません）`);
    return;
  }

  // 3) シートへ一括書き込み（先に書き込み → 成功後にラベル付与。途中失敗でも取りこぼさない）
  if (drafts.length > 0) {
    const rows = drafts.map(({ record }) => recordToRow_(record));
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, INQUIRY_HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();
    saveState_(state);
  }

  // 4) Gmailラベル付与 & 担当者通知
  const processedLabel = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  drafts.forEach(({ thread, record }) => {
    thread.addLabel(getOrCreateLabel_(`${CONFIG.CATEGORY_LABEL_PARENT}/${record.category}`));
    thread.addLabel(processedLabel);
    if (CONFIG.NOTIFY_ASSIGNEE) {
      try {
        notifyAssignee_(record, ss.getUrl());
      } catch (e) {
        writeLog_('WARN', `通知メール送信に失敗 ${record.id}: ${e.message}`);
      }
    }
  });
  duplicateThreads.forEach((t) => t.addLabel(processedLabel));

  writeLog_('INFO', `取込 ${drafts.length}件 / 重複スキップ ${duplicateThreads.length}件`);
}

// ===== メール解析 =====
function parseMessage_(message, thread, rules) {
  const subject = message.getSubject() || '(件名なし)';
  const body = message.getPlainBody() || '';
  const fields = parseBody_(body);
  const from = parseAddress_(message.getFrom());
  const replyTo = parseAddress_(message.getReplyTo() || '');

  return {
    id: '',
    receivedAt: message.getDate(),
    category: classify_(subject, body, rules),
    status: STATUS_OPTIONS[0],
    assignee: '',
    assigneeEmail: '',
    subject,
    senderName: fields.name || from.name,
    // 優先順位: 本文の記載 > Reply-To > From（フォームサービス経由だとFromが送信元サービスになるため）
    senderEmail: fields.email || replyTo.email || from.email,
    company: fields.company,
    phone: fields.phone,
    excerpt: makeExcerpt_(body),
    threadUrl: thread.getPermalink(),
    messageId: message.getId(),
    importedAt: new Date(),
  };
}

function parseBody_(body) {
  const lines = String(body).split(/\r?\n/);
  const emailRaw = extractField_(lines, FIELD_LABELS.email);
  const emailMatch = emailRaw.match(/[\w.+-]+@[\w-]+(?:\.[\w-]+)+/);
  const phoneRaw = toHalfWidth_(extractField_(lines, FIELD_LABELS.phone));
  const phoneMatch = phoneRaw.match(/\+?\d[\d\-() ]{8,}\d/);
  return {
    name: extractField_(lines, FIELD_LABELS.name),
    company: extractField_(lines, FIELD_LABELS.company),
    email: emailMatch ? emailMatch[0] : '',
    phone: phoneMatch ? phoneMatch[0].trim() : '',
  };
}

/**
 * 「お名前：山田」「【お名前】山田」「■氏名: 山田」などの行から値を取り出す。
 * 値が次の行にあるフォーマット（【お名前】改行 山田）にも対応。
 */
function extractField_(lines, labels) {
  const re = new RegExp(
    '^[\\s　]*[【\\[■◆●・]?[\\s　]*(?:' + labels.join('|') + ')' +
    '(?:[】\\]][\\s　]*[：:]?|[\\s　]*[：:])[\\s　]*(.*)$', 'i');
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(re);
    if (!m) continue;
    if (m[1].trim()) return m[1].trim();
    const next = lines.slice(i + 1).find((l) => l.trim());
    if (!next || /^[\s　]*[【\[■◆●]/.test(next) || /[：:]/.test(next)) return '';
    return next.trim();
  }
  return '';
}

function parseAddress_(raw) {
  const s = String(raw || '').trim();
  const m = s.match(/^"?([^"<]*?)"?\s*<([^>]+)>$/);
  if (m) return { name: m[1].trim(), email: m[2].trim() };
  return { name: '', email: s };
}

function makeExcerpt_(body) {
  const cleaned = String(body)
    .split(/\r?\n/)
    .filter((l) => !/^\s*>/.test(l)) // 引用行を除外
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  const n = CONFIG.BODY_EXCERPT_LENGTH;
  return cleaned.length > n ? cleaned.slice(0, n) + '…' : cleaned;
}

function toHalfWidth_(s) {
  return String(s)
    .replace(/[０-９（）＋－]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[ー―‐]/g, '-');
}

// ===== 振り分け・アサイン =====
function classify_(subject, body, rules) {
  for (const rule of rules) {
    const target = rule.target === '件名' ? subject
      : rule.target === '本文' ? body
      : `${subject}\n${body}`;
    const text = target.toLowerCase();
    if (rule.keywords.some((k) => text.includes(k.toLowerCase()))) return rule.category;
  }
  return CONFIG.DEFAULT_CATEGORY;
}

/** カテゴリを担当できる有効メンバーの中で、順番に割り当てる（ラウンドロビン） */
function assign_(category, members, state) {
  const candidates = members.filter((m) => m.categories.includes('*') || m.categories.includes(category));
  if (candidates.length === 0) return null;
  const key = `rr_${category}`;
  const idx = (state[key] || 0) % candidates.length;
  state[key] = (idx + 1) % candidates.length;
  return candidates[idx];
}

function notifyAssignee_(r, sheetUrl) {
  if (!r.assigneeEmail) return;
  // 件名に元の件名を含めない：含めると通知メール自体が検索条件に一致して再取り込みされるため
  const subject = `[${CONFIG.NOTIFY_KEYWORD}] ${r.id}`;
  const body = [
    `${r.assignee} さん`,
    '',
    '新しい問い合わせが割り当てられました。',
    '',
    `受付ID　：${r.id}`,
    `カテゴリ：${r.category}`,
    `受信日時：${Utilities.formatDate(r.receivedAt, CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm')}`,
    `件名　　：${r.subject}`,
    `差出人　：${r.senderName}（${r.senderEmail}）`,
    `会社名　：${r.company || '-'}`,
    '',
    '--- 本文抜粋 ---',
    r.excerpt,
    '',
    `Gmail：${r.threadUrl}`,
    `管理シート：${sheetUrl}`,
  ].join('\n');
  GmailApp.sendEmail(r.assigneeEmail, subject, body);
}

// ===== シート読み書き =====
function recordToRow_(r) {
  return [
    r.id, r.receivedAt, r.category, r.status, r.assignee, r.assigneeEmail,
    r.subject, r.senderName, r.senderEmail, r.company, r.phone, r.excerpt,
    r.threadUrl, r.messageId, r.importedAt,
  ].map(sanitize_);
}

/** 外部から来た文字列が数式として解釈されないようにする（数式インジェクション対策） */
function sanitize_(v) {
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? `'${v}` : v;
}

function loadRules_(ss) {
  const sh = getSheetOrThrow_(ss, CONFIG.SHEETS.RULES);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, RULE_HEADERS.length).getValues()
    .filter((r) => r[4] === true && r[1] && r[2])
    .map((r) => ({
      priority: Number(r[0]) || 999,
      category: String(r[1]).trim().replace(/\//g, '／'), // スラッシュはラベル階層と衝突するため置換
      keywords: splitList_(r[2]),
      target: String(r[3]).trim() || '両方',
    }))
    .sort((a, b) => a.priority - b.priority);
}

function loadMembers_(ss) {
  const sh = getSheetOrThrow_(ss, CONFIG.SHEETS.MEMBERS);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, MEMBER_HEADERS.length).getValues()
    .filter((r) => r[3] === true && r[0] && /@/.test(String(r[1])))
    .map((r) => ({
      name: String(r[0]).trim(),
      email: String(r[1]).trim(),
      categories: splitList_(r[2] || '*'),
    }));
}

function loadExistingMessageIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();
  const col = INQUIRY_HEADERS.indexOf('メッセージID') + 1;
  return new Set(
    sheet.getRange(2, col, lastRow - 1, 1).getValues().map((r) => String(r[0])).filter(Boolean)
  );
}

function splitList_(v) {
  return String(v).split(/[,、，]/).map((s) => s.trim()).filter(Boolean);
}

function getSheetOrThrow_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シート「${name}」がありません。先に setup() を実行してください`);
  return sh;
}

// ===== 状態管理（受付ID連番・ラウンドロビン位置） =====
function loadState_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const state = {};
  Object.keys(props).forEach((k) => {
    if (k === 'seq' || k.startsWith('rr_')) state[k] = Number(props[k]) || 0;
  });
  return state;
}

function saveState_(state) {
  const out = {};
  Object.keys(state).forEach((k) => { out[k] = String(state[k]); });
  PropertiesService.getScriptProperties().setProperties(out);
}

// ===== Gmailラベル =====
const labelCache_ = {};
function getOrCreateLabel_(name) {
  if (labelCache_[name]) return labelCache_[name];
  const parts = name.split('/');
  let label = null;
  for (let i = 1; i <= parts.length; i++) {
    const path = parts.slice(0, i).join('/'); // 親ラベルから順に作成
    label = GmailApp.getUserLabelByName(path) || GmailApp.createLabel(path);
  }
  labelCache_[name] = label;
  return label;
}

function buildQuery_() {
  return `${CONFIG.BASE_QUERY} -label:${CONFIG.PROCESSED_LABEL} -subject:${CONFIG.NOTIFY_KEYWORD}`;
}

// ===== ログ =====
function writeLog_(level, message) {
  console.log(`[${level}] ${message}`);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOG);
    if (sh) sh.appendRow([new Date(), level, message]);
  } catch (e) {
    console.error('ログ書き込み失敗: ' + e.message);
  }
}

// ============================================================
// セットアップ・運用用の関数（エディタから手動実行）
// ============================================================

/** ① 初回に1回実行：必要なシート・見出し・サンプルデータを作成 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const maxDataRows = 999;

  const inquiry = ensureSheet_(ss, CONFIG.SHEETS.INQUIRIES, INQUIRY_HEADERS);
  const phoneCol = INQUIRY_HEADERS.indexOf('電話番号') + 1;
  inquiry.getRange(2, phoneCol, maxDataRows, 1).setNumberFormat('@'); // 先頭の0を保持
  const statusCol = INQUIRY_HEADERS.indexOf('ステータス') + 1;
  inquiry.getRange(2, statusCol, maxDataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS, true).build()
  );
  inquiry.getRange(2, INQUIRY_HEADERS.indexOf('受信日時') + 1, maxDataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  inquiry.getRange(2, INQUIRY_HEADERS.indexOf('取込日時') + 1, maxDataRows, 1).setNumberFormat('yyyy/mm/dd hh:mm');

  const rules = ensureSheet_(ss, CONFIG.SHEETS.RULES, RULE_HEADERS);
  if (rules.getLastRow() < 2) {
    const samples = [
      [1, '見積', '見積,お見積り,料金,費用,価格', '両方', true],
      [2, '不具合', '不具合,エラー,動かない,ログインできない', '両方', true],
      [3, '採用', '採用,求人,応募', '件名', true],
      [4, '営業', 'ご提案,サービスのご案内,営業', '件名', true],
    ];
    rules.getRange(2, 5, samples.length, 1).insertCheckboxes();
    rules.getRange(2, 1, samples.length, RULE_HEADERS.length).setValues(samples);
  }

  const members = ensureSheet_(ss, CONFIG.SHEETS.MEMBERS, MEMBER_HEADERS);
  if (members.getLastRow() < 2) {
    const me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    const samples = [
      ['担当A', me, '*', true],
      ['担当B', me, '見積,不具合', true],
    ];
    members.getRange(2, 4, samples.length, 1).insertCheckboxes();
    members.getRange(2, 1, samples.length, MEMBER_HEADERS.length).setValues(samples);
  }

  ensureSheet_(ss, CONFIG.SHEETS.LOG, LOG_HEADERS);
  console.log('セットアップ完了：「担当者マスタ」「振り分けルール」を実際の内容に書き換えてください');
}

function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** ② 定期実行トリガーを設定（重複登録しないよう既存を削除してから作成） */
function setupTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('processInquiries').timeBased().everyMinutes(CONFIG.TRIGGER_MINUTES).create();
  console.log(`${CONFIG.TRIGGER_MINUTES}分ごとのトリガーを設定しました`);
}

function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'processInquiries')
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

// ============================================================
// テスト用
// ============================================================

const SAMPLE_BODY = [
  'お問い合わせフォームより以下の内容が送信されました。',
  '',
  '【お名前】山田 太郎',
  '【会社名】株式会社サンプル',
  '【メールアドレス】yamada@example.com',
  '【電話番号】０３－１２３４－５６７８',
  '【お問い合わせ内容】',
  '営業管理の効率化について、お見積りをお願いしたいです。',
  '現在はExcelで管理しており、担当者ごとにファイルが分かれています。',
].join('\n');

/** テスト1：本文パースとカテゴリ判定だけを確認（Gmail・シートは変更しない） */
function testParse() {
  const subject = '【お問い合わせ】お見積りのご依頼';
  console.log('抽出結果:', JSON.stringify(parseBody_(SAMPLE_BODY), null, 2));
  const rules = loadRules_(SpreadsheetApp.getActiveSpreadsheet());
  console.log('判定カテゴリ:', classify_(subject, SAMPLE_BODY, rules));
}

/** テスト2：自分宛てにサンプル問い合わせメールを送信 */
function sendTestInquiry() {
  const me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  GmailApp.sendEmail(me, '【お問い合わせ】お見積りのご依頼（テスト）', SAMPLE_BODY);
  console.log(`${me} にテストメールを送信しました。1〜2分後に testDryRun() を実行してください`);
}

/** テスト3：何も変更せず「取り込まれる予定の内容」をログで確認 */
function testDryRun() {
  CONFIG.DRY_RUN = true;
  processInquiries();
}
