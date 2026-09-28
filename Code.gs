/**
 * 問い合わせ管理 自動化デモ（Google Apps Script）
 *
 * Googleフォームに問い合わせが届くと、
 *   受付番号の付与 → 種別・優先度の自動判定 → 担当の振り分け → 管理シートへ登録
 *   → 担当者へ通知 → 問い合わせ者へ自動返信
 * までを自動で行う。毎朝、24時間以上「未対応」の問い合わせをまとめて通知する。
 *
 * 使い方は README.md を参照。
 */

const CONFIG = {
  // true の間は、すべてのメールを自分（スクリプトの実行者）宛てに送る。
  // 本番で問い合わせ者・担当者に送るときだけ false にする。
  DEMO_MODE: true,
  SHOP_NAME: 'サンプル商店（架空）',
  SHEET_MANAGE: '問い合わせ管理',
  SHEET_SUMMARY: '集計',
  REMIND_HOURS: 24,
  REMIND_AT_HOUR: 9,
};

const COLUMNS = ['受付番号', '受付日時', 'お名前', 'メールアドレス', '種別', '優先度', 'お問い合わせ内容', '担当', 'ステータス', '対応メモ', '完了日時'];
const COL = { ID: 0, RECEIVED: 1, NAME: 2, EMAIL: 3, CATEGORY: 4, PRIORITY: 5, BODY: 6, ASSIGNEE: 7, STATUS: 8, MEMO: 9, DONE: 10 };
const STATUSES = ['未対応', '対応中', '完了'];

// 種別はキーワードの一致数で判定する。どれにも当たらなければ最後の「その他」
const CATEGORIES = [
  { name: '予約・日程', keys: ['予約', '日程', '空き', 'キャンセル', '日時', '変更', '来店'] },
  { name: '料金・お支払い', keys: ['料金', '価格', '値段', '見積', '支払', '請求', '領収', '振込'] },
  { name: '不具合・トラブル', keys: ['不具合', 'エラー', '届かない', '届いていない', '壊れ', '動かない', 'できない', '返品', '破損'] },
  { name: 'その他', keys: [] },
];

// 種別ごとの担当。DEMO_MODE 中は email に関係なく自分宛てに届く
const ASSIGNEES = {
  '予約・日程': { name: '受付担当', email: '' },
  '料金・お支払い': { name: '経理担当', email: '' },
  '不具合・トラブル': { name: 'サポート担当', email: '' },
  'その他': { name: '受付担当', email: '' },
};

const URGENT_KEYS = ['至急', '急ぎ', '今日中', '本日中', 'すぐに', '早急'];

const FORM_TITLES = { NAME: 'お名前', EMAIL: 'メールアドレス', BODY: 'お問い合わせ内容' };


// ───────── 判定ロジック（Google のサービスを使わない部分。test.js で検証） ─────────

function classify(text) {
  const t = String(text || '');
  let best = CATEGORIES[CATEGORIES.length - 1].name;
  let bestScore = 0;
  CATEGORIES.forEach(c => {
    const score = c.keys.filter(k => t.includes(k)).length;
    if (score > bestScore) {
      bestScore = score;
      best = c.name;
    }
  });
  return best;
}

function detectPriority(text) {
  const t = String(text || '');
  return URGENT_KEYS.some(k => t.includes(k)) ? '高' : '通常';
}

// dateKey は 'yyyyMMdd'。同じ日の番号の最大値 + 1 を振る（行が削除されても重複しない）
function buildId(dateKey, existingIds) {
  const prefix = 'INQ-' + dateKey + '-';
  let max = 0;
  existingIds.forEach(id => {
    const s = String(id);
    if (s.startsWith(prefix)) {
      const n = parseInt(s.slice(prefix.length), 10);
      if (n > max) max = n;
    }
  });
  return prefix + String(max + 1).padStart(3, '0');
}

function findOverdue(rows, now, hours) {
  const limit = hours * 60 * 60 * 1000;
  return rows.filter(r =>
    r[COL.STATUS] === '未対応' &&
    r[COL.RECEIVED] instanceof Date &&
    now.getTime() - r[COL.RECEIVED].getTime() >= limit
  );
}


// ───────── メニュー・初期設定 ─────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('問い合わせ管理デモ')
    .addItem('① 初期設定（シート・フォーム・トリガー作成）', 'setup')
    .addItem('② サンプルデータを入れる', 'createSampleData')
    .addItem('放置チェックを今すぐ実行', 'dailyReminder')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupManageSheet_(ss);
  const form = setupForm_(ss);
  setupSummarySheet_(ss, form);
  setupTriggers_(ss);
  console.log('初期設定が完了しました。フォームURL: ' + form.getPublishedUrl());
}

function setupManageSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_MANAGE) || ss.insertSheet(CONFIG.SHEET_MANAGE, 0);
  sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS])
    .setFontWeight('bold').setBackground('#2f5d8a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  const widths = [150, 130, 100, 190, 120, 60, 320, 100, 80, 200, 130];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange('G:G').setWrap(true);
  sheet.getRange('B:B').setNumberFormat('yyyy/MM/dd HH:mm');
  sheet.getRange('K:K').setNumberFormat('yyyy/MM/dd HH:mm');

  const rows = sheet.getMaxRows() - 1;
  const statusRange = sheet.getRange(2, COL.STATUS + 1, rows, 1);
  statusRange.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build()
  );

  const all = sheet.getRange(2, 1, rows, COLUMNS.length);
  const priority = sheet.getRange(2, COL.PRIORITY + 1, rows, 1);
  // 1つのセルには最初に当てはまったルールだけが効くため、優先度「高」を先に置く
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($F2="高",$I2<>"完了")').setBackground('#fde8e8').setFontColor('#d93025').setBold(true).setRanges([priority]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I2="未対応"').setBackground('#fde8e8').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I2="対応中"').setBackground('#fff4d6').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I2="完了"').setBackground('#f1f3f4').setFontColor('#80868b').setRanges([all]).build(),
  ]);
  return sheet;
}

function setupForm_(ss) {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('FORM_ID');
  if (savedId) {
    try {
      return FormApp.openById(savedId);
    } catch (err) {
      // フォームが削除されていたら作り直す
    }
  }
  const form = FormApp.create(CONFIG.SHOP_NAME + ' お問い合わせフォーム');
  form.setDescription('内容を確認のうえ、担当よりご連絡いたします。');
  form.addTextItem().setTitle(FORM_TITLES.NAME).setRequired(true);
  form.addTextItem().setTitle(FORM_TITLES.EMAIL).setRequired(true)
    .setValidation(FormApp.createTextValidation().requireTextIsEmail().build());
  form.addParagraphTextItem().setTitle(FORM_TITLES.BODY).setRequired(true);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  props.setProperty('FORM_ID', form.getId());
  return form;
}

function setupSummarySheet_(ss, form) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_SUMMARY) || ss.insertSheet(CONFIG.SHEET_SUMMARY, 1);
  sheet.clear();
  const m = "'" + CONFIG.SHEET_MANAGE + "'!";

  sheet.getRange('A1:B1').setValues([['種別', '件数']]);
  CATEGORIES.forEach((c, i) => {
    sheet.getRange(i + 2, 1, 1, 2).setValues([[c.name, '=COUNTIF(' + m + 'E:E,A' + (i + 2) + ')']]);
  });

  sheet.getRange('D1:E1').setValues([['ステータス', '件数']]);
  STATUSES.forEach((s, i) => {
    sheet.getRange(i + 2, 4, 1, 2).setValues([[s, '=COUNTIF(' + m + 'I:I,D' + (i + 2) + ')']]);
  });

  sheet.getRange('G1:H1').setValues([['要注意', '件数']]);
  sheet.getRange('G2:H3').setValues([
    ['未対応のまま' + CONFIG.REMIND_HOURS + '時間以上', '=COUNTIFS(' + m + 'I:I,"未対応",' + m + 'B:B,"<"&(NOW()-' + CONFIG.REMIND_HOURS / 24 + '))'],
    ['優先度「高」で未完了', '=COUNTIFS(' + m + 'F:F,"高",' + m + 'I:I,"<>完了")'],
  ]);

  sheet.getRange('A8:B8').setValues([['お問い合わせフォーム', form.getPublishedUrl()]]);
  [sheet.getRange('A1:B1'), sheet.getRange('D1:E1'), sheet.getRange('G1:H1'), sheet.getRange('A8')]
    .forEach(r => r.setFontWeight('bold').setBackground('#e8f0fe'));
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(7, 210);
}

function setupTriggers_(ss) {
  ScriptApp.getProjectTriggers()
    .filter(t => ['handleFormSubmit', 'dailyReminder'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(CONFIG.REMIND_AT_HOUR).create();
}


// ───────── 自動処理 ─────────

// フォーム送信時（インストール型トリガー）
function handleFormSubmit(e) {
  const v = e.namedValues;
  const pick = title => (v[title] && v[title][0]) || '';
  processInquiry({
    receivedAt: new Date(),
    name: pick(FORM_TITLES.NAME),
    email: pick(FORM_TITLES.EMAIL),
    body: pick(FORM_TITLES.BODY),
  }, { sendMail: true });
}

function processInquiry(inq, opts) {
  opts = opts || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MANAGE);
    const last = sheet.getLastRow();
    const ids = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]) : [];
    const dateKey = Utilities.formatDate(inq.receivedAt, Session.getScriptTimeZone(), 'yyyyMMdd');

    const record = {
      id: buildId(dateKey, ids),
      category: classify(inq.body),
      priority: detectPriority(inq.body),
      status: opts.status || '未対応',
    };
    record.assignee = ASSIGNEES[record.category];

    sheet.appendRow([
      record.id, inq.receivedAt, inq.name, inq.email, record.category, record.priority,
      inq.body, record.assignee.name, record.status, opts.memo || '',
      record.status === '完了' ? (opts.doneAt || new Date()) : '',
    ]);

    if (opts.sendMail) {
      notifyStaff_(inq, record);
      autoReply_(inq, record);
    }
    return record.id;
  } finally {
    lock.releaseLock();
  }
}

// ステータスを「完了」にしたら完了日時を入れる（シンプルトリガー）
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_MANAGE) return;
  if (range.getColumn() !== COL.STATUS + 1 || range.getRow() < 2 || range.getNumRows() > 1) return;
  const doneCell = sheet.getRange(range.getRow(), COL.DONE + 1);
  if (range.getValue() === '完了') {
    if (!doneCell.getValue()) doneCell.setValue(new Date());
  } else {
    doneCell.clearContent();
  }
}

// 毎朝：24時間以上「未対応」の問い合わせをまとめて通知
function dailyReminder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MANAGE);
  const last = sheet.getLastRow();
  if (last < 2) return;
  const rows = sheet.getRange(2, 1, last - 1, COLUMNS.length).getValues();
  const overdue = findOverdue(rows, new Date(), CONFIG.REMIND_HOURS);
  if (overdue.length === 0) return;

  const tz = Session.getScriptTimeZone();
  const lines = overdue.map(r =>
    '・' + r[COL.ID] + '（' + r[COL.CATEGORY] + ' / 担当：' + r[COL.ASSIGNEE] + '）' +
    Utilities.formatDate(r[COL.RECEIVED], tz, 'M/d HH:mm') + ' 受付　' + r[COL.NAME] + ' 様'
  );
  sendMail_(
    ownerEmail_(),
    '【要対応】未対応の問い合わせが' + overdue.length + '件あります',
    CONFIG.REMIND_HOURS + '時間以上「未対応」のままの問い合わせです。\n\n' + lines.join('\n') +
    '\n\n管理シート：' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  );
}


// ───────── メール ─────────

function notifyStaff_(inq, record) {
  const urgent = record.priority === '高' ? '至急・' : '';
  sendMail_(
    record.assignee.email || ownerEmail_(),
    '【' + urgent + '新規問い合わせ】' + record.id + '（' + record.category + '）',
    [
      record.assignee.name + ' 様',
      '',
      '新しい問い合わせが届きました。',
      '',
      '受付番号：' + record.id,
      '種別　　：' + record.category,
      '優先度　：' + record.priority,
      'お名前　：' + inq.name + ' 様',
      'メール　：' + inq.email,
      '',
      '【内容】',
      inq.body,
      '',
      '管理シート：' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    ].join('\n')
  );
}

function autoReply_(inq, record) {
  if (!inq.email) return;
  sendMail_(
    inq.email,
    '【' + CONFIG.SHOP_NAME + '】お問い合わせを受け付けました（' + record.id + '）',
    [
      inq.name + ' 様',
      '',
      'お問い合わせいただき、ありがとうございます。',
      '以下の内容で受け付けました。担当より順次ご連絡いたします。',
      '',
      '受付番号：' + record.id,
      '',
      '【お問い合わせ内容】',
      inq.body,
      '',
      '※このメールは自動送信です。',
      CONFIG.SHOP_NAME,
    ].join('\n')
  );
}

function sendMail_(to, subject, body) {
  if (CONFIG.DEMO_MODE) {
    body = '（デモモード：本来の宛先は ' + to + ' です）\n\n' + body;
    to = ownerEmail_();
    subject = '[DEMO] ' + subject;
  }
  MailApp.sendEmail(to, subject, body);
}

function ownerEmail_() {
  return Session.getEffectiveUser().getEmail();
}


// ───────── サンプルデータ（スクリーンショット用・メールは送らない） ─────────

function createSampleData() {
  const hoursAgo = h => new Date(Date.now() - h * 60 * 60 * 1000);
  const samples = [
    { h: 96, name: '山田 花子', email: 'hanako@example.com', body: '来週土曜日の14時に予約できますか？2名です。', status: '完了', memo: '予約確定の返信済み' },
    { h: 80, name: '佐藤 一郎', email: 'ichiro@example.com', body: '先月分の請求書と領収書を発行してほしいです。', status: '完了', memo: 'PDF送付済み' },
    { h: 70, name: '鈴木 美咲', email: 'misaki@example.com', body: '注文した商品がまだ届いていないのですが、確認できますか。', status: '対応中', memo: '配送業者に確認中' },
    { h: 50, name: '高橋 健', email: 'ken@example.com', body: '法人向けの料金プランと見積もりをお願いします。', status: '未対応' },
    { h: 30, name: '田中 さくら', email: 'sakura@example.com', body: '予約の日時を変更したいです。', status: '未対応' },
    { h: 6, name: '伊藤 大輔', email: 'daisuke@example.com', body: 'ログインするとエラーが出て動かないです。至急確認をお願いします。', status: '未対応' },
    { h: 3, name: '渡辺 ゆい', email: 'yui@example.com', body: 'お店の近くに駐車場はありますか？', status: '未対応' },
    { h: 1, name: '中村 翔', email: 'sho@example.com', body: '振込先の口座を教えてください。今日中に支払いたいです。', status: '未対応' },
  ];
  samples.forEach(s => {
    processInquiry(
      { receivedAt: hoursAgo(s.h), name: s.name, email: s.email, body: s.body },
      { sendMail: false, status: s.status, memo: s.memo, doneAt: s.status === '完了' ? hoursAgo(s.h - 2) : '' }
    );
  });
}


// Node.js でのテスト用（Apps Script では module が無いので無視される）
if (typeof module !== 'undefined') {
  module.exports = { classify, detectPriority, buildId, findOverdue, COL, CATEGORIES };
}
