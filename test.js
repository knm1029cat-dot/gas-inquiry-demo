// 判定ロジックのテスト：node test.js
const assert = require('assert');
const { classify, detectPriority, buildId, findOverdue, COL } = require('./Code.gs');

// 種別判定
assert.strictEqual(classify('来週土曜日の14時に予約できますか？'), '予約・日程');
assert.strictEqual(classify('予約の日時を変更したいです'), '予約・日程');
assert.strictEqual(classify('先月分の請求書と領収書を発行してほしい'), '料金・お支払い');
assert.strictEqual(classify('法人向けの料金プランと見積もりをお願いします'), '料金・お支払い');
assert.strictEqual(classify('注文した商品がまだ届いていないのですが'), '不具合・トラブル');
assert.strictEqual(classify('ログインするとエラーが出て動かないです'), '不具合・トラブル');
assert.strictEqual(classify('お店の近くに駐車場はありますか？'), 'その他');
assert.strictEqual(classify(''), 'その他');
assert.strictEqual(classify(undefined), 'その他');

// 優先度
assert.strictEqual(detectPriority('至急確認をお願いします'), '高');
assert.strictEqual(detectPriority('今日中に支払いたいです'), '高');
assert.strictEqual(detectPriority('駐車場はありますか'), '通常');

// 受付番号：日付ごとに連番、他の日付や欠番の影響を受けない
assert.strictEqual(buildId('20260926', []), 'INQ-20260926-001');
assert.strictEqual(buildId('20260926', ['INQ-20260926-001', 'INQ-20260926-002']), 'INQ-20260926-003');
assert.strictEqual(buildId('20260926', ['INQ-20260925-007', 'INQ-20260926-001']), 'INQ-20260926-002');
assert.strictEqual(buildId('20260926', ['INQ-20260926-001', 'INQ-20260926-005']), 'INQ-20260926-006');
assert.strictEqual(buildId('20260926', ['受付番号', '', null]), 'INQ-20260926-001');

// 放置検知：未対応かつ24時間以上のみ
const now = new Date('2026-09-26T09:00:00+09:00');
const row = (status, hoursAgo) => {
  const r = new Array(11).fill('');
  r[COL.STATUS] = status;
  r[COL.RECEIVED] = new Date(now.getTime() - hoursAgo * 3600 * 1000);
  return r;
};
const rows = [row('未対応', 30), row('未対応', 24), row('未対応', 23.9), row('対応中', 50), row('完了', 90)];
const overdue = findOverdue(rows, now, 24);
assert.strictEqual(overdue.length, 2);
assert.ok(overdue.every(r => r[COL.STATUS] === '未対応'));
const blank = new Array(11).fill('');
assert.strictEqual(findOverdue([blank], now, 24).length, 0);

console.log('all tests passed');
