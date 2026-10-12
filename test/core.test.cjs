'use strict';
/* core.js 单元测试：node --test test/ */
const test = require('node:test');
const assert = require('node:assert');
const path = require('path');
const FX = require(path.join(__dirname, '..', 'public', 'core.js'));

test('toNumber：识别千分位、货币符号、括号负数', () => {
  assert.strictEqual(FX.toNumber(14.59), 14.59);
  assert.strictEqual(FX.toNumber('1,234.56'), 1234.56);
  assert.strictEqual(FX.toNumber('(123.45)'), -123.45);
  assert.strictEqual(FX.toNumber('¥1,000'), 1000);
  assert.strictEqual(FX.toNumber('-0.21'), -0.21);
  assert.strictEqual(FX.toNumber('$ 8.83'), 8.83);
  assert.strictEqual(FX.toNumber('0'), 0);
  assert.strictEqual(FX.toNumber(''), null);
  assert.strictEqual(FX.toNumber(null), null);
  assert.strictEqual(FX.toNumber('--'), null);
  assert.strictEqual(FX.toNumber('abc'), null);
});

test('parseDateValue：Excel 序列号 / 多种字符串格式', () => {
  assert.strictEqual(FX.toISODate(FX.parseDateValue(46234)), '2026-07-31');
  assert.strictEqual(FX.toISODate(FX.parseDateValue(46233)), '2026-07-30');
  assert.strictEqual(FX.toISODate(FX.parseDateValue('2026-07-31')), '2026-07-31');
  assert.strictEqual(FX.toISODate(FX.parseDateValue('2026/7/31')), '2026-07-31');
  assert.strictEqual(FX.toISODate(FX.parseDateValue('2026年7月31日')), '2026-07-31');
  assert.strictEqual(FX.toISODate(FX.parseDateValue('7/31/2026')), '2026-07-31');
  assert.strictEqual(FX.toISODate(FX.parseDateValue('20260731')), '2026-07-31');
  assert.strictEqual(FX.toISODate(FX.parseDateValue(new Date(2026, 6, 31, 15, 30))), '2026-07-31');
  assert.strictEqual(FX.toISODate(FX.parseDateValue('2026-07-31T12:00:00Z')), '2026-07-31');
  assert.strictEqual(FX.parseDateValue('不是日期'), null);
  assert.strictEqual(FX.parseDateValue(20261301), null);
  assert.strictEqual(FX.parseDateValue(''), null);
});

test('Excel 1904 日期系统', () => {
  assert.strictEqual(FX.toISODate(FX.parseDateValue(44772, true)), '2026-07-31');
});

test('roundTo / formatNumber / formatThousands', () => {
  assert.strictEqual(FX.roundTo(8.83 * 6.7468, 2), 59.57);
  assert.strictEqual(FX.roundTo(1.005, 2), 1.01);
  assert.strictEqual(FX.roundTo(-1.005, 2), -1.01);
  assert.strictEqual(FX.formatNumber(6.74684, 4), '6.7468');
  assert.strictEqual(FX.formatNumber(6.7, 2), '6.70');
  assert.strictEqual(FX.formatThousands(10994.05, 2), '10,994.05');
  assert.strictEqual(FX.formatNumber(null, 2), '');
});

test('addDaysISO 跨月跨年', () => {
  assert.strictEqual(FX.addDaysISO('2026-07-31', 1), '2026-08-01');
  assert.strictEqual(FX.addDaysISO('2026-12-31', 1), '2027-01-01');
  assert.strictEqual(FX.addDaysISO('2026-03-01', -1), '2026-02-28');
});

test('detectHeaderRow：跳过标题行', () => {
  const aoa = [
    ['2026 年 7 月交易明细', null, null],
    [null, null, null],
    ['日期', '订单编号', '总计 (USD)'],
    [46234, '112-1', 8.83]
  ];
  assert.strictEqual(FX.detectHeaderRow(aoa), 2);
});

test('tableFromAOA：构建对象行、跳过全空行', () => {
  const aoa = [
    ['日期', '总计 (USD)'],
    [46234, 8.83],
    [null, null],
    [46233, 12.64]
  ];
  const t = FX.tableFromAOA(aoa);
  assert.deepStrictEqual(t.headers, ['日期', '总计 (USD)']);
  assert.strictEqual(t.rows.length, 2);
  assert.strictEqual(t.rows[1]['总计 (USD)'], 12.64);
});

test('列识别：本项目的真实表头', () => {
  const headers = ['日期', '交易状态', '交易类型', '订单编号', '商品详情',
    '商品价格总额', '促销返点总额', '亚马逊所收费用', '其他', '总计 (USD)'];
  assert.strictEqual(FX.detectDateColumn(headers), 0);
  assert.strictEqual(FX.detectAmountColumn(headers), 9);
});

test('列识别：值层面兜底', () => {
  const headers = ['A', 'B', 'C'];
  const rows = [
    { A: 46234, B: '订单付款', C: 8.83 },
    { A: 46233, B: '订单付款', C: 12.64 },
    { A: 46232, B: '退款', C: -3.5 }
  ];
  const guess = FX.detectColumnsByValues(rows, headers);
  assert.strictEqual(guess.date, 0);
  assert.strictEqual(guess.amount, 2);
});

test('buildRateMap + lookupRate：非交易日沿用前一工作日', () => {
  const idx = FX.buildRateMap([
    { date: '2026-07-29', rate: 6.7631 },
    { date: '2026-07-30', rate: 6.7545 },
    { date: '2026-07-31', rate: 6.7468 },
    { date: '2026-08-03', rate: 6.7469 }
  ]);
  const exact = FX.lookupRate(idx, '2026-07-31');
  assert.strictEqual(exact.rate, 6.7468);
  assert.strictEqual(exact.exact, true);

  const weekend = FX.lookupRate(idx, '2026-08-01');
  assert.strictEqual(weekend.rate, 6.7468);
  assert.strictEqual(weekend.rateDate, '2026-07-31');
  assert.strictEqual(weekend.exact, false);

  const before = FX.lookupRate(idx, '2026-07-01');
  assert.strictEqual(before.rateDate, '2026-07-29');
  assert.strictEqual(before.exact, false);
});

test('fetchRates：解析 v2 数组格式', async () => {
  const calls = [];
  const fakeFetch = async (url) => {
    calls.push(url);
    assert.match(url, /\/v2\/rates\?/);
    assert.match(url, /base=USD/);
    assert.match(url, /quotes=CNY/);
    return {
      ok: true,
      json: async () => ([
        { date: '2026-07-30', base: 'USD', quote: 'CNY', rate: 6.7545 },
        { date: '2026-07-31', base: 'USD', quote: 'CNY', rate: 6.7468 }
      ])
    };
  };
  const series = await FX.fetchRates({
    base: 'USD', quote: 'CNY', from: '2026-07-30', to: '2026-07-31', fetchImpl: fakeFetch
  });
  assert.strictEqual(calls.length, 1);
  assert.deepStrictEqual(series, [
    { date: '2026-07-30', rate: 6.7545 },
    { date: '2026-07-31', rate: 6.7468 }
  ]);
});

test('fetchRates：v2 失败时回退 v1 结构', async () => {
  let n = 0;
  const fakeFetch = async (url) => {
    n++;
    if (url.indexOf('/v2/') !== -1) throw new Error('v2 挂了');
    assert.match(url, /\/v1\/2026-07-30\.\.2026-07-31\?base=USD&symbols=CNY/);
    return { ok: true, json: async () => ({ rates: { '2026-07-31': { CNY: 6.7513 } } }) };
  };
  const series = await FX.fetchRates({
    base: 'USD', quote: 'CNY', from: '2026-07-30', to: '2026-07-31', fetchImpl: fakeFetch
  });
  assert.strictEqual(n, 2);
  assert.deepStrictEqual(series, [{ date: '2026-07-31', rate: 6.7513 }]);
});

test('fetchRates：长区间自动分片（每片不超过 366 天）', async () => {
  const calls = [];
  const fakeFetch = async (url) => {
    calls.push(url);
    const m = url.match(/from=([\d-]+)&to=([\d-]+)/);
    const start = FX.fromISODate(m[1]);
    const end = FX.fromISODate(m[2]);
    const days = Math.round((end - start) / 86400000) + 1;
    assert.ok(days <= 366, `分片天数 ${days} 超出上限`);
    return { ok: true, json: async () => ([{ date: m[1], rate: 7 }]) };
  };
  await FX.fetchRates({
    base: 'USD', quote: 'CNY', from: '2026-01-01', to: '2027-03-01', fetchImpl: fakeFetch
  });
  assert.strictEqual(calls.length, 2);
});

test('fetchRates：HTTP 错误会抛异常', async () => {
  const fakeFetch = async () => ({ ok: false, status: 422, json: async () => ({}) });
  await assert.rejects(
    () => FX.fetchRates({ base: 'USD', quote: 'CNY', from: '2026-07-01', to: '2026-07-02', fetchImpl: fakeFetch }),
    /HTTP 422/
  );
});

test('summarize：合计与缺失统计', () => {
  const records = [
    { amt: 8.83, rmb: 59.57, __isoDate: '2026-07-31', __rateInfo: { rateDate: '2026-07-31' } },
    { amt: 12.64, rmb: 85.28, __isoDate: '2026-08-31', __rateInfo: { rateDate: '2026-08-31' } },
    { amt: null, rmb: null, __isoDate: '2026-09-30', __rateInfo: null }
  ];
  const s = FX.summarize(records, 'amt', 'rmb');
  assert.strictEqual(s.count, 3);
  assert.strictEqual(FX.roundTo(s.totalBase, 2), 21.47);
  assert.strictEqual(FX.roundTo(s.totalTarget, 2), 144.85);
  assert.strictEqual(s.missingAmount, 1);
  assert.strictEqual(s.minDate, '2026-07-31');
  assert.strictEqual(s.maxDate, '2026-09-30');
  assert.deepStrictEqual(s.rateDates, ['2026-07-31', '2026-08-31']);
});
