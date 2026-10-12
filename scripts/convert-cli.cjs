#!/usr/bin/env node
/**
 * convert-cli.cjs — 命令行版：把 XLSX / CSV 交易明细转成带
 * 「汇率」与「总计 (RMB)」两列的新文件。
 *
 * 用法：
 *   node scripts/convert-cli.cjs <输入文件> [-o 输出文件] [选项]
 *
 * 选项：
 *   -o, --out <path>          输出文件路径（默认 <输入名>-CNY换算.xlsx）
 *       --base <CODE>         原始币种，默认 USD
 *       --quote <CODE>        目标币种，默认 CNY
 *       --rate-decimals <n>   汇率小数位，默认 4
 *       --amount-decimals <n> 金额小数位，默认 2
 *       --date-col <名称>     指定日期列（默认自动识别）
 *       --amount-col <名称>   指定金额列（默认自动识别）
 *       --json                只输出 JSON 结果，不写文件
 */
'use strict';

const fs = require('fs');
const path = require('path');
const FX = require(path.join(__dirname, '..', 'public', 'core.js'));
const XLSX = require(path.join(__dirname, '..', 'public', 'vendor', 'xlsx.full.min.js'));

function parseArgs(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '-o' || a === '--out') out.out = argv[++i];
    else if (a === '--base') out.base = argv[++i];
    else if (a === '--quote') out.quote = argv[++i];
    else if (a === '--rate-decimals') out.rateDecimals = parseInt(argv[++i], 10);
    else if (a === '--amount-decimals') out.amountDecimals = parseInt(argv[++i], 10);
    else if (a === '--date-col') out.dateCol = argv[++i];
    else if (a === '--amount-col') out.amountCol = argv[++i];
    else if (a === '--json') out.json = true;
    else out._.push(a);
  }
  return out;
}

function decodeText(buf) {
  const u8 = new Uint8Array(buf);
  let text = new TextDecoder('utf-8', { fatal: false }).decode(u8);
  if (text.indexOf('\uFFFD') !== -1) {
    try { text = new TextDecoder('gbk').decode(u8); } catch (e) { /* ignore */ }
  }
  return text.replace(/^\uFEFF/, '');
}

function readWorkbook(filePath) {
  const buf = fs.readFileSync(filePath);
  const isDelimited = /\.(csv|tsv|txt)$/i.test(filePath);
  const wb = isDelimited
    ? XLSX.read(decodeText(buf), { type: 'string', raw: true, cellDates: false })
    : XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false });
  const date1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
  return { wb, date1904 };
}

function pickColumn(headers, explicit, detectFn, byValueIdx) {
  if (explicit) {
    const i = headers.findIndex((h) => String(h).trim() === explicit || String(h).includes(explicit));
    if (i === -1) throw new Error(`找不到列「${explicit}」，可选列：${headers.join(' | ')}`);
    return i;
  }
  const auto = detectFn(headers);
  if (auto >= 0) return auto;
  return byValueIdx;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const input = args._[0];
  if (!input) {
    console.error('用法：node scripts/convert-cli.cjs <输入文件> [-o 输出文件] [选项]');
    process.exit(2);
  }
  if (!fs.existsSync(input)) {
    console.error('文件不存在：' + input);
    process.exit(2);
  }

  const base = (args.base || 'USD').toUpperCase();
  const quote = (args.quote || 'CNY').toUpperCase();
  const rateDecimals = Number.isFinite(args.rateDecimals) ? args.rateDecimals : 4;
  const amountDecimals = Number.isFinite(args.amountDecimals) ? args.amountDecimals : 2;

  const { wb, date1904 } = readWorkbook(input);

  // 解析每个工作表
  const sheets = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    if (!aoa.length) continue;
    const table = FX.tableFromAOA(aoa);
    if (!table.headers.length || !table.rows.length) continue;

    const byValue = FX.detectColumnsByValues(table.rows, table.keys, date1904);
    const dateCol = pickColumn(table.headers, args.dateCol, FX.detectDateColumn, byValue.date);
    let amountCol = pickColumn(table.headers, args.amountCol, FX.detectAmountColumn, byValue.amount);
    if (amountCol === dateCol) amountCol = -1;
    if (dateCol < 0 || amountCol < 0) {
      console.warn(`跳过工作表「${name}」：未能识别日期列或金额列`);
      continue;
    }
    sheets.push({ name, headers: table.headers, keys: table.keys, rows: table.rows, dateCol, amountCol });
  }

  if (!sheets.length) {
    console.error('没有找到可处理的数据表');
    process.exit(1);
  }

  // 收集日期
  const dateSet = new Set();
  for (const sh of sheets) {
    const dk = sh.keys[sh.dateCol];
    for (const row of sh.rows) {
      const d = FX.parseDateValue(row[dk], date1904);
      if (d) dateSet.add(FX.toISODate(d));
    }
  }
  const needed = Array.from(dateSet).sort();
  if (!needed.length) {
    console.error('没有解析出任何有效日期');
    process.exit(1);
  }
  const from = needed[0];
  const to = needed[needed.length - 1];

  console.log(`读取 ${path.basename(input)}：${sheets.length} 个工作表`);
  console.log(`币种 ${base} → ${quote}，日期区间 ${from} ~ ${to}，共 ${needed.length} 个数据日期`);
  console.log('正在获取汇率…');

  const series = await FX.fetchRates({ base, quote, from, to });
  if (!series.length) {
    console.error('汇率接口没有返回数据');
    process.exit(1);
  }
  console.log(`取得 ${series.length} 个交易日的汇率（来源：api.frankfurter.dev）`);

  const rateIndex = FX.buildRateMap(series);
  const RATE_KEY = '汇率';
  const RATE_DATE_KEY = '汇率取值日';
  const RMB_KEY = quote === 'CNY' ? '总计 (RMB)' : `总计 (${quote})`;

  const outWb = XLSX.utils.book_new();
  const detailAll = [];
  const summaryRows = [];
  let grandBase = 0, grandTarget = 0, nonExact = 0, invalid = 0;

  for (const sh of sheets) {
    const dateKey = sh.keys[sh.dateCol];
    const amountKey = sh.keys[sh.amountCol];
    const header = [].concat(sh.headers, [RATE_KEY, RATE_DATE_KEY, RMB_KEY]);
    const aoa = [header];
    let sBase = 0, sTarget = 0, sNonExact = 0;

    for (const row of sh.rows) {
      const d = FX.parseDateValue(row[dateKey], date1904);
      const iso = d ? FX.toISODate(d) : '';
      const amount = FX.toNumber(row[amountKey]);
      const info = iso ? FX.lookupRate(rateIndex, iso) : null;

      const line = sh.keys.map((k) => {
        const v = row[k];
        if (k === dateKey) return iso || (v == null ? '' : String(v));
        if (typeof v === 'number') return v;
        const n = FX.toNumber(v);
        return n !== null && /^[-+]?[\d,.\s¥￥$€£]+$/.test(String(v)) ? n : (v == null ? '' : String(v));
      });

      if (info && amount !== null) {
        const rmb = FX.roundTo(amount * info.rate, amountDecimals);
        line.push(FX.roundTo(info.rate, rateDecimals), info.rateDate, rmb);
        sBase += amount;
        sTarget += rmb;
        if (!info.exact) { sNonExact++; nonExact++; }
      } else {
        line.push('', '', '');
        invalid++;
      }
      aoa.push(line);
      detailAll.push(line);
    }

    grandBase += sBase;
    grandTarget += sTarget;
    summaryRows.push([sh.name, sh.rows.length, FX.roundTo(sBase, amountDecimals), FX.roundTo(sTarget, amountDecimals), sNonExact]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = header.map(() => ({ wch: 18 }));
    XLSX.utils.book_append_sheet(outWb, ws, String(sh.name).replace(/[\\/?*[\]:]/g, '_').slice(0, 31));
  }

  const summaryAoa = [
    ['项目', '数值'],
    ['数据来源文件', path.basename(input)],
    ['记录数', detailAll.length],
    ['日期区间', `${from} ~ ${to}`],
    ['币种对', `${base} → ${quote}`],
    ['使用汇率天数', series.length],
    ['汇率数据源', 'https://frankfurter.dev/ (api.frankfurter.dev)'],
    ['非当日取值行数', nonExact],
    ['未能换算行数', invalid],
    [],
    ['工作表', '记录数', `总计 (${base})`, `总计 (${quote === 'CNY' ? 'RMB' : quote})`, '非当日取值']
  ].concat(summaryRows)
    .concat([['合计', detailAll.length, FX.roundTo(grandBase, amountDecimals), FX.roundTo(grandTarget, amountDecimals), nonExact]]);

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryAoa);
  wsSummary['!cols'] = [{ wch: 22 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(outWb, wsSummary, '汇总', true);

  const result = {
    input: path.resolve(input),
    sheets: sheets.map((s) => s.name),
    base, quote,
    dateRange: [from, to],
    tradingDays: series.length,
    rows: detailAll.length,
    nonExact,
    invalid,
    totalBase: FX.roundTo(grandBase, amountDecimals),
    totalTarget: FX.roundTo(grandTarget, amountDecimals),
    perSheet: summaryRows.map((r) => ({ sheet: r[0], rows: r[1], totalBase: r[2], totalTarget: r[3], nonExact: r[4] }))
  };

  if (args.json) {
    console.log(JSON.stringify(result, null, 2));
    return;
  }

  const defOut = path.join(
    path.dirname(input),
    path.basename(input).replace(/\.[^.]+$/, '') + `-${quote}换算.xlsx`
  );
  const outPath = args.out || defOut;
  const buf = XLSX.write(outWb, { bookType: 'xlsx', type: 'buffer', compression: true });
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, buf);

  console.log('');
  console.log('换算完成');
  console.log(`  记录数      ${result.rows}`);
  console.log(`  总计 (${base})   ${result.totalBase}`);
  console.log(`  总计 (${quote})   ${result.totalTarget}`);
  console.log(`  非当日取值  ${nonExact} 行`);
  console.log(`  未能换算    ${invalid} 行`);
  console.log(`  输出文件    ${outPath}`);
}

main().catch((err) => {
  console.error('出错：' + (err && err.message ? err.message : err));
  process.exit(1);
});
