#!/usr/bin/env node
/* 读回生成的 XLSX，打印汇总与抽样行，用于人工/自动校验 */
'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'public', 'vendor', 'xlsx.full.min.js'));

const file = process.argv[2];
if (!file) { console.error('用法：node test/verify-output.cjs <xlsx>'); process.exit(2); }

const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
console.log('工作表：', wb.SheetNames.join(' | '));
console.log('');
console.log('=== 汇总 ===');
XLSX.utils.sheet_to_json(wb.Sheets['汇总'], { header: 1 })
  .slice(0, 16)
  .forEach((r) => console.log(r.join('  |  ')));
console.log('');
for (const name of wb.SheetNames) {
  if (name === '汇总') continue;
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
  console.log(`=== ${name}（${rows.length - 1} 行）表头 ===`);
  console.log(rows[0].join(' | '));
  console.log('--- 前 3 行 ---');
  rows.slice(1, 4).forEach((r) => console.log(r.join(' | ')));
  console.log('');
  break;
}
