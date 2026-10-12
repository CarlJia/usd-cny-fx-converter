/*!
 * app.js — 页面交互与编排
 * 依赖：core.js（FXCore）、vendor/xlsx.full.min.js（SheetJS）
 */
(function () {
  'use strict';

  var FX = window.FXCore;
  var XLSX = window.XLSX;

  var CURRENCIES = ['USD', 'CNY', 'EUR', 'GBP', 'JPY', 'HKD', 'TWD', 'AUD', 'CAD', 'SGD',
    'KRW', 'CHF', 'NZD', 'THB', 'MYR', 'PHP', 'IDR', 'VND', 'INR', 'AED', 'SAR', 'TRY',
    'RUB', 'BRL', 'MXN', 'ZAR', 'SEK', 'NOK', 'DKK', 'PLN'];

  var MAX_RENDER_ROWS = 1500;
  var CACHE_KEY = 'fx-rate-cache-v1';

  var state = {
    files: [],       // { id, name, size, sheetCount }
    sheets: [],      // 见 parseSheet 返回值
    result: null,    // { columns, records, bySheet, summary, settings, sourceName }
    activeTab: '__all__',
    onlyNonExact: false
  };

  var $ = function (sel) { return document.querySelector(sel); };

  /* ================================================================== *
   * 货币下拉
   * ================================================================== */
  function fillCurrencySelects() {
    [['#baseCurrency', 'USD'], ['#quoteCurrency', 'CNY']].forEach(function (pair) {
      var sel = $(pair[0]);
      if (!sel) return;
      sel.innerHTML = '';
      CURRENCIES.forEach(function (c) {
        var o = document.createElement('option');
        o.value = c;
        o.textContent = c;
        sel.appendChild(o);
      });
      sel.value = pair[1];
    });
  }

  /* ================================================================== *
   * 文件解析
   * ================================================================== */
  function decodeText(buffer) {
    var u8 = new Uint8Array(buffer);
    var text = new TextDecoder('utf-8', { fatal: false }).decode(u8);
    if (text.indexOf('\uFFFD') !== -1) {
      try { text = new TextDecoder('gbk').decode(u8); } catch (e) { /* 忽略 */ }
    }
    return text.replace(/^\uFEFF/, '');
  }

  function readWorkbook(file) {
    return file.arrayBuffer().then(function (buf) {
      var isDelimited = /\.(csv|tsv|txt)$/i.test(file.name);
      var wb;
      try {
        if (isDelimited) {
          wb = XLSX.read(decodeText(buf), { type: 'string', raw: true, cellDates: false });
        } else {
          wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false });
        }
      } catch (e) {
        if (isDelimited) {
          wb = XLSX.read(new Uint8Array(buf), { type: 'array', raw: true, cellDates: false });
        } else {
          throw new Error('无法解析该文件，请确认是有效的 XLSX / CSV');
        }
      }
      var date1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
      return { wb: wb, date1904: date1904 };
    });
  }

  function parseSheet(ws, sheetName, fileId, fileName, date1904) {
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    if (!aoa.length) return null;

    var table = FX.tableFromAOA(aoa);
    if (!table.headers.length || !table.rows.length) return null;

    var dateCol = FX.detectDateColumn(table.headers);
    var amountCol = FX.detectAmountColumn(table.headers);

    if (dateCol === -1 || amountCol === -1) {
      var byValue = FX.detectColumnsByValues(table.rows, table.keys, date1904);
      if (dateCol === -1) dateCol = byValue.date;
      if (amountCol === -1) amountCol = byValue.amount;
      if (amountCol === dateCol) amountCol = -1;
    }

    return {
      id: fileId + '::' + sheetName,
      fileId: fileId,
      fileName: fileName,
      name: sheetName,
      headers: table.headers,
      keys: table.keys,
      rows: table.rows,
      headerRowIndex: table.headerRowIndex,
      date1904: date1904,
      include: true,
      dateCol: dateCol,
      amountCol: amountCol
    };
  }

  function addFiles(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return Promise.resolve();

    var tasks = files.map(function (file) {
      var id = 'f' + Date.now() + Math.random().toString(36).slice(2, 7);
      return readWorkbook(file).then(function (out) {
        var added = [];
        out.wb.SheetNames.forEach(function (name) {
          var ws = out.wb.Sheets[name];
          if (!ws) return;
          var sh = parseSheet(ws, name, id, file.name, out.date1904);
          if (sh) added.push(sh);
        });
        if (!added.length) {
          state.files.push({ id: id, name: file.name, size: file.size, ok: false, error: '未找到可解析的数据表' });
          return;
        }
        state.files.push({ id: id, name: file.name, size: file.size, ok: true, sheetCount: added.length });
        state.sheets = state.sheets.concat(added);
      }).catch(function (err) {
        state.files.push({ id: id, name: file.name, size: file.size, ok: false, error: err.message || String(err) });
      });
    });

    return Promise.all(tasks).then(function () {
      state.result = null;
      renderFileList();
      renderMapping();
      updateRunState();
    });
  }

  /* ================================================================== *
   * 渲染：文件列表
   * ================================================================== */
  function formatSize(bytes) {
    if (!bytes && bytes !== 0) return '';
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1024 / 1024).toFixed(2) + ' MB';
  }

  function renderFileList() {
    var box = $('#fileList');
    box.innerHTML = '';
    state.files.forEach(function (f) {
      var row = document.createElement('div');
      row.className = 'file-item' + (f.ok ? '' : ' error');
      var meta = f.ok ? (formatSize(f.size) + ' · ' + f.sheetCount + ' 个工作表') : f.error;
      row.innerHTML =
        '<span>📄</span>' +
        '<span class="fname"></span>' +
        '<span class="fmeta"></span>' +
        '<button class="fdel" title="移除">✕</button>';
      row.querySelector('.fname').textContent = f.name;
      row.querySelector('.fmeta').textContent = meta;
      row.querySelector('.fdel').addEventListener('click', function () {
        state.files = state.files.filter(function (x) { return x.id !== f.id; });
        state.sheets = state.sheets.filter(function (x) { return x.fileId !== f.id; });
        state.result = null;
        $('#resultCard').hidden = true;
        renderFileList();
        renderMapping();
        updateRunState();
      });
      box.appendChild(row);
    });
  }

  /* ================================================================== *
   * 渲染：字段映射
   * ================================================================== */
  function renderMapping() {
    var card = $('#mappingCard');
    var body = $('#mappingBody');
    var settings = $('#settingsCard');

    if (!state.sheets.length) {
      card.hidden = true;
      settings.hidden = true;
      return;
    }
    card.hidden = false;
    settings.hidden = false;
    body.innerHTML = '';

    state.sheets.forEach(function (sh) {
      var tr = document.createElement('tr');

      // 参与
      var tdOn = document.createElement('td');
      tdOn.className = 'col-on';
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = sh.include;
      cb.addEventListener('change', function () { sh.include = cb.checked; updateRunState(); });
      tdOn.appendChild(cb);

      // 名称
      var tdName = document.createElement('td');
      var nm = document.createElement('div');
      nm.className = 'sheet-label';
      nm.textContent = sh.name;
      var of = document.createElement('div');
      of.className = 'sheet-of';
      of.textContent = sh.fileName;
      tdName.appendChild(nm);
      tdName.appendChild(of);

      // 行数
      var tdCount = document.createElement('td');
      tdCount.className = 'num';
      tdCount.textContent = sh.rows.length;

      // 表头行
      var tdHead = document.createElement('td');
      tdHead.textContent = '第 ' + (sh.headerRowIndex + 1) + ' 行';

      // 日期列
      var tdDate = document.createElement('td');
      tdDate.appendChild(columnSelect(sh, 'dateCol'));
      // 金额列
      var tdAmt = document.createElement('td');
      tdAmt.appendChild(columnSelect(sh, 'amountCol'));

      tr.appendChild(tdOn);
      tr.appendChild(tdName);
      tr.appendChild(tdCount);
      tr.appendChild(tdHead);
      tr.appendChild(tdDate);
      tr.appendChild(tdAmt);
      body.appendChild(tr);
    });
  }

  function columnSelect(sheet, field) {
    var sel = document.createElement('select');
    var empty = document.createElement('option');
    empty.value = '-1';
    empty.textContent = '— 未选择 —';
    sel.appendChild(empty);

    sheet.headers.forEach(function (h, i) {
      var o = document.createElement('option');
      o.value = String(i);
      o.textContent = (i + 1) + '. ' + h;
      sel.appendChild(o);
    });
    sel.value = String(sheet[field]);

    var valid = sheet[field] >= 0 && sheet[field] !== sheet[field === 'dateCol' ? 'amountCol' : 'dateCol'];
    if (!valid) sel.classList.add('warn');
    sel.addEventListener('change', function () {
      sheet[field] = parseInt(sel.value, 10);
      renderMapping();
      updateRunState();
    });
    return sel;
  }

  function updateRunState() {
    var usable = state.sheets.filter(function (s) {
      return s.include && s.dateCol >= 0 && s.amountCol >= 0 && s.dateCol !== s.amountCol;
    });
    $('#runBtn').disabled = usable.length === 0;
    var skipped = state.sheets.filter(function (s) { return s.include; }).length - usable.length;
    var status = $('#runStatus');
    if (!state.sheets.length) { status.textContent = ''; return; }
    var rows = usable.reduce(function (a, s) { return a + s.rows.length; }, 0);
    status.className = 'run-status';
    status.textContent = usable.length
      ? '将换算 ' + usable.length + ' 个工作表 / ' + rows + ' 行' + (skipped ? '（有 ' + skipped + ' 个表缺少日期或金额列，已跳过）' : '')
      : '请为至少一个工作表指定「日期列」和「美元金额列」';
  }

  /* ================================================================== *
   * 汇率缓存
   * ================================================================== */
  function loadCache() {
    try { return JSON.parse(localStorage.getItem(CACHE_KEY)) || {}; } catch (e) { return {}; }
  }
  function saveCache(cache) {
    try {
      // 控制体积：每组最多保留最近的 4000 条
      Object.keys(cache).forEach(function (k) {
        var dates = Object.keys(cache[k].rates || {});
        if (dates.length > 4000) {
          dates.sort();
          dates.slice(0, dates.length - 4000).forEach(function (d) { delete cache[k].rates[d]; });
        }
        if (cache[k].ranges && cache[k].ranges.length > 60) {
          cache[k].ranges = cache[k].ranges.slice(-60);
        }
      });
      localStorage.setItem(CACHE_KEY, JSON.stringify(cache));
    } catch (e) { /* 忽略配额错误 */ }
  }

  function isCovered(entry, from, to) {
    return (entry.ranges || []).some(function (r) { return r[0] <= from && r[1] >= to; });
  }

  async function getRateSeries(base, quote, from, to, useCache, onProgress) {
    var key = base + '_' + quote;
    var cache = loadCache();
    var entry = cache[key] || { ranges: [], rates: {} };

    if (useCache && isCovered(entry, from, to)) {
      var series = Object.keys(entry.rates)
        .filter(function (d) { return d >= from && d <= to; })
        .sort()
        .map(function (d) { return { date: d, rate: entry.rates[d] }; });
      if (series.length) {
        onProgress && onProgress('已命中本地缓存');
        return series;
      }
    }

    var fetched = await FX.fetchRates({
      base: base,
      quote: quote,
      from: from,
      to: to,
      onProgress: onProgress
    });

    if (useCache) {
      fetched.forEach(function (r) { entry.rates[r.date] = r.rate; });
      entry.ranges = (entry.ranges || []).concat([[from, to]]);
      cache[key] = entry;
      saveCache(cache);
    }
    return fetched;
  }

  /* ================================================================== *
   * 主流程
   * ================================================================== */
  function collectNeededDates(sheets) {
    var dates = {};
    sheets.forEach(function (sh) {
      sh.rows.forEach(function (raw) {
        var d = FX.parseDateValue(raw[sh.keys[sh.dateCol]], sh.date1904);
        if (d) dates[FX.toISODate(d)] = true;
      });
    });
    return Object.keys(dates).sort();
  }

  async function run() {
    var runBtn = $('#runBtn');
    var status = $('#runStatus');
    var sheets = state.sheets.filter(function (s) {
      return s.include && s.dateCol >= 0 && s.amountCol >= 0 && s.dateCol !== s.amountCol;
    });
    if (!sheets.length) return;

    var base = $('#baseCurrency').value;
    var quote = $('#quoteCurrency').value;
    var rateDecimals = clampInt($('#rateDecimals').value, 0, 8, 4);
    var amountDecimals = clampInt($('#amountDecimals').value, 0, 6, 2);
    var showRateDate = $('#showRateDate').checked;
    var useCache = $('#useCache').checked;

    if (base === quote) {
      status.className = 'run-status error';
      status.textContent = '原始币种与目标币种不能相同';
      return;
    }

    runBtn.disabled = true;
    status.className = 'run-status';
    status.textContent = '正在整理数据…';

    try {
      var needed = collectNeededDates(sheets);
      if (!needed.length) throw new Error('没有解析出任何有效日期，请检查「日期列」是否选对');
      var from = needed[0];
      var to = needed[needed.length - 1];

      status.textContent = '正在获取 ' + base + '→' + quote + ' 汇率（' + from + ' ~ ' + to + '）…';

      var series = await getRateSeries(base, quote, from, to, useCache, function (p) {
        if (p && p.from) status.textContent = '正在获取汇率区间 ' + p.from + ' ~ ' + p.to + '…';
        else if (typeof p === 'string') status.textContent = p;
      });

      if (!series.length) throw new Error('汇率接口没有返回任何数据，请稍后重试或检查币种代码');
      var rateIndex = FX.buildRateMap(series);

      var rateKey = '汇率';
      var rateDateKey = '汇率取值日';
      var rmbKey = quote === 'CNY' ? '总计 (RMB)' : '总计 (' + quote + ')';

      // 列描述
      var columns = [];
      var seenKeys = {};
      sheets.forEach(function (sh) {
        sh.keys.forEach(function (k, i) {
          if (seenKeys[k]) return;
          seenKeys[k] = true;
          columns.push({
            key: k,
            label: sh.headers[i],
            type: k === sh.keys[sh.dateCol] ? 'date' : 'raw'
          });
        });
      });
      // 把日期列类型标记清楚（跨表可能有同名不同列）
      columns.forEach(function (c) {
        if (sheets.some(function (sh) { return sh.keys[sh.dateCol] === c.key; })) c.type = 'date';
      });

      var newColumns = [
        { key: rateKey, label: rateKey, isNew: true, decimals: rateDecimals },
        { key: rateDateKey, label: rateDateKey, isNew: true, optional: true },
        { key: rmbKey, label: rmbKey, isNew: true, decimals: amountDecimals }
      ];

      var missingDates = {};
      var bySheet = {};
      var records = [];
      var summaryRows = [];
      var unparsed = 0;

      sheets.forEach(function (sh) {
        var dateKey = sh.keys[sh.dateCol];
        var amountKey = sh.keys[sh.amountCol];
        var amountDecimalsOrig = decimalsOf(sh, amountKey);
        var sheetRecords = [];
        var invalidCount = 0;

        sh.rows.forEach(function (raw) {
          var rec = {};
          sh.keys.forEach(function (k) { rec[k] = raw[k]; });

          var dObj = FX.parseDateValue(raw[dateKey], sh.date1904);
          var iso = dObj ? FX.toISODate(dObj) : '';
          var amount = FX.toNumber(raw[amountKey]);
          var info = iso ? FX.lookupRate(rateIndex, iso) : null;

          rec[dateKey] = iso || (raw[dateKey] === null || raw[dateKey] === undefined ? '' : String(raw[dateKey]));
          rec.__source = sh.name;
          rec.__isoDate = iso;
          rec.__sheetId = sh.id;
          rec.__amountKey = amountKey;
          rec.__amountDecimals = amountDecimalsOrig;

          if (info && amount !== null) {
            rec[rateKey] = FX.roundTo(info.rate, rateDecimals);
            rec[rateDateKey] = info.rateDate;
            rec[rmbKey] = FX.roundTo(amount * info.rate, amountDecimals);
            rec.__rateInfo = info;
            if (!info.exact) rec.__nonExact = true;
          } else {
            rec[rateKey] = '';
            rec[rateDateKey] = '';
            rec[rmbKey] = '';
            invalidCount++;
            if (!iso) { missingDates[String(raw[dateKey])] = true; unparsed++; }
          }
          sheetRecords.push(rec);
          records.push(rec);
        });

        var sum = FX.summarize(sheetRecords, amountKey, rmbKey);
        bySheet[sh.id] = { name: sh.name, fileName: sh.fileName, records: sheetRecords, columns: columns, invalidCount: invalidCount };
        summaryRows.push({ name: sh.name, fileName: sh.fileName, summary: sum, invalidCount: invalidCount });
      });

      state.result = {
        columns: columns,
        newColumns: newColumns,
        rateKey: rateKey,
        rateDateKey: rateDateKey,
        rmbKey: rmbKey,
        base: base,
        quote: quote,
        rateDecimals: rateDecimals,
        amountDecimals: amountDecimals,
        showRateDate: showRateDate,
        records: records,
        bySheet: bySheet,
        summaryRows: summaryRows,
        rateIndex: rateIndex,
        series: series,
        minDate: from,
        maxDate: to,
        unparsedDateRows: unparsed,
        sources: state.files.filter(function (f) { return f.ok; }).map(function (f) { return f.name; })
      };

      state.activeTab = '__all__';
      state.onlyNonExact = false;
      $('#onlyNonExact').checked = false;

      renderResult();

      status.textContent = '完成：' + records.length + ' 行，使用了 ' + series.length + ' 个交易日的汇率';
      $('#resultCard').hidden = false;
      $('#resultCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (err) {
      status.className = 'run-status error';
      status.textContent = '出错：' + (err && err.message ? err.message : String(err));
    } finally {
      runBtn.disabled = false;
    }
  }

  function clampInt(v, min, max, dflt) {
    var n = parseInt(v, 10);
    if (!isFinite(n)) return dflt;
    return Math.max(min, Math.min(max, n));
  }

  function decimalsOf(sheet, key) {
    var vals = sheet.rows.slice(0, 200);
    var maxDec = 2;
    vals.forEach(function (r) {
      var s = String(r[key] === null || r[key] === undefined ? '' : r[key]);
      var m = s.match(/\.(\d+)$/);
      if (m) maxDec = Math.max(maxDec, Math.min(6, m[1].length));
    });
    return maxDec;
  }

  /* ================================================================== *
   * 渲染：结果
   * ================================================================== */
  function renderResult() {
    var res = state.result;
    if (!res) return;

    $('#resultRange').textContent = '汇率区间 ' + res.minDate + ' ~ ' + res.maxDate +
      '　·　' + res.base + ' → ' + res.quote;

    // 汇总卡片
    var totalBase = 0, totalTarget = 0, invalid = 0;
    res.summaryRows.forEach(function (r) {
      totalBase += r.summary.totalBase;
      totalTarget += r.summary.totalTarget;
      invalid += r.invalidCount;
    });
    var nonExact = res.records.filter(function (r) { return r.__nonExact; }).length;

    var cards = [
      { k: '记录数', v: FX.formatThousands(res.records.length, 0) + ' 行', s: res.summaryRows.length + ' 个工作表' },
      { k: '总计 (USD)', v: '$' + FX.formatThousands(totalBase, res.amountDecimals), s: '原币种合计' },
      { k: '总计 (' + (res.quote === 'CNY' ? 'RMB' : res.quote) + ')', v: '¥' + FX.formatThousands(totalTarget, res.amountDecimals), s: '按逐日汇率换算后合计', accent: true },
      { k: '使用汇率日', v: FX.formatThousands(res.series.length, 0) + ' 天', s: res.minDate + ' ~ ' + res.maxDate }
    ];
    if (nonExact) cards.push({ k: '非当日取值', v: nonExact + ' 行', s: '当天无牌价，沿用最近工作日' });
    if (invalid) cards.push({ k: '未能换算', v: invalid + ' 行', s: '日期或金额缺失' });

    var box = $('#summaryCards');
    box.innerHTML = '';
    cards.forEach(function (c) {
      var d = document.createElement('div');
      d.className = 'sum-card' + (c.accent ? ' accent' : '');
      var k = document.createElement('div'); k.className = 'k'; k.textContent = c.k;
      var v = document.createElement('div'); v.className = 'v'; v.textContent = c.v;
      var s = document.createElement('div'); s.className = 's'; s.textContent = c.s;
      d.appendChild(k); d.appendChild(v); d.appendChild(s);
      box.appendChild(d);
    });

    // 按工作表汇总
    var tb = $('#sheetSummaryBody');
    tb.innerHTML = '';
    res.summaryRows.forEach(function (r) {
      var tr = document.createElement('tr');
      tr.innerHTML = '<td></td><td class="num"></td><td class="num"></td><td class="num"></td>';
      tr.children[0].textContent = r.name + '（' + r.fileName + '）';
      tr.children[1].textContent = FX.formatThousands(r.summary.count, 0);
      tr.children[2].textContent = '$' + FX.formatThousands(r.summary.totalBase, res.amountDecimals);
      tr.children[3].textContent = '¥' + FX.formatThousands(r.summary.totalTarget, res.amountDecimals);
      tb.appendChild(tr);
    });
    var trAll = document.createElement('tr');
    trAll.innerHTML = '<td><b>合计</b></td><td class="num"><b></b></td><td class="num"><b></b></td><td class="num"><b></b></td>';
    trAll.children[1].textContent = FX.formatThousands(res.records.length, 0);
    trAll.children[2].textContent = '$' + FX.formatThousands(totalBase, res.amountDecimals);
    trAll.children[3].textContent = '¥' + FX.formatThousands(totalTarget, res.amountDecimals);
    tb.appendChild(trAll);

    // 标签页
    var tabs = $('#tabs');
    tabs.innerHTML = '';
    var tabDefs = [{ id: '__all__', label: '全部（' + res.records.length + '）' }];
    Object.keys(res.bySheet).forEach(function (id) {
      var s = res.bySheet[id];
      tabDefs.push({ id: id, label: s.name + '（' + s.records.length + '）' });
    });

    tabDefs.forEach(function (t, i) {
      var b = document.createElement('button');
      b.className = 'tab' + (i === 0 ? ' active' : '');
      b.textContent = t.label;
      b.dataset.id = t.id;
      b.addEventListener('click', function () {
        state.activeTab = t.id;
        Array.prototype.forEach.call(tabs.children, function (x) { x.classList.remove('active'); });
        b.classList.add('active');
        renderTable();
      });
      tabs.appendChild(b);
    });

    renderTable();
  }

  function currentColumnsAndRecords() {
    var res = state.result;
    var cols = [];
    var records;

    if (state.activeTab === '__all__') {
      cols.push({ key: '__source', label: '来源', type: 'text' });
      records = res.records;
    } else {
      var sheet = res.bySheet[state.activeTab];
      records = sheet ? sheet.records : res.records;
    }

    res.columns.forEach(function (c) { cols.push(c); });

    res.newColumns.forEach(function (c) {
      if (c.optional && !res.showRateDate) return;
      cols.push(c);
    });

    return { cols: cols, records: records };
  }

  function renderTable() {
    var res = state.result;
    if (!res) return;
    var pack = currentColumnsAndRecords();
    var cols = pack.cols;
    var records = pack.records;

    if (state.onlyNonExact) {
      records = records.filter(function (r) { return r.__nonExact; });
    }

    var head = $('#resultHead');
    head.innerHTML = '';
    var htr = document.createElement('tr');
    cols.forEach(function (c) {
      var th = document.createElement('th');
      th.textContent = c.label;
      if (c.isNew) th.classList.add('newcol');
      if (c.decimals !== undefined || c.type === 'date' || c.type === 'text') {
        if (c.decimals !== undefined) th.classList.add('num');
      }
      htr.appendChild(th);
    });
    head.appendChild(htr);

    var body = $('#resultBody');
    body.innerHTML = '';
    var frag = document.createDocumentFragment();
    var limit = Math.min(records.length, MAX_RENDER_ROWS);

    for (var i = 0; i < limit; i++) {
      var rec = records[i];
      var tr = document.createElement('tr');
      if (rec.__nonExact) tr.className = 'fallback';
      if (!rec.__rateInfo || FX.toNumber(rec[res.rmbKey]) === null) tr.classList.add('invalid');

      cols.forEach(function (c) {
        var td = document.createElement('td');
        var v = rec[c.key];
        if (c.key === res.rateKey) {
          td.className = 'num mono';
          td.textContent = v === '' || v === null || v === undefined ? '' : FX.formatNumber(v, res.rateDecimals);
        } else if (c.key === res.rmbKey) {
          td.className = 'num newcol';
          var n = FX.toNumber(v);
          td.textContent = n === null ? '' : FX.formatThousands(n, res.amountDecimals);
        } else if (c.key === res.rateDateKey) {
          td.className = 'mono' + (rec.__nonExact ? '' : '');
          td.textContent = v || '';
          if (rec.__nonExact && v) td.title = '该日期无牌价，取 ' + v + ' 的汇率';
        } else if (c.type === 'date') {
          td.className = 'date-cell';
          td.textContent = v || '';
        } else if (c.type === 'text') {
          td.textContent = v === null || v === undefined ? '' : String(v);
        } else {
          var num = FX.toNumber(v);
          if (num !== null && typeof v === 'number') td.classList.add('num');
          var dec = rec.__amountDecimals !== undefined && c.key === rec.__amountKey ? rec.__amountDecimals : 2;
          td.textContent = num === null
            ? (v === null || v === undefined ? '' : String(v))
            : FX.formatNumber(num, dec);
        }
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    }
    body.appendChild(frag);

    var note = $('#rowNote');
    var parts = [];
    if (records.length > limit) {
      parts.push('共 ' + records.length + ' 行，表格仅显示前 ' + limit + ' 行；导出文件包含全部数据。');
    } else if (state.onlyNonExact) {
      parts.push('筛选出 ' + records.length + ' 行「汇率非当日取值」的记录。');
    } else {
      parts.push('共 ' + records.length + ' 行。');
    }
    if (res.unparsedDateRows) {
      parts.push('其中 ' + res.unparsedDateRows + ' 行日期无法识别，已置空并跳过换算（可尝试手动改「日期列」）。');
    }
    if (res.records.some(function (r) { return r.__nonExact; })) {
      parts.push('黄色底纹行表示当天没有牌价，沿用了最近一个工作日的汇率。');
    }
    note.textContent = parts.join(' ');
  }

  /* ================================================================== *
   * 导出
   * ================================================================== */
  function cellForExport(rec, c, res) {
    var v = rec[c.key];
    if (c.key === res.rateKey) {
      var r = FX.toNumber(v);
      return r === null ? '' : FX.roundTo(r, res.rateDecimals);
    }
    if (c.key === res.rmbKey) {
      var n = FX.toNumber(v);
      return n === null ? '' : FX.roundTo(n, res.amountDecimals);
    }
    if (c.key === res.rateDateKey) return v || '';
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return v;
    var num = FX.toNumber(v);
    if (num !== null && c.type === 'raw' && /^[-+]?[\d,.\s¥￥$€£]+$/.test(String(v))) return num;
    return String(v);
  }

  function exportColumns(res, includeSource) {
    var cols = [];
    if (includeSource) cols.push({ key: '__source', label: '来源', type: 'text' });
    res.columns.forEach(function (c) { cols.push(c); });
    res.newColumns.forEach(function (c) {
      if (c.optional && !res.showRateDate) return;
      cols.push(c);
    });
    return cols;
  }

  function baseName() {
    var res = state.result;
    var n = res.sources && res.sources.length ? res.sources[0] : '交易明细';
    return n.replace(/\.[^.]+$/, '');
  }

  function buildSummaryAOA(res) {
    var totalBase = 0, totalTarget = 0, invalid = 0, nonExact = 0;
    res.summaryRows.forEach(function (r) {
      totalBase += r.summary.totalBase;
      totalTarget += r.summary.totalTarget;
      invalid += r.invalidCount;
    });
    nonExact = res.records.filter(function (r) { return r.__nonExact; }).length;

    var aoa = [
      ['项目', '数值'],
      ['数据来源文件', (res.sources || []).join('、')],
      ['记录数', res.records.length],
      ['日期区间', res.minDate + ' ~ ' + res.maxDate],
      ['币种对', res.base + ' → ' + res.quote],
      ['使用汇率天数', res.series.length],
      ['汇率数据源', 'https://frankfurter.dev/ (api.frankfurter.dev)'],
      ['非当日取值行数', nonExact],
      ['未能换算行数', invalid],
      [],
      ['工作表', '记录数', '总计 (' + res.base + ')', '总计 (' + (res.quote === 'CNY' ? 'RMB' : res.quote) + ')']
    ];
    res.summaryRows.forEach(function (r) {
      aoa.push([r.name, r.summary.count,
        FX.roundTo(r.summary.totalBase, res.amountDecimals),
        FX.roundTo(r.summary.totalTarget, res.amountDecimals)]);
    });
    aoa.push(['合计', res.records.length,
      FX.roundTo(totalBase, res.amountDecimals),
      FX.roundTo(totalTarget, res.amountDecimals)]);
    return aoa;
  }

  function buildDetailAOA(records, cols, res) {
    var aoa = [cols.map(function (c) { return c.label; })];
    records.forEach(function (rec) {
      aoa.push(cols.map(function (c) { return cellForExport(rec, c, res); }));
    });
    return aoa;
  }

  function exportXlsx() {
    var res = state.result;
    if (!res) return;

    var wb = XLSX.utils.book_new();

    // 汇总
    var wsSum = XLSX.utils.aoa_to_sheet(buildSummaryAOA(res));
    wsSum['!cols'] = [{ wch: 22 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, wsSum, '汇总');

    // 全部明细
    var allCols = exportColumns(res, true);
    var wsAll = XLSX.utils.aoa_to_sheet(buildDetailAOA(res.records, allCols, res));
    wsAll['!cols'] = allCols.map(function (c) { return { wch: c.type === 'text' && c.key === '__source' ? 12 : 18 }; });
    XLSX.utils.book_append_sheet(wb, wsAll, '全部明细');

    // 各工作表
    Object.keys(res.bySheet).forEach(function (id) {
      var sheet = res.bySheet[id];
      var cols = exportColumns(res, false);
      var name = safeSheetName(sheet.name);
      var ws = XLSX.utils.aoa_to_sheet(buildDetailAOA(sheet.records, cols, res));
      ws['!cols'] = cols.map(function () { return { wch: 18 }; });
      XLSX.utils.book_append_sheet(wb, ws, name, true);
    });

    var fname = baseName() + '-' + res.quote + '换算.xlsx';
    XLSX.writeFile(wb, fname);
  }

  function safeSheetName(name) {
    var n = String(name || 'Sheet').replace(/[\\/?*\[\]:]/g, '_').slice(0, 31);
    return n || 'Sheet';
  }

  function exportCsv() {
    var res = state.result;
    if (!res) return;
    var cols = exportColumns(res, true);
    var lines = [cols.map(function (c) { return csvCell(c.label); }).join(',')];
    res.records.forEach(function (rec) {
      lines.push(cols.map(function (c) {
        var v = cellForExport(rec, c, res);
        if (typeof v === 'number') return String(v);
        return csvCell(v);
      }).join(','));
    });
    var blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    triggerDownload(blob, baseName() + '-' + res.quote + '换算.csv');
  }

  function csvCell(v) {
    var s = v === null || v === undefined ? '' : String(v);
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function triggerDownload(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  }

  /* ================================================================== *
   * 事件绑定
   * ================================================================== */
  function init() {
    if (!XLSX) {
      alert('解析库加载失败，请检查 vendor/xlsx.full.min.js 是否存在');
      return;
    }
    fillCurrencySelects();

    var dz = $('#dropzone');
    var input = $('#fileInput');

    dz.addEventListener('click', function () { input.click(); });
    dz.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
    });
    input.addEventListener('change', function () {
      addFiles(input.files);
      input.value = '';
    });

    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('drag'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('drag'); });
    });
    dz.addEventListener('drop', function (e) {
      if (e.dataTransfer && e.dataTransfer.files) addFiles(e.dataTransfer.files);
    });
    // 阻止整页拖放时浏览器直接打开文件
    window.addEventListener('dragover', function (e) { e.preventDefault(); });
    window.addEventListener('drop', function (e) { e.preventDefault(); });

    $('#runBtn').addEventListener('click', run);
    $('#resetBtn').addEventListener('click', function () {
      state = { files: [], sheets: [], result: null, activeTab: '__all__', onlyNonExact: false };
      $('#resultCard').hidden = true;
      $('#mappingCard').hidden = true;
      $('#settingsCard').hidden = true;
      $('#runStatus').textContent = '';
      renderFileList();
    });
    $('#onlyNonExact').addEventListener('change', function () {
      state.onlyNonExact = this.checked;
      renderTable();
    });
    $('#showRateDate').addEventListener('change', function () {
      if (state.result) {
        state.result.showRateDate = this.checked;
        renderResult();
      }
    });
    $('#exportXlsxBtn').addEventListener('click', exportXlsx);
    $('#exportCsvBtn').addEventListener('click', exportCsv);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
