/*!
 * core.js — 汇率换算核心逻辑
 * 纯函数，无 DOM 依赖，可在浏览器与 Node 中复用。
 * 汇率数据来源：https://frankfurter.dev/ (api.frankfurter.dev)
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FXCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var GLOBAL = typeof globalThis !== 'undefined'
    ? globalThis
    : (typeof self !== 'undefined' ? self : this);

  var MS_DAY = 86400000;

  /* ------------------------------------------------------------------ *
   * 数字处理
   * ------------------------------------------------------------------ */

  /**
   * 把单元格值转成数字，识别千分位、货币符号、括号负数、常见币种后缀。
   * @returns {number|null}
   */
  function toNumber(value) {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number') return isFinite(value) ? value : null;
    if (typeof value === 'boolean') return null;

    var s = String(value).trim();
    if (!s || s === '-' || s === '--' || s === '—' || s === 'N/A' || s === 'n/a') return null;

    var negative = false;
    if (/^\(.*\)$/.test(s)) {
      negative = true;
      s = s.slice(1, -1);
    }

    s = s
      .replace(/[\s\u00a0]/g, '')
      .replace(/[,，]/g, '')
      .replace(/[¥￥$€£₩₽]/g, '')
      .replace(/(USD|CNY|RMB|EUR|GBP|JPY|HKD|美元|人民币|欧元|日元|港币)/gi, '');

    if (s.charAt(0) === '-') {
      negative = !negative;
      s = s.slice(1);
    } else if (s.charAt(0) === '+') {
      s = s.slice(1);
    }

    if (!/^\d*\.?\d+$/.test(s)) return null;
    var n = parseFloat(s);
    if (!isFinite(n)) return null;
    return negative ? -n : n;
  }

  /** 四舍五入到指定小数位（避免浮点误差） */
  function roundTo(value, decimals) {
    if (value === null || value === undefined || !isFinite(value)) return null;
    var d = Math.max(0, Math.min(10, decimals | 0));
    var factor = Math.pow(10, d);
    return Math.round((value + (value >= 0 ? 1 : -1) * Number.EPSILON) * factor) / factor;
  }

  /** 格式化数字为固定小数位字符串，非法值返回空串 */
  function formatNumber(value, decimals) {
    var r = roundTo(value, decimals);
    return r === null ? '' : r.toFixed(Math.max(0, Math.min(10, decimals | 0)));
  }

  /** 带千分位的展示字符串 */
  function formatThousands(value, decimals) {
    var r = roundTo(value, decimals);
    if (r === null) return '';
    var fixed = r.toFixed(Math.max(0, Math.min(10, decimals | 0)));
    var parts = fixed.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.');
  }

  /* ------------------------------------------------------------------ *
   * 日期处理
   * ------------------------------------------------------------------ */

  function mkUTCDate(y, m, d) {
    if (!isFinite(y) || !isFinite(m) || !isFinite(d)) return null;
    if (m < 1 || m > 12 || d < 1 || d > 31) return null;
    var dt = new Date(Date.UTC(y, m - 1, d));
    if (isNaN(dt.getTime())) return null;
    // 校验回读，过滤 2 月 30 日这类非法日期
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return dt;
  }

  /** Excel 序列号 -> Date（以 UTC 表示当天零点） */
  function excelSerialToDate(serial, date1904) {
    var epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    return new Date(epoch + Math.round(serial * MS_DAY));
  }

  /**
   * 解析各种形态的日期：Date 对象、Excel 序列号、yyyy-mm-dd、mm/dd/yyyy、
   * yyyymmdd、中文年月日等。返回 UTC 零点 Date 或 null。
   */
  function parseDateValue(value, date1904) {
    if (value === null || value === undefined || value === '') return null;

    if (value instanceof Date) {
      if (isNaN(value.getTime())) return null;
      return mkUTCDate(value.getFullYear(), value.getMonth() + 1, value.getDate());
    }

    if (typeof value === 'number') {
      if (!isFinite(value)) return null;
      // 合理的 Excel 序列号区间：1900-01-01 ~ 2173-10-14
      if (value >= 1 && value <= 100000) {
        var d = excelSerialToDate(value, date1904);
        return isNaN(d.getTime()) ? null : d;
      }
      return null;
    }

    var s = String(value).trim();
    if (!s) return null;

    var m;

    // 2026-07-31 / 2026/7/31 / 2026.07.31 / 2026年7月31日
    m = s.match(/^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?/);
    if (m) return mkUTCDate(+m[1], +m[2], +m[3]);

    // 20260731
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return mkUTCDate(+m[1], +m[2], +m[3]);

    // 7/31/2026（按 月/日/年 解析）
    m = s.match(/^(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{4})/);
    if (m) return mkUTCDate(+m[3], +m[1], +m[2]);

    // 纯数字字符串形式的 Excel 序列号
    if (/^\d{5}(\.\d+)?$/.test(s)) {
      var n = parseFloat(s);
      if (n >= 20000 && n <= 80000) return excelSerialToDate(n, date1904);
    }

    // 含时区的完整时间串，例如 2026-07-31T12:00:00Z
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
    if (m) return mkUTCDate(+m[1], +m[2], +m[3]);

    var fallback = new Date(s);
    if (!isNaN(fallback.getTime())) {
      return mkUTCDate(fallback.getFullYear(), fallback.getMonth() + 1, fallback.getDate());
    }
    return null;
  }

  /** Date -> 'YYYY-MM-DD' */
  function toISODate(date) {
    if (!date || isNaN(date.getTime())) return '';
    return date.toISOString().slice(0, 10);
  }

  /** 'YYYY-MM-DD' -> Date(UTC) */
  function fromISODate(iso) {
    var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    return mkUTCDate(+m[1], +m[2], +m[3]);
  }

  /** ISO 日期字符串加减天数 */
  function addDaysISO(iso, days) {
    var d = fromISODate(iso);
    if (!d) return '';
    return toISODate(new Date(d.getTime() + days * MS_DAY));
  }

  /* ------------------------------------------------------------------ *
   * 表格结构识别
   * ------------------------------------------------------------------ */

  var EXCLUDE_HEADER_RE = /(汇率|rate|exchange|换算后|rmb|cny|人民币)/i;
  var DATE_HEADER_RE = /(日期|交易时间|记账日期|date|时间|time)/i;
  var TOTAL_HEADER_RE = /(总计|合计|总额|小计|净额|total|sum)/i;
  var AMOUNT_HEADER_RE = /(金额|amount|价格|费用|收入|price|value)/i;
  var BASE_CCY_HEADER_RE = /(usd|美元|\$|美金)/i;

  function isBlank(v) {
    return v === null || v === undefined || String(v).trim() === '';
  }

  /**
   * 在二维数组中定位表头行：前 15 行里「非空单元格最多且含有文本」的一行。
   * @returns {number} 索引，找不到返回 0
   */
  function detectHeaderRow(aoa) {
    var limit = Math.min(aoa.length, 15);
    var bestIdx = 0;
    var bestScore = -1;
    for (var i = 0; i < limit; i++) {
      var row = aoa[i] || [];
      var filled = 0;
      var textCount = 0;
      for (var j = 0; j < row.length; j++) {
        if (!isBlank(row[j])) {
          filled++;
          if (typeof row[j] === 'string') textCount++;
        }
      }
      if (filled < 2) continue;
      var score = filled + textCount * 0.5;
      if (score > bestScore) {
        bestScore = score;
        bestIdx = i;
      }
    }
    return bestIdx;
  }

  /**
   * 二维数组 -> { headers, rows }
   * rows 为对象数组，key 为列名（重复列名自动加后缀）。
   */
  function tableFromAOA(aoa, headerRowIndex) {
    if (!Array.isArray(aoa) || aoa.length === 0) return { headers: [], rows: [], keys: [] };
    var hIdx = typeof headerRowIndex === 'number' ? headerRowIndex : detectHeaderRow(aoa);
    if (hIdx >= aoa.length) hIdx = 0;

    var rawHeaders = aoa[hIdx] || [];
    var headers = [];
    var keys = [];
    var seen = {};
    for (var c = 0; c < rawHeaders.length; c++) {
      var name = isBlank(rawHeaders[c]) ? '列' + (c + 1) : String(rawHeaders[c]).trim();
      var key = name;
      if (Object.prototype.hasOwnProperty.call(seen, key)) {
        seen[key] += 1;
        key = name + ' (' + seen[name] + ')';
      } else {
        seen[key] = 1;
      }
      headers.push(name);
      keys.push(key);
    }

    var rows = [];
    for (var r = hIdx + 1; r < aoa.length; r++) {
      var src = aoa[r] || [];
      var obj = {};
      var hasValue = false;
      for (var i = 0; i < keys.length; i++) {
        var v = i < src.length ? src[i] : null;
        if (!isBlank(v)) hasValue = true;
        obj[keys[i]] = v;
      }
      if (hasValue) {
        obj.__row = r;
        obj.__sheet = null;
        rows.push(obj);
      }
    }
    return { headers: headers, keys: keys, rows: rows, headerRowIndex: hIdx };
  }

  /** 基于表头名猜测日期列索引，找不到返回 -1 */
  function detectDateColumn(headers) {
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || '');
      if (EXCLUDE_HEADER_RE.test(h)) continue;
      if (DATE_HEADER_RE.test(h)) return i;
    }
    return -1;
  }

  /** 基于表头名猜测金额列索引，找不到返回 -1 */
  function detectAmountColumn(headers) {
    var best = -1;
    var bestScore = 0;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || '');
      if (EXCLUDE_HEADER_RE.test(h)) continue;
      var score = 0;
      if (TOTAL_HEADER_RE.test(h)) score += 6;
      if (BASE_CCY_HEADER_RE.test(h)) score += 3;
      if (AMOUNT_HEADER_RE.test(h)) score += 2;
      if (score > bestScore) {
        bestScore = score;
        best = i;
      }
    }
    return bestScore > 0 ? best : -1;
  }

  /**
   * 值层面的兜底探测：找出「大多数行都能解析成日期 / 数字」的列。
   * @param {Array<Object>} rows 数据行
   * @param {Array<string>} keys 列 key
   * @returns {{date:number, amount:number}} 列索引，未命中为 -1
   */
  function detectColumnsByValues(rows, keys, date1904) {
    var cols = keys.length;
    var dateScore = new Array(cols).fill(0);
    var numScore = new Array(cols).fill(0);
    var sample = (rows || []).slice(0, Math.min(rows.length, 200));
    if (!sample.length || !cols) return { date: -1, amount: -1 };

    for (var r = 0; r < sample.length; r++) {
      for (var c = 0; c < cols; c++) {
        var v = sample[r][keys[c]];
        if (isBlank(v)) continue;
        var num = toNumber(v);
        // 只有「看起来像 Excel 序列号」或「字符串日期」才算日期列，
        // 避免 8.83 这类小额数字被误判成 1900-01-07
        var isDateLike = !!parseDateValue(v, date1904) && !(num !== null && Math.abs(num) < 20000);
        if (isDateLike) {
          dateScore[c]++;
          continue;
        }
        var raw = String(v);
        if (/[\d]/.test(raw) && num !== null && !/^\d{5}$/.test(raw)) numScore[c]++;
      }
    }

    var n0 = sample.length;
    var bestDate = -1, bestDateS = 0.6;
    var bestNum = -1, bestNumS = 0.6;
    for (var i = 0; i < cols; i++) {
      var ds = dateScore[i] / n0;
      var ns = numScore[i] / n0;
      if (ds > bestDateS) { bestDateS = ds; bestDate = i; }
      // 金额列优先取靠右的数值列（总计通常在最右）
      if (ns > bestNumS || (ns === bestNumS && bestNum >= 0 && ns > 0.6)) { bestNumS = ns; bestNum = i; }
    }
    return { date: bestDate, amount: bestNum };
  }

  /* ------------------------------------------------------------------ *
   * 汇率获取与匹配
   * ------------------------------------------------------------------ */

  var API_BASE = 'https://api.frankfurter.dev';

  /**
   * 拉取区间汇率（v2 接口，失败时回退到 v1）。
   * @param {{base:string, quote:string, from:string, to:string, fetchImpl?:Function, onProgress?:Function}} opts
   * @returns {Promise<Array<{date:string, rate:number}>>}
   */
  async function fetchRates(opts) {
    var base = (opts.base || 'USD').toUpperCase();
    var quote = (opts.quote || 'CNY').toUpperCase();
    var doFetch = opts.fetchImpl || (typeof GLOBAL.fetch === 'function' ? GLOBAL.fetch.bind(GLOBAL) : null);
    if (!doFetch) throw new Error('当前环境不支持 fetch');

    var out = [];
    var cursor = opts.from;

    while (cursor <= opts.to) {
      var end = addDaysISO(cursor, 365);
      if (end > opts.to || !end) end = opts.to;

      var result = null;
      var err = null;

      // 首选 v2
      try {
        result = await requestSeries(doFetch, API_BASE + '/v2/rates', {
          from: cursor, to: end, base: base, quotes: quote
        });
      } catch (e) {
        err = e;
      }

      // 回退 v1（参数名不同）
      if (!result) {
        try {
          result = await requestSeries(doFetch, API_BASE + '/v1/' + cursor + '..' + end, {
            base: base, symbols: quote
          });
        } catch (e2) {
          throw err || e2;
        }
      }

      for (var i = 0; i < result.length; i++) out.push(result[i]);
      if (typeof opts.onProgress === 'function') {
        opts.onProgress({ from: cursor, to: end });
      }
      cursor = addDaysISO(end, 1);
      if (!cursor) break;
    }

    // 去重 + 按日期升序
    var map = {};
    out.forEach(function (r) { map[r.date] = r.rate; });
    return Object.keys(map).sort().map(function (d) { return { date: d, rate: map[d] }; });
  }

  async function requestSeries(doFetch, url, params) {
    var qs = Object.keys(params)
      .map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); })
      .join('&');
    var res = await doFetch(url + '?' + qs, { headers: { Accept: 'application/json' } });
    if (!res.ok) throw new Error('汇率接口返回 HTTP ' + res.status);
    var json = await res.json();

    if (Array.isArray(json)) {
      return json.map(function (x) {
        return { date: String(x.date).slice(0, 10), rate: Number(x.rate) };
      }).filter(function (x) { return x.date && isFinite(x.rate); });
    }
    // v1 结构：{ rates: { '2026-07-31': { CNY: 6.75 } } }
    if (json && json.rates && typeof json.rates === 'object') {
      return Object.keys(json.rates).map(function (d) {
        var v = json.rates[d];
        var rate = v && typeof v === 'object' ? v[Object.keys(v)[0]] : v;
        return { date: String(d).slice(0, 10), rate: Number(rate) };
      }).filter(function (x) { return x.date && isFinite(x.rate); });
    }
    throw new Error('汇率接口返回了无法识别的数据格式');
  }

  /** 构建「日期 -> 汇率」索引，并按日期升序返回 */
  function buildRateMap(series) {
    var map = {};
    var dates = [];
    (series || []).forEach(function (r) {
      if (!r || !r.date) return;
      map[r.date] = r.rate;
      dates.push(r.date);
    });
    dates.sort();
    return { map: map, dates: dates };
  }

  /**
   * 查询某天的汇率。若当天无牌价（周末 / 节假日），沿用之前最近一个有牌价的日期。
   * @returns {{rate:number, rateDate:string, exact:boolean}|null}
   */
  function lookupRate(rateIndex, isoDate) {
    if (!rateIndex || !isoDate) return null;
    var map = rateIndex.map;
    var dates = rateIndex.dates;
    if (Object.prototype.hasOwnProperty.call(map, isoDate)) {
      return { rate: map[isoDate], rateDate: isoDate, exact: true };
    }
    for (var i = dates.length - 1; i >= 0; i--) {
      if (dates[i] < isoDate) return { rate: map[dates[i]], rateDate: dates[i], exact: false };
    }
    for (var j = 0; j < dates.length; j++) {
      if (dates[j] > isoDate) return { rate: map[dates[j]], rateDate: dates[j], exact: false };
    }
    return null;
  }

  /* ------------------------------------------------------------------ *
   * 汇总
   * ------------------------------------------------------------------ */

  function summarize(records, amountKey, rmbKey) {
    var usd = 0, rmb = 0, missing = 0;
    var minDate = '', maxDate = '';
    var rateSet = {};
    records.forEach(function (rec) {
      var a = toNumber(rec[amountKey]);
      if (a === null) missing++; else usd += a;
      var r = toNumber(rec[rmbKey]);
      if (r !== null) rmb += r;
      if (rec.__isoDate) {
        if (!minDate || rec.__isoDate < minDate) minDate = rec.__isoDate;
        if (!maxDate || rec.__isoDate > maxDate) maxDate = rec.__isoDate;
      }
      if (rec.__rateInfo) rateSet[rec.__rateInfo.rateDate] = true;
    });
    return {
      count: records.length,
      totalBase: usd,
      totalTarget: rmb,
      missingAmount: missing,
      minDate: minDate,
      maxDate: maxDate,
      rateDates: Object.keys(rateSet).sort()
    };
  }

  return {
    toNumber: toNumber,
    roundTo: roundTo,
    formatNumber: formatNumber,
    formatThousands: formatThousands,
    parseDateValue: parseDateValue,
    toISODate: toISODate,
    fromISODate: fromISODate,
    addDaysISO: addDaysISO,
    detectHeaderRow: detectHeaderRow,
    tableFromAOA: tableFromAOA,
    detectDateColumn: detectDateColumn,
    detectAmountColumn: detectAmountColumn,
    detectColumnsByValues: detectColumnsByValues,
    fetchRates: fetchRates,
    buildRateMap: buildRateMap,
    lookupRate: lookupRate,
    summarize: summarize,
    API_BASE: API_BASE
  };
});
