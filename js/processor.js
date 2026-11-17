/*
 * Report correction rules.
 *
 * Pure data transformation: takes the sheets of the uploaded workbook as
 * arrays of rows and returns a cleaned model plus a log of what changed.
 * Works in the browser (window.ReportProcessor) and in Node (module.exports).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ReportProcessor = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------- config
  // Header names are compared after normalising: lower-case, text in
  // brackets removed, everything except letters/digits removed.
  // e.g. "Width (mm)" -> "width", "Lot No." -> "lotno".
  var CONFIG = {
    itemName: ['itemname', 'item', 'itemdescription', 'itemdesc', 'productname', 'materialname', 'papername'],
    serialNo: ['sno', 'slno', 'srno', 'serialno', 'no', 'sn', 'sl'],
    rawAgingRemove: ['lotcode', 'lotno', 'lotnumber', 'supplierlotno', 'supplierlotnumber', 'category', 'gsm', 'qty', 'quantity', 'machineline'],
    tagEnvRemove: ['lotcode', 'lotno', 'lotnumber', 'supplierlotno', 'supplierlotnumber', 'width', 'gsm', 'sourcetype'],
    qtn: ['qtn', 'qtnkg', 'qtnmt', 'qtnton', 'qtntons'],
    // Quantity column renamed to "QTY (kg)" on the Raw Aging / TagEnv sheets, in this order of preference.
    qtyKg: ['qtn', 'qty', 'quantity'],
    holdingDays: ['holdingdays', 'holdingday', 'holddays', 'holdingdaysdays'],
    stock: ['stock', 'stockkg', 'closingstock', 'currentstock', 'stockqty', 'balancestock', 'availablestock', 'stockinhand', 'totalstock', 'stockinkg'],
    nullTokens: ['', 'null', 'nil', 'n/a', 'na', '#n/a', 'none', 'undefined', 'nan'],
    totalTokens: ['total', 'grandtotal', 'subtotal', 'overalltotal']
  };

  function norm(v) {
    if (v === null || v === undefined) return '';
    return String(v).toLowerCase().replace(/\([^)]*\)/g, '').replace(/\[[^\]]*\]/g, '').replace(/[^a-z0-9]/g, '');
  }

  function isEmpty(v) {
    if (v === null || v === undefined) return true;
    if (v instanceof Date) return false;
    if (typeof v === 'number') return isNaN(v);
    return CONFIG.nullTokens.indexOf(String(v).trim().toLowerCase()) !== -1;
  }

  function findCol(headers, aliases) {
    for (var i = 0; i < headers.length; i++) {
      if (aliases.indexOf(norm(headers[i])) !== -1) return i;
    }
    return -1;
  }

  function sheetKind(name) {
    var n = norm(name);
    if (n.indexOf('tagenv') !== -1 || (n.indexOf('all') !== -1 && n.indexOf('tag') !== -1)) return 'tagenv';
    if (n.indexOf('raw') !== -1 && (n.indexOf('aging') !== -1 || n.indexOf('ageing') !== -1)) return 'raw';
    return 'other';
  }

  // Numbers stored as text ("1,250.50") become numbers. Codes with leading
  // zeros ("00123") are left alone.
  function toNumberIfNumeric(v) {
    if (typeof v !== 'string') return v;
    var s = v.trim();
    if (!/^-?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?$/.test(s) || s === '' || s === '-') return v;
    if (/^-?0\d/.test(s)) return v;
    var n = Number(s.replace(/,/g, ''));
    return isNaN(n) ? v : n;
  }

  function roundHalfAway(n) {
    var r = Math.sign(n) * Math.round(Math.abs(n));
    return r === 0 ? 0 : r; // no -0
  }

  // -------------------------------------------------------- header detection
  function detectHeaderRow(rows) {
    var limit = Math.min(rows.length, 25);
    for (var r = 0; r < limit; r++) {
      if (findCol(rows[r] || [], CONFIG.itemName) !== -1) return r;
    }
    // Fallback: first row with the most text cells.
    var best = 0, bestCount = -1;
    for (var i = 0; i < limit; i++) {
      var c = (rows[i] || []).filter(function (v) { return typeof v === 'string' && v.trim() !== ''; }).length;
      if (c > bestCount) { best = i; bestCount = c; }
    }
    return best;
  }

  // ------------------------------------------------------------- main entry
  /**
   * @param {Array<{name:string, rows:Array<Array<any>>}>} sheets
   * @returns {{sheets:Array, log:Array<{sheet:string, message:string, level:string}>}}
   */
  function processWorkbook(sheets) {
    var log = [];
    var out = sheets.map(function (s, idx) { return processSheet(s, idx, log); });
    if (!out.some(function (s) { return s.kind === 'raw'; }))
      log.push({ sheet: '-', level: 'warn', message: 'No "Raw Aging Stock" sheet found; its column removals were skipped.' });
    if (!out.some(function (s) { return s.kind === 'tagenv'; }))
      log.push({ sheet: '-', level: 'warn', message: 'No "All TagEnv Aging Stock" sheet found; grouping was skipped.' });
    return { sheets: out, log: log };
  }

  function processSheet(sheet, index, log) {
    var name = sheet.name;
    var kind = sheetKind(name);
    var note = function (msg, level) { log.push({ sheet: name, message: msg, level: level || 'info' }); };

    var rows = (sheet.rows || []).map(function (r) { return (r || []).slice(); });
    // Trim fully empty trailing rows.
    while (rows.length && rows[rows.length - 1].every(isEmpty)) rows.pop();

    if (!rows.length) {
      note('Sheet is empty; copied as-is.');
      return { name: name, kind: kind, titleRows: [], headers: [], rows: [], numericCols: [] };
    }

    var h = detectHeaderRow(rows);
    var titleRows = rows.slice(0, h).filter(function (r) { return !r.every(isEmpty); })
      .map(function (r) { return r.filter(function (v) { return !isEmpty(v); }).join('   '); });
    var headers = rows[h].map(function (v) { return isEmpty(v) ? '' : String(v).trim(); });
    var width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, headers.length);
    while (headers.length < width) headers.push('');
    var data = rows.slice(h + 1).map(function (r) {
      var c = r.slice(); while (c.length < width) c.push(null);
      return c.map(toNumberIfNumeric);
    });

    // Drop columns that have neither a header nor any data.
    var keep = headers.map(function (hd, ci) {
      return hd !== '' || data.some(function (r) { return !isEmpty(r[ci]); });
    });
    var dropCols = function (pred) {
      var idx = [];
      headers.forEach(function (hd, ci) { if (pred(hd, ci)) idx.push(ci); });
      if (!idx.length) return [];
      var removed = idx.map(function (i) { return headers[i]; });
      headers = headers.filter(function (_, ci) { return idx.indexOf(ci) === -1; });
      data = data.map(function (r) { return r.filter(function (_, ci) { return idx.indexOf(ci) === -1; }); });
      return removed;
    };
    dropCols(function (_, ci) { return !keep[ci]; });

    // ---- Rule 3 / 4: column removal
    if (kind === 'raw' || kind === 'tagenv') {
      var list = kind === 'raw' ? CONFIG.rawAgingRemove : CONFIG.tagEnvRemove;
      var removed = dropCols(function (hd) { return list.indexOf(norm(hd)) !== -1; });
      note(removed.length ? 'Removed columns: ' + removed.join(', ') + '.' : 'None of the columns to remove were present.',
        removed.length ? 'info' : 'warn');
    }

    var itemCol = findCol(headers, CONFIG.itemName);
    if (itemCol === -1) note('No "Item Name" column found; empty-row removal and grouping skipped for this sheet.', 'warn');

    // ---- Rule 2: Day Consumption tons -> kg, recompute No. of Days
    var dayCol = -1, daysCol = -1;
    headers.forEach(function (hd, ci) {
      var n = norm(hd);
      if (n.indexOf('consumption') === -1) return;
      if (n.indexOf('noofdays') !== -1 || n.indexOf('days') !== -1) { if (daysCol === -1) daysCol = ci; }
      else if (dayCol === -1) dayCol = ci;
    });
    if (dayCol !== -1) {
      var stockCol = findCol(headers, CONFIG.stock);
      var converted = 0;
      data.forEach(function (r) {
        var t = r[dayCol];
        if (typeof t === 'number') { r[dayCol] = t * 1000; converted++; }
        if (daysCol !== -1) {
          var kg = r[dayCol];
          if (stockCol !== -1 && typeof r[stockCol] === 'number' && typeof kg === 'number') {
            r[daysCol] = kg === 0 ? null : r[stockCol] / kg;
          } else if (typeof r[daysCol] === 'number' && typeof t === 'number') {
            // Old value was stock / tons; using kg means dividing by 1000.
            r[daysCol] = r[daysCol] / 1000;
          }
        }
      });
      var oldHeader = headers[dayCol];
      headers[dayCol] = /\b(tons?|mt|tonnes?)\b/i.test(oldHeader)
        ? oldHeader.replace(/\b(tons?|mt|tonnes?)\b/ig, 'Kg')
        : oldHeader + ' (Kg)';
      note('Converted ' + converted + ' "' + oldHeader + '" values from tons to kg.');
      if (daysCol !== -1) {
        note(stockCol !== -1
          ? '"' + headers[daysCol] + '" recalculated as "' + headers[stockCol] + '" ÷ Day Consumption (Kg).'
          : '"' + headers[daysCol] + '" recalculated using Day Consumption in kg (no stock column found, so the existing value was scaled by 1/1000).');
      }
    }

    // ---- Rule 1: remove rows that only have an item name
    var serialCols = [];
    headers.forEach(function (hd, ci) { if (CONFIG.serialNo.indexOf(norm(hd)) !== -1) serialCols.push(ci); });
    var before = data.length;
    data = data.filter(function (r) {
      var others = r.filter(function (_, ci) { return ci !== itemCol && serialCols.indexOf(ci) === -1; });
      var hasOther = others.some(function (v) { return !isEmpty(v); });
      if (hasOther) return true;
      return false; // only item name (or nothing at all)
    });
    var removedRows = before - data.length;
    if (removedRows) note('Removed ' + removedRows + ' row(s) with no details apart from the item name (or fully blank).');

    // ---- Raw Aging Stock: Holding Days, highest first (Total rows stay last)
    if (kind === 'raw') {
      var hdc = findCol(headers, CONFIG.holdingDays);
      if (hdc === -1) note('No "Holding Days" column found, so rows were not sorted.', 'warn');
      else {
        var rawTotal = function (r) { return itemCol !== -1 && CONFIG.totalTokens.indexOf(norm(r[itemCol])) !== -1; };
        var body = data.filter(function (r) { return !rawTotal(r); })
          .map(function (r, i) { return { r: r, i: i }; })
          .sort(function (a, b) {
            var x = a.r[hdc], y = b.r[hdc], xn = typeof x === 'number', yn = typeof y === 'number';
            if (xn && yn && x !== y) return y - x;
            if (xn !== yn) return xn ? -1 : 1;
            return a.i - b.i; // stable for ties
          })
          .map(function (o) { return o.r; });
        data = body.concat(data.filter(rawTotal));
        note('Sorted by Holding Days (highest first).');
      }
    }

    // Renumber S.No after removals.
    serialCols.forEach(function (sc) {
      var n = 0;
      data.forEach(function (r) { if (typeof r[sc] === 'number') r[sc] = ++n; });
    });

    var numericCols = headers.map(function (_, ci) {
      var nums = 0, texts = 0;
      data.forEach(function (r) {
        if (typeof r[ci] === 'number') nums++;
        else if (!isEmpty(r[ci]) && !(r[ci] instanceof Date)) texts++;
      });
      return nums > 0 && nums >= texts;
    });

    var isTotal = function (r) { return itemCol !== -1 && CONFIG.totalTokens.indexOf(norm(r[itemCol])) !== -1; };

    // ---- Rule 4: grouping on TagEnv
    var outRows;
    if (kind === 'tagenv' && itemCol !== -1) {
      outRows = groupRows(data, headers, itemCol, isTotal, note);
    } else {
      outRows = data.map(function (r) { return { type: isTotal(r) ? 'total' : 'data', cells: r }; });
    }

    // ---- Rule 6: round decimals (after grouping so averages use exact values)
    var rounded = 0;
    outRows.forEach(function (row) {
      var r = row.cells;
      for (var ci = 0; ci < r.length; ci++) {
        if (typeof r[ci] === 'number' && !Number.isInteger(r[ci])) { r[ci] = roundHalfAway(r[ci]); rounded++; }
      }
    });
    if (rounded) note('Rounded ' + rounded + ' decimal value(s) to whole numbers.');

    // ---- Rule 5 / 7: zero -> "-"
    var zeros = 0;
    outRows.forEach(function (row) {
      row.cells = row.cells.map(function (v) {
        if (typeof v === 'number' && v === 0) { zeros++; return '-'; }
        if (typeof v === 'string' && /^\s*-?0+(\.0+)?\s*$/.test(v)) { zeros++; return '-'; }
        return v;
      });
    });
    if (zeros) note('Replaced ' + zeros + ' zero value(s) with "-".');

    // ---- Quantity heading becomes "QTY (kg)" (QTN preferred, then Qty, then Quantity)
    if (kind === 'raw' || kind === 'tagenv') {
      for (var qa = 0; qa < CONFIG.qtyKg.length; qa++) {
        var qc = findCol(headers, [CONFIG.qtyKg[qa]]);
        if (qc === -1) continue;
        if (headers[qc] !== 'QTY (kg)') {
          note('Renamed heading "' + headers[qc] + '" to "QTY (kg)".');
          headers[qc] = 'QTY (kg)';
        }
        break;
      }
    }

    return { name: name, kind: kind, index: index, titleRows: titleRows, headers: headers, rows: outRows, numericCols: numericCols, itemCol: itemCol };
  }

  function groupRows(data, headers, itemCol, isTotal, note) {
    var qtnCol = findCol(headers, CONFIG.qtn);
    var hdCol = findCol(headers, CONFIG.holdingDays);
    if (qtnCol === -1) note('No "QTN" column found for subtotals.', 'warn');
    if (hdCol === -1) note('No "Holding Days" column found for averages.', 'warn');

    var order = [], groups = {}, totals = [];
    data.forEach(function (r) {
      if (isTotal(r)) { totals.push(r); return; }
      var key = isEmpty(r[itemCol]) ? '\u0000blank' : String(r[itemCol]).trim().toLowerCase();
      if (!groups[key]) { groups[key] = []; order.push(key); }
      groups[key].push(r);
    });

    // Descending by Holding Days; rows without a value go last. Stable for ties.
    var hdCmp = function (a, b) {
      var x = hdCol === -1 ? null : a[hdCol], y = hdCol === -1 ? null : b[hdCol];
      var xn = typeof x === 'number', yn = typeof y === 'number';
      if (xn && yn) return y - x;
      return xn ? -1 : yn ? 1 : 0;
    };
    var blocks = [], groupId = 0, multi = 0;
    order.forEach(function (key) {
      var recs = groups[key];
      if (recs.length === 1) { blocks.push([{ type: 'data', cells: recs[0] }]); return; }
      multi++; groupId++;
      var summary = headers.map(function (_, ci) {
        if (ci === itemCol) return recs[0][itemCol];
        if (ci === qtnCol) {
          var nums = recs.map(function (r) { return r[ci]; }).filter(function (v) { return typeof v === 'number'; });
          return nums.length ? nums.reduce(function (a, b) { return a + b; }, 0) : null;
        }
        if (ci === hdCol) {
          var hv = recs.map(function (r) { return r[ci]; }).filter(function (v) { return typeof v === 'number'; });
          return hv.length ? hv.reduce(function (a, b) { return a + b; }, 0) / hv.length : null;
        }
        // Show a value on the summary row only when every record agrees.
        var first = recs[0][ci];
        var same = recs.every(function (r) { return String(r[ci]) === String(first); });
        return same && !(first instanceof Date) ? first : null;
      });
      var block = [{ type: 'summary', cells: summary, groupId: groupId, count: recs.length }];
      recs.slice().sort(hdCmp).forEach(function (r) { block.push({ type: 'detail', cells: r, groupId: groupId }); });
      blocks.push(block);
    });
    if (hdCol !== -1) blocks.sort(function (a, b) { return hdCmp(a[0].cells, b[0].cells); });
    var out = [];
    blocks.forEach(function (b) { out.push.apply(out, b); });
    totals.forEach(function (r) { out.push({ type: 'total', cells: r }); });
    note('Grouped ' + multi + ' repeated item name(s): QTN totalled, Holding Days averaged. ' +
      order.length + ' unique items shown' + (hdCol !== -1 ? ', sorted by Holding Days (highest first).' : '.'));
    return out;
  }

  return { processWorkbook: processWorkbook, CONFIG: CONFIG, _norm: norm, _sheetKind: sheetKind };
});
