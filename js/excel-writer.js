/*
 * Writes the processed model to a formatted .xlsx using ExcelJS.
 * Browser: window.ReportWriter.build(ExcelJS, model) ; Node: require(...).build
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ReportWriter = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var COLORS = {
    title: 'FF1F3864',
    headerFill: 'FF1F4E79',
    headerFont: 'FFFFFFFF',
    border: 'FFBFBFBF',
    band: 'FFF5F8FC',
    summaryFill: 'FFDDEBF7',
    detailFont: 'FF404040',
    totalFill: 'FFFCE4D6'
  };
  var thin = { style: 'thin', color: { argb: COLORS.border } };
  var BORDER = { top: thin, left: thin, bottom: thin, right: thin };

  function textLen(v) {
    if (v === null || v === undefined) return 0;
    if (v instanceof Date) return 11;
    if (typeof v === 'number') return v.toLocaleString('en-US').length;
    return String(v).split('\n').reduce(function (m, l) { return Math.max(m, l.length); }, 0);
  }

  function safeSheetName(name, used) {
    var n = String(name || 'Sheet').replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31).trim() || 'Sheet';
    var base = n, i = 2;
    while (used[n.toLowerCase()]) { n = base.slice(0, 28) + ' ' + i++; }
    used[n.toLowerCase()] = true;
    return n;
  }

  function build(ExcelJS, model) {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Report Correction Tool';
    wb.created = new Date();
    var used = {};
    model.sheets.forEach(function (s) { writeSheet(wb, s, safeSheetName(s.name, used)); });
    return wb;
  }

  function writeSheet(wb, s, sheetName) {
    var ws = wb.addWorksheet(sheetName, {
      properties: { outlineProperties: { summaryBelow: false, summaryRight: false } },
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 }
    });
    var ncol = Math.max(s.headers.length, 1);
    var r = 1;

    // Title rows (report name, date range, etc. found above the header)
    s.titleRows.forEach(function (t, i) {
      var row = ws.getRow(r);
      row.getCell(1).value = t;
      if (ncol > 1) ws.mergeCells(r, 1, r, ncol);
      row.getCell(1).font = { bold: true, size: i === 0 ? 14 : 11, color: { argb: COLORS.title } };
      row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      row.height = i === 0 ? 24 : 18;
      r++;
    });
    if (s.titleRows.length) r++; // spacer

    if (!s.headers.length) return;

    // Header
    var headerRowNum = r;
    var hr = ws.getRow(r);
    s.headers.forEach(function (h, ci) {
      var c = hr.getCell(ci + 1);
      c.value = h;
      c.font = { bold: true, color: { argb: COLORS.headerFont }, size: 11 };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.headerFill } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      c.border = BORDER;
    });
    hr.height = 32;
    r++;

    var widths = s.headers.map(function (h) { return Math.min(textLen(h), 22) + 4; });
    var band = 0;

    s.rows.forEach(function (row) {
      var xr = ws.getRow(r);
      if (row.type !== 'detail') band++;
      row.cells.forEach(function (v, ci) {
        if (ci >= ncol) return;
        var c = xr.getCell(ci + 1);
        var display = v;
        if (row.type === 'detail' && ci === s.itemCol) display = '    ' + (v === null || v === undefined ? '' : v);
        c.value = display === undefined ? null : display;
        c.border = BORDER;
        var isNum = typeof v === 'number';
        var isDash = v === '-';
        var isDate = v instanceof Date;
        if (isNum) c.numFmt = '#,##0';
        if (isDate) c.numFmt = 'dd-mmm-yyyy';
        c.alignment = {
          vertical: 'middle',
          horizontal: isNum ? 'right' : (isDash || isDate) ? 'center' : (s.numericCols[ci] && !isEmpty(v) ? 'right' : 'left'),
          wrapText: false,
          indent: (!isNum && !isDash && !isDate) ? 1 : 0
        };
        if (isDash && s.numericCols[ci]) c.alignment.horizontal = 'right';
        widths[ci] = Math.max(widths[ci] || 0, textLen(display) + 3);
      });
      // Style per row type
      for (var ci = 1; ci <= ncol; ci++) {
        var cell = xr.getCell(ci);
        if (row.type === 'summary') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.summaryFill } };
          cell.font = { bold: true };
        } else if (row.type === 'detail') {
          cell.font = { color: { argb: COLORS.detailFont }, italic: false, size: 10 };
        } else if (row.type === 'total') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.totalFill } };
          cell.font = { bold: true };
        } else if (band % 2 === 0) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.band } };
        }
      }
      if (row.type === 'summary') {
        var ic = xr.getCell((s.itemCol || 0) + 1);
        ic.value = '▸ ' + (row.cells[s.itemCol] === null ? '' : row.cells[s.itemCol]) + '  (' + row.count + ')';
        ic.note = 'Repeated item: ' + row.count + ' records. Click the [+] in the left margin to expand.';
        widths[s.itemCol] = Math.max(widths[s.itemCol], textLen(ic.value) + 3);
      }
      if (row.type === 'detail') { xr.outlineLevel = 1; xr.hidden = true; }
      xr.height = row.type === 'detail' ? 18 : 20;
      r++;
    });

    s.headers.forEach(function (_, ci) {
      ws.getColumn(ci + 1).width = Math.max(8, Math.min(widths[ci] || 10, 60));
    });

    ws.views = [{ state: 'frozen', ySplit: headerRowNum, xSplit: (s.itemCol > -1 && s.itemCol <= 2) ? s.itemCol + 1 : 0, showGridLines: false }];
    if (s.rows.length) {
      ws.autoFilter = { from: { row: headerRowNum, column: 1 }, to: { row: headerRowNum + s.rows.length, column: ncol } };
    }
    ws.pageSetup.printTitlesRow = headerRowNum + ':' + headerRowNum;
  }

  function isEmpty(v) { return v === null || v === undefined || v === ''; }

  return { build: build };
});
