(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var drop = $('drop'), fileInput = $('file'), statusEl = $('status');
  var result = $('result'), uploadCard = $('upload-card');
  var current = null; // { model, blob, fileName }
  var PREVIEW_LIMIT = 500;

  function setStatus(msg, isError) {
    statusEl.hidden = !msg;
    statusEl.textContent = msg || '';
    statusEl.className = 'status' + (isError ? ' error' : '');
  }

  ['dragenter', 'dragover'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); });
  });
  drop.addEventListener('drop', function (e) {
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', function () {
    if (fileInput.files.length) handleFile(fileInput.files[0]);
  });
  $('reset').addEventListener('click', function () {
    current = null; fileInput.value = '';
    result.hidden = true; uploadCard.hidden = false; setStatus('');
  });
  $('download').addEventListener('click', function () {
    if (!current) return;
    var a = document.createElement('a');
    a.href = URL.createObjectURL(current.blob);
    a.download = current.fileName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 2000);
  });

  function readSheets(buf) {
    var wb = XLSX.read(buf, { type: 'array', cellDates: true });
    return wb.SheetNames.map(function (name) {
      var ws = wb.Sheets[name];
      var rows = ws ? XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true }) : [];
      return { name: name, rows: rows };
    });
  }

  function handleFile(file) {
    if (!/\.(xlsx|xlsm|xls|csv)$/i.test(file.name)) {
      setStatus('Please choose an Excel file (.xlsx or .xls).', true);
      return;
    }
    setStatus('Processing "' + file.name + '"…');
    var reader = new FileReader();
    reader.onerror = function () { setStatus('Could not read the file.', true); };
    reader.onload = function () {
      // Let the status message paint before the heavy work.
      setTimeout(function () {
        try {
          var sheets = readSheets(new Uint8Array(reader.result));
          var model = ReportProcessor.processWorkbook(sheets);
          var wb = ReportWriter.build(ExcelJS, model);
          wb.xlsx.writeBuffer().then(function (out) {
            var base = file.name.replace(/\.[^.]+$/, '');
            current = {
              model: model,
              blob: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
              fileName: base + ' - Corrected.xlsx'
            };
            showResult(file.name);
          }).catch(fail);
        } catch (e) { fail(e); }
      }, 30);
    };
    reader.readAsArrayBuffer(file);
  }

  function fail(e) {
    console.error(e);
    setStatus('Something went wrong while processing this file: ' + (e && e.message ? e.message : e), true);
  }

  function showResult(originalName) {
    var m = current.model;
    uploadCard.hidden = true; result.hidden = false; setStatus('');
    $('result-sub').textContent = originalName + ' → ' + current.fileName + ' · ' + m.sheets.length + ' sheet(s)';

    var log = $('log'); log.innerHTML = '';
    m.log.forEach(function (l) {
      var li = document.createElement('li');
      li.className = l.level;
      var b = document.createElement('b'); b.textContent = l.sheet + ': ';
      li.appendChild(b); li.appendChild(document.createTextNode(l.message));
      log.appendChild(li);
    });

    var tabs = $('tabs'); tabs.innerHTML = '';
    m.sheets.forEach(function (s, i) {
      var t = document.createElement('button');
      t.className = 'tab'; t.type = 'button'; t.setAttribute('role', 'tab');
      t.textContent = s.name;
      t.addEventListener('click', function () { selectTab(i); });
      tabs.appendChild(t);
    });
    selectTab(0);
  }

  function fmt(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    if (typeof v === 'number') return v.toLocaleString('en-US');
    return String(v);
  }

  function selectTab(i) {
    var s = current.model.sheets[i];
    Array.prototype.forEach.call($('tabs').children, function (t, j) { t.setAttribute('aria-selected', String(i === j)); });
    var table = $('preview'); table.innerHTML = '';
    var hasGroups = s.rows.some(function (r) { return r.type === 'summary'; });
    var shown = Math.min(s.rows.length, PREVIEW_LIMIT);
    $('preview-hint').textContent =
      (hasGroups ? 'Click a highlighted item name to show its individual records. ' : '') +
      (s.rows.length > PREVIEW_LIMIT ? 'Showing the first ' + PREVIEW_LIMIT + ' of ' + s.rows.length + ' rows. The download has every row.' : s.rows.length + ' row(s).');

    if (!s.headers.length) { table.innerHTML = '<tbody><tr><td>Empty sheet</td></tr></tbody>'; return; }

    var thead = document.createElement('thead'), htr = document.createElement('tr');
    s.headers.forEach(function (h) { var th = document.createElement('th'); th.textContent = h; htr.appendChild(th); });
    thead.appendChild(htr); table.appendChild(thead);

    var tbody = document.createElement('tbody');
    s.rows.slice(0, shown).forEach(function (row) {
      var tr = document.createElement('tr');
      tr.className = row.type;
      if (row.groupId) tr.dataset.group = row.groupId;
      if (row.type === 'detail') tr.hidden = true;
      row.cells.forEach(function (v, ci) {
        if (ci >= s.headers.length) return;
        var td = document.createElement('td');
        td.textContent = fmt(v);
        if (typeof v === 'number') td.className = 'num';
        else if (v === '-') td.className = 'dash';
        if (ci === s.itemCol) td.classList.add('item');
        if (row.type === 'summary' && ci === s.itemCol) {
          var badge = document.createElement('span'); badge.className = 'badge';
          badge.textContent = row.count; badge.title = row.count + ' records';
          td.appendChild(badge);
          td.title = 'Show ' + row.count + ' records';
          td.addEventListener('click', function () { toggleGroup(tbody, tr, row.groupId); });
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
  }

  function toggleGroup(tbody, summaryTr, id) {
    var open = !summaryTr.classList.contains('open');
    summaryTr.classList.toggle('open', open);
    tbody.querySelectorAll('tr.detail[data-group="' + id + '"]').forEach(function (tr) { tr.hidden = !open; });
  }
})();
