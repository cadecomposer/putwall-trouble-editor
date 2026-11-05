/*
 * Putwall Trouble Spreadsheet Editor - browser glue.
 * Builds the UI, reads the OC Clearing file with SheetJS, processes it with
 * PWProcessor, and writes the edited workbook (with barcode images) via ExcelJS.
 */
(function () {
  'use strict';

  var LOCATIONS = [];
  for (var n = 1; n <= 20; n++) LOCATIONS.push('TRBL-PW-' + (n < 10 ? '0' + n : '' + n));

  var state = {
    grid: null,            // parsed 2D array from the uploaded file
    sourceFileName: '',
    selectedLocations: [],
    showAllColumns: false,
    printPages: 1,         // how many pages the trouble list is split into
    printBoundaries: [],   // cumulative row counts where each page ends
    result: null
  };

  var MAX_PRINT_PAGES = 20;

  function $(id) { return document.getElementById(id); }

  // ---------------------------------------------------------------- UI setup

  function buildLocations() {
    var grid = $('locGrid');
    LOCATIONS.forEach(function (loc) {
      var label = document.createElement('label');
      label.className = 'loc-item';
      label.setAttribute('data-loc', loc);

      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = loc;
      cb.className = 'loc-cb';
      cb.addEventListener('change', function () {
        label.classList.toggle('checked', cb.checked);
        updateLocCount();
        reprocess();
      });

      var span = document.createElement('span');
      span.textContent = loc;

      label.appendChild(cb);
      label.appendChild(span);
      grid.appendChild(label);
    });
  }

  function getSelectedLocations() {
    return Array.prototype.slice
      .call(document.querySelectorAll('.loc-cb:checked'))
      .map(function (cb) { return cb.value; });
  }

  function setAllLocations(checked) {
    Array.prototype.forEach.call(document.querySelectorAll('.loc-cb'), function (cb) {
      cb.checked = checked;
      cb.parentNode.classList.toggle('checked', checked);
    });
    updateLocCount();
    reprocess();
  }

  function updateLocCount() {
    var count = getSelectedLocations().length;
    $('locCountLabel').textContent = count + ' selected';
  }

  function setStatus(msg, kind) {
    var s = $('status');
    s.textContent = msg || '';
    s.className = 'status' + (kind ? ' ' + kind : '');
  }

  // ------------------------------------------------------------- File handling

  function onFiles(files) {
    if (!files || !files.length) return;
    handleFile(files[0]);
  }

  function handleFile(file) {
    var selected = getSelectedLocations();
    if (!selected.length) {
      setStatus('Select at least one trouble location in step 1 before adding a file.', 'error');
      return;
    }
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      setStatus('That is not an Excel workbook. Please choose a .xlsx file.', 'error');
      return;
    }

    setStatus('Reading ' + file.name + ' \u2026', '');

    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        if (!wb.SheetNames.length) throw new Error('No worksheet was found in this file.');
        var ws = wb.Sheets[wb.SheetNames[0]];
        var grid = XLSX.utils.sheet_to_json(ws, {
          header: 1, raw: true, defval: '', blankrows: true
        });
        state.grid = grid;
        state.sourceFileName = file.name;
        state.printPages = 1;
        state.printBoundaries = [];
        reprocess();
      } catch (err) {
        console.error(err);
        setStatus(err && err.message ? err.message : 'Could not read that file.', 'error');
      }
    };
    reader.onerror = function () { setStatus('Could not read that file.', 'error'); };
    reader.readAsArrayBuffer(file);
  }

  function reprocess() {
    if (!state.grid) return;
    var selected = getSelectedLocations();
    if (!selected.length) {
      setStatus('Select at least one trouble location in step 1.', 'error');
      return;
    }
    try {
      var result = PWProcessor.processData(state.grid, selected);
      if (!result.rows.length) {
        setStatus('No rows matched the selected locations. Double-check the locations you picked.', 'error');
        $('resultsCard').classList.add('hidden');
        state.result = null;
        return;
      }
      state.result = result;
      state.selectedLocations = selected;
      setStatus('Done \u2014 ' + result.stats.keptRows + ' row(s) kept from ' +
        result.stats.inputRows + ' putwall-related rows.', 'ok');
      renderResults(result);
    } catch (err) {
      console.error(err);
      setStatus(err && err.message ? err.message : 'Could not process the spreadsheet.', 'error');
      $('resultsCard').classList.add('hidden');
      state.result = null;
    }
  }

  // ---------------------------------------------------------------- Rendering

  function statCard(value, label) {
    var d = document.createElement('div');
    d.className = 'stat';
    var v = document.createElement('div');
    v.className = 'value';
    v.textContent = value;
    var l = document.createElement('div');
    l.className = 'label';
    l.textContent = label;
    d.appendChild(v);
    d.appendChild(l);
    return d;
  }

  var TABLE_LIMIT = 500;

  function visibleColumnIndices(result) {
    if (state.showAllColumns) {
      return result.header.map(function (_, i) { return i; });
    }
    var start = result.columns.oLPN;
    var end = result.columns.reserveLocn2;
    if (start === undefined || end === undefined || end < start) {
      return result.header.map(function (_, i) { return i; });
    }
    var idx = [];
    for (var i = start; i <= end; i++) idx.push(i);
    return idx;
  }

  function renderTable(tableEl, indices, rowIdxs, result, breaksAfter) {
    var labelCol = result.columns.label;
    var needCol = result.columns.need;
    tableEl.innerHTML = '';

    var thead = document.createElement('thead');
    var htr = document.createElement('tr');
    indices.forEach(function (ci) {
      var th = document.createElement('th');
      th.textContent = result.header[ci];
      if (ci === labelCol) th.className = 'label-col';
      htr.appendChild(th);
    });
    thead.appendChild(htr);
    tableEl.appendChild(thead);

    var tbody = document.createElement('tbody');
    var limit = Math.min(rowIdxs.length, TABLE_LIMIT);
    for (var n = 0; n < limit; n++) {
      var r = rowIdxs[n];
      var row = result.rows[r];
      var tr = document.createElement('tr');
      if (result.duplicates[r]) tr.className = 'dup';
      indices.forEach(function (ci) {
        var td = document.createElement('td');
        var val = row[ci];
        td.textContent = val == null ? '' : val;
        var classes = [];
        if (ci === labelCol) classes.push('label-col');
        if (ci === needCol && result.multiNeed[r]) classes.push('needs-multi');
        if (classes.length) td.className = classes.join(' ');
        tr.appendChild(td);
      });
      tbody.appendChild(tr);

      if (breaksAfter && breaksAfter[n] && n < limit - 1) {
        var brk = document.createElement('tr');
        brk.className = 'page-break-row';
        var btd = document.createElement('td');
        btd.colSpan = indices.length;
        btd.textContent = 'Page ' + breaksAfter[n] + ' starts here';
        brk.appendChild(btd);
        tbody.appendChild(brk);
      }
    }
    tableEl.appendChild(tbody);
    return rowIdxs.length > limit;
  }

  function buildSection(def, result, cols) {
    var wrap = document.createElement('div');
    var h = document.createElement('h3');
    h.className = 'subhead' + (def.accent ? ' subhead-warn' : '');
    h.appendChild(document.createTextNode(def.title + ' '));
    var count = document.createElement('span');
    count.className = 'muted';
    count.textContent = '(' + def.rows.length + ')';
    h.appendChild(count);
    wrap.appendChild(h);

    if (def.note) {
      var p = document.createElement('p');
      p.className = 'hint';
      p.style.margin = '-4px 0 10px';
      p.textContent = def.note;
      wrap.appendChild(p);
    }

    if (!def.rows.length) {
      var none = document.createElement('div');
      none.className = 'preview-note';
      none.textContent = 'None.';
      wrap.appendChild(none);
      return wrap;
    }

    var pw = document.createElement('div');
    pw.className = 'preview-wrap';
    var table = document.createElement('table');
    table.className = 'preview';
    pw.appendChild(table);
    wrap.appendChild(pw);

    var truncated = renderTable(table, cols, def.rows, result);
    var note = document.createElement('div');
    note.className = 'preview-note';
    note.textContent = truncated
      ? 'Showing the first ' + TABLE_LIMIT + ' of ' + def.rows.length + '.'
      : '';
    wrap.appendChild(note);
    return wrap;
  }

  // ---- Print pagination helpers -------------------------------------------

  function effectivePages(total, pages) {
    if (total <= 0) return 1;
    return Math.max(1, Math.min(pages, total, MAX_PRINT_PAGES));
  }

  function evenBoundaries(total, pages) {
    if (pages <= 1 || total <= 1) return [];
    var b = [];
    for (var i = 1; i < pages; i++) {
      b.push(Math.round(i * total / pages));
    }
    return b;
  }

  // Returns [[start, end), ...] for each page, clamped to valid boundaries.
  function chunkRanges(total, pages, boundaries) {
    pages = effectivePages(total, pages);
    if (pages <= 1 || total <= 0) return [[0, total]];
    var ranges = [];
    var prev = 0;
    for (var i = 0; i < pages - 1; i++) {
      var remaining = pages - 1 - i;           // pages still to come after this one
      var minEnd = prev + 1;
      var maxEnd = total - remaining;          // leave >= 1 row for each later page
      var raw = boundaries && boundaries[i] != null
        ? boundaries[i]
        : Math.round((i + 1) * total / pages);
      var end = Math.max(minEnd, Math.min(maxEnd, raw));
      ranges.push([prev, end]);
      prev = end;
    }
    ranges.push([prev, total]);
    return ranges;
  }

  function currentRanges(result) {
    var total = result.groups.normal.length;
    return chunkRanges(total, state.printPages, state.printBoundaries);
  }

  function breaksAfterMap(ranges) {
    var map = {};
    var acc = 0;
    for (var i = 0; i < ranges.length - 1; i++) {
      acc += (ranges[i][1] - ranges[i][0]);
      map[acc - 1] = i + 2;                    // next page number
    }
    return map;
  }

  function dividerNote(text) {
    var s = document.createElement('span');
    s.className = 'divider-note';
    s.textContent = text;
    return s;
  }

  function renderMainTableOnly(result) {
    var cols = visibleColumnIndices(result);
    var normalRows = result.groups.normal;
    var ranges = currentRanges(result);
    var breaks = ranges.length > 1 ? breaksAfterMap(ranges) : null;
    var truncated = renderTable($('previewTable'), cols, normalRows, result, breaks);
    $('troublesCount').textContent = '(' + normalRows.length + ')';
    $('previewNote').textContent = truncated
      ? 'Showing the first ' + TABLE_LIMIT + ' of ' + normalRows.length +
        ' rows. The download and print include all of them.'
      : '';
  }

  var mainTablePending = false;
  function scheduleMainTable() {
    if (mainTablePending) return;
    mainTablePending = true;
    requestAnimationFrame(function () {
      mainTablePending = false;
      if (state.result) renderMainTableOnly(state.result);
    });
  }

  function updateDividerControls(ranges, total, pages) {
    var rows = $('dividerControls').querySelectorAll('.divider-row');
    for (var i = 0; i < rows.length && i < ranges.length; i++) {
      var range = ranges[i];
      var slider = rows[i].querySelector('input[type="range"]');
      if (slider) {
        slider.min = String(range[0] + 1);
        slider.max = String(total - (pages - 1 - i));
        slider.value = String(range[1]);
      }
      var info = rows[i].querySelector('.divider-range');
      if (info) info.textContent = 'rows ' + (range[0] + 1) + '\u2013' + range[1];
    }
  }

  function onBoundaryInput(e) {
    if (!state.result) return;
    var idx = Number(e.target.dataset.boundary);
    state.printBoundaries = state.printBoundaries.slice();
    state.printBoundaries[idx] = Number(e.target.value);

    var total = state.result.groups.normal.length;
    var pages = effectivePages(total, state.printPages);
    var ranges = chunkRanges(total, pages, state.printBoundaries);
    state.printBoundaries = ranges.slice(0, -1).map(function (r) { return r[1]; });
    updateDividerControls(ranges, total, pages);
    scheduleMainTable();
  }

  function renderPrintLayout(result) {
    var total = result.groups.normal.length;
    var pages = effectivePages(total, state.printPages);
    state.printPages = pages;
    if (state.printBoundaries.length !== Math.max(0, pages - 1)) {
      state.printBoundaries = evenBoundaries(total, pages);
    }
    var ranges = chunkRanges(total, pages, state.printBoundaries);
    state.printBoundaries = ranges.slice(0, -1).map(function (r) { return r[1]; });

    $('pageCountLabel').textContent = String(pages);
    $('pageMinus').disabled = pages <= 1;
    $('pagePlus').disabled = pages >= effectivePages(total, MAX_PRINT_PAGES);

    var host = $('dividerControls');
    host.innerHTML = '';
    if (total === 0) {
      host.appendChild(dividerNote('No printable troubles.'));
      return;
    }
    if (pages === 1) {
      host.appendChild(dividerNote(
        'Printing as a single list. A long list spills onto more pages at the same font size.'));
      return;
    }

    ranges.forEach(function (range, idx) {
      var row = document.createElement('div');
      row.className = 'divider-row';

      var label = document.createElement('span');
      label.className = 'divider-page';
      label.textContent = 'Page ' + (idx + 1);
      row.appendChild(label);

      var mid = document.createElement('span');
      if (idx < pages - 1) {
        var slider = document.createElement('input');
        slider.type = 'range';
        slider.min = String(range[0] + 1);
        slider.max = String(total - (pages - 1 - idx));
        slider.value = String(range[1]);
        slider.dataset.boundary = String(idx);
        slider.setAttribute('aria-label', 'Last row on page ' + (idx + 1));
        slider.addEventListener('input', onBoundaryInput);
        mid.appendChild(slider);
      }
      row.appendChild(mid);

      var info = document.createElement('span');
      info.className = 'divider-range';
      info.textContent = 'rows ' + (range[0] + 1) + '\u2013' + range[1];
      row.appendChild(info);
      host.appendChild(row);
    });
    host.appendChild(dividerNote('Each divider is where the next page begins.'));
  }

  function setPrintPages(pages) {
    if (!state.result) return;
    var total = state.result.groups.normal.length;
    var p = effectivePages(total, pages);
    state.printPages = p;
    state.printBoundaries = evenBoundaries(total, p);
    renderTables(state.result);
  }

  function evenSplit() {
    if (!state.result) return;
    var total = state.result.groups.normal.length;
    state.printBoundaries = evenBoundaries(total, state.printPages);
    renderTables(state.result);
  }

  function renderTables(result) {
    var cols = visibleColumnIndices(result);

    renderMainTableOnly(result);

    var host = $('otherSections');
    host.innerHTML = '';
    host.appendChild(buildSection({
      title: 'Alt form \u2014 EO pickup',
      note: 'No active location, but a reserve location exists. Send an alt form so an EO can grab it. Not printed.',
      rows: result.groups.alt
    }, result, cols));
    host.appendChild(buildSection({
      title: 'Split cancel',
      note: 'No active or reserve location, and no other item on the order is findable. Safe to split cancel. Not printed.',
      rows: result.groups.split
    }, result, cols));
    host.appendChild(buildSection({
      title: 'Do NOT split cancel yet',
      accent: true,
      note: 'No active or reserve location, but another item on the same order (same oLPN) is findable. Hold these until that item is found. Not printed.',
      rows: result.groups.hold
    }, result, cols));
    host.appendChild(buildSection({
      title: 'Duplicates',
      note: 'Exact repeat rows (later copies of an earlier row). Not printed.',
      rows: result.groups.duplicate
    }, result, cols));

    $('columnsBtn').textContent = state.showAllColumns ? 'Show fewer columns' : 'Show all columns';
    renderPrintLayout(result);
  }

  function renderResults(result) {
    var s = result.stats;
    $('resultsSummary').textContent =
      'Kept ' + s.keptRows + ' of ' + s.inputRows + ' rows \u2014 ' +
      s.normalRows + ' trouble(s) to print, ' + s.altRows + ' alt form, ' +
      s.splitRows + ' split cancel, ' + s.holdRows + ' hold, ' +
      s.duplicateRows + ' duplicate.';

    var box = $('statsBox');
    box.innerHTML = '';
    box.appendChild(statCard(String(s.normalRows), 'Troubles (print)'));
    box.appendChild(statCard(String(s.altRows), 'Alt form (EO)'));
    box.appendChild(statCard(String(s.splitRows), 'Split cancel'));
    box.appendChild(statCard(String(s.holdRows), 'Hold (no split)'));
    box.appendChild(statCard(String(s.duplicateRows), 'Duplicates'));
    box.appendChild(statCard(String(s.multiNeedRows), 'Need > 1'));
    box.appendChild(statCard(String(s.distinctTotes), 'Barcodes'));

    renderTables(result);
    renderBarcodePreview(result);
    $('resultsCard').classList.remove('hidden');
  }

  // ------------------------------------------------------- Barcode preview

  var SVG_NS = 'http://www.w3.org/2000/svg';

  function makeBarcodeSvg(value) {
    var svg = document.createElementNS(SVG_NS, 'svg');
    JsBarcode(svg, value, {
      format: 'CODE128',
      width: 2,
      height: 60,
      margin: 0,
      displayValue: false,
      background: '#ffffff',
      lineColor: '#000000'
    });
    return svg;
  }

  function makeBarcodeSvgMarkup(value) {
    return new XMLSerializer().serializeToString(makeBarcodeSvg(value));
  }

  function renderBarcodePreview(result) {
    var grid = $('barcodeGrid');
    grid.innerHTML = '';
    var cap = 300;
    if (!result.barcodes.length) {
      $('bcCount').textContent = '(0)';
      $('bcNote').textContent = 'No tote numbers were found in the selected rows, so there are no barcodes.';
      return;
    }
    var shown = Math.min(result.barcodes.length, cap);
    for (var i = 0; i < shown; i++) {
      var b = result.barcodes[i];
      var card = document.createElement('div');
      card.className = 'barcode-card';

      var num = document.createElement('div');
      num.className = 'bc-num';
      num.textContent = b.label;

      var tote = document.createElement('div');
      tote.className = 'bc-tote';
      tote.textContent = b.tote;

      var svg = makeBarcodeSvg(b.tote);
      svg.setAttribute('class', 'bc-svg');

      card.appendChild(num);
      card.appendChild(svg);
      card.appendChild(tote);
      grid.appendChild(card);
    }
    $('bcCount').textContent = '(' + result.barcodes.length + ')';
    $('bcNote').textContent = result.barcodes.length > cap
      ? 'Showing the first ' + cap + ' of ' + result.barcodes.length +
        '. The printed document and the downloaded file contain all of them.'
      : '';
  }

  // ------------------------------------------------------------- Printing

  // Columns to print, per the Putwall team's convention:
  // from Tote Gid through the first Reserve Locn (not Reserve Locn 2).
  var PRINT_COLUMNS = [
    'Tote Gid', 'Label', 'Item', 'Need',
    'Active Locn', 'Active Locn 2', 'Active OH', 'Reserve Locn'
  ];

  var PRINT_CSS = [
    '@page { size: landscape; margin: 8mm; }',
    '* { box-sizing: border-box; }',
    'body { margin: 0; color: #000; font-family: Arial, Helvetica, sans-serif; }',
    'h1 { font-size: 12pt; margin: 0 0 6px; }',
    'table.sheet { border-collapse: collapse; width: 100%; font-size: 10pt; }',
    'table.sheet th, table.sheet td { border: 1px solid #7f7f7f; padding: 2px 4px; text-align: left; white-space: nowrap; }',
    'table.sheet th { background: #e6e6e6; }',
    'table.sheet thead { display: table-header-group; }',
    'table.sheet tr { page-break-inside: avoid; }',
    'table.sheet td.lbl { text-align: center; font-weight: bold; }',
    'table.sheet tr.dup td { background: #ffc7ce; }',
    'table.sheet tr td.need-multi { background: #ffeb3b; color: #000; font-weight: bold; }',
    '.sheet-section.page-break { page-break-before: always; }',
    '.sheet-sub { font-size: 9pt; color: #444444; margin: 0 0 4px; }',
    '.labels-section { page-break-before: always; }',
    '.label { display: inline-block; width: 60mm; padding: 3mm; margin: 0; text-align: center; vertical-align: top; border: 1px dashed #b0b0b0; page-break-inside: avoid; }',
    '.label .num { font-size: 20pt; font-weight: bold; line-height: 1; }',
    '.label svg { display: block; margin: 2mm auto 1mm; max-width: 100%; height: 16mm; }',
    '.label .tote { font-family: Consolas, "Liberation Mono", monospace; font-size: 9pt; }'
  ].join('\n');

  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (ch) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch];
    });
  }

  function buildPrintHtml(result) {
    var headerIdx = {};
    result.header.forEach(function (h, i) { headerIdx[h] = i; });
    var cols = PRINT_COLUMNS
      .map(function (name) { return { name: name, i: headerIdx[name] }; })
      .filter(function (c) { return c.i !== undefined; });

    var thead = '<thead><tr>' + cols.map(function (c) {
      return '<th>' + escapeHtml(c.name) + '</th>';
    }).join('') + '</tr></thead>';

    var head = 'Putwall Troubles \u2014 ' + escapeHtml((state.selectedLocations || []).join(', '));

    var total = result.groups.normal.length;
    var pages = effectivePages(total, state.printPages);
    var ranges = chunkRanges(total, pages, state.printBoundaries);

    var sections = ranges.map(function (range, idx) {
      var rows = result.groups.normal.slice(range[0], range[1]);
      var tbody = '<tbody>' + rows.map(function (ri) {
        var row = result.rows[ri];
        var tds = cols.map(function (c) {
          var cls = '';
          if (c.name === 'Label') cls = ' class="lbl"';
          else if (c.name === 'Need' && result.multiNeed[ri]) cls = ' class="need-multi"';
          return '<td' + cls + '>' + escapeHtml(row[c.i]) + '</td>';
        }).join('');
        return '<tr>' + tds + '</tr>';
      }).join('') + '</tbody>';

      var sub = pages > 1
        ? '<div class="sheet-sub">Page ' + (idx + 1) + ' of ' + pages +
          ' \u00b7 rows ' + (range[0] + 1) + '\u2013' + range[1] + '</div>'
        : '';
      return '<div class="sheet-section' + (idx ? ' page-break' : '') + '">' +
        '<h1>' + head + '</h1>' + sub +
        '<table class="sheet">' + thead + tbody + '</table></div>';
    }).join('');

    var labels = result.barcodes.map(function (b) {
      var svg = makeBarcodeSvgMarkup(b.tote);
      return '<div class="label"><div class="num">' + escapeHtml(b.label) +
        '</div>' + svg + '<div class="tote">' + escapeHtml(b.tote) + '</div></div>';
    }).join('');

    return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + head +
      '</title><style>' + PRINT_CSS + '</style></head><body>' +
      sections +
      '<div class="labels-section">' + labels + '</div>' +
      '</body></html>';
  }

  function openPrintFallback(html) {
    var w = window.open('', '_blank');
    if (!w) {
      setStatus('Could not open the print view. Allow pop-ups for this page and try again.', 'error');
      return;
    }
    w.document.open();
    w.document.write(html);
    w.document.close();
    w.focus();
    setTimeout(function () { try { w.print(); } catch (e) { /* user can print manually */ } }, 400);
  }

  function printDocument() {
    if (!state.result) return;
    var html;
    try {
      html = buildPrintHtml(state.result);
    } catch (err) {
      console.error(err);
      setStatus('Could not build the print view: ' + (err && err.message ? err.message : err), 'error');
      return;
    }

    // Print via a hidden iframe so no pop-up blocker is involved.
    var frame = document.createElement('iframe');
    frame.setAttribute('aria-hidden', 'true');
    frame.style.position = 'fixed';
    frame.style.right = '0';
    frame.style.bottom = '0';
    frame.style.width = '0';
    frame.style.height = '0';
    frame.style.border = '0';
    document.body.appendChild(frame);

    var idoc = frame.contentWindow.document;
    idoc.open();
    idoc.write(html);
    idoc.close();

    var done = false;
    function cleanup() {
      setTimeout(function () { if (frame.parentNode) frame.parentNode.removeChild(frame); }, 1500);
    }
    function go() {
      if (done) return;
      done = true;
      try {
        frame.contentWindow.focus();
        frame.contentWindow.print();
        frame.contentWindow.onafterprint = cleanup;
        cleanup();
      } catch (err) {
        console.error(err);
        if (frame.parentNode) frame.parentNode.removeChild(frame);
        openPrintFallback(html);
      }
    }
    setTimeout(go, 300);
  }

  // --------------------------------------------------------- Workbook writing

  function makeBarcode(value) {
    var canvas = document.createElement('canvas');
    JsBarcode(canvas, value, {
      format: 'CODE128',
      width: 2,
      height: 56,
      margin: 4,
      displayValue: false,
      background: '#ffffff',
      lineColor: '#000000'
    });
    return {
      base64: canvas.toDataURL('image/png').split(',')[1],
      width: canvas.width,
      height: canvas.height
    };
  }

  var COLUMN_WIDTHS = {
    'Time in OC': 16, 'Order Date': 12, 'Wave': 14, 'oLPN': 22, 'OC Mod': 9,
    'Location': 14, 'LPN Type': 10, 'Tote Gid': 14, 'Label': 7, 'Item': 12,
    'Need': 6, 'Active Locn': 14, 'Active Locn 2': 14, 'Active OH': 9,
    'Reserve Locn': 14, 'Reserve Locn 2': 14, 'iLPN': 22, 'iLPN2': 22,
    'Staging Locn': 14
  };

  function styleHeaderRow(headerRow) {
    headerRow.height = 20;
    headerRow.eachCell(function (cell) {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002E6D' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
  }

  function addDataSheet(wb, name, result, rowIdxs, markDuplicates) {
    var ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    styleHeaderRow(ws.addRow(result.header));

    rowIdxs.forEach(function (ri) {
      var row = ws.addRow(result.rows[ri]);
      if (markDuplicates) {
        row.eachCell({ includeEmpty: true }, function (cell) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
          cell.font = { color: { argb: 'FF9C0006' } };
        });
      }
      var labelCell = row.getCell(result.columns.label + 1);
      labelCell.alignment = { horizontal: 'center' };
      if (!markDuplicates) labelCell.font = { bold: true };

      if (result.multiNeed[ri]) {
        var needCell = row.getCell(result.columns.need + 1);
        needCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB3B' } };
        needCell.font = { bold: true, color: { argb: 'FF1C2530' } };
      }
    });

    result.header.forEach(function (h, i) {
      ws.getColumn(i + 1).width = COLUMN_WIDTHS[h] || 14;
    });
    return ws;
  }

  function buildWorkbook(result) {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Putwall Trouble Spreadsheet Editor';
    wb.created = new Date();
    wb.modified = new Date();

    // ---- one sheet per section ----
    addDataSheet(wb, 'Troubles', result, result.groups.normal, false);
    addDataSheet(wb, 'Alt Form', result, result.groups.alt, false);
    addDataSheet(wb, 'Split Cancel', result, result.groups.split, false);
    addDataSheet(wb, 'Split Cancel HOLD', result, result.groups.hold, false);
    addDataSheet(wb, 'Duplicates', result, result.groups.duplicate, true);

    // ---- barcode sheet (printable troubles only) ----
    var bws = wb.addWorksheet('Barcodes', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });
    bws.getColumn(1).width = 8;   // label number
    bws.getColumn(2).width = 16;  // tote gid text
    bws.getColumn(3).width = 2;   // spacer
    bws.getColumn(4).width = 30;  // barcode image sits here and to the right

    var bhead = bws.addRow(['Label', 'Tote Gid', '', 'Barcode']);
    bhead.height = 20;
    bhead.eachCell(function (cell) {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002E6D' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    result.barcodes.forEach(function (b, idx) {
      var r = idx + 2;
      bws.getRow(r).height = 46;

      var lc = bws.getCell(r, 1);
      lc.value = b.label;
      lc.font = { bold: true, size: 18 };
      lc.alignment = { horizontal: 'center', vertical: 'middle' };

      var tc = bws.getCell(r, 2);
      tc.value = b.tote;
      tc.font = { name: 'Consolas', size: 11 };
      tc.alignment = { horizontal: 'center', vertical: 'middle' };

      var img = makeBarcode(b.tote);
      var imgId = wb.addImage({ base64: img.base64, extension: 'png' });
      var scale = 54 / img.height;
      bws.addImage(imgId, {
        tl: { col: 3.1, row: r - 1 + 0.08 },
        ext: { width: Math.round(img.width * scale), height: Math.round(img.height * scale) }
      });
    });

    return wb.xlsx.writeBuffer();
  }

  function outputFileName() {
    var base = (state.sourceFileName || 'OC-Clearing').replace(/\.[^.]+$/, '');
    return base + ' - Putwall Edited.xlsx';
  }

  function saveBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
  }

  function setBusy(busy) {
    $('downloadBtn').disabled = busy;
  }

  function downloadEdited() {
    if (!state.result) return;
    setBusy(true);
    setStatus('Building the workbook \u2026', '');
    buildWorkbook(state.result).then(function (buffer) {
      var blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      saveBlob(blob, outputFileName());
      setStatus('Downloaded: ' + outputFileName(), 'ok');
    }).catch(function (err) {
      console.error(err);
      setStatus('Could not build the workbook: ' + (err && err.message ? err.message : err), 'error');
    }).then(function () { setBusy(false); });
  }

  function reset() {
    state.grid = null;
    state.result = null;
    state.sourceFileName = '';
    state.printPages = 1;
    state.printBoundaries = [];
    $('resultsCard').classList.add('hidden');
    $('fileInput').value = '';
    setStatus('');
  }

  // ------------------------------------------------------------------- Theme

  var THEME_KEY = 'pwTheme';
  var THEME_MODES = ['auto', 'light', 'dark'];
  var THEME_LABELS = { auto: 'Auto', light: 'Light', dark: 'Dark' };
  var THEME_ICONS = { auto: '\u25D0', light: '\u2600\uFE0F', dark: '\u263E' };

  function getThemeMode() {
    try { return localStorage.getItem(THEME_KEY) || 'auto'; } catch (e) { return 'auto'; }
  }

  function themePrefersDark(mode) {
    return mode === 'dark' ||
      (mode === 'auto' && window.matchMedia &&
        window.matchMedia('(prefers-color-scheme: dark)').matches);
  }

  function applyTheme(mode) {
    var root = document.documentElement;
    root.setAttribute('data-theme', themePrefersDark(mode) ? 'dark' : 'light');
    root.setAttribute('data-theme-mode', mode);

    var name = THEME_LABELS[mode] || THEME_LABELS.auto;
    var icon = $('themeIcon');
    var label = $('themeLabel');
    var btn = $('themeToggle');
    if (icon) icon.textContent = THEME_ICONS[mode] || THEME_ICONS.auto;
    if (label) label.textContent = name;
    if (btn) {
      btn.setAttribute('title', 'Theme: ' + name + ' (click to change)');
      btn.setAttribute('aria-label',
        'Theme: ' + name + '. Click to switch between auto, light and dark.');
    }
  }

  function cycleTheme() {
    var mode = getThemeMode();
    var next = THEME_MODES[(THEME_MODES.indexOf(mode) + 1) % THEME_MODES.length];
    try { localStorage.setItem(THEME_KEY, next); } catch (e) { /* ignore */ }
    applyTheme(next);
  }

  // ------------------------------------------------------------------ Wiring

  function init() {
    buildLocations();
    updateLocCount();

    applyTheme(getThemeMode());
    $('themeToggle').addEventListener('click', cycleTheme);
    if (window.matchMedia) {
      var schemeQuery = window.matchMedia('(prefers-color-scheme: dark)');
      var onSchemeChange = function () {
        if (getThemeMode() === 'auto') applyTheme('auto');
      };
      if (schemeQuery.addEventListener) schemeQuery.addEventListener('change', onSchemeChange);
      else if (schemeQuery.addListener) schemeQuery.addListener(onSchemeChange);
    }

    $('selectAllBtn').addEventListener('click', function () { setAllLocations(true); });
    $('clearAllBtn').addEventListener('click', function () { setAllLocations(false); });
    $('browseBtn').addEventListener('click', function (e) {
      e.stopPropagation();
      $('fileInput').click();
    });
    $('fileInput').addEventListener('change', function (e) { onFiles(e.target.files); });

    var dz = $('dropzone');
    dz.addEventListener('click', function () { $('fileInput').click(); });
    dz.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('fileInput').click(); }
    });

    function highlightDropzone(on) { dz.classList.toggle('dragover', on); }

    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) {
        e.preventDefault();
        highlightDropzone(true);
      });
    });
    dz.addEventListener('dragleave', function (e) {
      // Only clear the highlight when the pointer truly leaves the dropzone.
      if (!dz.contains(e.relatedTarget)) highlightDropzone(false);
    });

    // Accept a drop anywhere in the window. Handling it here (rather than on
    // the dropzone) means dropping directly on the box is not swallowed.
    window.addEventListener('dragover', function (e) {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
    });
    window.addEventListener('drop', function (e) {
      e.preventDefault();
      highlightDropzone(false);
      var dt = e.dataTransfer;
      if (dt && dt.files && dt.files.length) onFiles(dt.files);
    });
    window.addEventListener('dragend', function () { highlightDropzone(false); });

    $('downloadBtn').addEventListener('click', downloadEdited);
    $('printBtn').addEventListener('click', printDocument);
    $('columnsBtn').addEventListener('click', function () {
      state.showAllColumns = !state.showAllColumns;
      if (state.result) renderTables(state.result);
    });
    $('pageMinus').addEventListener('click', function () { setPrintPages(state.printPages - 1); });
    $('pagePlus').addEventListener('click', function () { setPrintPages(state.printPages + 1); });
    $('pageEvenBtn').addEventListener('click', evenSplit);
    $('resetBtn').addEventListener('click', reset);
  }

  // Exposed so the print/preview can be inspected and tested.
  window.PWApp = {
    buildPrintHtml: function () { return state.result ? buildPrintHtml(state.result) : ''; },
    state: state,
    getThemeMode: getThemeMode,
    applyTheme: applyTheme,
    cycleTheme: cycleTheme
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
