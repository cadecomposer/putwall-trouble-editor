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
    result: null
  };

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

  function renderResults(result) {
    var s = result.stats;
    $('resultsSummary').textContent =
      'Kept ' + s.keptRows + ' of ' + s.inputRows + ' rows. ' +
      s.duplicateRows + ' duplicate row(s) marked red. ' +
      s.distinctTotes + ' distinct tote(s) got barcodes.';

    var box = $('statsBox');
    box.innerHTML = '';
    box.appendChild(statCard(String(s.keptRows), 'Rows kept'));
    box.appendChild(statCard(String(s.filteredOut), 'Rows filtered out'));
    box.appendChild(statCard(String(s.duplicateRows), 'Duplicate rows'));
    box.appendChild(statCard(String(s.distinctTotes), 'Barcodes / totes'));

    renderPreview(result);
    renderBarcodePreview(result);
    $('resultsCard').classList.remove('hidden');
  }

  function renderPreview(result) {
    var table = $('previewTable');
    table.innerHTML = '';
    var labelCol = result.columns.label;

    var thead = document.createElement('thead');
    var htr = document.createElement('tr');
    result.header.forEach(function (h, i) {
      var th = document.createElement('th');
      th.textContent = h;
      if (i === labelCol) th.className = 'label-col';
      htr.appendChild(th);
    });
    thead.appendChild(htr);
    table.appendChild(thead);

    var tbody = document.createElement('tbody');
    var limit = Math.min(result.rows.length, 400);
    for (var r = 0; r < limit; r++) {
      var row = result.rows[r];
      var tr = document.createElement('tr');
      if (result.duplicates[r]) tr.className = 'dup';
      for (var c = 0; c < result.header.length; c++) {
        var td = document.createElement('td');
        var val = row[c];
        td.textContent = val == null ? '' : val;
        if (c === labelCol) td.className = 'label-col';
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);

    $('previewNote').textContent = result.rows.length > limit
      ? 'Previewing the first ' + limit + ' of ' + result.rows.length + ' rows. The downloaded file has them all.'
      : '';
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
    '@page labels { size: portrait; margin: 8mm; }',
    '* { box-sizing: border-box; }',
    'body { margin: 0; color: #000; font-family: Arial, Helvetica, sans-serif; }',
    'h1 { font-size: 12pt; margin: 0 0 6px; }',
    'table.sheet { border-collapse: collapse; width: 100%; font-size: 10pt; }',
    'table.sheet th, table.sheet td { border: 1px solid #7f7f7f; padding: 2px 4px; text-align: left; white-space: nowrap; }',
    'table.sheet th { background: #e6e6e6; }',
    'table.sheet td.lbl { text-align: center; font-weight: bold; }',
    'table.sheet tr.dup td { background: #ffc7ce; }',
    '.labels-section { page: labels; page-break-before: always; }',
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

    var tbody = '<tbody>' + result.rows.map(function (row, ri) {
      var tds = cols.map(function (c) {
        var cls = (c.name === 'Label') ? ' class="lbl"' : '';
        return '<td' + cls + '>' + escapeHtml(row[c.i]) + '</td>';
      }).join('');
      return '<tr' + (result.duplicates[ri] ? ' class="dup"' : '') + '>' + tds + '</tr>';
    }).join('') + '</tbody>';

    var labels = result.barcodes.map(function (b) {
      var svg = makeBarcodeSvgMarkup(b.tote);
      return '<div class="label"><div class="num">' + escapeHtml(b.label) +
        '</div>' + svg + '<div class="tote">' + escapeHtml(b.tote) + '</div></div>';
    }).join('');

    var head = 'Putwall Troubles \u2014 ' + escapeHtml((state.selectedLocations || []).join(', '));

    return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + head +
      '</title><style>' + PRINT_CSS + '</style></head><body>' +
      '<div class="sheet-section"><h1>' + head + '</h1>' +
      '<table class="sheet">' + thead + tbody + '</table></div>' +
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

  function buildWorkbook(result) {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Putwall Trouble Spreadsheet Editor';
    wb.created = new Date();
    wb.modified = new Date();

    // ---- main sheet ----
    var ws = wb.addWorksheet('Troubles', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    var header = ws.addRow(result.header);
    header.height = 20;
    header.eachCell(function (cell) {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002E6D' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    result.rows.forEach(function (rowData, i) {
      var row = ws.addRow(rowData);
      if (result.duplicates[i]) {
        row.eachCell({ includeEmpty: true }, function (cell) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
          cell.font = { color: { argb: 'FF9C0006' } };
        });
      }
      var labelCell = row.getCell(result.columns.label + 1);
      labelCell.alignment = { horizontal: 'center' };
      if (!result.duplicates[i]) labelCell.font = { bold: true };
    });

    result.header.forEach(function (h, i) {
      ws.getColumn(i + 1).width = COLUMN_WIDTHS[h] || 14;
    });

    // ---- barcode sheet ----
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
    $('resultsCard').classList.add('hidden');
    $('fileInput').value = '';
    setStatus('');
  }

  // ------------------------------------------------------------------ Wiring

  function init() {
    buildLocations();
    updateLocCount();

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
    $('resetBtn').addEventListener('click', reset);
  }

  // Exposed so the print/preview can be inspected and tested.
  window.PWApp = {
    buildPrintHtml: function () { return state.result ? buildPrintHtml(state.result) : ''; },
    state: state
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
