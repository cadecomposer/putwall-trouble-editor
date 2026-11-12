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
    // Alt-form (EO pickup) extras. `entries` is keyed by a stable per-row key
    // so the typed values survive re-renders and re-prints until "Start over".
    altForm: { name: '', entries: {} },
    // Rows the user has hidden, keyed by the stable source-row id from the
    // processor. Kept across filter changes and re-imports; cleared only by
    // "Start over".
    hiddenRows: {},
    // Rows the user manually added to the alt form, keyed by the same stable
    // source-row id. These are additions on top of the rows the processor
    // classified as alt; the source row stays in its original section.
    manualAltRows: {},
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
    state.selectedLocations = selected;

    // No locations ticked yet is a valid state: the file is loaded and waiting
    // for a filter. Show a neutral prompt instead of hiding everything as an
    // error. Ticking any location immediately runs the full pipeline again.
    if (!selected.length) {
      state.result = null;
      $('resultsCard').classList.add('hidden');
      setStatus(
        'Loaded ' + (state.sourceFileName || 'the spreadsheet') +
        '. Select one or more trouble locations in step 1 to see results.', '');
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
    for (var i = start; i <= end; i++) {
      // OC Mod and Reserve Locn 2 are only shown with "Show all columns".
      if (i === result.columns.ocMod || i === result.columns.reserveLocn2) continue;
      idx.push(i);
    }
    return idx;
  }

  function isRowHidden(result, rowIndex) {
    if (!result || !result.rowIds) return false;
    return !!state.hiddenRows[result.rowIds[rowIndex]];
  }

  // Rows moved to the alt form are hidden from the Troubles list/print (but
  // not from the alt form itself) until the move is undone with "-alt".
  function isHiddenInNormalTable(result, rowIndex) {
    return isRowHidden(result, rowIndex) || isManualAltRow(result, rowIndex);
  }

  function toggleRowHidden(result, rowIndex) {
    if (!result || !result.rowIds) return;
    var id = result.rowIds[rowIndex];
    if (state.hiddenRows[id]) delete state.hiddenRows[id];
    else state.hiddenRows[id] = true;

    // Only re-render: do not re-process. Labels, sections and barcodes keep
    // exactly the values they already had.
    renderTables(result);
    renderBarcodePreview(result);
  }

  function isManualAltRow(result, rowIndex) {
    if (!result || !result.rowIds) return false;
    return !!state.manualAltRows[result.rowIds[rowIndex]];
  }

  // Alt-form rows are the ones the processor classified as alt, plus any rows
  // the user chose to add by hand. Manual additions stay in their original
  // section too; the alt form is an extra output for that row.
  function manualAltRowIndices(result) {
    var out = [];
    result.rows.forEach(function (_, i) {
      if (result.categories[i] !== 'alt' && isManualAltRow(result, i)) out.push(i);
    });
    return out;
  }

  function allAltRows(result) {
    return result.groups.alt.concat(manualAltRowIndices(result));
  }

  function toggleManualAltRow(result, rowIndex) {
    if (!result || !result.rowIds) return;
    var id = result.rowIds[rowIndex];
    if (state.manualAltRows[id]) delete state.manualAltRows[id];
    else state.manualAltRows[id] = true;

    renderTables(result);
    renderStats(result);
    renderBarcodePreview(result);
  }

  function eyeIconMarkup(hidden) {
    return '<svg class="row-eye-icon" viewBox="0 0 24 24" width="14" height="14" ' +
      'fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" ' +
      'stroke-linejoin="round" aria-hidden="true">' +
      '<path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z"></path>' +
      '<circle cx="12" cy="12" r="3"></circle>' +
      (hidden ? '<line x1="3" y1="3" x2="21" y2="21"></line>' : '') +
      '</svg>';
  }

  function renderTable(tableEl, indices, rowIdxs, result, breaksAfter, showAltToggle, hideMovedToAlt) {
    var labelCol = result.columns.label;
    var needCol = result.columns.need;
    var activeOHCol = result.columns.activeOH;
    tableEl.innerHTML = '';

    var thead = document.createElement('thead');
    var htr = document.createElement('tr');
    var actionTh = document.createElement('th');
    actionTh.className = 'row-actions-col';
    actionTh.textContent = 'Actions';
    htr.appendChild(actionTh);
    indices.forEach(function (ci) {
      var th = document.createElement('th');
      th.textContent = result.header[ci];
      if (ci === labelCol) th.className = 'label-col';
      htr.appendChild(th);
    });
    thead.appendChild(htr);
    tableEl.appendChild(thead);

    var tbody = document.createElement('tbody');
    for (var n = 0; n < rowIdxs.length; n++) {
      var r = rowIdxs[n];
      var row = result.rows[r];
      var movedToAlt = hideMovedToAlt && isManualAltRow(result, r);
      var hidden = isRowHidden(result, r) || movedToAlt;
      var tr = document.createElement('tr');
      tr.dataset.rowPos = String(n);
      var rowClasses = [];
      if (result.duplicates[r]) rowClasses.push('dup');
      if (hidden) rowClasses.push('row-hidden');
      if (rowClasses.length) tr.className = rowClasses.join(' ');

      var actionTd = document.createElement('td');
      actionTd.className = 'row-actions-col';
      var actions = document.createElement('div');
      actions.className = 'row-actions';

      var hideBtn = document.createElement('button');
      hideBtn.type = 'button';
      hideBtn.className = 'row-icon-btn row-eye-btn';
      hideBtn.innerHTML = eyeIconMarkup(hidden);
      if (movedToAlt) {
        hideBtn.disabled = true;
        hideBtn.title = 'This row is on the alt form. Use -alt to bring it back to Troubles.';
        hideBtn.setAttribute('aria-label', 'Hidden in Troubles because it is on the alt form');
      } else {
        hideBtn.title = hidden ? 'Show this row again' : 'Hide this row (not printed)';
        hideBtn.setAttribute('aria-label', hidden ? 'Show this row' : 'Hide this row');
      }
      hideBtn.setAttribute('aria-pressed', hidden ? 'true' : 'false');
      (function (rowIndex) {
        hideBtn.addEventListener('click', function () { toggleRowHidden(result, rowIndex); });
      })(r);
      actions.appendChild(hideBtn);

      // Adding to the alt form is only offered on the Troubles rows. Rows the
      // processor already put on the alt form need no button.
      if (showAltToggle && result.categories[r] !== 'alt') {
        var added = isManualAltRow(result, r);
        var altBtn = document.createElement('button');
        altBtn.type = 'button';
        altBtn.className = 'row-icon-btn row-alt-btn' + (added ? ' is-added' : '');
        altBtn.textContent = added ? '-alt' : '+alt';
        altBtn.title = added
          ? 'Remove this row from the alt form'
          : 'Also put this row on the alt form';
        altBtn.setAttribute('aria-label', added ? 'Remove from alt form' : 'Add to alt form');
        altBtn.setAttribute('aria-pressed', added ? 'true' : 'false');
        (function (rowIndex) {
          altBtn.addEventListener('click', function () { toggleManualAltRow(result, rowIndex); });
        })(r);
        actions.appendChild(altBtn);
      }

      actionTd.appendChild(actions);
      tr.appendChild(actionTd);

      indices.forEach(function (ci) {
        var td = document.createElement('td');
        var val = row[ci];
        td.textContent = val == null ? '' : val;
        var classes = [];
        if (ci === labelCol) classes.push('label-col');
        if (ci === needCol && result.multiNeed[r]) classes.push('needs-multi');
        if (ci === activeOHCol && result.lowActiveOH[r]) classes.push('active-oh-low');
        if (classes.length) td.className = classes.join(' ');
        tr.appendChild(td);
      });
      tbody.appendChild(tr);

      if (breaksAfter && breaksAfter[n] && n < rowIdxs.length - 1) {
        var brk = document.createElement('tr');
        brk.className = 'page-break-row';
        brk.dataset.boundary = String(breaksAfter[n] - 2);
        brk.dataset.nextPage = String(breaksAfter[n]);
        var btd = document.createElement('td');
        btd.colSpan = indices.length + 1;
        btd.textContent = 'Page ' + breaksAfter[n] + ' starts here';
        brk.appendChild(btd);
        tbody.appendChild(brk);
      }
    }
    tableEl.appendChild(tbody);
  }

  function buildSection(def, result, cols) {
    var wrap = document.createElement('div');
    if (def.controls) wrap.appendChild(def.controls);
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

    renderTable(table, cols, def.rows, result, null, false, false);
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

  function visibleNormalRows(result) {
    return result.groups.normal.filter(function (r) { return !isHiddenInNormalTable(result, r); });
  }

  function visibleAltRows(result) {
    return allAltRows(result).filter(function (r) { return !isRowHidden(result, r); });
  }

  function printableBarcodes(result) {
    var toteIdx = result.columns.toteGid;
    var visibleTotes = {};
    visibleNormalRows(result).forEach(function (r) {
      var tote = String(result.rows[r][toteIdx] == null ? '' : result.rows[r][toteIdx]).trim();
      if (tote !== '') visibleTotes[tote] = true;
    });
    return result.barcodes.filter(function (b) { return visibleTotes[b.tote]; });
  }

  function currentRanges(result) {
    var total = visibleNormalRows(result).length;
    return chunkRanges(total, state.printPages, state.printBoundaries);
  }

  // Page-break markers are based on printable (not hidden) rows, but the table
  // still renders hidden rows. Map each printable page start back to the row
  // position in the full table so the marker/handle sits in the right place.
  function mainBreaks(result) {
    var ranges = currentRanges(result);
    if (ranges.length <= 1) return null;
    var fullRows = result.groups.normal;
    var breaks = {};
    for (var i = 0; i < ranges.length - 1; i++) {
      var nextStart = ranges[i + 1][0];
      var count = 0;
      var fullPos = -1;
      for (var p = 0; p < fullRows.length; p++) {
        if (isHiddenInNormalTable(result, fullRows[p])) continue;
        if (count === nextStart) { fullPos = p; break; }
        count++;
      }
      if (fullPos > 0) breaks[fullPos - 1] = i + 2;
    }
    return breaks;
  }

  // Convert a visible table row position into a printable-row count, so that
  // dragging a page handle ignores the hidden rows it passes over.
  function printableCountBefore(result, fullPos) {
    var count = 0;
    var rows = result.groups.normal;
    for (var p = 0; p < fullPos && p < rows.length; p++) {
      if (!isHiddenInNormalTable(result, rows[p])) count++;
    }
    return count;
  }

  // ---- Draggable page-break handles ---------------------------------------
  //
  // A handle sits in the left rail at each "Page X starts here" line. Drag it
  // up or down (or use the up/down arrow keys) to move where that page starts.

  var pendingHandleFocus = null;

  function layoutPageHandles(result) {
    var rail = $('pageHandleRail');
    if (!rail) return;
    rail.innerHTML = '';
    var table = $('previewTable');
    var wrap = rail.parentNode;
    if (!table || !wrap) return;

    var breakRows = table.querySelectorAll('tbody tr.page-break-row');
    wrap.classList.toggle('has-handles', breakRows.length > 0);
    if (!breakRows.length) return;

    var wrapTop = wrap.getBoundingClientRect().top;
    Array.prototype.forEach.call(breakRows, function (brk) {
      var boundary = Number(brk.dataset.boundary);
      var nextPage = Number(brk.dataset.nextPage);
      var rect = brk.getBoundingClientRect();

      var handle = document.createElement('button');
      handle.type = 'button';
      handle.className = 'page-handle';
      handle.dataset.boundary = String(boundary);
      handle.title = 'Drag up or down to move where page ' + nextPage + ' starts';
      handle.setAttribute('aria-label',
        'Start of page ' + nextPage + '. Drag, or press the up and down arrow keys to move it.');
      var grip = document.createElement('span');
      grip.className = 'page-handle-grip';
      grip.setAttribute('aria-hidden', 'true');
      handle.appendChild(grip);
      handle.style.top = (rect.top - wrapTop + rect.height / 2) + 'px';

      handle.addEventListener('pointerdown', function (e) { startBoundaryDrag(e, boundary); });
      handle.addEventListener('keydown', function (e) { onBoundaryKey(e, boundary); });
      rail.appendChild(handle);
    });

    if (pendingHandleFocus !== null) {
      var b = pendingHandleFocus;
      pendingHandleFocus = null;
      var el = rail.querySelector('.page-handle[data-boundary="' + b + '"]');
      if (el) el.focus({ preventScroll: true });
    }
  }

  function renderMainTableOnly(result) {
    var cols = visibleColumnIndices(result);
    var normalRows = result.groups.normal;
    var breaks = mainBreaks(result);
    renderTable($('previewTable'), cols, normalRows, result, breaks, true, true);
    $('troublesCount').textContent = '(' + normalRows.length + ')';
    $('previewNote').textContent = '';
    layoutPageHandles(result);
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

  // Set page-break `b` (0-based) to the given raw row count, clamp/cascade it
  // through chunkRanges, then re-render the table and handles.
  function applyBoundary(b, raw) {
    if (!state.result) return;
    var total = visibleNormalRows(state.result).length;
    var pages = effectivePages(total, state.printPages);
    state.printBoundaries = state.printBoundaries.slice();
    state.printBoundaries[b] = raw;
    var ranges = chunkRanges(total, pages, state.printBoundaries);
    state.printBoundaries = ranges.slice(0, -1).map(function (r) { return r[1]; });
    scheduleMainTable();
  }

  // The row position the boundary should snap to for a given pointer Y: the
  // top half of a row keeps the break above it, the bottom half moves it below.
  function rowPosAtClientY(clientY) {
    var rows = $('previewTable').querySelectorAll('tbody tr[data-row-pos]');
    if (!rows.length) return null;
    for (var i = 0; i < rows.length; i++) {
      var rect = rows[i].getBoundingClientRect();
      if (clientY < rect.top + rect.height / 2) {
        return Number(rows[i].dataset.rowPos);
      }
    }
    return Number(rows[rows.length - 1].dataset.rowPos) + 1;
  }

  var dragBoundary = null;

  function onBoundaryDragMove(e) {
    if (dragBoundary === null) return;
    e.preventDefault();
    var pos = rowPosAtClientY(e.clientY);
    if (pos !== null) applyBoundary(dragBoundary, printableCountBefore(state.result, pos));
  }

  function endBoundaryDrag() {
    dragBoundary = null;
    document.body.classList.remove('dragging-handle');
    window.removeEventListener('pointermove', onBoundaryDragMove);
    window.removeEventListener('pointerup', endBoundaryDrag);
    window.removeEventListener('pointercancel', endBoundaryDrag);
  }

  function startBoundaryDrag(e, boundary) {
    if (!state.result || e.button !== 0) return;
    e.preventDefault();
    dragBoundary = boundary;
    document.body.classList.add('dragging-handle');
    window.addEventListener('pointermove', onBoundaryDragMove);
    window.addEventListener('pointerup', endBoundaryDrag);
    window.addEventListener('pointercancel', endBoundaryDrag);
  }

  function onBoundaryKey(e, boundary) {
    if (e.key !== 'ArrowUp' && e.key !== 'ArrowDown') return;
    e.preventDefault();
    if (!state.result) return;
    var total = visibleNormalRows(state.result).length;
    var pages = effectivePages(total, state.printPages);
    var ranges = chunkRanges(total, pages, state.printBoundaries);
    var cur = ranges[boundary][1];
    pendingHandleFocus = boundary;
    applyBoundary(boundary, e.key === 'ArrowUp' ? cur - 1 : cur + 1);
  }

  function renderPrintLayout(result) {
    var total = visibleNormalRows(result).length;
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

    var hint = $('printLayoutHint');
    if (hint) {
      if (total === 0) {
        hint.textContent = 'No printable troubles.';
      } else if (pages === 1) {
        hint.textContent = 'Printing as a single list. Use + to split it across pages, then ' +
          'drag the handles beside the table to set where each page starts.';
      } else {
        hint.textContent = 'Drag the handles beside the troubles table to set where each page starts.';
      }
    }
  }

  function setPrintPages(pages) {
    if (!state.result) return;
    var total = visibleNormalRows(state.result).length;
    var p = effectivePages(total, pages);
    state.printPages = p;
    state.printBoundaries = evenBoundaries(total, p);
    renderTables(state.result);
  }

  function evenSplit() {
    if (!state.result) return;
    var total = visibleNormalRows(state.result).length;
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
      note: 'Rows with no active location but a reserve location, plus Troubles rows added with "+alt" (those leave the Troubles print). Send an alt form so an EO can grab the item. Use the buttons above to fill in the location and case for each item, then print.',
      controls: buildAltFormControls(),
      rows: allAltRows(result)
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
    updateAltFormButtons();
  }

  function renderStats(result) {
    var s = result.stats;
    var altCount = allAltRows(result).length;
    $('resultsSummary').textContent =
      'Kept ' + s.keptRows + ' of ' + s.inputRows + ' rows \u2014 ' +
      s.normalRows + ' trouble(s) to print, ' + altCount + ' alt form, ' +
      s.splitRows + ' split cancel, ' + s.holdRows + ' hold, ' +
      s.duplicateRows + ' duplicate.';

    var box = $('statsBox');
    box.innerHTML = '';
    box.appendChild(statCard(String(s.normalRows), 'Troubles (print)'));
    box.appendChild(statCard(String(altCount), 'Alt form (EO)'));
    box.appendChild(statCard(String(s.splitRows), 'Split cancel'));
    box.appendChild(statCard(String(s.holdRows), 'Hold (no split)'));
    box.appendChild(statCard(String(s.duplicateRows), 'Duplicates'));
    box.appendChild(statCard(String(s.multiNeedRows), 'Need > 1'));
    box.appendChild(statCard(String(s.lowActiveOHRows), 'Active OH < 1'));
    box.appendChild(statCard(String(s.distinctTotes), 'Barcodes'));
  }

  function renderResults(result) {
    renderStats(result);
    renderTables(result);
    renderBarcodePreview(result);
    $('resultsCard').classList.remove('hidden');
  }

  // ----------------------------------------------------- Alt form (EO pickup)
  //
  // The alt-form section lists items that have no active location but do have
  // a reserve location, plus rows the operator adds by hand (for example when
  // the system says stock exists but it cannot be found). The operator fills
  // in a name plus a Location and Case for each item (the fields that are not
  // on the OC Clearing export) in a dialog, then prints a simple form. The
  // typed values are remembered until "Start over" so the same form can be
  // printed at several locations.

  var ALT_ROW_SEP = '\u241f';   // record separator; cannot appear in a cell

  function cellText(v) { return String(v == null ? '' : v).trim(); }

  // Last N characters of a value, ignoring spaces. "Last 4 digits" of an oLPN.
  function lastChars(value, n) {
    var v = String(value == null ? '' : value).replace(/\s+/g, '');
    return v.length <= n ? v : v.slice(-n);
  }

  // Readable local date + time for the top of the printed alt form.
  function formatTimestamp(date) {
    try {
      return date.toLocaleString(undefined, {
        year: 'numeric', month: 'short', day: 'numeric',
        hour: 'numeric', minute: '2-digit'
      });
    } catch (e) {
      return date.toString();
    }
  }

  function altFormState() {
    if (!state.altForm) state.altForm = { name: '', entries: {} };
    return state.altForm;
  }

  // A stable identity for an alt row, so saved values reattach after a render.
  function altRowKey(result, rowIndex) {
    var c = result.columns;
    var row = result.rows[rowIndex] || [];
    return [
      cellText(row[c.oLPN]),
      cellText(row[c.item]),
      cellText(row[c.toteGid]),
      cellText(row[c.need])
    ].join(ALT_ROW_SEP);
  }

  function altEntry(result, rowIndex) {
    var data = altFormState();
    var key = altRowKey(result, rowIndex);
    if (!data.entries[key]) data.entries[key] = { location: '', caseNum: '' };
    return data.entries[key];
  }

  // Every alt row needs a location and a case, and the form needs a name.
  function altFormComplete(result) {
    if (!result) return false;
    var altRows = visibleAltRows(result);
    if (!altRows.length) return false;
    var data = altFormState();
    if (!cellText(data.name)) return false;
    for (var i = 0; i < altRows.length; i++) {
      var e = data.entries[altRowKey(result, altRows[i])];
      if (!e || !cellText(e.location) || !cellText(e.caseNum)) return false;
    }
    return true;
  }

  function buildAltFormControls() {
    var wrap = document.createElement('div');
    wrap.className = 'alt-form-controls';

    var fillBtn = document.createElement('button');
    fillBtn.type = 'button';
    fillBtn.id = 'altFormFillBtn';
    fillBtn.className = 'primary';
    fillBtn.textContent = '\u270E Fill in alt form info';
    fillBtn.addEventListener('click', openAltFormDialog);

    var printBtn = document.createElement('button');
    printBtn.type = 'button';
    printBtn.id = 'altFormPrintBtn';
    printBtn.textContent = '\uD83D\uDDA8 Print alt form';
    printBtn.addEventListener('click', printAltForm);

    var note = document.createElement('span');
    note.className = 'hint alt-form-status';
    note.id = 'altFormStatus';

    wrap.appendChild(fillBtn);
    wrap.appendChild(printBtn);
    wrap.appendChild(note);
    return wrap;
  }

  function updateAltFormButtons() {
    var fill = $('altFormFillBtn');
    var print = $('altFormPrintBtn');
    var note = $('altFormStatus');
    if (!fill || !print) return;

    var result = state.result;
    var altRows = result ? visibleAltRows(result) : [];
    var has = altRows.length > 0;
    var complete = has && altFormComplete(result);

    // Both buttons are dead when there is nothing to fill in.
    fill.disabled = !has;
    // Print stays dead until everything required has been saved.
    print.disabled = !complete;

    if (!note) return;
    if (!has) {
      note.textContent = result && allAltRows(result).length
        ? 'All alt-form rows are hidden.'
        : '';
    } else if (complete) note.textContent = 'Ready to print.';
    else note.textContent = 'Fill in the information before printing.';
  }

  function closeAltFormDialog() {
    var overlay = $('altFormOverlay');
    if (overlay && overlay.parentNode) overlay.parentNode.removeChild(overlay);
    document.body.classList.remove('modal-open');
    document.removeEventListener('keydown', onAltDialogKey);
  }

  function onAltDialogKey(e) {
    if (e.key === 'Escape') closeAltFormDialog();
  }

  function openAltFormDialog() {
    var result = state.result;
    var altRows = result ? visibleAltRows(result) : [];
    if (!altRows.length) return;

    closeAltFormDialog();
    var data = altFormState();
    var c = result.columns;

    var overlay = document.createElement('div');
    overlay.className = 'modal-overlay';
    overlay.id = 'altFormOverlay';

    var modal = document.createElement('div');
    modal.className = 'modal alt-form-modal';
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-modal', 'true');
    modal.setAttribute('aria-labelledby', 'altFormDialogTitle');

    var head = document.createElement('div');
    head.className = 'modal-head';
    var title = document.createElement('h3');
    title.id = 'altFormDialogTitle';
    title.textContent = 'Alt form \u2014 extra information';
    var closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'modal-close';
    closeBtn.setAttribute('aria-label', 'Close');
    closeBtn.textContent = '\u2715';
    closeBtn.addEventListener('click', closeAltFormDialog);
    head.appendChild(title);
    head.appendChild(closeBtn);

    var hint = document.createElement('p');
    hint.className = 'hint';
    hint.textContent = 'Enter your name, then the location and case for each item. ' +
      'The greyed-out fields come from the spreadsheet and cannot be changed.';

    var body = document.createElement('div');
    body.className = 'modal-body';

    var nameField = document.createElement('label');
    nameField.className = 'alt-field alt-name-field';
    nameField.appendChild(document.createTextNode('Name'));
    var nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.id = 'altFormName';
    nameInput.value = data.name || '';
    nameInput.autocomplete = 'name';
    nameInput.placeholder = 'Your name';
    nameField.appendChild(nameInput);
    body.appendChild(nameField);

    var list = document.createElement('div');
    list.className = 'alt-items';

    altRows.forEach(function (rowIndex, i) {
      var entry = altEntry(result, rowIndex);
      var row = result.rows[rowIndex];

      var item = document.createElement('div');
      item.className = 'alt-item';

      var num = document.createElement('div');
      num.className = 'alt-item-num';
      num.textContent = (i + 1) + '.';

      var fields = document.createElement('div');
      fields.className = 'alt-item-fields';

      function addField(labelText, value, opts) {
        opts = opts || {};
        var field = document.createElement('label');
        field.className = 'alt-field' + (opts.auto ? ' auto' : '');
        field.appendChild(document.createTextNode(labelText));
        var input = document.createElement('input');
        input.type = 'text';
        input.value = value;
        if (opts.auto) {
          input.readOnly = true;
          input.tabIndex = -1;
          input.setAttribute('aria-readonly', 'true');
        }
        if (opts.field) {
          input.dataset.field = opts.field;
          input.dataset.rowIndex = String(rowIndex);
        }
        field.appendChild(input);
        fields.appendChild(field);
        return input;
      }

      // Required order: SKU, Location, Case, Qty, oLPN.
      addField('SKU', cellText(row[c.item]), { auto: true });
      addField('Location', entry.location || '', { field: 'location' });
      addField('Case', entry.caseNum || '', { field: 'case' });
      addField('Qty', cellText(row[c.need]), { auto: true });
      addField('oLPN', lastChars(cellText(row[c.oLPN]), 4), { auto: true });

      item.appendChild(num);
      item.appendChild(fields);
      list.appendChild(item);
    });

    body.appendChild(list);

    var foot = document.createElement('div');
    foot.className = 'modal-foot';
    var error = document.createElement('span');
    error.className = 'status';
    error.id = 'altFormError';
    var cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'Cancel';
    cancelBtn.addEventListener('click', closeAltFormDialog);
    var saveBtn = document.createElement('button');
    saveBtn.type = 'button';
    saveBtn.className = 'primary';
    saveBtn.textContent = 'Save';
    saveBtn.addEventListener('click', saveAltFormDialog);
    foot.appendChild(error);
    foot.appendChild(cancelBtn);
    foot.appendChild(saveBtn);

    modal.appendChild(head);
    modal.appendChild(hint);
    modal.appendChild(body);
    modal.appendChild(foot);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    document.body.classList.add('modal-open');

    overlay.addEventListener('click', function (e) {
      if (e.target === overlay) closeAltFormDialog();
    });
    document.addEventListener('keydown', onAltDialogKey);

    nameInput.focus();
  }

  function joinWordList(items) {
    if (items.length <= 1) return items[0] || '';
    return items.slice(0, -1).join(', ') + ' and ' + items[items.length - 1];
  }

  function saveAltFormDialog() {
    var result = state.result;
    if (!result) return;

    var nameInput = $('altFormName');
    var name = nameInput ? nameInput.value.trim() : '';

    var byRow = {};
    Array.prototype.forEach.call(
      document.querySelectorAll('#altFormOverlay [data-field]'),
      function (input) {
        var ri = input.dataset.rowIndex;
        if (!byRow[ri]) byRow[ri] = {};
        byRow[ri][input.dataset.field] = input.value.trim();
      }
    );

    // Work out everything first and only touch the saved state once the whole
    // form is valid, so a failed save never half-clears it.
    var missing = [];
    if (!name) missing.push('your name');
    var pending = visibleAltRows(result).map(function (rowIndex, i) {
      var got = byRow[String(rowIndex)] || {};
      var location = got.location || '';
      var caseNum = got.case || '';
      if (!location) missing.push('location for item ' + (i + 1));
      if (!caseNum) missing.push('case for item ' + (i + 1));
      return { rowIndex: rowIndex, location: location, caseNum: caseNum };
    });

    if (missing.length) {
      var error = $('altFormError');
      if (error) {
        error.textContent = 'Please fill in ' + joinWordList(missing) + '.';
        error.className = 'status error';
      }
      return;
    }

    var data = altFormState();
    data.name = name;
    pending.forEach(function (p) {
      var key = altRowKey(result, p.rowIndex);
      data.entries[key] = { location: p.location, caseNum: p.caseNum };
    });

    closeAltFormDialog();
    updateAltFormButtons();
  }

  function buildAltFormPrintHtml(result) {
    var data = altFormState();
    var c = result.columns;

    var rows = visibleAltRows(result).map(function (rowIndex, i) {
      var entry = data.entries[altRowKey(result, rowIndex)] || {};
      var row = result.rows[rowIndex];
      return '<tr>' +
        '<td class="num">' + (i + 1) + '.</td>' +
        '<td>' + escapeHtml(cellText(row[c.item])) + '</td>' +
        '<td>' + escapeHtml(cellText(entry.location)) + '</td>' +
        '<td>' + escapeHtml(cellText(entry.caseNum)) + '</td>' +
        '<td>' + escapeHtml(cellText(row[c.need])) + '</td>' +
        '<td>' + escapeHtml(lastChars(cellText(row[c.oLPN]), 4)) + '</td>' +
        '</tr>';
    }).join('');

    var title = 'Alt Form';
    var locations = (state.selectedLocations || []).join(', ');
    var head = title + (locations ? ' \u2014 ' + locations : '');
    var stamp = formatTimestamp(new Date());

    return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' +
      escapeHtml(head) + '</title><style>' + ALT_FORM_PRINT_CSS + '</style></head><body>' +
      '<div class="alt-head"><h1>' + escapeHtml(head) + '</h1>' +
      '<div class="alt-stamp">Printed: ' + escapeHtml(stamp) + '</div></div>' +
      '<div class="alt-name">Name: ' + escapeHtml(cellText(data.name)) + '</div>' +
      '<table class="alt"><thead><tr>' +
      '<th class="num">#</th><th>SKU</th><th>Location</th><th>Case</th>' +
      '<th>Qty</th><th>oLPN</th></tr></thead><tbody>' + rows + '</tbody></table>' +
      '<p class="alt-thanks">Thank you!</p>' +
      '</body></html>';
  }

  var ALT_FORM_PRINT_CSS = [
    '@page { size: portrait; margin: 12mm; }',
    '* { box-sizing: border-box; }',
    'body { margin: 0; color: #000; font-family: Arial, Helvetica, sans-serif; font-size: 24px; }',
    '.alt-head { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; margin: 0 0 6px; }',
    'h1 { font-size: 28px; margin: 0; }',
    '.alt-stamp { font-size: 18px; color: #333; white-space: nowrap; }',
    '.alt-name { font-size: 24px; font-weight: 700; margin: 0 0 14px; padding-bottom: 6px; border-bottom: 2px solid #000; }',
    // Width auto keeps each column sized to its contents instead of stretching
    // the table across the whole page. Auto side margins centre the table.
    'table.alt { border-collapse: collapse; width: auto; margin-left: auto; margin-right: auto; }',
    'table.alt th, table.alt td { border: 1px solid #555; padding: 10px 8px; text-align: left; font-size: 24px; }',
    'table.alt th { background: #ececec; }',
    'table.alt td.num { text-align: right; white-space: nowrap; width: 1%; }',
    'table.alt tr { page-break-inside: avoid; }',
    '.alt-thanks { margin-top: 18px; font-size: 20px; }'
  ].join('\n');

  function printAltForm() {
    if (!state.result || !altFormComplete(state.result)) return;
    var html;
    try {
      html = buildAltFormPrintHtml(state.result);
    } catch (err) {
      console.error(err);
      setStatus('Could not build the alt form: ' + (err && err.message ? err.message : err), 'error');
      return;
    }
    printHtmlDocument(html);
  }

  // ------------------------------------------------------- Barcode preview

  var SVG_NS = 'http://www.w3.org/2000/svg';

  function makeBarcodeSvg(value) {
    var svg = document.createElementNS(SVG_NS, 'svg');
    JsBarcode(svg, value, {
      format: 'CODE128',
      width: 2,
      height: 60,
      margin: 10,            // white quiet zone so scanners ignore the page background
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
    var visibleBarcodes = printableBarcodes(result);
    if (!visibleBarcodes.length) {
      $('bcCount').textContent = '(0)';
      $('bcNote').textContent = result.barcodes.length
        ? 'All printable rows with barcodes are hidden, so there are no barcodes to print.'
        : 'No tote numbers were found in the selected rows, so there are no barcodes.';
      return;
    }
    for (var i = 0; i < visibleBarcodes.length; i++) {
      var b = visibleBarcodes[i];
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
    $('bcCount').textContent = '(' + visibleBarcodes.length + ')';
    $('bcNote').textContent = '';
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
    // Width auto keeps each column sized to its contents instead of stretching
    // the table across the whole page. Auto side margins centre the table.
    'table.sheet { border-collapse: collapse; width: auto; margin-left: auto; margin-right: auto; font-size: 10pt; }',
    'table.sheet th, table.sheet td { border: 1px solid #7f7f7f; padding: 2px 4px; text-align: left; white-space: nowrap; }',
    'table.sheet th { background: #e6e6e6; }',
    'table.sheet thead { display: table-header-group; }',
    'table.sheet tr { page-break-inside: avoid; }',
    'table.sheet td.lbl { text-align: center; font-weight: bold; }',
    'table.sheet tr.dup td { background: #ffc7ce; }',
    'table.sheet tr td.need-multi { background: #ffeb3b; color: #000; font-weight: bold; }',
    'table.sheet tr td.oh-low { background: #ffe0b2; color: #000; font-weight: bold; }',
    '.sheet-section.page-break { page-break-before: always; }',
    '.sheet-sub { font-size: 9pt; color: #444444; margin: 0 0 4px; }',
    '.labels-section { page-break-before: always; }',
    '.labels-section.first { page-break-before: auto; }',
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

  // mode: 'all' (sheet + barcodes), 'sheet', or 'barcodes'.
  function buildPrintHtml(result, mode) {
    mode = mode || 'all';
    var includeSheet = mode !== 'barcodes';
    var includeLabels = mode !== 'sheet';
    var headerIdx = {};
    result.header.forEach(function (h, i) { headerIdx[h] = i; });
    var cols = PRINT_COLUMNS
      .map(function (name) { return { name: name, i: headerIdx[name] }; })
      .filter(function (c) { return c.i !== undefined; });

    var thead = '<thead><tr>' + cols.map(function (c) {
      return '<th>' + escapeHtml(c.name) + '</th>';
    }).join('') + '</tr></thead>';

    var head = 'Putwall Troubles \u2014 ' + escapeHtml((state.selectedLocations || []).join(', '));

    var printableRows = visibleNormalRows(result);
    var total = printableRows.length;
    var pages = effectivePages(total, state.printPages);
    var ranges = chunkRanges(total, pages, state.printBoundaries);

    var sections = includeSheet ? ranges.map(function (range, idx) {
      var rows = printableRows.slice(range[0], range[1]);
      var tbody = '<tbody>' + rows.map(function (ri) {
        var row = result.rows[ri];
        var tds = cols.map(function (c) {
          var cls = '';
          if (c.name === 'Label') cls = ' class="lbl"';
          else if (c.name === 'Need' && result.multiNeed[ri]) cls = ' class="need-multi"';
          else if (c.name === 'Active OH' && result.lowActiveOH[ri]) cls = ' class="oh-low"';
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
    }).join('') : '';

    var labels = includeLabels ? printableBarcodes(result).map(function (b) {
      var svg = makeBarcodeSvgMarkup(b.tote);
      return '<div class="label"><div class="num">' + escapeHtml(b.label) +
        '</div>' + svg + '<div class="tote">' + escapeHtml(b.tote) + '</div></div>';
    }).join('') : '';

    var labelsSection = includeLabels
      ? '<div class="labels-section' + (includeSheet ? '' : ' first') + '">' + labels + '</div>'
      : '';

    return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + head +
      '</title><style>' + PRINT_CSS + '</style></head><body>' +
      sections + labelsSection +
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

  function printDocument(mode) {
    if (!state.result) return;
    var html;
    try {
      html = buildPrintHtml(state.result, mode);
    } catch (err) {
      console.error(err);
      setStatus('Could not build the print view: ' + (err && err.message ? err.message : err), 'error');
      return;
    }
    printHtmlDocument(html);
  }

  // Print a finished HTML document through a hidden iframe, so no pop-up
  // blocker is involved. Shared by the troubles print and the alt form.
  function printHtmlDocument(html) {
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

      if (result.lowActiveOH[ri] && result.columns.activeOH !== undefined) {
        var ohCell = row.getCell(result.columns.activeOH + 1);
        ohCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE0B2' } };
        ohCell.font = { bold: true, color: { argb: 'FF6B3B00' } };
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
    addDataSheet(wb, 'Alt Form', result, allAltRows(result), false);
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
    state.altForm = { name: '', entries: {} };
    state.hiddenRows = {};
    state.manualAltRows = {};
    closeAltFormDialog();
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
    $('printBtn').addEventListener('click', function () { printDocument('all'); });
    $('printSheetBtn').addEventListener('click', function () { printDocument('sheet'); });
    $('printBarcodesBtn').addEventListener('click', function () { printDocument('barcodes'); });
    $('columnsBtn').addEventListener('click', function () {
      state.showAllColumns = !state.showAllColumns;
      if (state.result) renderTables(state.result);
    });
    $('pageMinus').addEventListener('click', function () { setPrintPages(state.printPages - 1); });
    $('pagePlus').addEventListener('click', function () { setPrintPages(state.printPages + 1); });
    $('pageEvenBtn').addEventListener('click', evenSplit);
    $('resetBtn').addEventListener('click', reset);

    // Keep the page-break handles lined up with their rows on resize.
    window.addEventListener('resize', function () {
      if (state.result) layoutPageHandles(state.result);
    });
  }

  // Exposed so the print/preview can be inspected and tested.
  window.PWApp = {
    buildPrintHtml: function (mode) { return state.result ? buildPrintHtml(state.result, mode) : ''; },
    buildAltFormPrintHtml: function () {
      return state.result ? buildAltFormPrintHtml(state.result) : '';
    },
    altFormComplete: function () { return altFormComplete(state.result); },
    allAltRows: function () { return state.result ? allAltRows(state.result) : []; },
    toggleManualAltRow: function (rowIndex) {
      if (state.result) toggleManualAltRow(state.result, rowIndex);
    },
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
