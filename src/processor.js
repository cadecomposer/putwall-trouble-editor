/*
 * Putwall Trouble Spreadsheet - core processing logic.
 *
 * This file is intentionally free of DOM and third-party dependencies so it
 * can run both in the browser (as window.PWProcessor) and in Node for tests.
 *
 * It works on a "grid": an array of rows, where each row is an array of cell
 * values (usually strings). The OC Clearing report has a few title rows at the
 * top, then a header row, then data. We detect the header row automatically.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.PWProcessor = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Known header labels -> internal column key. Compared case/space-insensitively.
  var HEADER_ALIASES = {
    'time in oc': 'timeInOC',
    'order date': 'orderDate',
    'wave': 'wave',
    'olpn': 'oLPN',
    'oc mod': 'ocMod',
    'location': 'location',
    'lpn type': 'lpnType',
    'tote gid': 'toteGid',
    'item': 'item',
    'need': 'need',
    'active locn': 'activeLocn',
    'active locn 2': 'activeLocn2',
    'active oh': 'activeOH',
    'reserve locn': 'reserveLocn',
    'reserve locn 2': 'reserveLocn2',
    'ilpn': 'iLPN',
    'ilpn2': 'iLPN2',
    'staging locn': 'stagingLocn'
  };

  function normLabel(s) {
    return String(s == null ? '' : s)
      .replace(/\u00a0/g, ' ')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ');
  }

  // Convert any spreadsheet cell value into the plain text we use internally.
  function cellToString(v) {
    if (v == null) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') {
      if (!isFinite(v)) return '';
      // Avoid scientific notation for ordinary numbers.
      return String(v);
    }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (v instanceof Date) {
      // Keep it simple and stable: ISO date part.
      var y = v.getFullYear();
      var m = String(v.getMonth() + 1).padStart(2, '0');
      var d = String(v.getDate()).padStart(2, '0');
      return y + '-' + m + '-' + d;
    }
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) {
        return v.richText.map(function (r) { return r.text || ''; }).join('');
      }
      if (v.text != null) return String(v.text);
      if (v.result != null) return cellToString(v.result);
      if (v instanceof Date) return cellToString(v);
    }
    return String(v);
  }

  function rowToNormalized(row) {
    var out = {};
    for (var i = 0; i < row.length; i++) {
      var key = HEADER_ALIASES[normLabel(row[i])];
      if (key && out[key] === undefined) out[key] = i;
    }
    return out;
  }

  // Find the header row: the first row that contains both "Tote Gid" and
  // "Active Locn". Returns -1 when not found.
  function findHeaderRow(rows) {
    for (var r = 0; r < rows.length; r++) {
      var cols = rowToNormalized(rows[r] || []);
      if (cols.toteGid !== undefined && cols.activeLocn !== undefined) return r;
    }
    return -1;
  }

  function maxRowLength(rows, from, to) {
    var max = 0;
    for (var i = from; i < to; i++) {
      if (rows[i] && rows[i].length > max) max = rows[i].length;
    }
    return max;
  }

  function padRow(row, width) {
    var out = row.slice(0, width);
    while (out.length < width) out.push('');
    return out;
  }

  // Case-insensitive, whitespace-trimmed location compare.
  function normalizeLocation(s) {
    return String(s == null ? '' : s).replace(/\u00a0/g, ' ').trim().toUpperCase();
  }

  /**
   * Run the full pipeline.
   *
   * @param {Array<Array<any>>} rawRows  grid including title rows + header row
   * @param {Array<string>} selectedLocations  e.g. ['TRBL-PW-01','TRBL-PW-05']
   * @returns {object} result
   */
  function processData(rawRows, selectedLocations) {
    var rows = rawRows || [];
    var headerIndex = findHeaderRow(rows);
    if (headerIndex < 0) {
      throw new Error(
        'Could not find the header row. Expected a row containing both ' +
        '"Tote Gid" and "Active Locn". Is this an OC Clearing spreadsheet?'
      );
    }

    var srcCols = rowToNormalized(rows[headerIndex]);
    if (srcCols.location === undefined) {
      throw new Error('Could not find a "Location" column in the header row.');
    }

    var header = padRow((rows[headerIndex] || []).map(cellToString), maxRowLength(rows, headerIndex, rows.length));
    var dataRows = rows.slice(headerIndex + 1).map(function (r) { return padRow(r || [], header.length); });

    // 1. Filter to the selected trouble locations.
    var wanted = {};
    (selectedLocations || []).forEach(function (loc) {
      wanted[normalizeLocation(loc)] = true;
    });
    var filtered = dataRows.filter(function (row) {
      return wanted[normalizeLocation(row[srcCols.location])];
    });

    // 2. Sort by the first "Active Locn" column, Z-A. Blank values go last.
    var activeIdx = srcCols.activeLocn;
    var decorated = filtered.map(function (row, i) {
      return { row: row, i: i, key: String(row[activeIdx] == null ? '' : row[activeIdx]).trim() };
    });
    decorated.sort(function (a, b) {
      var ak = a.key.toUpperCase();
      var bk = b.key.toUpperCase();
      if (ak === '' && bk === '') return a.i - b.i;
      if (ak === '') return 1;   // blanks last
      if (bk === '') return -1;
      if (ak < bk) return 1;     // Z-A
      if (ak > bk) return -1;
      return a.i - b.i;          // stable for ties
    });
    var sorted = decorated.map(function (d) { return d.row; });

    // 3. Insert the new "Label" column to the right of Tote Gid.
    var toteIdx = srcCols.toteGid;
    var outCols = {};
    Object.keys(srcCols).forEach(function (k) {
      var idx = srcCols[k];
      outCols[k] = idx > toteIdx ? idx + 1 : idx;
    });
    outCols.label = toteIdx + 1;

    var outHeader = header.slice(0, toteIdx + 1)
      .concat(['Label'])
      .concat(header.slice(toteIdx + 1));

    // Assign labels: number distinct tote numbers 1..N in top-to-bottom order.
    var labelByTote = {};
    var nextLabel = 1;
    var outRows = sorted.map(function (row) {
      var tote = String(row[toteIdx] == null ? '' : row[toteIdx]).trim();
      var label = '';
      if (tote !== '') {
        if (labelByTote[tote] === undefined) {
          labelByTote[tote] = nextLabel++;
        }
        label = labelByTote[tote];
      }
      return row.slice(0, toteIdx + 1).concat([label]).concat(row.slice(toteIdx + 1));
    });

    // 4. Mark duplicate rows (entire original row identical). First is kept
    //    un-marked; every later repetition is marked.
    var seen = {};
    var duplicates = outRows.map(function (row) {
      var key = JSON.stringify(row.slice(0, toteIdx + 1).concat(row.slice(toteIdx + 2)));
      if (seen[key]) return true;
      seen[key] = true;
      return false;
    });

    // 5. Classify each row:
    //    - duplicate: a later exact copy of an earlier row
    //    - normal:    has an active location (the printable troubles)
    //    - alt:       no active location, but a reserve location exists
    //                 (send an alt form so EO can grab it)
    //    - split:     no active location and no reserve location (split cancel)
    //    - hold:      a would-be split cancel, but another item on the SAME
    //                 order (same oLPN) is findable, so do NOT split cancel yet.
    var activeIdx = outCols.activeLocn;
    var reserveIdx = outCols.reserveLocn;
    var olpnIdx = outCols.oLPN;
    var srcActive = srcCols.activeLocn;
    var srcReserve = srcCols.reserveLocn;
    var srcOlpn = srcCols.oLPN;
    var trimText = function (v) { return String(v == null ? '' : v).trim(); };

    // Which orders (oLPN) have at least one findable item, across the WHOLE
    // file (not just the selected locations), so we never cancel prematurely.
    var orderHasFindable = {};
    dataRows.forEach(function (row) {
      var olpn = trimText(row[srcOlpn]);
      if (olpn === '') return;
      if (trimText(row[srcActive]) !== '' || trimText(row[srcReserve]) !== '') {
        orderHasFindable[olpn] = true;
      }
    });

    var categories = outRows.map(function (row, i) {
      if (duplicates[i]) return 'duplicate';
      if (trimText(row[activeIdx]) !== '') return 'normal';
      if (trimText(row[reserveIdx]) !== '') return 'alt';
      return orderHasFindable[trimText(row[olpnIdx])] ? 'hold' : 'split';
    });

    var groups = { normal: [], alt: [], split: [], hold: [], duplicate: [] };
    categories.forEach(function (cat, i) { groups[cat].push(i); });

    // 6. Barcode list: distinct totes on the printable (normal) rows only,
    //    in top-to-bottom (label) order.
    var barcodes = [];
    var emitted = {};
    outRows.forEach(function (row, i) {
      if (categories[i] !== 'normal') return;
      var tote = trimText(row[toteIdx]);
      if (tote === '' || emitted[tote]) return;
      emitted[tote] = true;
      barcodes.push({ label: labelByTote[tote], tote: tote, rowIndex: i });
    });

    var duplicateCount = duplicates.filter(Boolean).length;

    // 7. Flag rows whose "Need" is more than 1 (these get highlighted yellow).
    var needIdx = outCols.need;
    var multiNeed = outRows.map(function (row) {
      if (needIdx === undefined) return false;
      var raw = row[needIdx];
      var v = String(raw == null ? '' : raw).replace(/,/g, '').trim();
      if (v === '') return false;
      var num = Number(v);
      return isFinite(num) && num > 1;
    });
    var multiNeedCount = multiNeed.filter(Boolean).length;

    return {
      header: outHeader,
      rows: outRows,
      columns: outCols,
      duplicates: duplicates,
      multiNeed: multiNeed,
      categories: categories,
      groups: groups,
      barcodes: barcodes,
      stats: {
        headerIndex: headerIndex,
        inputRows: dataRows.length,
        keptRows: outRows.length,
        filteredOut: dataRows.length - outRows.length,
        normalRows: groups.normal.length,
        altRows: groups.alt.length,
        splitRows: groups.split.length,
        holdRows: groups.hold.length,
        duplicateRows: duplicateCount,
        multiNeedRows: multiNeedCount,
        distinctTotes: barcodes.length
      }
    };
  }

  return {
    processData: processData,
    findHeaderRow: findHeaderRow,
    cellToString: cellToString,
    normalizeLocation: normalizeLocation,
    normLabel: normLabel,
    HEADER_ALIASES: HEADER_ALIASES
  };
});
