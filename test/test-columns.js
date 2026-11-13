/*
 * Automated test for the compact/expanded web-table columns.
 *
 * The web tables hide a few columns by default ("compact" mode) and the
 * "Show all columns" button reveals every column. This pins down exactly which
 * columns are hidden so a change to that list cannot slip through unnoticed.
 *
 * Usage:  node test/test-columns.js
 */
'use strict';

const path = require('path');
const assert = require('assert');

// src/app.js is the browser glue. Stub just enough of the DOM for it to load in
// Node; the column helper only needs the result object and state.
global.window = {};
global.document = { readyState: 'loading', addEventListener: function () {} };

require(path.join(__dirname, '..', 'src', 'app.js'));

const app = global.window.PWApp;
const header = [
  'Time in OC', 'Order Date', 'Wave', 'oLPN', 'OC Mod', 'Location', 'LPN Type',
  'Tote Gid', 'Item', 'Need', 'Active Locn', 'Active Locn 2', 'Active OH',
  'Reserve Locn', 'Reserve Locn 2', 'iLPN', 'iLPN2', 'Staging Locn'
];
const columns = {
  timeInOC: 0, orderDate: 1, wave: 2, oLPN: 3, ocMod: 4, location: 5,
  lpnType: 6, toteGid: 7, item: 8, need: 9, activeLocn: 10, activeLocn2: 11,
  activeOH: 12, reserveLocn: 13, reserveLocn2: 14, iLPN: 15, iLPN2: 16,
  stagingLocn: 17
};
const result = { header: header, columns: columns };

function labels(indices) {
  return indices.map(function (i) { return header[i]; });
}

// Compact mode (the default) still skips OC Mod, LPN Type and Reserve Locn 2.
app.state.showAllColumns = false;
assert.deepStrictEqual(
  labels(app.visibleColumnIndices(result)),
  ['oLPN', 'Location', 'Tote Gid', 'Item', 'Need', 'Active Locn',
    'Active Locn 2', 'Active OH', 'Reserve Locn'],
  'compact columns'
);
assert(!labels(app.visibleColumnIndices(result)).includes('LPN Type'),
  'compact mode hides LPN Type');

// Expanded mode shows every column in order.
app.state.showAllColumns = true;
assert.deepStrictEqual(
  labels(app.visibleColumnIndices(result)),
  header,
  'expanded columns'
);

console.log('All column tests passed.');
app.state.showAllColumns = false;
console.log('  compact list: ' + labels(app.visibleColumnIndices(result)).join(', '));
