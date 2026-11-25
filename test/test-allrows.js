/*
 * Automated test for the "all rows" processor option used by the split-cancel
 * "Print all" / download. With { allRows: true } the location filter is skipped,
 * so every split cancel in the report is returned regardless of the TRBL-PW
 * locations the user has ticked.
 *
 * Usage:  node test/test-allrows.js
 */
'use strict';

const path = require('path');
const assert = require('assert');

const PW = require(path.join(__dirname, '..', 'src', 'processor.js'));

const HEADER = [
  'Time in OC', 'Order Date', 'Wave', 'oLPN', 'OC Mod', 'Location', 'LPN Type',
  'Tote Gid', 'Item', 'Need', 'Active Locn', 'Active Locn 2', 'Active OH',
  'Reserve Locn', 'Reserve Locn 2', 'iLPN', 'iLPN2', 'Staging Locn'
];

function row(olpn, tote, item, need, active, reserve, location) {
  return [
    '<2 Hours', '11/05', 'W1', olpn, '', location, 'OC',
    tote, item, String(need), active, active, '1', reserve, '', '', '', ''
  ];
}

const L02 = 'TRBL-PW-02';
const L03 = 'TRBL-PW-03';

const grid = [
  ['OC Clearing'],
  ['EFC3 - 819'],
  HEADER,
  row('OLPN-A', 'T100', '111', 1, 'B100-01-A1', '', L02),  // normal at L02
  row('OLPN-C', 'T300', '333', 1, '', '', L02),             // split at L02
  row('OLPN-G', 'T700', '777', 1, '', '', L03),             // split at L03
  row('OLPN-H', 'T800', '888', 1, '', '', L02)              // hold: order H findable? no -> split
];

function splitTotes(res) {
  return res.groups.split
    .map(i => String(res.rows[i][res.columns.toteGid]))
    .sort();
}

// Normal run honours the location filter.
{
  const res = PW.processData(grid, [L02]);
  assert.deepStrictEqual(splitTotes(res), ['T300', 'T800'],
    'filtered run keeps only the selected location');
}

// allRows ignores the filter, so the L03 split cancel shows up too.
{
  const res = PW.processData(grid, [L02], { allRows: true });
  assert.deepStrictEqual(splitTotes(res), ['T300', 'T700', 'T800'],
    'allRows includes split cancels from every location');
  assert.strictEqual(res.stats.keptRows, 4, 'allRows keeps every data row');
}

// An empty selection with allRows still returns everything.
{
  const res = PW.processData(grid, [], { allRows: true });
  assert.strictEqual(res.stats.keptRows, 4, 'allRows works with no selection');
  assert.strictEqual(splitTotes(res).length, 3, 'allRows finds all three split cancels');
}

console.log('All all-rows tests passed.');
console.log('  filtered L02 split cancels: T300, T800');
console.log('  allRows split cancels:      T300, T700, T800');
