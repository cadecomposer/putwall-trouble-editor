/*
 * Automated test for the split-cancel "Print all" set.
 *
 * "Print all" processes the whole report with every TRBL-PW location selected,
 * so it covers every putwall split cancel while leaving out rows at unrelated
 * (non-putwall) locations. This pins that down.
 *
 * Usage:  node test/test-split-all.js
 */
'use strict';

const path = require('path');
const assert = require('assert');

const PW = require(path.join(__dirname, '..', 'src', 'processor.js'));

// The same 20 locations the page offers.
const LOCATIONS = [];
for (let n = 1; n <= 20; n++) LOCATIONS.push('TRBL-PW-' + (n < 10 ? '0' + n : '' + n));

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
  row('OLPN-C', 'T300', '333', 1, '', '', L02),             // putwall split at L02
  row('OLPN-G', 'T700', '777', 1, '', '', L03),             // putwall split at L03
  row('OLPN-X', 'T900', '999', 1, '', '', 'SHIP-01'),       // split, NOT putwall
  row('OLPN-Y', 'T901', '998', 1, '', '', '')               // split, blank location
];

function splitTotes(res) {
  return res.groups.split
    .map(i => String(res.rows[i][res.columns.toteGid]))
    .sort();
}

// A single selected location only keeps that location's rows.
{
  const res = PW.processData(grid, [L02]);
  assert.deepStrictEqual(splitTotes(res), ['T300'],
    'a single selection keeps only that location');
}

// "Print all" = every TRBL-PW location, so both putwall splits show up but the
// non-putwall and blank-location splits are left out.
{
  const res = PW.processData(grid, LOCATIONS);
  assert.deepStrictEqual(splitTotes(res), ['T300', 'T700'],
    'all putwall locations keep every TRBL-PW split cancel');
  assert(!splitTotes(res).includes('T900'), 'a non-putwall location is filtered out');
  assert(!splitTotes(res).includes('T901'), 'a blank location is filtered out');
  assert.strictEqual(res.stats.keptRows, 3,
    'only the putwall rows are kept (T100, T300, T700)');
}

console.log('All split "print all" tests passed.');
console.log('  all putwall locations -> split cancels: T300, T700');
console.log('  non-putwall (SHIP-01) and blank-location rows are excluded.');
