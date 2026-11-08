/*
 * Automated test for the dynamic location filter.
 *
 * The same imported grid is processed with different location selections, and
 * the test checks that EVERYTHING is recalculated each time: sections (normal /
 * alt / split / hold), duplicates, tote-number labels and barcodes.
 *
 * This is the behaviour the browser relies on when location ticks are changed
 * after a file has already been imported.
 *
 * Usage:  node test/test-filter.js
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

// Two locations. Some rows deliberately depend on the OTHER location to prove
// that the filter is applied before labels/duplicates but after the whole-file
// "order is findable" rule.
const grid = [
  ['OC Clearing'],
  ['EFC3 - 819'],
  HEADER,
  row('OLPN-A', 'T100', '111', 1, 'B100-01-A1', '', L02),  // normal
  row('OLPN-B', 'T200', '222', 2, '', 'R200-02-B2', L02),   // alt
  row('OLPN-C', 'T300', '333', 1, '', '', L02),             // split
  row('OLPN-D', 'T401', '444', 1, 'E100-05-E5', '', L02),   // normal (order D findable)
  row('OLPN-D', 'T400', '555', 1, '', '', L02),             // hold
  row('OLPN-A', 'T100', '111', 1, 'B100-01-A1', '', L02),   // duplicate of first row
  row('OLPN-E', 'T500', '555', 1, 'B999-99-Z9', '', L03),   // normal
  row('OLPN-A', 'T100', '111', 1, 'B100-01-A1', '', L03),   // duplicate only if L02 is also selected
  row('OLPN-F', 'T600', '666', 1, '', '', L02),             // hold: order F is findable at L03
  row('OLPN-F', 'T601', '667', 1, 'E200-02-E2', '', L03)    // normal
];

function run(selected) {
  return PW.processData(grid, selected);
}

function findRow(res, predicate) {
  for (let i = 0; i < res.rows.length; i++) {
    if (predicate(res.rows[i], res.columns)) return i;
  }
  return -1;
}

function categoriesFor(res) {
  const out = {};
  res.rows.forEach((r, i) => {
    const c = res.columns;
    const key = [
      r[c.oLPN], r[c.toteGid], r[c.location]
    ].join('/');
    out[key] = res.categories[i];
  });
  return out;
}

// --- L02 only ------------------------------------------------------------
{
  const res = run([L02]);
  const s = res.stats;
  assert.strictEqual(s.keptRows, 7, 'L02 kept rows');
  assert.strictEqual(s.normalRows, 2, 'L02 normal');
  assert.strictEqual(s.altRows, 1, 'L02 alt');
  assert.strictEqual(s.splitRows, 1, 'L02 split');
  assert.strictEqual(s.holdRows, 2, 'L02 hold (one depends on unselected L03)');
  assert.strictEqual(s.duplicateRows, 1, 'L02 duplicate');
  assert.deepStrictEqual(
    res.barcodes.map(b => b.tote + '=' + b.label),
    ['T401=1', 'T100=2'],
    'L02 labels/barcodes'
  );

  const cats = categoriesFor(res);
  assert.strictEqual(cats['OLPN-B/T200/' + L02], 'alt', 'L02 alt form');
  assert.strictEqual(cats['OLPN-C/T300/' + L02], 'split', 'L02 split cancel');
  assert.strictEqual(cats['OLPN-D/T400/' + L02], 'hold', 'L02 hold from same-location findable item');
  assert.strictEqual(cats['OLPN-F/T600/' + L02], 'hold', 'L02 hold from an UNSELECTED location');
}

// --- L03 only ------------------------------------------------------------
{
  const res = run([L03]);
  const s = res.stats;
  assert.strictEqual(s.keptRows, 3, 'L03 kept rows');
  assert.strictEqual(s.normalRows, 3, 'L03 normal');
  assert.strictEqual(s.altRows, 0, 'L03 alt (alt row was L02)');
  assert.strictEqual(s.splitRows, 0, 'L03 split');
  assert.strictEqual(s.holdRows, 0, 'L03 hold');
  assert.strictEqual(s.duplicateRows, 0, 'L03 duplicate (T100 is the first row here)');
  assert.deepStrictEqual(
    res.barcodes.map(b => b.tote + '=' + b.label),
    ['T601=1', 'T500=2', 'T100=3'],
    'L03 labels/barcodes recalculated'
  );

  const cats = categoriesFor(res);
  assert.strictEqual(cats['OLPN-A/T100/' + L03], 'normal', 'T100 is not a duplicate on its own');
  assert.strictEqual(cats['OLPN-E/T500/' + L03], 'normal', 'L03 normal');
}

// --- both locations ------------------------------------------------------
{
  const res = run([L02, L03]);
  const s = res.stats;
  assert.strictEqual(s.keptRows, 10, 'both kept rows');
  assert.strictEqual(s.normalRows, 4, 'both normal');
  assert.strictEqual(s.altRows, 1, 'both alt');
  assert.strictEqual(s.splitRows, 1, 'both split');
  assert.strictEqual(s.holdRows, 2, 'both hold');
  assert.strictEqual(s.duplicateRows, 2, 'both duplicates (T100 now repeats across locations)');
  assert.deepStrictEqual(
    res.barcodes.map(b => b.tote + '=' + b.label),
    ['T601=1', 'T401=2', 'T500=3', 'T100=4'],
    'both labels/barcodes recalculated'
  );

  const cats = categoriesFor(res);
  assert.strictEqual(cats['OLPN-A/T100/' + L03], 'duplicate',
    'T100 at L03 becomes a duplicate when L02 is added');
}

// --- no locations selected ----------------------------------------------
{
  const res = run([]);
  assert.strictEqual(res.stats.keptRows, 0, 'no filter = no rows kept');
  assert.strictEqual(res.groups.normal.length, 0, 'no normal rows');
  assert.strictEqual(res.groups.alt.length, 0, 'no alt rows');
  assert.strictEqual(res.groups.split.length, 0, 'no split rows');
  assert.strictEqual(res.groups.hold.length, 0, 'no hold rows');
  assert.strictEqual(res.groups.duplicate.length, 0, 'no duplicate rows');
  assert.strictEqual(res.barcodes.length, 0, 'no barcodes');
}

console.log('All dynamic-filter tests passed.');
console.log('  L02 only: normal=2 alt=1 split=1 hold=2 duplicate=1');
console.log('  L03 only: normal=3 alt=0 split=0 hold=0 duplicate=0');
console.log('  both:     normal=4 alt=1 split=1 hold=2 duplicate=2');
console.log('  labels and barcodes change with the filter in every case.');