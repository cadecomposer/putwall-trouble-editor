/*
 * Automated test for the section logic (normal / alt / split / hold / duplicate).
 *
 * Uses the synthetic file from test/make-sample-sections.js and asserts how each
 * row is classified. Regenerates the .xlsx if it is missing.
 *
 * Usage:  node test/test-sections.js
 */
'use strict';

const fs = require('fs');
const path = require('path');
const assert = require('assert');

const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const PW = require(path.join(__dirname, '..', 'src', 'processor.js'));
const sample = require(path.join(__dirname, 'make-sample-sections.js'));

if (!fs.existsSync(sample.OUT)) {
  sample.write();
}

const wb = XLSX.read(fs.readFileSync(sample.OUT), { type: 'buffer' });
const grid = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
  header: 1, raw: true, defval: '', blankrows: true
});
const res = PW.processData(grid, ['TRBL-PW-02']);
const c = res.columns;

function catOf(oLPN, tote) {
  for (let i = 0; i < res.rows.length; i++) {
    if (String(res.rows[i][c.oLPN]) === oLPN && String(res.rows[i][c.toteGid]) === tote) {
      return res.categories[i];
    }
  }
  return '(missing)';
}

// --- expectations ---------------------------------------------------------
assert.strictEqual(res.stats.normalRows, 3, 'normal rows');
assert.strictEqual(res.stats.altRows, 1, 'alt rows');
assert.strictEqual(res.stats.splitRows, 1, 'split rows');
assert.strictEqual(res.stats.holdRows, 1, 'hold rows');
assert.strictEqual(res.stats.duplicateRows, 1, 'duplicate rows');

assert.strictEqual(catOf('OLPN-A', 'T100'), 'normal', 'A/T100 normal');
assert.strictEqual(catOf('OLPN-B', 'T200'), 'alt', 'B/T200 alt');
assert.strictEqual(catOf('OLPN-C', 'T300'), 'split', 'C/T300 split');
assert.strictEqual(catOf('OLPN-D', 'T401'), 'normal', 'D/T401 normal');
assert.strictEqual(catOf('OLPN-D', 'T400'), 'hold', 'D/T400 hold (order is findable)');
assert.strictEqual(catOf('OLPN-B', 'T201'), 'normal', 'B/T201 normal');

// The alt row keeps its reserve location and has no active location.
const altIdx = res.groups.alt[0];
assert.strictEqual(String(res.rows[altIdx][c.activeLocn]), '', 'alt has no active locn');
assert.strictEqual(String(res.rows[altIdx][c.reserveLocn]), 'R200-02-B2', 'alt has a reserve locn');

// Barcodes cover only the printable (normal) totes, in sorted order.
assert.deepStrictEqual(
  res.barcodes.map(b => b.tote),
  ['T201', 'T401', 'T100'],
  'barcodes are normal totes only'
);

console.log('All section tests passed:');
console.log('  normal=' + res.stats.normalRows +
  ' alt=' + res.stats.altRows +
  ' split=' + res.stats.splitRows +
  ' hold=' + res.stats.holdRows +
  ' duplicate=' + res.stats.duplicateRows);
console.log('  barcodes: ' + res.barcodes.map(b => b.label + ':' + b.tote).join(', '));
