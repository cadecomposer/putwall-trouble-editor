/*
 * Node test for the core processing logic (src/processor.js).
 *
 * Runs the same functions the browser uses, against a real OC Clearing export,
 * and prints the filtered / sorted / labeled / duplicate-marked result.
 *
 * Usage:
 *   node test/test-processor.js [path-to-xlsx] [TRBL-PW-01,TRBL-PW-05,...]
 *
 * With no arguments it looks for an "OC-CLEARING*.xlsx" file in the project
 * root and defaults to TRBL-PW-01,TRBL-PW-04,TRBL-PW-05.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const P = require(path.join(__dirname, '..', 'src', 'processor.js'));

const root = path.join(__dirname, '..');

function findSample() {
  const files = fs.readdirSync(root).filter(f => /^OC-CLEARING.*\.xlsx$/i.test(f));
  if (!files.length) return null;
  return path.join(root, files[0]);
}

const file = process.argv[2] || findSample();
const selected = (process.argv[3] || 'TRBL-PW-01,TRBL-PW-04,TRBL-PW-05').split(',');

if (!file) {
  console.error('No spreadsheet given and no OC-CLEARING*.xlsx found in the project root.');
  process.exit(1);
}

const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
const ws = wb.Sheets[wb.SheetNames[0]];
const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });

const result = P.processData(grid, selected);
const c = result.columns;

console.log('File:      ', path.basename(file));
console.log('Locations: ', selected.join(', '));
console.log('Stats:     ', JSON.stringify(result.stats));
console.log();
console.log(result.header.map(h => String(h).padEnd(12)).join(''));
console.log('-'.repeat(12 * result.header.length));
result.rows.forEach((r, i) => {
  const cells = result.header.map((_, ci) => String(r[ci] == null ? '' : r[ci]).padEnd(12));
  console.log(cells.join('') + (result.duplicates[i] ? '  <-- DUP' : ''));
});
console.log();
console.log('Barcodes:  ' + result.barcodes.map(b => b.label + '=' + b.tote).join(', '));
console.log('Label col: column ' + (c.label + 1) + ' (right of Tote Gid at column ' + (c.toteGid + 1) + ')');
