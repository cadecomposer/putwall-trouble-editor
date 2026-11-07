/*
 * Builds test/sample-sections.xlsx - a tiny, synthetic spreadsheet that covers
 * every section the editor can produce:
 *
 *   normal   OLPN-A / T100   active location present
 *   ALT      OLPN-B / T200   no active location, but a reserve location
 *   normal   OLPN-D / T401   active location present (makes order D findable)
 *   HOLD     OLPN-D / T400   no active/reserve, but order D is findable
 *   split    OLPN-C / T300   no active/reserve, order C has nothing findable
 *   dup      OLPN-A / T100   exact repeat of the first normal row
 *   dup      OLPN-A / T100   same oLPN/tote/item/need/active as the first row,
 *                            but a different "Time in OC" - still a duplicate
 *   normal   OLPN-A / T100   same fields but a DIFFERENT active location, so
 *                            it is NOT a duplicate
 *   normal   OLPN-B / T201   active location present
 *
 * It contains NO real order data, but it is still written as .xlsx and therefore
 * ignored by .gitignore (the tool's job is to handle spreadsheets; this sample
 * stays on your machine).
 *
 * Usage:  node test/make-sample-sections.js
 */
'use strict';

const fs = require('fs');
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));

const OUT = path.join(__dirname, 'sample-sections.xlsx');

const HEADER = [
  'Time in OC', 'Order Date', 'Wave', 'oLPN', 'OC Mod', 'Location', 'LPN Type',
  'Tote Gid', 'Item', 'Need', 'Active Locn', 'Active Locn 2', 'Active OH',
  'Reserve Locn', 'Reserve Locn 2', 'iLPN', 'iLPN2', 'Staging Locn'
];

function row(olpn, tote, item, need, active, reserve, timeInOC) {
  return [
    timeInOC || '<2 Hours', '11/05', 'W1', olpn, '', 'TRBL-PW-02', 'OC',
    tote, item, String(need), active, active, '1', reserve, '', '', '', ''
  ];
}

function buildAoa() {
  return [
    ['OC Clearing'],
    ['EFC3 - 819'],
    HEADER,
    row('OLPN-A', 'T100', '111', 1, 'B100-01-A1', ''),  // normal
    row('OLPN-B', 'T200', '222', 2, '', 'R200-02-B2'),   // ALT
    row('OLPN-C', 'T300', '333', 1, '', ''),             // split
    row('OLPN-D', 'T401', '444', 1, 'E100-05-E5', ''),   // normal (order D findable)
    row('OLPN-D', 'T400', '555', 1, '', ''),             // HOLD
    row('OLPN-A', 'T100', '111', 1, 'B100-01-A1', ''),   // duplicate of first normal
    row('OLPN-A', 'T100', '111', 1, 'B100-01-A1', '', '4 Hours'), // duplicate: only Time in OC differs
    row('OLPN-A', 'T100', '111', 1, 'B999-99-Z9', ''),   // NOT a duplicate: active location differs
    row('OLPN-B', 'T201', '777', 1, 'F100-06-F6', '')    // normal
  ];
}

function write() {
  const ws = XLSX.utils.aoa_to_sheet(buildAoa());
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Waving_1');
  fs.writeFileSync(OUT, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  return OUT;
}

if (require.main === module) {
  console.log('wrote ' + write());
}

module.exports = { write: write, buildAoa: buildAoa, OUT: OUT };
