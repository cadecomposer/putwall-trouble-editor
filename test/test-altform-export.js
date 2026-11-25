/*
 * Automated test for the alt-form print/docx export and the split-cancel print.
 *
 * These are browser helpers in src/app.js, so this stubs just enough of the DOM
 * for the file to load in Node, then calls the builders directly through
 * window.PWApp.
 *
 * Usage:  node test/test-altform-export.js
 */
'use strict';

const path = require('path');
const assert = require('assert');

global.window = {};
global.document = { readyState: 'loading', addEventListener: function () {} };

require(path.join(__dirname, '..', 'src', 'app.js'));
const app = global.window.PWApp;

const HEADER = [
  'Time in OC', 'Order Date', 'Wave', 'oLPN', 'OC Mod', 'Location', 'LPN Type',
  'Tote Gid', 'Label', 'Item', 'Need', 'Active Locn', 'Active Locn 2',
  'Active OH', 'Reserve Locn', 'Reserve Locn 2', 'iLPN', 'iLPN2', 'Staging Locn'
];
const columns = {
  timeInOC: 0, orderDate: 1, wave: 2, oLPN: 3, ocMod: 4, location: 5, lpnType: 6,
  toteGid: 7, label: 8, item: 9, need: 10, activeLocn: 11, activeLocn2: 12,
  activeOH: 13, reserveLocn: 14, reserveLocn2: 15, iLPN: 16, iLPN2: 17,
  stagingLocn: 18
};

function row(oLPN, tote, item, need, reserve) {
  return ['<2 Hours', '11/05', 'W1', oLPN, '', 'TRBL-PW-02', 'OC', tote, '1',
    item, String(need), '', '', '1', reserve, '', '', '', ''];
}

const result = {
  header: HEADER,
  columns: columns,
  rows: [
    row('000123456789', 'T200', '222', 2, 'R200-02-B2'),  // alt
    row('000987654321', 'T300', '333', 1, '')             // split
  ],
  rowIds: [1, 2],
  groups: { normal: [], alt: [0], split: [1], hold: [], duplicate: [] },
  categories: ['alt', 'split'],
  duplicates: [false, false],
  multiNeed: [true, false],
  lowActiveOH: [false, false]
};

// The alt-form entries are keyed by oLPN/item/tote/need joined with the record
// separator the app uses internally.
const SEP = '\u241f';
const key = ['000123456789', '222', 'T200', '2'].join(SEP);

app.state.result = result;
app.state.allSplitResult = result;
app.state.selectedLocations = ['TRBL-PW-02'];
app.state.altForm = {
  name: 'Alex',
  priority: false,
  entries: {}
};
app.state.altForm.entries[key] = { location: 'R200-02-B2', caseNum: '12345' };

// --- alt form print ------------------------------------------------------
{
  const html = app.buildAltFormPrintHtml();
  assert(html.includes('Putwall Alt Form'), 'title is "Putwall Alt Form"');
  assert(html.includes('class="alt-head"'), 'title block present');
  assert(!html.includes('TRBL-PW-02'), 'the TRBL-PW locations are not on the form');
  assert(html.includes('Alex'), 'name is on the form');
  assert(html.includes('R200-02-B2'), 'location is on the form');
  assert(html.includes('12345'), 'case is on the form');
  assert(!html.includes('PRIORITY'), 'no priority suffix by default');
}

// --- priority ------------------------------------------------------------
{
  app.state.altForm.priority = true;
  const html = app.buildAltFormPrintHtml();
  assert(html.includes('Putwall Alt Form - <strong>PRIORITY</strong>'),
    'priority title is bold');
  app.state.altForm.priority = false;
}

// --- alt form docx -------------------------------------------------------
{
  const bytes = app.buildAltFormDocx();
  assert(bytes && bytes.length > 0, 'docx bytes produced');
  assert.strictEqual(bytes[0], 0x50, 'docx starts with the ZIP signature P');
  assert.strictEqual(bytes[1], 0x4b, 'docx starts with the ZIP signature K');

  // Stored (uncompressed) entries, so the XML is literally in the archive.
  const asText = Buffer.from(bytes).toString('latin1');
  assert(asText.includes('word/document.xml'), 'document part is in the package');
  assert(asText.includes('[Content_Types].xml'), 'content types part is in the package');
  assert(asText.includes('Putwall Alt Form'), 'document XML has the title');
  assert(asText.includes('R200-02-B2'), 'document XML has the location');

  const xml = app.buildAltFormDocumentXml();
  assert(xml.includes('<w:jc w:val="center"/>'), 'title is centred in Word');
  assert(xml.includes('<w:b/>'), 'priority run can be bold');
}

// --- split cancel print columns -----------------------------------------
{
  const names = app.splitPrintColumns().map(c => c.name);
  assert.strictEqual(names[0], 'oLPN', 'split columns start at oLPN');
  assert.strictEqual(names[names.length - 1], 'Reserve Locn 2',
    'split columns end at Reserve Locn 2');
  assert(names.includes('Reserve Locn 2'), 'Reserve Locn 2 is included');

  const html = app.buildSplitPrintHtml('visible');
  assert(html.includes('Putwall Split Cancels'), 'split print has a title');
  assert(html.includes('TRBL-PW-02'), 'visible print names the selected locations');
  assert(html.includes('000987654321'), 'visible print includes the split row');
  assert(html.includes('Reserve Locn 2'), 'visible print has the wide column set');
  assert(html.includes('font-size: 7pt'), 'split print uses the smaller font');

  const all = app.buildSplitPrintHtml('all');
  assert(all.includes('All locations'), 'all print is labelled');
}

console.log('All alt-form/split export tests passed.');
console.log('  split print columns: ' + app.splitPrintColumns().map(c => c.name).join(', '));
