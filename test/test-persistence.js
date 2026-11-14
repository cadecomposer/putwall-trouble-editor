/*
 * Automated test for the remembered session (localStorage).
 *
 * The browser keeps the imported grid and the small UI/alt-form settings
 * between refreshes and restarts. This pins down the saved shape, the
 * version/corruption guards and the sanitising done when reading it back, so
 * an old or damaged entry can never break startup.
 *
 * Usage:  node test/test-persistence.js
 */
'use strict';

const path = require('path');
const assert = require('assert');

// src/app.js is the browser glue. Stub just enough of the DOM for it to load
// in Node; the storage helpers only need state and a localStorage.
let store = {};
let failWrites = false;

global.localStorage = {
  getItem: function (key) {
    return Object.prototype.hasOwnProperty.call(store, key) ? store[key] : null;
  },
  setItem: function (key, value) {
    if (failWrites) throw new Error('QuotaExceededError');
    store[key] = String(value);
  },
  removeItem: function (key) { delete store[key]; }
};
global.window = {};
global.document = { readyState: 'loading', addEventListener: function () {} };

require(path.join(__dirname, '..', 'src', 'app.js'));

const app = global.window.PWApp;
const storage = app.storage;
const state = app.state;

// 1. Settings round-trip, including the alt-form entries keyed by row.
state.selectedLocations = ['TRBL-PW-01', 'TRBL-PW-04'];
state.showAllColumns = true;
state.printPages = 3;
state.printBoundaries = [2, 5];
state.altForm = {
  name: 'Pat',
  entries: { 'oLPN\u241fSKU\u241fT1\u241f3': { location: 'A1', caseNum: 'C9' } }
};
state.hiddenRows = { '12': true };
state.manualAltRows = { '15': true };
state.rowOrder = [7, 3, 9];

assert.strictEqual(storage.saveSettingsNow(), true);
const savedSettings = JSON.parse(store[storage.settingsKey]);
assert.strictEqual(savedSettings.v, 1);
assert.deepStrictEqual(savedSettings.selectedLocations, ['TRBL-PW-01', 'TRBL-PW-04']);
assert.strictEqual(savedSettings.showAllColumns, true);
assert.strictEqual(savedSettings.printPages, 3);
assert.deepStrictEqual(savedSettings.printBoundaries, [2, 5]);
assert.deepStrictEqual(savedSettings.altForm, state.altForm);
assert.deepStrictEqual(savedSettings.hiddenRows, { '12': true });
assert.deepStrictEqual(savedSettings.manualAltRows, { '15': true });
assert.deepStrictEqual(savedSettings.rowOrder, [7, 3, 9]);

// Wipe state, then restore exactly what was saved.
state.selectedLocations = [];
state.showAllColumns = false;
state.printPages = 1;
state.printBoundaries = [];
state.altForm = { name: '', entries: {} };
state.hiddenRows = {};
state.manualAltRows = {};
state.rowOrder = [];

assert.strictEqual(storage.applySettings(storage.loadSettings()), true);
assert.deepStrictEqual(state.selectedLocations, ['TRBL-PW-01', 'TRBL-PW-04']);
assert.strictEqual(state.showAllColumns, true);
assert.strictEqual(state.printPages, 3);
assert.deepStrictEqual(state.printBoundaries, [2, 5]);
assert.deepStrictEqual(state.altForm, savedSettings.altForm);
assert.deepStrictEqual(state.hiddenRows, { '12': true });
assert.deepStrictEqual(state.manualAltRows, { '15': true });
assert.deepStrictEqual(state.rowOrder, [7, 3, 9]);

// 2. The grid lives under its own key.
state.grid = [['Tote Gid', 'Active Locn'], ['T1', 'A1']];
state.sourceFileName = 'OC-CLEARING.xlsx';
assert.strictEqual(storage.saveSessionNow(), true);
const savedSession = JSON.parse(store[storage.sessionKey]);
assert.strictEqual(savedSession.v, 1);
assert.strictEqual(savedSession.sourceFileName, 'OC-CLEARING.xlsx');
assert.deepStrictEqual(savedSession.grid, state.grid);

state.grid = null;
state.sourceFileName = '';
assert.strictEqual(storage.applySession(storage.loadSession()), true);
assert.deepStrictEqual(state.grid, savedSession.grid);
assert.strictEqual(state.sourceFileName, 'OC-CLEARING.xlsx');

// 3. Version and corruption guards: an unusable entry is ignored, not thrown.
store[storage.settingsKey] = JSON.stringify({ v: 999, selectedLocations: ['TRBL-PW-01'] });
assert.strictEqual(storage.loadSettings(), null);
store[storage.settingsKey] = '{not json';
assert.strictEqual(storage.loadSettings(), null);
store[storage.sessionKey] = JSON.stringify({ v: 1, grid: 'not a grid' });
assert.strictEqual(storage.applySession(storage.loadSession()), false);
assert.strictEqual(state.grid, null);

// 4. Values from storage are sanitised before they reach the app.
const badFlags = JSON.parse('{"5":1,"__proto__":{"polluted":true}}');
storage.applySettings({
  v: 1,
  selectedLocations: ['TRBL-PW-02', 'BOGUS', 42],
  showAllColumns: 'yes',
  printPages: 'nonsense',
  printBoundaries: [4, -1, 'x', 9],
  altForm: { name: 7, entries: { ok: { location: 'A1', caseNum: 3 }, bad: 'x' } },
  hiddenRows: ['not', 'a', 'map'],
  manualAltRows: badFlags,
  rowOrder: ['8', 'nope', 2]
});
assert.deepStrictEqual(state.selectedLocations, ['TRBL-PW-02']);
assert.strictEqual(state.showAllColumns, true);
assert.strictEqual(state.printPages, 1);
assert.deepStrictEqual(state.printBoundaries, [4, 9]);
assert.deepStrictEqual(state.altForm, { name: '', entries: { ok: { location: 'A1', caseNum: '' } } });
assert.deepStrictEqual(state.hiddenRows, {});
assert.deepStrictEqual(state.manualAltRows, { '5': true });
assert.deepStrictEqual(state.rowOrder, [8, 2]);
assert.strictEqual({}.polluted, undefined);

// 5. A full quota must not throw; the caller gets false and can carry on.
state.grid = [['x']];
failWrites = true;
assert.strictEqual(storage.saveSettingsNow(), false);
assert.strictEqual(storage.saveSessionNow(), false);
// The previous workbook is dropped so a refresh cannot restore it by mistake.
assert.strictEqual(store[storage.sessionKey], undefined);
failWrites = false;
assert.strictEqual(storage.saveSessionNow(), true);

// 6. "Start over" removes everything that was remembered.
storage.clearSavedState();
assert.strictEqual(store[storage.settingsKey], undefined);
assert.strictEqual(store[storage.sessionKey], undefined);
assert.strictEqual(storage.loadSettings(), null);
assert.strictEqual(storage.loadSession(), null);

console.log('All persistence tests passed.');
console.log('  settings and grid are stored separately, versioned and sanitised.');
console.log('  corrupt entries are ignored; a full quota fails gracefully.');
