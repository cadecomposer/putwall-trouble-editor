/*
 * Automated test for custom Troubles row ordering.
 *
 * The browser keeps a user-dragged order as stable source-row ids. This tests
 * that the order is applied to the visible/printable Troubles rows and that a
 * reset is detected correctly.
 *
 * Usage:  node test/test-order.js
 */
'use strict';

const path = require('path');
const assert = require('assert');

// src/app.js is the browser glue. Stub just enough of the DOM for it to load in
// Node; the ordering helpers only need the result object and state.
global.window = {};
global.document = { readyState: 'loading', addEventListener: function () {} };

require(path.join(__dirname, '..', 'src', 'app.js'));

const app = global.window.PWApp;
const result = {
  groups: { normal: [0, 1, 2, 3] },
  rowIds: ['a', 'b', 'c', 'd']
};

function order() { return app.orderedNormalRows(result).join(','); }

// No custom order = the processor's own row order.
app.state.rowOrder = [];
assert.strictEqual(order(), '0,1,2,3', 'default order');
assert.strictEqual(app.hasCustomRowOrder(result), false, 'default is not custom');

// A custom order is applied top to bottom.
app.state.rowOrder = ['d', 'b', 'a', 'c'];
assert.strictEqual(order(), '3,1,0,2', 'custom order');
assert.strictEqual(app.hasCustomRowOrder(result), true, 'custom order detected');

// Hidden rows keep their relative position and are left out of print order.
app.state.hiddenRows = { b: true };
assert.strictEqual(app.visibleNormalRows(result).join(','), '3,0,2',
  'hidden row skipped in custom order');
app.state.hiddenRows = {};

// Returning to the default order counts as reset (reset button disabled).
app.state.rowOrder = ['a', 'b', 'c', 'd'];
assert.strictEqual(app.hasCustomRowOrder(result), false, 'default order after dragging');

console.log('All row-order tests passed.');
console.log('  custom order: ' + ['d', 'b', 'a', 'c'].join(',') + ' -> rows 3,1,0,2');
