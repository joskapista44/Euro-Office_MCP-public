'use strict';
/**
 * GH #134 (P1): read_print_area -- the product READ mirror of the released set_print_area / clear_print_area.
 * The projection (families/print-area-read.cjs) over the driver's defined-names read, the request it makes, and the catalogue entry.
 */
const test = require('node:test');
const assert = require('node:assert');
const { OPERATIONS, STATUS, validateReadRequest, xlsxRequest } = require('../../src/server/read-tools.cjs');
const { projectPrintAreaRead } = require('../../src/adapters/xlsx/families/print-area-read.cjs');

const inv = (names, extra = {}) => ({ sheets: [{ name: 'Budget', index: 0 }, { name: 'My Sheet', index: 1 }, { name: 'Other', index: 2 }], definedNames: names, ...extra });
const pa = (scopeSheet, ref) => ({ name: 'Print_Area', ref, scopeSheet });

test('GH #134: read_print_area is an ACCEPTED xlsx read operation taking { sheet }', () => {
  assert.strictEqual(OPERATIONS.xlsx.read_print_area.status, STATUS.ACCEPTED);
  const ok = validateReadRequest('xlsx', { fileId: '1', operation: 'read_print_area', sheet: 'Budget' });
  assert.strictEqual(ok.ok, true);
  const extra = validateReadRequest('xlsx', { fileId: '1', operation: 'read_print_area', sheet: 'Budget', range: 'A1' });
  assert.strictEqual(extra.ok, false);
  const missing = validateReadRequest('xlsx', { fileId: '1', operation: 'read_print_area' });
  assert.strictEqual(missing.ok, false);
});

test('GH #134: the request asks only for the defined-names read (no cell window)', () => {
  assert.deepStrictEqual(xlsxRequest('read_print_area', { sheet: 'Budget' }), { windows: [], dependents: true });
});

test('GH #134: one rectangle -> { set, range }; two or more -> { set, areas } in the stored order', () => {
  assert.deepStrictEqual(projectPrintAreaRead(inv([pa('Budget', 'Budget!$A$1:$D$10')]), { sheet: 'Budget' }), { sheet: 'Budget', printArea: { set: true, range: 'A1:D10' } });
  assert.deepStrictEqual(projectPrintAreaRead(inv([pa('Budget', '=Budget!$F$2:$G$3,Budget!$A$1:$B$2')]), { sheet: 'Budget' }), { sheet: 'Budget', printArea: { set: true, areas: ['F2:G3', 'A1:B2'] } });
  assert.deepStrictEqual(projectPrintAreaRead(inv([pa('Budget', 'Budget!$C$5')]), { sheet: 'Budget' }), { sheet: 'Budget', printArea: { set: true, range: 'C5' } });
  assert.deepStrictEqual(projectPrintAreaRead(inv([pa('My Sheet', "'My Sheet'!$A$1:$B$2")]), { sheet: 'My Sheet' }), { sheet: 'My Sheet', printArea: { set: true, range: 'A1:B2' } });
});

test('GH #134: absence is { set: false } -- another sheet\'s or a workbook-level name does not count', () => {
  const names = [pa('Other', 'Other!$A$1:$B$2'), { name: 'Print_Area', ref: 'Budget!$A$1', scopeSheet: null }, { name: 'Total', ref: 'Budget!$A$1', scopeSheet: 'Budget' }];
  assert.deepStrictEqual(projectPrintAreaRead(inv(names), { sheet: 'Budget' }), { sheet: 'Budget', printArea: { set: false } });
  assert.deepStrictEqual(projectPrintAreaRead(inv([]), { sheet: 'Budget' }), { sheet: 'Budget', printArea: { set: false } });
});

test('GH #134: a shape the writer has no word for is NAMED, never a value', () => {
  assert.deepStrictEqual(projectPrintAreaRead(inv([pa('Budget', 'Budget!$A:$B')]), { sheet: 'Budget' }), { sheet: 'Budget', printArea: 'UNOBSERVABLE:not-a-plain-rectangle', ref: 'Budget!$A:$B' });
  assert.deepStrictEqual(projectPrintAreaRead(inv([pa('Budget', 'Other!$A$1:$B$2')]), { sheet: 'Budget' }), { sheet: 'Budget', printArea: 'UNOBSERVABLE:part-on-another-sheet', ref: 'Other!$A$1:$B$2' });
  const two = projectPrintAreaRead(inv([pa('Budget', 'Budget!$A$1'), pa('Budget', 'Budget!$B$2')]), { sheet: 'Budget' });
  assert.strictEqual(two.printArea, 'UNOBSERVABLE:more-than-one-print-area');
});

test('GH #134: the names reader not read -> its own named gap (never { set: false }); a missing sheet -> SHEET_NOT_FOUND', () => {
  const gap = inv(null, { dependencies: { unobservable: [{ axis: 'defined-names', why: 'UNOBSERVABLE:bridge-capability-missing', host: null }] } });
  assert.deepStrictEqual(projectPrintAreaRead(gap, { sheet: 'Budget' }), { sheet: 'Budget', printArea: 'UNOBSERVABLE:bridge-capability-missing' });
  assert.deepStrictEqual(projectPrintAreaRead(inv(null), { sheet: 'Budget' }), { sheet: 'Budget', printArea: 'UNOBSERVABLE:defined-names-not-read' });
  assert.deepStrictEqual(projectPrintAreaRead(inv([]), { sheet: 'Nope' }), { sheet: 'Nope', error: 'SHEET_NOT_FOUND' });
});
