'use strict';
/**
 * GH #134 (P2): read_defined_names -- every defined name the driver reads, each the MIRROR of a set_defined_name step; Print_Titles in the set_print_titles shape.
 * The projection (families/defined-names-read.cjs) over the driver's defined-names read, the request it makes, and the catalogue entry.
 */
const test = require('node:test');
const assert = require('node:assert');
const { OPERATIONS, STATUS, validateReadRequest, xlsxRequest } = require('../../src/server/read-tools.cjs');
const { projectDefinedNamesRead, MAX_NAMES } = require('../../src/adapters/xlsx/families/defined-names-read.cjs');

const inv = (names, extra = {}) => ({ sheets: [{ name: 'Budget', index: 0 }, { name: 'My Sheet', index: 1 }], definedNames: names, ...extra });
const n = (name, ref, scopeSheet = null) => ({ name, ref, scopeSheet });

test('GH #134 P2: read_defined_names is an ACCEPTED xlsx read operation with no arguments', () => {
  assert.strictEqual(OPERATIONS.xlsx.read_defined_names.status, STATUS.ACCEPTED);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_defined_names' }).ok, true);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_defined_names', sheet: 'Budget' }).ok, false);
});

test('GH #134 P2: the request asks only for the defined-names read (no cell window)', () => {
  assert.deepStrictEqual(xlsxRequest('read_defined_names', {}), { windows: [], dependents: true });
});

test('GH #134 P2: each name mirrors a set_defined_name step -- refersTo with its =, scope only at sheet level, the driver order kept', () => {
  const r = projectDefinedNamesRead(inv([n('Rate', 'Budget!$B$2'), n('Local', "='My Sheet'!$A$1:$B$2", 'My Sheet')]));
  assert.deepStrictEqual(r, { names: [{ name: 'Rate', refersTo: '=Budget!$B$2' }, { name: 'Local', refersTo: "='My Sheet'!$A$1:$B$2", scope: 'My Sheet' }], count: 2 });
  assert.deepStrictEqual(projectDefinedNamesRead(inv([])), { names: [], count: 0 }, 'an empty workbook is an empty list, read');
});

test('GH #134 P2: Print_Titles -> builtIn + titles in the set_print_titles shape (columns / rows / both)', () => {
  const r = projectDefinedNamesRead(inv([
    n('Print_Titles', 'Budget!$1:$2', 'Budget'),
    n('Print_Titles', "'My Sheet'!$A:$B,'My Sheet'!$1:$1", 'My Sheet'),
    n('Print_Area', 'Budget!$A$1:$D$9', 'Budget'),
  ]));
  assert.deepStrictEqual(r.names[0], { name: 'Print_Titles', refersTo: '=Budget!$1:$2', scope: 'Budget', builtIn: 'Print_Titles', titles: { rows: '1:2' } });
  assert.deepStrictEqual(r.names[1].titles, { columns: 'A:B', rows: '1:1' });
  const rowsFirst = projectDefinedNamesRead(inv([n('Print_Titles', 'Budget!$1:$1,Budget!$A:$B', 'Budget')])).names[0].titles;
  assert.strictEqual(JSON.stringify(rowsFirst), JSON.stringify({ columns: 'A:B', rows: '1:1' }), 'the stored member order does not change the titles (fixed key order)');
  assert.deepStrictEqual(r.names[2], { name: 'Print_Area', refersTo: '=Budget!$A$1:$D$9', scope: 'Budget', builtIn: 'Print_Area' });
});

test('GH #134 P2: a Print_Titles shape the writer has no word for is NAMED, the raw refersTo kept', () => {
  const t = ref => projectDefinedNamesRead(inv([n('Print_Titles', ref, 'Budget')])).names[0];
  assert.deepStrictEqual(t('Budget!$A$1:$B$2'), { name: 'Print_Titles', refersTo: '=Budget!$A$1:$B$2', scope: 'Budget', builtIn: 'Print_Titles', titles: 'UNOBSERVABLE:not-a-titles-shape' });
  assert.strictEqual(t('Other!$1:$1').titles, 'UNOBSERVABLE:part-on-another-sheet');
  assert.strictEqual(t('Budget!$1:$1,Budget!$3:$3').titles, 'UNOBSERVABLE:not-a-titles-shape', 'two row bands are not one set_print_titles step');
  // a workbook-level name that happens to be called Print_Titles is an ordinary name, not the built-in
  assert.deepStrictEqual(projectDefinedNamesRead(inv([n('Print_Titles', 'Budget!$1:$1')])).names[0], { name: 'Print_Titles', refersTo: '=Budget!$1:$1' });
});

test('GH #134 P2: a names gap is its own named reason (never an empty list); over MAX_NAMES is named, never partial', () => {
  const gap = inv(null, { dependencies: { unobservable: [{ axis: 'defined-names', why: 'UNOBSERVABLE:bridge-capability-missing', host: null }] } });
  assert.deepStrictEqual(projectDefinedNamesRead(gap), { names: 'UNOBSERVABLE:bridge-capability-missing' });
  assert.deepStrictEqual(projectDefinedNamesRead(inv(null)), { names: 'UNOBSERVABLE:defined-names-not-read' });
  const many = Array.from({ length: MAX_NAMES + 1 }, (_, i) => n(`N${i}`, 'Budget!$A$1'));
  assert.deepStrictEqual(projectDefinedNamesRead(inv(many)), { names: `UNOBSERVABLE:more-than-${MAX_NAMES}-names`, count: MAX_NAMES + 1 });
});
