'use strict';

// L-09 unit 1: the read_layout projection (src/adapters/xlsx/families/layout-read.cjs). The ENGINE side is simulated in the shape the driver's `layout` read returns it
// (xlsx-driver.cjs: per column GetColumnWidth / GetHidden, per row GetRowHeight / GetHidden; a hidden column reads width 0, a hidden row height null -- W4b LIVE,
// layout.cjs header), from the writer's OWN lowered layout_range operation, so a round trip request -> engine entry -> projection must reproduce the request.
// The LIVE acceptance of the same is test/live/release-l09-layout-read-product.live.cjs (through the product, EO_BIN).

const test = require('node:test');
const assert = require('node:assert/strict');

const layout = require('../../src/adapters/xlsx/families/layout.cjs');
const { projectLayoutRead, layoutExtent, CEILING } = require('../../src/adapters/xlsx/families/layout-read.cjs');
const rt = require('../../src/server/read-tools.cjs');

const inv = (sheetLayout, name = 'S') => ({ sheets: [{ name, index: 0, used: 'A1:E5' }], layout: sheetLayout === undefined ? null : { [name]: sheetLayout } });
const read = (sheetLayout, range, name = 'S') => projectLayoutRead(inv(sheetLayout, name), { sheet: name, range });
/** What the engine answers for a column / row the writer set. */
const engine = { col: (w, hidden = false) => ({ width: hidden ? 0 : w, hidden }), row: (h, hidden = false) => ({ height: hidden ? null : h, hidden }) };

test('ROUND TRIP: a width / height the writer lowered reads back as the same number (the read mirrors layout_range)', () => {
  const w = layout.plan({ intent: 'layout_range', sheet: 'S', columns: 'B:D', width: 20 }, 0).operation;
  const cols = Object.fromEntries(w.items.map(c => [c, engine.col(w.value)]));
  assert.deepEqual(read({ cols, rows: {} }, 'B1:D1').columns, { B: { width: 20, hidden: false }, C: { width: 20, hidden: false }, D: { width: 20, hidden: false } });
  const h = layout.plan({ intent: 'layout_range', sheet: 'S', rows: '2:3', height: 30 }, 0).operation;
  const rows = Object.fromEntries(h.items.map(r => [r, engine.row(h.value)]));
  assert.deepEqual(read({ cols: {}, rows }, 'A2:A3').rows, { 2: { height: 30, hidden: false }, 3: { height: 30, hidden: false } });
});

test('hidden: true / false round trip -- a hidden column / row is NAMED (its size reads 0 / null), never reported as a number', () => {
  const hid = layout.plan({ intent: 'layout_range', sheet: 'S', columns: 'B', hidden: true }, 0).operation;
  assert.equal(hid.value, true);
  const out = read({ cols: { B: engine.col(0, true), C: engine.col(12) }, rows: { 2: engine.row(0, true), 3: engine.row(18) } }, 'B2:C3');
  assert.deepEqual(out.columns.B, { hidden: true, width: null, reason: out.columns.B.reason });
  assert.match(out.columns.B.reason, /^HIDDEN:.*width/);
  assert.deepEqual(out.columns.C, { width: 12, hidden: false });
  assert.deepEqual(out.rows[2], { hidden: true, height: null, reason: out.rows[2].reason });
  assert.match(out.rows[2].reason, /^HIDDEN:.*height/);
  assert.deepEqual(out.rows[3], { height: 18, hidden: false });
});

test('an UNSET row height (null, not hidden) is named, not guessed; the caveat about the default is on the result', () => {
  const out = read({ cols: {}, rows: { 4: { height: null, hidden: false } } }, 'A4:A4');
  assert.equal(out.rows[4].height, null);
  assert.match(out.rows[4].reason, /^UNOBSERVABLE:.*not set/);
  assert.ok(out.caveats.some(c => /never-set row .*default/.test(c)));
});

test('values the write schema has no word for are NAMED (unrepresentable), the entry stays: width over 255, height over 409, not a number, a non-boolean hidden', () => {
  const out = read({ cols: { A: { width: 300, hidden: false }, B: { width: '12', hidden: false }, C: { width: 10, hidden: 'yes' } }, rows: { 1: { height: 500, hidden: false }, 2: { height: -1, hidden: false } } }, 'A1:C2');
  assert.deepEqual(out.columns.A.unrepresentable, [{ field: 'width', engine: 300, reason: 'outside the writable (0, 255]' }]);
  assert.equal(out.columns.A.width, null);
  assert.deepEqual(out.columns.B.unrepresentable, [{ field: 'width', engine: '12', reason: 'not a positive number' }]);
  assert.match(out.columns.C.reason, /^UNOBSERVABLE:hidden-/);
  assert.deepEqual(out.rows[1].unrepresentable, [{ field: 'height', engine: 500, reason: 'outside the writable (0, 409]' }]);
  assert.deepEqual(out.rows[2].unrepresentable, [{ field: 'height', engine: -1, reason: 'not a positive number' }]);
});

test('a missing entry inside the range is a named gap for THAT column / row, the rest still read', () => {
  const out = read({ cols: { A: engine.col(9) }, rows: { 1: engine.row(15) } }, 'A1:B1');
  assert.deepEqual(out.columns.A, { width: 9, hidden: false });
  assert.deepEqual(out.columns.B, { reason: 'UNOBSERVABLE:not-read' });
});

test('whole-sheet gaps are strings on the result, never an empty list: not read, the driver\'s own reason, an unexpected shape', () => {
  assert.deepEqual(read(undefined, 'A1:B2'), { sheet: 'S', range: 'A1:B2', columns: null, rows: null, reason: 'UNOBSERVABLE:not-read' });
  assert.equal(read('UNOBSERVABLE:layout-truncated', 'A1:B2').reason, 'UNOBSERVABLE:layout-truncated');
  assert.equal(read({ cols: {} }, 'A1:B2').reason, 'UNOBSERVABLE:unexpected-shape');
  assert.equal(projectLayoutRead({ sheets: [{ name: 'S' }], layout: {} }, { sheet: 'S', range: 'A1' }).reason, 'UNOBSERVABLE:not-read');
});

test('a sheet that does not exist is SHEET_NOT_FOUND; a range over the ceiling is named BEFORE any value is read', () => {
  assert.deepEqual(projectLayoutRead(inv({ cols: {}, rows: {} }), { sheet: 'Nope', range: 'A1' }), { sheet: 'Nope', range: 'A1', error: 'SHEET_NOT_FOUND' });
  assert.equal(CEILING, 500);
  const wide = read({ cols: {}, rows: {} }, 'A1:XX1');
  assert.deepEqual([wide.columns, wide.rows], [null, null]);
  assert.match(wide.reason, /^RANGE_OVER_CEILING: 648 columns/);
  assert.match(read({ cols: {}, rows: {} }, 'A1:A501').reason, /^RANGE_OVER_CEILING: 1 columns x 501 rows/);
  const edge = read({ cols: {}, rows: {} }, 'A1:A500');
  assert.equal(edge.reason, undefined);
  assert.equal(Object.keys(edge.rows).length, 500);
});

test('layoutExtent: the driver reads from 1 to the range\'s last column / row, so the extent is the range\'s far corner ($ and case tolerated)', () => {
  assert.deepEqual(layoutExtent('B2:D5'), { maxCol: 4, maxRow: 5 });
  assert.deepEqual(layoutExtent('$c$3'), { maxCol: 3, maxRow: 3 });
  assert.equal(layoutExtent('nonsense'), null);
});

test('TOOL SURFACE: read_layout is ACCEPTED (sheet + range), the request asks the driver for the layout, and the one remaining non-ACCEPTED xlsx op (NOT_MEASURED) still refuses', () => {
  assert.equal(rt.OPERATIONS.xlsx.read_layout.status, rt.STATUS.ACCEPTED);
  assert.ok(/L-09/.test(rt.OPERATIONS.xlsx.read_layout.evidence));
  assert.deepEqual(rt.xlsxRequest('read_layout', { sheet: 'S', range: 'B2:D5' }), { windows: [], layout: [{ sheet: 'S', maxCol: 4, maxRow: 5 }] });
  assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_layout', sheet: 'S', range: 'B2:D5' }).ok, true);
  assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_layout', sheet: 'S' }).ok, false, 'range is required');
  assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_layout', sheet: 'S', range: 'B2:D5', width: 5 }).ok, false, 'no unknown field');
  // GH #140 U2: read_pivots moved PARTIAL -> ACCEPTED; xlsx has no PARTIAL operation left, so this
  // negative control uses read_print_rendering (NOT_MEASURED, the only remaining non-ACCEPTED xlsx op).
  assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_print_rendering' }).code, 'REFUSED_NOT_MEASURED');
});
