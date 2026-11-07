'use strict';

/**
 * GH #298 D: create_table (families/tables-write.cjs) -- the family's own decisions on hand-built fresh observations
 * (the observation shape is the adapter's: cells keyed by inventory.key, `tableReads` = read_tables' axis, `objects` =
 * the per-sheet objects read with the pivots), and the product surface through the real validateWriteRequest. The
 * engine path itself (FormatAsTable) is proven LIVE: test/live/release-gh298-d-create-table-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const fam = require('../../src/adapters/xlsx/families/tables-write.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { key } = require('../../src/adapters/xlsx/inventory.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const S = 'Data';
const RANGE = 'A1:B3';
const VALUES = { A1: 'Name', B1: 'Qty', A2: 'x', B2: 1, A3: 'y', B3: 2, C1: 'side', A5: 'other' };

function obs({ values = VALUES, formulas = {}, merges = {}, tables = [], pivots = [], tableRead, objects } = {}) {
  const cells = {};
  for (const [a, v] of Object.entries(values)) cells[key(S, a)] = { v, f: formulas[a] || '', merge: merges[a] || a };
  for (const a of ['A1', 'B1', 'A2', 'B2', 'A3', 'B3']) if (!cells[key(S, a)]) cells[key(S, a)] = { v: null, f: '', merge: merges[a] || a };
  return { format: 'xlsx', complete: true, sheets: [{ name: S }], cells,
    tableReads: { [S]: tableRead !== undefined ? tableRead : { tables } }, objects: objects !== undefined ? objects : { [S]: { pivots } } };
}
const created = (extra = {}) => ({ name: 'Table1', ref: RANGE, headerRow: true, totalsRow: false, columns: ['Name', 'Qty'], style: null, autoFilter: true, ...extra });
const opOf = (step = {}) => { const p = fam.plan({ intent: 'create_table', sheet: S, range: RANGE, ...step }, 0); assert.ok(p.ok, p.detail); return p.operation; };
const reasonOf = r => r.evidence.reason;

test('plan: a header row and one data row at least, the cell ceiling, no unknown field', () => {
  assert.equal(opOf().selector.range, RANGE);
  assert.match(fam.plan({ intent: 'create_table', sheet: S, range: 'A1:C1' }, 0).detail, /header row and at least one data row/);
  assert.match(fam.plan({ intent: 'create_table', sheet: S, range: 'A1:B501' }, 0).detail, /over the 1000 ceiling/);
  assert.match(fam.plan({ intent: 'create_table', sheet: S, range: RANGE, name: 'T' }, 0).detail, /does not support field\(s\): name/);
  assert.match(fam.plan({ intent: 'create_table', sheet: S, range: 'nope' }, 0).detail, /plain A1 range/);
});

test('classify: a clean range is APPLY, its header texts named', () => {
  const r = fam.classify(obs(), opOf());
  assert.equal(r.state, 'APPLY');
  assert.deepEqual(r.evidence.columns, ['Name', 'Qty']);
});

test('classify: the terminal state on a fresh read is SATISFIED (retry, 0 write); a different table on the range is refused', () => {
  assert.equal(fam.classify(obs({ tables: [created()] }), opOf()).state, 'SATISFIED');
  const other = fam.classify(obs({ tables: [created({ columns: ['A', 'B'] })] }), opOf());
  assert.equal(other.state, 'FAIL_CLOSED');
  assert.equal(reasonOf(other), 'xlsx-table-range-holds-a-different-table');
});

test('classify: every conflict the engine would store is refused, named, before the write', () => {
  const cases = [
    [obs({ tables: [{ ...created(), name: 'T0', ref: 'B2:C4' }] }), 'xlsx-table-overlaps-table'],
    [obs({ merges: { A2: 'A2:B2', B2: 'A2:B2' } }), 'xlsx-table-over-merge'],
    [obs({ pivots: [{ name: 'P1', table: 'B3:D6' }] }), 'xlsx-table-overlaps-pivot'],
    [obs({ values: { ...VALUES, B1: '' } }), 'xlsx-table-header-blank'],
    [obs({ values: { ...VALUES, B1: '   ' } }), 'xlsx-table-header-blank'],
    [obs({ formulas: { B1: '="Qty"' } }), 'xlsx-table-header-formula-unmeasured'],
    [obs({ values: { ...VALUES, B1: 'NAME' } }), 'xlsx-table-header-duplicate'],
  ];
  for (const [p, reason] of cases) {
    const r = fam.classify(p, opOf());
    assert.equal(r.state, 'FAIL_CLOSED', reason);
    assert.equal(reasonOf(r), reason);
  }
  // NEG: a table, a merge and a pivot NEXT TO the range are not conflicts
  const near = obs({ tables: [{ ...created(), name: 'T0', ref: 'D1:E3' }], merges: { A5: 'A5:B5' }, pivots: [{ name: 'P1', table: 'D5:F8' }] });
  assert.equal(fam.classify(near, opOf()).state, 'APPLY');
});

test('classify: what cannot be read is FAIL_CLOSED, never APPLY', () => {
  const cases = [
    [obs({ tableRead: 'UNOBSERVABLE:capability-missing' }), 'xlsx-tables-unobservable'],
    [obs({ tableRead: null }), 'xlsx-tables-unobservable'],
    [obs({ objects: { [S]: { pivots: 'UNOBSERVABLE:threw' } } }), 'xlsx-pivots-unobservable'],
    [obs({ pivots: [{ name: 'P1', table: 'UNOBSERVABLE:threw' }] }), 'xlsx-pivots-unobservable'],
    [{ ...obs(), complete: false }, 'xlsx-observation-unreadable-getters'],
  ];
  for (const [p, reason] of cases) assert.equal(reasonOf(fam.classify(p, opOf())), reason);
  const unread = obs(); delete unread.cells[key(S, 'B3')];
  assert.equal(reasonOf(fam.classify(unread, opOf())), 'xlsx-target-cell-unread');
});

test('lower: one callCommand, an exact-name sheet check, then FormatAsTable on the range', () => {
  const l = fam.lower(opOf());
  assert.equal(l.kind, 'callCommand');
  assert.match(l.body, /ws\.GetName\(\) !== name/);
  assert.match(l.body, /ws\.FormatAsTable\("A1:B3"\);/);
});

test('verify: L1 = the table on the fresh read; L3 = every other table and every cell unchanged', () => {
  const op = opOf();
  const before = { primary: obs({ tables: [{ ...created(), name: 'T0', ref: 'D1:E3' }] }) };
  const after = { primary: obs({ tables: [{ ...created(), name: 'T0', ref: 'D1:E3' }, created()] }) };
  assert.equal(fam.verifyPrimitive(op, before, after).ok, true);
  assert.equal(fam.verifyPrimitive(op, before, before).ok, false, 'no table -> L1 FAIL');
  assert.equal(fam.verifyPreservation(op, before, after).ok, true);
  const otherChanged = { primary: obs({ tables: [{ ...created(), name: 'T0', ref: 'D1:E4' }, created()] }) };
  assert.equal(fam.verifyPreservation(op, before, otherChanged).ok, false, 'another table changed -> L3 FAIL');
  const cellChanged = { primary: obs({ values: { ...VALUES, A2: 'changed' }, tables: [{ ...created(), name: 'T0', ref: 'D1:E3' }, created()] }) };
  assert.equal(fam.verifyPreservation(op, before, cellChanged).ok, false, 'a cell value changed -> L3 FAIL');
  assert.equal(fam.wholeTaskChecks(op, after)[0].ok, true);
});

test('liveScope: the range window, read_tables axis and the pivot objects of the sheet', () => {
  const s = liveScope({ steps: [{ intent: 'create_table', sheet: S, range: RANGE }] });
  assert.deepEqual(s.windows, [{ sheet: S, range: RANGE }]);
  assert.deepEqual(s.tableSheets, [S]);
  assert.deepEqual(s.objectSheets, [S]);
});

test('product surface: create_table is accepted by the real validateWriteRequest; an unsupported shape is refused before dispatch', () => {
  const ok = W.validateWriteRequest(req([{ intent: 'create_table', sheet: S, range: RANGE }]));
  assert.equal(ok.ok, true, JSON.stringify(ok));
  const bad = W.validateWriteRequest(req([{ intent: 'create_table', sheet: S, range: RANGE, style: 'TableStyleMedium2' }]));
  assert.equal(W.validateWriteRequest(req([{ intent: 'create_table', sheet: S, range: 'A1:C1' }])).ok, false, 'one row: refused by the family plan before dispatch');
  assert.equal(bad.ok, false);
});
