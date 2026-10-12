'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const layout = require('../../src/adapters/xlsx/families/layout.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 4: layout_range (width / height / hide / show) through the target path on the fake transport. The
 * fake follows W4b: per column / row PUBLIC getters, a hidden column reads width 0, a hidden row height null,
 * an unset row height null. AutoFit (M7 X-L01/X-L02, owner option B) has its own test file,
 * xlsx-m7-p56-autofit.test.cjs.
 */

const book = () => [
  { name: 'Data', cells: { A1: { v: 'Name' }, B1: { v: 'Qty' }, A2: { v: 'x' }, B2: { v: 5 }, D4: { v: 'end' } },
    layout: { cols: { C: { hidden: true } }, rows: { '3': { height: 22 } } } },
  { name: 'Other', cells: { A1: { v: 'keep me' } }, layout: { cols: { A: { width: 12 } } } },
];

const step = extra => ({ intent: 'layout_range', sheet: 'Data', ...extra });
const task = extra => ({ id: 'm6-w4-layout', steps: [step(extra)] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const L = store => store.durable.find(s => s.name === 'Data').layout;
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

for (const [name, spec, check] of [
  ['column width', { columns: 'A:B', width: 20 }, l => assert.deepEqual([l.cols.A.width, l.cols.B.width], [20, 20])],
  ['row height', { rows: '1:2', height: 30 }, l => assert.deepEqual([l.rows['1'].height, l.rows['2'].height], [30, 30])],
  // X-L04 (S18 LIVE m7-s18-probe2): the row axis is MEASURED to round-trip within
  // TOL (0.05) for a fractional height -- 15.37 is a maintainer's own measured value (set -> R1 15.37,
  // reopen 15.35, both within 0.05). The column (width) axis stays a SEPARATE, forbidden item
  // (an unmeasured set-side px-quantization rule, not a tolerance problem) -- deliberately NOT
  // touched here.
  ['fractional row height (X-L04, row axis only)', { rows: '1:2', height: 15.37 }, l => assert.deepEqual([l.rows['1'].height, l.rows['2'].height], [15.37, 15.37])],
  ['hide columns', { columns: 'E:F', hidden: true }, l => assert.deepEqual([l.cols.E.hidden, l.cols.F.hidden], [true, true])],
  ['show a hidden column', { columns: 'C', hidden: false }, l => assert.equal(l.cols.C.hidden, false)],
  ['hide rows', { rows: '5:6', hidden: true }, l => assert.deepEqual([l.rows['5'].hidden, l.rows['6'].hidden], [true, true])],
]) {
  test(`${name}: FIRST APPLIED (1 command, L1 + L3, barrier), reopened retry SATISFIED 0 write / 0 command`, async () => {
    const store = createStore(book());
    const { result, m, driver } = await run(store, task(spec));
    assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
    assert.equal(m.writes, 1);
    assert.equal(m.commands, 1);
    assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
    assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
    assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
    check(L(store));
    assert.deepEqual(driver.reads[0].windows, [], 'no cell window: a whole column is a million cells');
    assert.ok(Array.isArray(driver.reads[0].layout), 'the layout extent is read');
    const again = await run(store, task(spec));
    assert.equal(again.result.ok, true);
    assert.equal(again.m.writes, 0);
    assert.equal(again.m.commands, 0);
    assert.equal(again.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
    assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  });
}

test('a target beyond the used range is read (the extent), e.g. column H', async () => {
  const store = createStore(book());
  const { result } = await run(store, task({ columns: 'H', width: 11 }));
  assert.equal(result.ok, true);
  assert.equal(L(store).cols.H.width, 11);
});

for (const [name, t, script, why] of [
  ['a width on a hidden column (its width reads 0: unverifiable)', task({ columns: 'B:C', width: 15 }), {}, 'xlsx-layout-target-hidden'],
  ['a missing sheet', { id: 'x', steps: [{ intent: 'layout_range', sheet: 'Nope', columns: 'A', width: 9 }] }, {}, 'xlsx-sheet-not-found'],
  ['an unobservable layout', task({ columns: 'A', width: 9 }), { layoutUnobservable: true }, 'xlsx-layout-unobservable'],
  ['a truncated read', task({ columns: 'A', width: 9 }), { readTruncated: true }, 'xlsx-observation-truncated'],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const store = createStore(book());
    const { result, driver } = await run(store, t, script);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
    assert.deepEqual(store.durable, book());
  });
}

test('RED: ignored / partial / drifted -> L1 FAIL, nothing persisted', async () => {
  for (const [spec, fault] of [[{ columns: 'A:B', width: 20 }, 'ignore'], [{ columns: 'A:B', width: 20 }, 'partial'], [{ rows: '1:2', hidden: true }, 'partial'], [{ rows: '1', height: 30 }, 'layout-drifts']]) {
    const store = createStore(book());
    const { result, driver } = await run(store, task(spec), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
    assert.deepEqual(store.durable, book(), fault);
  }
});

// X-L04 (S18 LIVE m7-s18-probe2): the EXACT measured row-axis drift (15.37 set,
// 15.35 after reopen -- 0.02 off) is WITHIN TOL (0.05) and must still APPLY, not FAIL -- the
// mirror image of the RED test above, which proves TOL is exceeded at +0.5.
test('GREEN: a row height drift within TOL (0.02, zero\'s own measured 15.37 -> 15.35) still holds -- L1 PASS', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task({ rows: '1:2', height: 15.37 }), { fault: 'layout-drifts-within-tol' });
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.writes, 1);
  assert.deepEqual([L(store).rows['1'].height, L(store).rows['2'].height], [15.35, 15.35]);
});

test('RED: a side effect -- another column\'s width, a cell -> L3 FAIL', async () => {
  // W4e LIVE: an auto row reading its single-line default (14.25) after the write is the SAME state -- not an L3 change;
  // any other height on a row that was unset IS one
  const mat = await run(createStore(book()), task({ columns: 'A', width: 20 }), { fault: 'auto-row-materialized', sideRow: '2' });
  assert.equal(mat.result.ok, true, JSON.stringify(mat.result.execution.preservationVerification));
  const fixedSide = await run(createStore(book()), task({ columns: 'A', width: 20 }), { fault: 'side-row-height', sideRow: '2' });
  assert.equal(fixedSide.result.execution.preservationVerification[0].verdict, 'FAIL');
  const hiddenSide = await run(createStore(book()), task({ columns: 'A', width: 20 }), { fault: 'auto-row-hidden', sideRow: '2' });
  assert.equal(hiddenSide.result.execution.preservationVerification[0].verdict, 'FAIL', 'a row that got HIDDEN is a change even at its default height');
  const side = await run(createStore(book()), task({ columns: 'A', width: 20 }), { fault: 'side-layout', sideColumn: 'D' });
  assert.equal(side.result.ok, false);
  const pv = side.result.execution.preservationVerification[0];
  assert.equal(pv.verdict, 'FAIL');
  assert.deepEqual(pv.evidence.evidence.layoutChanged.map(x => x.column), ['D']);
  const cell = await run(createStore(book()), task({ columns: 'A', width: 20 }), { fault: 'side-effect' });
  assert.equal(cell.result.ok, false);
  assert.equal(cell.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('planner: shapes, one property, the ceiling', () => {
  const refuse = (s, re) => { const p = planOperations({ id: 'x', steps: [step(s)] }); assert.equal(p.ok, false, JSON.stringify(s)); assert.match(p.detail, re); };
  refuse({ columns: 'A', autofit: false }, /autofit must be true/);
  refuse({ columns: 'A', rows: '1', width: 9 }, /exactly one of columns/);
  refuse({ width: 9 }, /exactly one of columns/);
  refuse({ columns: 'A', width: 9, hidden: true }, /exactly one of width, height, hidden/);
  refuse({ rows: '1', width: 9 }, /width applies to columns/);
  refuse({ columns: 'A', height: 9 }, /height applies to rows/);
  refuse({ columns: 'A', hidden: 'yes' }, /hidden must be true or false/);
  refuse({ columns: 'A', width: 0 }, /width must be a number/);
  refuse({ columns: 'A', width: 300 }, /width must be a number/);
  refuse({ columns: 'a:b', width: 9 }, /plain column letters/);
  refuse({ columns: 'C:A', width: 9 }, /plain column letters/);
  refuse({ rows: '0', hidden: true }, /plain row numbers/);
  refuse({ rows: '1:600', hidden: true }, /over the 500 ceiling/);
  refuse({ columns: 'A', width: 9, colour: 1 }, /does not support field/);
  const both = planOperations({ id: 'x', steps: [step({ columns: 'B', width: 9 }), { intent: 'write_range', sheet: 'Data', range: 'B7', values: [['v']] }] });
  // M7 S3: a width and a write on the same column touch disjoint axes -> planned, the write after a barrier
  assert.equal(both.ok, true, 'a width and a write on the same column: disjoint axes (M7 S3)');
  assert.deepEqual(both.operations[1].dependsOn, [both.operations[0].id]);
  const twice = planOperations({ id: 'x', steps: [step({ columns: 'B', width: 9 }), step({ columns: 'B:C', hidden: true })] });
  assert.equal(twice.ok, false, 'two column steps on one column share the axis');
  assert.match(twice.detail, /overlapping-writes-refused: .* \(same axis: columns\)/);
});

test('the command bodies compile: the public entire-column / entire-row setter, exact-name sheet', () => {
  for (const [spec, want] of [[{ columns: 'B:C', width: 20 }, ['B1:C1', 'col', 'SetColumnWidth', 20]], [{ rows: '2:3', height: 30 }, ['A2:A3', 'row', 'SetRowHeight', 30]], [{ rows: '4', hidden: true }, ['A4:A4', 'row', 'SetHidden', true]]]) {
    const cmd = layout.lower(planOperations(task(spec)).operations[0]);
    const calls = [];
    const entire = (a, kind) => ({ SetColumnWidth: v => calls.push([a, kind, 'SetColumnWidth', v]), SetRowHeight: v => calls.push([a, kind, 'SetRowHeight', v]), SetHidden: v => calls.push([a, kind, 'SetHidden', v]) });
    const ws = { GetName: () => 'Data', GetRange: a => ({ GetEntireColumn: () => entire(a, 'col'), GetEntireRow: () => entire(a, 'row') }) };
    assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ws }), { ok: true });
    assert.deepEqual(calls, [want]);
    assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ({ GetName: () => 'data' }) }), { ok: false, stage: 'sheet-identity' });
  }
});

test('the driver layout read compiles, stays read-only and reads every sheet over max(used, extent)', () => {
  const q = readQuery([], { layout: [{ sheet: 'Data', maxCol: 3, maxRow: 1 }] });
  const log = [];
  const entire = (ws, a, kind) => ({ GetColumnWidth: () => (log.push([ws, 'w', a]), 8.43), GetRowHeight: () => (log.push([ws, 'h', a]), null), GetHidden: () => false, _kind: kind });
  const rng = (ws, a) => ({ GetAddress: () => a, ForEach: () => {}, GetEntireColumn: () => entire(ws, a, 'col'), GetEntireRow: () => entire(ws, a, 'row') });
  const sheet = (name, used) => ({ GetName: () => name, GetIndex: () => 0, GetUsedRange: () => (used ? rng(name, used) : null), GetRange: a => rng(name, a) });
  const out = new Function('Api', q)({ GetSheets: () => [sheet('Data', 'A1:B2'), sheet('Empty', null)], GetSheet: () => null, WorksheetFunction: { COUNT: () => 0 } });
  assert.equal(out.ok, true);
  assert.deepEqual(Object.keys(out.layout.Data.cols), ['A', 'B', 'C'], 'extent 3 > used 2');
  assert.deepEqual(Object.keys(out.layout.Data.rows), ['1', '2'], 'used 2 > extent 1');
  assert.deepEqual(out.layout.Empty, { cols: {}, rows: {} });
  assert.deepEqual(out.layout.Data.rows['1'], { height: null, hidden: false });
  assert.equal(new Function('Api', readQuery())({ GetSheets: () => [], GetSheet: () => null, WorksheetFunction: { COUNT: () => 0 } }).layout, undefined, 'no layout unless asked');
  const huge = new Function('Api', readQuery([], { layout: [{ sheet: 'Data', maxCol: 16384, maxRow: 1 }] }))({ GetSheets: () => [sheet('Data', 'A1')], GetSheet: () => null, WorksheetFunction: { COUNT: () => 0 } });
  assert.equal(huge.layout.Data, 'UNOBSERVABLE:layout-truncated');
});

test('fail closed when a target column\'s hidden state is not a boolean (a getter answered null)', () => {
  const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
  const op = planOperations(task({ columns: 'A', width: 9 })).operations[0];
  const primary = normalizeInventory({ sheets: [{ name: 'Data', index: 0, used: 'A1' }], cells: [{ sheet: 'Data', a: 'A1', v: 'x', f: 'x', count: 0, merge: 'A1' }],
    truncated: false, unsupported: [], layout: { Data: { cols: { A: { width: 8.43, hidden: null } }, rows: { 1: { height: null, hidden: false } } } } });
  assert.deepEqual(layout.classify(primary, op), { state: 'FAIL_CLOSED', evidence: { ok: false, reason: 'xlsx-layout-target-unread', item: 'A' } });
});
