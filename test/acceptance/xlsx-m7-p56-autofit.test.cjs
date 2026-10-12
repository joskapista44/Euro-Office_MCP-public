'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const layout = require('../../src/adapters/xlsx/families/layout.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 backlog X-L01 (row autofit) / X-L02 (column autofit) --. This REPLACES the owner's
 * prior #ref "AutoFit is not supported" refusal with owner option B (owner channel,
 * 2026-10-05, docs/M7-AUTOFIT-CONTRACT-DRAFT.md): the end state still cannot be read without
 * re-applying it, so classify() NEVER treats autofit as already-satisfied -- it always re-dispatches
 * the SAME command. L1 is deliberately weak: only that the row / column now reads a numeric size and
 * is not hidden (there is no reference to compare a "correct fit" against; a reference write was the
 * rejected option C). Idempotency / non-cumulative-ness across a retry is proven externally, by the
 * LIVE harness calling the REAL driver twice -- the fake here always returns the SAME fitted value, so
 * it cannot and does not claim to test that.
 */

// The merge sits on rows 10-11, away from the 1:2 range the other row tests use -- so "a merge in the
// TARGET range" and "a merge outside it" stay genuinely distinct fixtures, not an accidental overlap.
const book = () => [
  { name: 'Data', cells: { A1: { v: 'Name' }, B1: { v: 'Qty' }, A2: { v: 'x' }, B2: { v: 5 }, D4: { v: 'end' }, F10: { v: 'merged-header' } },
    layout: { cols: { C: { hidden: true } }, rows: { '3': { height: 22 } } }, merges: ['F10:G11'] },
];

const step = extra => ({ intent: 'layout_range', sheet: 'Data', ...extra });
const task = extra => ({ id: 'm7-p56-autofit', steps: [step(extra)] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const L = store => store.durable.find(s => s.name === 'Data').layout;
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

for (const [name, spec, , items, read] of [
  ['row', { rows: '1:2', autofit: true }, 'height', ['1', '2'], l => [l.rows['1'].height, l.rows['2'].height]],
  ['column', { columns: 'A:B', autofit: true }, 'width', ['A', 'B'], l => [l.cols.A.width, l.cols.B.width]],
]) {
  test(`${name} autofit: FIRST APPLIED (1 command, L1 + L3, barrier), reopened RETRY *also* APPLIED (owner option B: never SATISFIED)`, async () => {
    const store = createStore(book());
    const { result, m, driver } = await run(store, task(spec));
    assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
    assert.equal(m.writes, 1);
    assert.equal(m.commands, 1);
    assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
    assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
    assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
    const sizes = read(L(store));
    assert.ok(sizes.every(v => typeof v === 'number'), JSON.stringify(sizes));
    assert.deepEqual(result.execution.primitiveVerification[0].evidence.evidence.sizes, Object.fromEntries(items.map((it, idx) => [it, sizes[idx]])), 'verifyPrimitive echoes back the SAME sizes it just read, for a caller (the LIVE harness) comparing a retry');
    assert.deepEqual(driver.reads[0].windows, [], 'no cell window: the default used-range read already carries merge info');
    assert.ok(Array.isArray(driver.reads[0].layout), 'the layout extent is read');
    const again = await run(store, task(spec));
    assert.equal(again.result.ok, true);
    assert.equal(again.m.writes, 1, 'owner option B: a retry is NEVER 0-command');
    assert.equal(again.m.commands, 1);
    assert.equal(again.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
    assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
    assert.deepEqual(read(L(store)), sizes, 'the fake always fits to the same value -- idempotent across this one call pair (true idempotency / non-cumulativeness is the LIVE harness\'s own proof, against the real engine)');
  });
}

test('NEG: a hidden row / column in the target range is refused before dispatch (the existing hidden check, shared with width / height)', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task({ columns: 'B:C', autofit: true }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-layout-target-hidden');
  assert.deepEqual(store.durable, book());
});

for (const spec of [{ columns: 'F:G', autofit: true }, { rows: '10:11', autofit: true }]) {
  test(`NEG: a merge anywhere among the target ${'columns' in spec ? 'columns' : 'rows'}' used cells is refused before dispatch`, async () => {
    const store = createStore(book());
    const { result, driver } = await run(store, task(spec));
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), 'xlsx-layout-autofit-merged-range');
    assert.deepEqual(store.durable, book());
  });
}

test('a merge OUTSIDE the target range does not block it', async () => {
  const store = createStore(book());
  const { result } = await run(store, task({ rows: '1:2', autofit: true }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
});

test('RED: the size stays unread after the write -> L1 FAIL, nothing persisted', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task({ rows: '1:2', autofit: true }), { fault: 'autofit-no-size' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
  assert.deepEqual(store.durable, book());
});

test('RED: a side effect -- another row, a cell -> L3 FAIL', async () => {
  const side = await run(createStore(book()), task({ rows: '1:2', autofit: true }), { fault: 'side-layout', sideColumn: 'D' });
  assert.equal(side.result.ok, false);
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
  const cell = await run(createStore(book()), task({ rows: '1:2', autofit: true }), { fault: 'side-effect' });
  assert.equal(cell.result.ok, false);
  assert.equal(cell.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('planner: autofit shapes', () => {
  const refuse = (s, re) => { const p = planOperations({ id: 'x', steps: [step(s)] }); assert.equal(p.ok, false, JSON.stringify(s)); assert.match(p.detail, re); };
  refuse({ columns: 'A', autofit: false }, /autofit must be true/);
  refuse({ columns: 'A', autofit: 'yes' }, /autofit must be true/);
  refuse({ columns: 'A', autofit: true, width: 9 }, /exactly one of width, height, hidden, autofit/);
  refuse({ rows: '1', columns: 'A', autofit: true }, /exactly one of columns/);
  const ok = planOperations(task({ rows: '2:5', autofit: true }));
  assert.equal(ok.ok, true, JSON.stringify(ok));
  assert.equal(ok.operations[0].prop, 'autofit');
  assert.equal(ok.operations[0].value, true);
  assert.deepEqual(ok.operations[0].footprint, { rows: [2, 5] });
  const colFootprint = planOperations(task({ columns: 'B:C', autofit: true }));
  assert.deepEqual(colFootprint.operations[0].footprint, { columns: [2, 3], rows: '*' }, 'a column autofit re-fits its own width, the same side effect a manual SetColumnWidth claims');
});

test('the command bodies compile: per-item SetRowHeight(current)+AutoFit(true,false), or SetColumnWidth(current +/- 1)+AutoFit(false,true), looped inside one command', () => {
  const calls = [];
  // a DIFFERENT height per row (A2 != A3) -- a command that hard-codes one literal value instead of
  // reading each row's OWN current height would pass a mock returning one constant unnoticed.
  const heightOf = { A2: 15, A3: 42 };
  const widthOf = { B1: 8.43, C1: 254.5 };
  const entire = a => ({
    GetColumnWidth: () => widthOf[a],
    SetColumnWidth: v => calls.push(['SetColumnWidth', a, v]),
    GetRowHeight: () => heightOf[a],
    SetRowHeight: v => calls.push(['SetRowHeight', a, v]),
    AutoFit: (...args) => calls.push(['AutoFit', a, ...args]),
  });
  const ws = { GetName: () => 'Data', GetRange: a => ({ GetEntireRow: () => entire(a), GetEntireColumn: () => entire(a) }) };
  const rowCmd = layout.lower(planOperations(task({ rows: '2:3', autofit: true })).operations[0]);
  assert.deepEqual(new Function('Api', rowCmd.body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(calls, [['SetRowHeight', 'A2', 15], ['AutoFit', 'A2', true, false], ['SetRowHeight', 'A3', 42], ['AutoFit', 'A3', true, false]]);
  calls.length = 0;
  const colCmd = layout.lower(planOperations(task({ columns: 'B:C', autofit: true })).operations[0]);
  assert.deepEqual(new Function('Api', colCmd.body)({ GetSheet: () => ws }), { ok: true });
  // one width step away first (up; down at the 255 ceiling), so the AutoFit always writes (p56 LIVE: a retry on an
  // already-fitted column was a no-op dispatch the save contract fails)
  assert.deepEqual(calls, [['SetColumnWidth', 'B1', 9.43], ['AutoFit', 'B1', false, true], ['SetColumnWidth', 'C1', 253.5], ['AutoFit', 'C1', false, true]]);
});

test('classify(): autofit NEVER reads as SATISFIED, even if an axis entry coincidentally carried its own numeric "autofit" field matching op.value -- owner option B must not depend on holds() missing a match', () => {
  const op = planOperations(task({ rows: '2', autofit: true })).operations[0];
  // no real or fake read ever produces an 'autofit' axis field today (width / height / hidden only) --
  // this constructs the coincidence directly, because that absence is exactly what would make the
  // dedicated branch below look removable without any test noticing.
  const primary = { format: 'xlsx', complete: true, sheets: [{ name: 'Data' }], layout: { Data: { cols: {}, rows: { '2': { height: 15, hidden: false, autofit: 1 } } } }, cells: {} };
  assert.equal(layout.classify(primary, op).state, 'APPLY');
});

test('the command reports a sheet-identity mismatch, not a throw', () => {
  const ws = { GetName: () => 'Other' };
  const cmd = layout.lower(planOperations(task({ rows: '2', autofit: true })).operations[0]);
  assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ws }), { ok: false, stage: 'sheet-identity' });
});
