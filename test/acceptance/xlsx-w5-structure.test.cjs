'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const structure = require('../../src/adapters/xlsx/families/structure.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 5: insert / delete rows / columns through the target path on the fake transport. The retry is decided
 * from the ANCHOR's position (no token -- the source's X1 defect: an identical retry inserted / deleted again).
 */

const book = () => [
  { name: 'P', cells: {
    A1: { v: 'r1' }, A2: { v: 'r2' }, A3: { v: 'r3' }, A4: { v: 'r4' }, A5: { v: 'r5' },
    B1: { v: 1 }, B2: { v: 2 }, B3: { v: 3 }, B4: { v: 4 }, B5: { v: 5 },
    C1: { v: 15, f: '=SUM(B1:B5)' }, C2: { v: 10, f: '=B5*2' }, C5: { v: 'B5x', f: '="B5"&"x"' },
  }, merges: ['E1:F1'] },
  { name: 'Q', cells: { A1: { v: 5, f: '=P!B5' }, A2: { v: 1, f: '=P!B1' }, A3: { v: 'keep' } } },
];

const step = (intent, extra) => ({ id: 'm6-w5', steps: [{ intent, sheet: 'P', ...extra }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const P = store => store.durable.find(s => s.name === 'P').cells;
const Q = store => store.durable.find(s => s.name === 'Q').cells;
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('insert 2 rows at 3 (anchor r3): APPLIED -- cells moved, formulas on BOTH sheets re-pointed, the string literal untouched; retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const t = step('insert_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r3' } });
  const { result, m } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.preservationVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(P(store).A5.v, 'r3');
  assert.equal(P(store).A3, undefined, 'the new rows are blank');
  assert.equal(P(store).C1.f, '=SUM(B1:B7)');
  assert.equal(P(store).C2.f, '=B7*2');
  assert.equal(P(store).C7.f, '="B5"&"x"');
  assert.equal(Q(store).A1.f, '=P!B7');
  assert.equal(Q(store).A2.f, '=P!B1');
  const again = await run(store, t);
  assert.equal(again.m.writes, 0);
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('delete 1 row at 3 (surviving anchor r4): APPLIED; retry SATISFIED 0/0 -- an identical retry does NOT delete again', async () => {
  const store = createStore(book());
  const t = step('delete_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r4' } });
  const { result } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(P(store).A3.v, 'r4');
  assert.equal(P(store).C1.f, '=SUM(B1:B4)');
  assert.equal(Q(store).A1.f, '=P!B4');
  const again = await run(store, t);
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(P(store).A3.v, 'r4', 'still there: no second delete');
});

test('columns: insert 1 column at B (anchor: row 1 holds 1 in B) and delete it again through its surviving anchor', async () => {
  const store = createStore(book());
  const ins = step('insert_columns', { at: 'B', count: 1, anchor: { row: 1, value: 1 } });
  const r1 = await run(store, ins);
  assert.equal(r1.result.ok, true, JSON.stringify(r1.result.execution.primitiveVerification));
  assert.equal(P(store).C1.v, 1);
  assert.equal(P(store).D1.f, '=SUM(C1:C5)');
  assert.equal(Q(store).A1.f, '=P!C5');
  assert.equal(r1.driver.calls.dispatch, 1);
  const del = step('delete_columns', { at: 'B', count: 1, anchor: { row: 1, value: 1 } });
  const r2 = await run(store, del);
  assert.equal(r2.result.ok, true, JSON.stringify(r2.result.execution.primitiveVerification));
  assert.equal(P(store).B1.v, 1);
});

for (const [name, t, why, mutate] of [
  ['the anchor is neither before nor after', step('insert_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r5' } }), 'xlsx-structure-anchor-state-conflict'],
  ['the anchor is not unique', step('insert_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'dup' } }), 'xlsx-structure-anchor-not-unique', b => { b[0].cells.A3 = { v: 'dup' }; b[0].cells.A4 = { v: 'dup' }; }],
  ['the anchor is missing', step('insert_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'nope' } }), 'xlsx-structure-anchor-not-found'],
  ['a delete whose block a formula references (it would become #REF!)', step('delete_rows', { at: 5, count: 1, anchor: { column: 'A', value: 'r6' } }), 'xlsx-structure-delete-referenced-by-formulas', b => { b[0].cells.A6 = { v: 'r6' }; }],
  // M7 S5 R5: an insert inside a merge now GROWS it (xlsx-m7-s5.test.cjs); a delete of its top-left stays refused
  ['a delete removing a merge\'s top-left', step('delete_columns', { at: 'E', count: 1, anchor: { row: 1, value: 'f1' } }), 'xlsx-structure-merge-top-left-deleted', b => { b[0].cells.F1 = { v: 'f1' }; b[0].merges = ['E2:F2']; b[0].cells.E2 = { v: 'm' }; }],
  ['a missing sheet', { id: 'x', steps: [{ intent: 'insert_rows', sheet: 'Nope', at: 1, count: 1, anchor: { column: 'A', value: 'x' } }] }, 'xlsx-sheet-not-found'],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const b = book();
    if (mutate) mutate(b);
    const store = createStore(b);
    const { result, driver } = await run(store, t);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

// M7 S4 R1: a single-cell absolute defined name no longer refuses -- it follows (test/acceptance/xlsx-m7-s4-names.test.cjs)
test('refused: a dependent the probe has not measured (a conditional format reaching the sheet)', async () => {
  const b = book();
  b[0].rules = [{ kind: 'conditional', range: 'E2:E5', formula: '=$B2>2' }];
  const withRule = await run(createStore(b), step('insert_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r3' } }));
  assert.equal(reason(withRule.result), 'xlsx-structure-dependent-not-measured');
  const unproven = await run(createStore(book()), step('insert_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r3' } }), { dependencyBuild: false });
  assert.equal(reason(unproven.result), 'xlsx-structure-dependency-observation-incomplete');
});

test('RED: ignored / one row short / references not followed -> L1 FAIL; another sheet shifted -> L3 FAIL; nothing persisted', async () => {
  const t = step('insert_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r3' } });
  for (const fault of ['ignore', 'partial', 'no-reference-follow']) {
    const store = createStore(book());
    const { result, driver } = await run(store, t, { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(driver.calls.forceSave, 0, fault);
    assert.deepEqual(store.durable, book(), fault);
  }
  const nr = await run(createStore(book()), t, { fault: 'no-reference-follow' });
  assert.equal(nr.result.execution.primitiveVerification[0].verdict, 'FAIL');
  const other = await run(createStore(book()), t, { fault: 'other-sheet-moves' });
  assert.equal(other.result.ok, false);
  assert.equal(other.result.execution.preservationVerification[0].verdict, 'FAIL');
  const side = await run(createStore(book()), t, { fault: 'side-effect' });
  assert.equal(side.result.ok, false);
});

test('planner: shapes, the anchor, the only step of its task', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse(step('insert_rows', { at: 0, count: 1, anchor: { column: 'A', value: 'x' } }), /at must be a row number/);
  refuse(step('insert_columns', { at: 2, count: 1, anchor: { row: 1, value: 'x' } }), /at must be a column letter/);
  refuse(step('insert_rows', { at: 3, count: 0, anchor: { column: 'A', value: 'x' } }), /count must be an integer/);
  refuse(step('insert_rows', { at: 3, count: 1 }), /anchor must be/);
  refuse(step('insert_rows', { at: 3, count: 1, anchor: { row: 1, value: 'x' } }), /anchor must be/);
  refuse(step('insert_rows', { at: 3, count: 1, anchor: { column: 'A', value: '' } }), /anchor.value/);
  refuse(step('insert_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'x' }, retryToken: 't' }), /does not support field/);
  refuse({ id: 'x', steps: [{ intent: 'insert_rows', sheet: 'P', at: 3, count: 1, anchor: { column: 'A', value: 'r3' } }, { intent: 'write_range', sheet: 'Q', range: 'Z9', values: [['v']] }] }, /only step of its task/);
});

test('the command body: exact-name sheet, the entire-span public Insert / Delete', () => {
  for (const [intent, extra, want] of [
    ['insert_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'x' } }, ['A3:XFD4', 'Insert', 'down']],
    ['delete_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'x' } }, ['A3:XFD3', 'Delete', 'up']],
    ['insert_columns', { at: 'B', count: 2, anchor: { row: 1, value: 'x' } }, ['B1:C1048576', 'Insert', 'right']],
    ['delete_columns', { at: 'B', count: 1, anchor: { row: 1, value: 'x' } }, ['B1:B1048576', 'Delete', 'left']],
  ]) {
    const cmd = structure.lower(planOperations(step(intent, extra)).operations[0]);
    const calls = [];
    const ws = { GetName: () => 'P', GetRange: a => ({ Insert: d => calls.push([a, 'Insert', d]), Delete: d => calls.push([a, 'Delete', d]) }) };
    assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ws }), { ok: true });
    assert.deepEqual(calls, [want], intent);
  }
});
