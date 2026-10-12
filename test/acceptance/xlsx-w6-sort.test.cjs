'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const sort = require('../../src/adapters/xlsx/families/sort.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M6 wave 6: sort_range on the fake -- row integrity is part of the proof (the source checked the key only). */

const book = () => [{ name: 'D', cells: {
  A1: { v: 'name' }, B1: { v: 'n' }, A2: { v: 'delta' }, B2: { v: 4 }, A3: { v: 'alpha' }, B3: { v: 1 }, A4: { v: 'charlie' }, B4: { v: 3 }, A5: { v: 'bravo' }, B5: { v: 2 },
  E1: { v: 10, f: '=SUM(B2:B5)' },
} }, { name: 'O', cells: { A1: { v: 'keep' } } }];

const task = (extra = {}) => ({ id: 'm6-w6s', steps: [{ intent: 'sort_range', sheet: 'D', range: 'A1:B5', key: 'A', order: 'asc', hasHeaders: true, ...extra }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const D = store => store.durable.find(s => s.name === 'D').cells;
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('text key asc: APPLIED, rows moved whole, header kept, the SUM outside by its text; retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task());
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.deepEqual(['A2', 'A3', 'A4', 'A5'].map(a => D(store)[a].v), ['alpha', 'bravo', 'charlie', 'delta']);
  assert.deepEqual(['B2', 'B3', 'B4', 'B5'].map(a => D(store)[a].v), [1, 2, 3, 4], 'rows moved with their keys');
  assert.equal(D(store).A1.v, 'name');
  const again = await run(store, task());
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('numeric key desc: APPLIED; retry SATISFIED', async () => {
  const store = createStore(book());
  const { result } = await run(store, task({ key: 'B', order: 'desc' }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual(['B2', 'B3', 'B4', 'B5'].map(a => D(store)[a].v), [4, 3, 2, 1]);
  assert.equal((await run(store, task({ key: 'B', order: 'desc' }))).m.commands, 0);
});

for (const [name, mutate, why] of [
  ['a blank key (placement unmeasured, audit D03)', b => { delete b[0].cells.A3; }, 'xlsx-sort-blank-key-unmeasured'],
  ['numbers mixed with text in the key', b => { b[0].cells.A3 = { v: 5 }; }, 'xlsx-sort-mixed-key-types-unmeasured'],
  ['case-only differences (the collation would decide)', b => { b[0].cells.A4 = { v: 'Beta' }; }, 'xlsx-sort-collation-unmeasured'],
  ['number-like text keys (10 vs 9)', b => { b[0].cells.A2 = { v: '10' }; b[0].cells.A3 = { v: '9' }; b[0].cells.A4 = { v: '2' }; b[0].cells.A5 = { v: '1' }; }, 'xlsx-sort-collation-unmeasured'],
  ['a formula inside the block', b => { b[0].cells.B3 = { v: 2, f: '=B2-2' }; }, 'xlsx-sort-formula-in-block-unmeasured'],
  ['a non-ASCII key', b => { b[0].cells.A3 = { v: 'éclair' }; }, 'xlsx-sort-collation-unmeasured'],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const b = book();
    mutate(b);
    const { result, driver } = await run(createStore(b), task());
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('RED: the key moved alone (rows broken) / the wrong order / ignored -> L1 FAIL; a side effect -> L3 FAIL', async () => {
  for (const fault of ['sort-key-only', 'sort-reverse', 'ignore']) {
    const store = createStore(book());
    const { result, driver } = await run(store, task(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
    assert.deepEqual(store.durable, book(), fault);
  }
  const k = await run(createStore(book()), task(), { fault: 'sort-key-only' });
  assert.equal(k.result.execution.primitiveVerification[0].evidence.evidence.rowsKept, false, 'the source would have passed this: its key is ordered');
  const side = await run(createStore(book()), task(), { fault: 'side-effect' });
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('planner: key inside the range, explicit order and hasHeaders, two data rows', () => {
  const refuse = (s, re) => { const p = planOperations(task(s)); assert.equal(p.ok, false, JSON.stringify(s)); assert.match(p.detail, re); };
  refuse({ key: 'C' }, /key must be a column letter inside the range/);
  refuse({ order: 'up' }, /order must be/);
  refuse({ hasHeaders: undefined }, /hasHeaders must be true or false/);
  refuse({ range: 'A1:B2' }, /fewer than two data rows/);
  refuse({ keys: ['A'] }, /does not support field/);
});

test('the command body: SetSort with the key column over the whole block, the header flag', () => {
  const cmd = sort.lower(planOperations(task({ key: 'B', order: 'desc', hasHeaders: false })).operations[0]);
  const log = [];
  const ws = { GetName: () => 'D', GetRange: a => ({ a, SetSort: (...args) => log.push([a, args[0].a, ...args.slice(1)]) }) };
  assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(log, [['A1:B5', 'B1:B5', 'xlDescending', null, null, null, null, 'xlNo', 'xlSortColumns']]);
});
