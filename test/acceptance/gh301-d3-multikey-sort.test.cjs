'use strict';

/**
 * GH #301 D3: sort_range with up to three keys (`keys: [{ column, order }]`), the public SetSort(key1, order1, key2,
 * order2, key3, order3, header, orientation) shape LIVE-measured in GH #301 S-02 (correct order, same-row formulas
 * follow). `keyColumn` + `order` (one key) stays exactly as before; `keys` is the alternative, never both. Every key
 * column is held to the SAME measured contract as the single key (no formula, no mixed types, no collation guess);
 * a blank cell is measured only in a DESCENDING PRIMARY key (R21 -- blanks last), so a blank in a secondary key, or
 * in an ascending key, is refused before the write.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const sort = require('../../src/adapters/xlsx/families/sort.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

// region (text) / amount (number) / name (text): two regions tie, the second key decides inside them
const book = () => [{ name: 'D', cells: {
  A1: { v: 'region' }, B1: { v: 'amount' }, C1: { v: 'name' },
  A2: { v: 'west' }, B2: { v: 5 }, C2: { v: 'echo' },
  A3: { v: 'east' }, B3: { v: 2 }, C3: { v: 'bravo' },
  A4: { v: 'west' }, B4: { v: 9 }, C4: { v: 'alpha' },
  A5: { v: 'east' }, B5: { v: 7 }, C5: { v: 'delta' },
  A6: { v: 'east' }, B6: { v: 7 }, C6: { v: 'charlie' },
} }];
const stepOf = (extra = {}) => ({ intent: 'sort_range', sheet: 'D', range: 'A1:C6', hasHeaders: true, keys: [{ column: 'A', order: 'asc' }, { column: 'B', order: 'desc' }, { column: 'C', order: 'asc' }], ...extra });
const task = step => ({ id: 'gh301-d3', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const col = (store, c) => [2, 3, 4, 5, 6].map(r => store.durable[0].cells[`${c}${r}`].v);

test('D3 schema: keys (1..3) accepted; keyColumn/order still accepted; both, neither, 4 keys, a duplicate or an outside column refused', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(stepOf()).ok, true);
  assert.equal(v({ intent: 'sort_range', sheet: 'D', range: 'A1:C6', hasHeaders: true, keyColumn: 'A', order: 'asc' }).ok, true, 'the one-key shape is unchanged');
  assert.equal(v(stepOf({ keyColumn: 'A', order: 'asc' })).ok, false, 'keys AND keyColumn together');
  assert.equal(v({ intent: 'sort_range', sheet: 'D', range: 'A1:C6', hasHeaders: true }).ok, false, 'no key at all');
  assert.equal(v(stepOf({ keys: [{ column: 'A', order: 'asc' }, { column: 'B', order: 'asc' }, { column: 'C', order: 'asc' }, { column: 'A', order: 'desc' }] })).ok, false, 'four keys');
  assert.equal(v(stepOf({ keys: [{ column: 'A', order: 'asc' }, { column: 'A', order: 'desc' }] })).ok, false, 'a duplicate key column');
  assert.equal(v(stepOf({ keys: [{ column: 'A', order: 'asc' }, { column: 'F', order: 'asc' }] })).ok, false, 'a key column outside the range');
});

test('D3: three keys -- APPLIED, rows ordered by region asc, amount desc, name asc, rows moved whole; retry SATISFIED 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(stepOf()));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.deepEqual(col(store, 'A'), ['east', 'east', 'east', 'west', 'west']);
  assert.deepEqual(col(store, 'B'), [7, 7, 2, 9, 5]);
  assert.deepEqual(col(store, 'C'), ['charlie', 'delta', 'bravo', 'alpha', 'echo']);
  const again = await run(store, task(stepOf()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('D3: the lowered command passes all three keys to SetSort, in order', () => {
  const plan = sort.plan(stepOf(), 0);
  assert.equal(plan.ok, true, plan.detail);
  const body = sort.lower(plan.operation).body;
  assert.match(body, /SetSort\(ws\.GetRange\("A1:A6"\), "xlAscending", ws\.GetRange\("B1:B6"\), "xlDescending", ws\.GetRange\("C1:C6"\), "xlAscending", "xlYes", 'xlSortColumns'\)/);
});

test('D3: a one-key keys[] lowers byte-identically to keyColumn/order', () => {
  const a = sort.plan({ intent: 'sort_range', sheet: 'D', range: 'A1:C6', hasHeaders: true, keyColumn: 'B', order: 'desc' }, 0).operation;
  const b = sort.plan(stepOf({ keys: [{ column: 'B', order: 'desc' }] }), 0).operation;
  assert.equal(sort.lower(b).body, sort.lower(a).body);
});

for (const [name, mutate, why] of [
  ['a blank cell in a SECONDARY key (placement unmeasured)', b => { delete b[0].cells.B3; }, 'xlsx-sort-blank-key-unmeasured'],
  ['a formula in a secondary key column', b => { b[0].cells.B3 = { v: 2, f: '=1+1' }; }, 'xlsx-sort-formula-key-unmeasured'],
  ['mixed types in a secondary key', b => { b[0].cells.B3 = { v: 'two' }; }, 'xlsx-sort-mixed-key-types-unmeasured'],
  ['a collation-dependent tertiary key', b => { b[0].cells.C5 = { v: 'Delta' }; }, 'xlsx-sort-collation-unmeasured'],
]) {
  test(`D3 refused before the write, 0 commands: ${name}`, async () => {
    const b = book(); mutate(b);
    const store = createStore(b);
    const { result, m } = await run(store, task(stepOf()));
    assert.equal(result.ok, false);
    assert.equal(m.commands, 0);
    assert.match(JSON.stringify(result.execution), new RegExp(why));
  });
}

test('D3: a fake that sorts by the PRIMARY key only (ties left in source order) fails the verify, not passes it', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(stepOf()), { fault: 'sort-primary-only' });
  assert.equal(m.commands, 1, 'the write was dispatched -- this is the VERIFY failing, not a refusal');
  assert.equal(result.ok, false, 'the secondary order is part of the proof');
});

test('D3: planOperations keeps keys on the operation (and the first key as keyColumn / order)', () => {
  const p = planOperations(task(stepOf()));
  assert.equal(p.ok, true, p.detail);
  const op = p.operations[0];
  assert.deepEqual(op.keys.map(k => [k.column, k.order]), [[1, 'asc'], [2, 'desc'], [3, 'asc']]);
  assert.equal(op.keyColumn, 1);
  assert.equal(op.order, 'asc');
});
