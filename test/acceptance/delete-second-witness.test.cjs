'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { STEP_SCHEMAS } = require('../../src/server/write-tools.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * Owner decision B (2026-11-02, owner channel): a delete_rows / delete_columns step carries a SECOND, required witness
 * -- the value of the FIRST REMOVED row (column), read where it is now (at `at`) -- beside the anchor (the first
 * SURVIVING one, now at `at + count`). Both are read before the first write. The defect it closes (the reported case): an
 * anchor given as the first REMOVED row's value already sat at `at`, the anchor-only rule read that as "already done"
 * and returned SATISFIED -- 0 rows deleted, reported as success. A delete now has no SATISFIED path: the after-shape
 * (anchor at `at`, witness gone) is refused as unprovable, so an identical retry still never deletes twice.
 */

const book = () => [
  { name: 'P', cells: {
    A1: { v: 'h' }, A2: { v: 'r2' }, A3: { v: 'r3' }, A4: { v: 'r4' }, A5: { v: 'r5' }, A6: { v: 'r6' },
    B1: { v: 'id' }, B2: { v: 102 }, B3: { v: 103 }, B4: { v: 104 }, B5: { v: 105 }, B6: { v: 106 },
    C1: { v: 'c1' }, D1: { v: 'd1' }, E1: { v: 'e1' }, F1: { v: 'f1' },
    C2: { v: 'x' }, D2: { v: 'y' }, E2: { v: 'z' },
  } },
];

const task = (intent, extra) => ({ id: 'del-witness', steps: [{ intent, sheet: 'P', ...extra }] });

async function run(store, t) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, {});
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}

const P = store => store.durable.find(s => s.name === 'P').cells;
const snapshot = store => JSON.stringify(P(store));
const first = result => result.execution.appliedOperations[0];
const reasonOf = result => first(result).evidence.evidence.reason;

test('the reported case, RED before decision B: the anchor given as the first REMOVED row (r3 at row 3) is no longer SATISFIED -- refused, 0 writes, the sheet unchanged', async () => {
  const store = createStore(book());
  const before = snapshot(store);
  // delete rows 3-4: the anchor SHOULD be r5 (first surviving); the caller gave r3 (the first removed row) -- before
  // decision B this read as "anchor at its terminal position" -> SATISFIED, nothing deleted, reported as success
  const t = task('delete_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r3' }, witness: { column: 'B', value: 103 } });
  const { result, m } = await run(store, t);
  assert.equal(result.ok, false);
  assert.notEqual(first(result).status, OPERATION_STATUS.SATISFIED);
  assert.equal(first(result).status, OPERATION_STATUS.FAILED);
  assert.equal(reasonOf(result), 'xlsx-structure-delete-witnesses-disagree');
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0);
  assert.equal(snapshot(store), before);
});

test('the anchor-only call of the old API is refused before dispatch -- the witness is required (plan and MCP schema)', () => {
  for (const [intent, at, anchor] of [['delete_rows', 3, { column: 'A', value: 'r5' }], ['delete_columns', 'C', { row: 1, value: 'e1' }]]) {
    const p = planOperations(task(intent, { at, count: 2, anchor }));
    assert.equal(p.ok, false, intent);
    assert.match(p.detail, /witness must be .*FIRST (row|column) the delete removes/);
  }
  assert.deepEqual(STEP_SCHEMAS.delete_rows.required.slice().sort(), ['anchor', 'at', 'count', 'intent', 'sheet', 'witness']);
  assert.deepEqual(STEP_SCHEMAS.delete_columns.required.slice().sort(), ['anchor', 'at', 'count', 'intent', 'sheet', 'witness']);
  assert.equal(STEP_SCHEMAS.insert_rows.properties.witness, undefined, 'an insert has no witness');
  assert.equal(STEP_SCHEMAS.insert_columns.properties.witness, undefined);
});

test('plan: a witness on an insert, a witness equal to the anchor, a malformed witness -- each refused before dispatch', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse(task('insert_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r3' }, witness: { column: 'A', value: 'r2' } }), /does not support field\(s\): witness/);
  refuse(task('delete_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r4' }, witness: { column: 'A', value: 'r4' } }), /same value in the same column/);
  refuse(task('delete_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r4' }, witness: { row: 3, value: 'r3' } }), /witness must be \{ column, value \}/);
  refuse(task('delete_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r4' }, witness: { column: 'A', value: '' } }), /witness.value must be/);
  refuse(task('delete_columns', { at: 'C', count: 1, anchor: { row: 1, value: 'd1' }, witness: { column: 'C', value: 'c1' } }), /witness must be \{ row, value \}/);
});

test('a genuine delete: both witnesses where the step says -> APPLIED, the witness gone (L1); the identical retry is refused as unprovable, 0 writes, nothing deleted twice', async () => {
  const store = createStore(book());
  const t = task('delete_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r5' }, witness: { column: 'B', value: 103 } });
  const { result, m } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations));
  assert.equal(first(result).status, OPERATION_STATUS.APPLIED);
  assert.equal(m.commands, 1);
  assert.equal(result.execution.primitiveVerification[0].evidence.evidence.witnessGone, true);
  assert.equal(P(store).A3.v, 'r5');
  assert.equal(P(store).B3.v, 105);
  assert.equal(P(store).A4.v, 'r6');
  const after = snapshot(store);
  // the retry: the sheet now reads exactly as this delete leaves it -- but matching data proves nothing about who made it
  const again = await run(store, t);
  assert.equal(again.result.ok, false);
  assert.equal(first(again.result).status, OPERATION_STATUS.FAILED);
  assert.equal(reasonOf(again.result), 'xlsx-structure-delete-prior-execution-unprovable');
  assert.equal(first(again.result).evidence.evidence.readsAsAlreadyApplied, true);
  assert.equal(again.m.writes, 0);
  assert.equal(again.m.commands, 0);
  assert.equal(snapshot(store), after, 'an identical retry deletes nothing');
});

test('the after-shape with a witness that never existed is NOT success either: unprovable, 0 writes (the false-SATISFIED shape the old rule accepted)', async () => {
  const store = createStore(book());
  const before = snapshot(store);
  // anchor r3 sits at row 3 (= at), the witness value is nowhere in its column: reads like "done", is not provable
  const t = task('delete_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r3' }, witness: { column: 'B', value: 999 } });
  const { result, m } = await run(store, t);
  assert.equal(result.ok, false);
  assert.equal(reasonOf(result), 'xlsx-structure-delete-prior-execution-unprovable');
  assert.equal(m.writes, 0);
  assert.equal(snapshot(store), before);
});

test('PRE_DISPATCH refusals: the witness not at `at`, the witness not unique, the anchor at neither position -- each named, 0 writes', async () => {
  const cases = [
    // count off by one: anchor r5 is at at+count only for count 2; with count 3 it is at neither -> conflict
    [{ at: 3, count: 3, anchor: { column: 'A', value: 'r5' }, witness: { column: 'B', value: 103 } }, 'xlsx-structure-anchor-state-conflict', null],
    // the anchor right, the witness one row off (104 is row 4, not at=3): the block is not the one on the sheet
    [{ at: 3, count: 2, anchor: { column: 'A', value: 'r5' }, witness: { column: 'B', value: 104 } }, 'xlsx-structure-delete-witnesses-disagree', null],
    // the anchor right, the witness gone: not the before-state either
    [{ at: 3, count: 2, anchor: { column: 'A', value: 'r5' }, witness: { column: 'B', value: 999 } }, 'xlsx-structure-delete-witnesses-disagree', null],
    // the witness value twice in its lane
    [{ at: 3, count: 2, anchor: { column: 'A', value: 'r5' }, witness: { column: 'B', value: 103 } }, 'xlsx-structure-witness-not-unique', b => { b[0].cells.B6 = { v: 103 }; }],
  ];
  for (const [extra, want, mutate] of cases) {
    const b = book();
    if (mutate) mutate(b);
    const store = createStore(b);
    const before = snapshot(store);
    const { result, m } = await run(store, task('delete_rows', extra));
    assert.equal(result.ok, false, want);
    assert.equal(reasonOf(result), want);
    assert.equal(m.writes, 0, want);
    assert.equal(m.commands, 0, want);
    assert.equal(snapshot(store), before, want);
  }
});

test('delete_columns follows the same contract: APPLIED with both witnesses, the identical retry unprovable, the wrong anchor refused', async () => {
  const store = createStore(book());
  // delete C:D -- the first removed column C (C1 = c1), the first surviving E (E1 = e1)
  const t = task('delete_columns', { at: 'C', count: 2, anchor: { row: 1, value: 'e1' }, witness: { row: 2, value: 'x' } });
  const { result } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations));
  assert.equal(P(store).C1.v, 'e1');
  assert.equal(P(store).D1.v, 'f1');
  const after = snapshot(store);
  const again = await run(store, t);
  assert.equal(reasonOf(again.result), 'xlsx-structure-delete-prior-execution-unprovable');
  assert.equal(again.m.writes, 0);
  assert.equal(snapshot(store), after);

  const fresh = createStore(book());
  const before = snapshot(fresh);
  // the anchor given as the first REMOVED column (c1, at C) -- the old rule's SATISFIED
  const wrong = await run(fresh, task('delete_columns', { at: 'C', count: 2, anchor: { row: 1, value: 'c1' }, witness: { row: 2, value: 'x' } }));
  assert.equal(reasonOf(wrong.result), 'xlsx-structure-delete-witnesses-disagree');
  assert.equal(wrong.m.writes, 0);
  assert.equal(snapshot(fresh), before);
});

test('inserts keep their SATISFIED retry (decision B is about deletes; an insert\'s second witness is the blank gap)', async () => {
  const store = createStore(book());
  const t = task('insert_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r3' } });
  const { result } = await run(store, t);
  assert.equal(result.ok, true);
  const again = await run(store, t);
  assert.equal(first(again.result).status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.m.writes, 0);
});
