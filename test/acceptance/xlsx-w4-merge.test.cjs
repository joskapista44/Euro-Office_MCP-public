'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const merge = require('../../src/adapters/xlsx/families/merge.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 4: merge_range / unmerge_range through the target path on the fake transport. The fake follows the
 * W4b LIVE probe: a merge absorbs the merges it intersects and SILENTLY empties every non-top-left cell (so the
 * content guard refuses before the write), an unmerge dissolves the merges it intersects. The readback is the
 * per-cell MergeArea (B-L18: the source's top-left-only readback gave a false "already unmerged").
 */

const book = () => [
  { name: 'Data', cells: {
    A1: { v: 'Head' }, A2: { v: 'x' }, B2: { v: 5 }, C3: { v: '', f: '=""' }, E1: { v: 'inner' },
    A6: { v: 'm1' }, D6: { v: 'm2' }, H1: { v: 'wide' },
  }, merges: ['E1:F1', 'H1:I2', 'A6:B6', 'D6:E6'] },
  { name: 'Other', cells: { A1: { v: 'keep me' } } },
];

const task = (intent, range, sheet = 'Data') => ({ id: 'm6-w4-merge', steps: [{ intent, sheet, range }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const data = store => store.durable.find(s => s.name === 'Data');
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('merge FIRST: APPLY -> 1 dispatch -> every cell reports the area, the top-left keeps its value -> L3 -> barrier', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task('merge_range', 'A1:C1'));
  assert.equal(result.ok, true);
  assert.equal(m.writes, 1);
  assert.equal(m.commands, 1);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'PASS');
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.ok(data(store).merges.includes('A1:C1'));
  assert.equal(data(store).cells.A1.v, 'Head');
});

test('merge SECOND (reopened retry): SATISFIED, 0 write, 0 command, no barrier', async () => {
  const store = createStore(book());
  await run(store, task('merge_range', 'A1:C1'));
  const { result, m } = await run(store, task('merge_range', 'A1:C1'));
  assert.equal(result.ok, true);
  assert.equal(result.execution.noOp, true);
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('merge over an inner merge absorbs it (the measured engine behaviour); the retry is SATISFIED 0/0', async () => {
  const store = createStore(book());
  const { result } = await run(store, task('merge_range', 'E1:G1'));
  assert.equal(result.ok, true);
  assert.equal(data(store).merges.includes('E1:F1'), false);
  assert.ok(data(store).merges.includes('E1:G1'));
  assert.equal(data(store).cells.E1.v, 'inner', 'the shared top-left keeps its value');
  const again = await run(store, task('merge_range', 'E1:G1'));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('unmerge FIRST dissolves every inner merge (per-cell readback), the content stays; retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task('unmerge_range', 'A6:E6'));
  assert.equal(result.ok, true);
  assert.equal(m.commands, 1);
  assert.deepEqual(data(store).merges, ['E1:F1', 'H1:I2']);
  assert.equal(data(store).cells.A6.v, 'm1');
  assert.equal(data(store).cells.D6.v, 'm2');
  const again = await run(store, task('unmerge_range', 'A6:E6'));
  assert.equal(again.m.writes, 0);
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('unmerge where nothing is merged: SATISFIED on the first run, 0 command', async () => {
  const { result, m } = await run(createStore(book()), task('unmerge_range', 'A1:C2'));
  assert.equal(result.ok, true);
  assert.equal(m.commands, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

for (const [name, t, script, why] of [
  ['merge over a non-blank non-top-left VALUE (the engine would empty it)', task('merge_range', 'A2:B2'), {}, 'xlsx-merge-data-loss-risk'],
  ['merge over a FORMULA whose result is empty (a formula is never blank)', task('merge_range', 'B3:C3'), {}, 'xlsx-merge-data-loss-risk'],
  ['merge that would absorb an inner merge whose top-left holds a value', task('merge_range', 'D1:F1'), {}, 'xlsx-merge-data-loss-risk'],
  ['merge whose existing merge reaches outside the target', task('merge_range', 'H1:H2'), {}, 'xlsx-merge-area-reaches-outside-the-target'],
  ['unmerge of part of a merge (the engine would unmerge more)', task('unmerge_range', 'H1'), {}, 'xlsx-merge-area-reaches-outside-the-target'],
  ['merge on a missing sheet', task('merge_range', 'A1:B1', 'Nope'), {}, 'xlsx-sheet-not-found'],
  ['merge on a truncated read', task('merge_range', 'A1:C1'), { readTruncated: true }, 'xlsx-observation-truncated'],
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

test('the data-loss refusal names the cells and says which is a formula', async () => {
  const { result } = await run(createStore(book()), task('merge_range', 'A3:C3'));
  assert.deepEqual(result.execution.appliedOperations[0].evidence.evidence.cells, [{ cell: 'C3', formula: true }]);
});

test('RED: ignored / partial / top-left dropped / landed elsewhere -> L1 FAIL; nothing persisted', async () => {
  for (const [t, script] of [
    [task('merge_range', 'A1:C1'), { fault: 'ignore' }],
    [task('merge_range', 'F3:G4'), { fault: 'partial' }],
    [task('unmerge_range', 'A6:E6'), { fault: 'partial' }],
    [task('merge_range', 'A1:C1'), { fault: 'merge-drops-top-left' }],
    [task('merge_range', 'A1:C1'), { fault: 'merge-lands-elsewhere', elsewhere: 'A4:C4' }],
  ]) {
    const store = createStore(book());
    const { result, driver } = await run(store, t, script);
    const label = `${t.steps[0].intent} ${script.fault}`;
    assert.equal(result.ok, false, label);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', label);
    assert.notEqual(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, label);
    assert.equal(driver.calls.forceSave, 0, label);
    assert.deepEqual(store.durable, book(), label);
  }
  const { result } = await run(createStore(book()), task('merge_range', 'A1:C1'), { fault: 'merge-drops-top-left' });
  assert.deepEqual(result.execution.primitiveVerification[0].evidence.evidence.contentChanged, ['A1']);
});

test('RED: a side effect -- a value elsewhere, or ANOTHER merge -> L3 FAIL (merge areas are preserved too)', async () => {
  const cell = await run(createStore(book()), task('merge_range', 'A1:C1'), { fault: 'side-effect' });
  assert.equal(cell.result.ok, false);
  assert.equal(cell.result.execution.preservationVerification[0].verdict, 'FAIL');
  const other = await run(createStore(book()), task('merge_range', 'A1:C1'), { fault: 'side-merge', sideMerge: 'A2:B2' });
  assert.equal(other.result.ok, false);
  const pv = other.result.execution.preservationVerification[0];
  assert.equal(pv.verdict, 'FAIL');
  assert.deepEqual(pv.evidence.evidence.mergeChanged.map(x => x.cell), ['Data!A2', 'Data!B2']);
});

test('fail closed when a target cell\'s merge area was not read', () => {
  const op = planOperations(task('merge_range', 'A1:B1')).operations[0];
  const primary = normalizeInventory({ sheets: [{ name: 'Data', index: 0, used: 'A1:B1' }], cells: [
    { sheet: 'Data', a: 'A1', v: 'x', f: 'x', count: 0, merge: 'A1' }, { sheet: 'Data', a: 'B1', v: '', f: '', count: 0, merge: null },
  ], truncated: false, unsupported: [] });
  assert.deepEqual(merge.classify(primary, op), { state: 'FAIL_CLOSED', evidence: { ok: false, reason: 'xlsx-target-cell-unread', cell: 'B1' } });
});

test('planner: fields, range shape, a one-cell merge, ceiling, overlap with a write', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false); assert.match(p.detail, re); };
  // X-L06: 'across' is now a supported field on merge_range (xlsx-xl06-merge-across.test.cjs) --
  // an unrelated unknown field still refuses.
  refuse({ id: 'x', steps: [{ intent: 'merge_range', sheet: 'Data', range: 'A1:B1', nope: true }] }, /does not support field/);
  refuse(task('merge_range', 'Data!A1:B1'), /plain A1 range/);
  refuse(task('merge_range', 'A1'), /at least two cells/);
  refuse(task('unmerge_range', 'A1:AL30'), /over the 1000 ceiling/);
  refuse({ id: 'x', steps: [{ intent: 'write_range', sheet: 'Data', range: 'B1', values: [['a']] }, { intent: 'merge_range', sheet: 'Data', range: 'A1:C1' }] }, /overlapping-writes-refused/);
  assert.equal(planOperations(task('unmerge_range', 'A1')).ok, true, 'a one-cell unmerge is a valid request');
});

test('the command bodies compile: Merge(false) / UnMerge() on the exact-name sheet only', () => {
  for (const [intent, want] of [['merge_range', [['Merge', false]]], ['unmerge_range', [['UnMerge']]]]) {
    const cmd = merge.lower(planOperations(task(intent, 'A1:C1')).operations[0]);
    const calls = [];
    const ws = { GetName: () => 'Data', GetRange: a => ({ Merge: x => calls.push(['Merge', x, a]), UnMerge: () => calls.push(['UnMerge', a]) }) };
    assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ws }), { ok: true });
    assert.deepEqual(calls.map(c => c.slice(0, want[0].length)), want);
    assert.equal(calls[0][calls[0].length - 1], 'A1:C1');
    assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ({ GetName: () => 'data' }) }), { ok: false, stage: 'sheet-identity' });
    // X-L06: meta carries `across` too (false here -- plain merge_range / unmerge_range).
    assert.deepEqual(cmd.meta, { family: 'merge', action: intent, sheet: 'Data', range: 'A1:C1', across: false });
  }
});
