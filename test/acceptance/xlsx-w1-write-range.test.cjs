'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { validateAdapter } = require('../../src/core/adapter/contract.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const cells = require('../../src/adapters/xlsx/families/cells.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 1: write_range (values) through the TARGET path -- M1 lifecycle -> M2 observation /
 * verification -> M3 executor -> XLSX adapter, on a fake transport with DURABLE and LIVE state apart.
 * The LIVE acceptance on the real editor is a separate harness (gate G3/G4 in
 * docs/evidence/M6-XLSX-CAPABILITY-MATRIX.md).
 */

const book = () => [
  { name: 'Data', cells: { A1: { v: 'Name' }, B1: { v: 'Qty' }, A2: { v: 'Alpha' }, B2: { v: 3 }, C2: { v: null, f: '=B2*2' } }, merges: ['E1:F1'] },
  { name: 'Other', cells: { A1: { v: 'keep me' } } },
];

const task = (range = 'D1:E2', values = [['Total', 7], [1.5, 'x']], sheet = 'Data') => ({ id: 'm6-w1', steps: [{ intent: 'write_range', sheet, range, values }] });
const noMergeTask = () => task('D1:D2', [['Total'], [7]]);

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  assert.equal(driver.calls.open, 1, 'OPEN_COUNT=1');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT=1');
  return { result, driver, m: result.metrics };
}

test('the XLSX adapter satisfies the M0 contract and owns no lifecycle member', () => {
  assert.deepEqual(validateAdapter(xlsxAdapter), { ok: true, problems: [] });
});

test('FIRST session: APPLY -> 1 dispatch -> L1 -> L3 -> L2 -> barrier -> close; DURABLE has the values', async () => {
  const store = createStore(book());
  const { result, driver, m } = await run(store, noMergeTask());
  assert.equal(result.ok, true);
  assert.equal(m.writes, 1);
  assert.equal(m.commands, 1, 'one callCommand for the whole range');
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.APPLIED);
  assert.equal(e.verification, 'LIVE_VERIFY');
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS', 'L3: nothing outside D1:D2 changed');
  assert.equal(result.execution.wholeTaskVerification.failedCount, 0);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(driver.calls.forceSave, 1);
  const data = store.durable.find(s => s.name === 'Data').cells;
  assert.equal(data.D1.v, 'Total');
  assert.equal(data.D2.v, 7);
  assert.equal(store.durable.find(s => s.name === 'Other').cells.A1.v, 'keep me');
});

test('the live read covers the target window even outside every used range', async () => {
  const store = createStore(book());
  const { driver } = await run(store, task('H10:H11', [['a'], [2]]));
  // X-T01: the target window is always read with its number format ALONE (a Text '=...' literal is told from a formula by nf '@')
  assert.deepEqual(driver.reads[0].windows, [{ sheet: 'Data', range: 'H10:H11', nf: true }]);
});

test('SECOND session (reopened retry): SATISFIED, 0 write, 0 command, no barrier', async () => {
  const store = createStore(book());
  await run(store, noMergeTask());
  const { result, driver, m } = await run(store, noMergeTask());
  assert.equal(result.ok, true);
  assert.equal(result.execution.noOp, true);
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0);
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('NEGATIVE CONTROL for the retry authority: a write that never reached DURABLE is applied again', async () => {
  const store = createStore(book());
  let n = 0;
  const driver = createFakeXlsxDriver(store);
  driver.forceSave = async () => ({ ok: true });
  driver.canSave = async () => ({ ok: true, canSave: false });
  await withSession({ driver, fileId: 'd', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } },
    xlsxAdapter, createExecutor(noMergeTask(), { observation: { live: liveScope(noMergeTask()) } }));
  const { result, m } = await run(store, noMergeTask());
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(m.writes, 1);
});

test('a number stored as text is NOT the requested number (COUNT proof), so the retry applies', async () => {
  const s = book();
  s[0].cells.D2 = { v: '7' }; // text "7", COUNT = 0
  s[0].cells.D1 = { v: 'Total' };
  const { result } = await run(createStore(s), noMergeTask());
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
});

test('side effect outside the range -> NOT APPLIED (L3 integral), uncertain, no barrier', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, noMergeTask(), { fault: 'side-effect' });
  const e = result.execution.appliedOperations[0];
  assert.equal(result.ok, false);
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(store.durable.find(x => x.name === 'Data').cells.D1, undefined, 'durable untouched');
});

test('partial write -> L1 FAIL, never APPLIED', async () => {
  const { result } = await run(createStore(book()), noMergeTask(), { fault: 'partial' });
  const e = result.execution.appliedOperations[0];
  assert.equal(result.ok, false);
  assert.notEqual(e.status, OPERATION_STATUS.APPLIED);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('ignored command -> L1 FAIL, nothing persisted', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, noMergeTask(), { fault: 'ignore' });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.forceSave, 0);
});

for (const [name, t, script, reason] of [
  // alice-1: the anchor E1 of E1:F1 is writable now (test/acceptance/mona1-failure-reason.test.cjs); a NON-anchor cell is refused
  ['into a merged area', task('F1:F1', [['m']]), {}, 'xlsx-write-into-merged-area-refused'],
  ['over a formula', task('C2', [[5]]), {}, 'xlsx-write-over-formula-refused'],
  ['on a missing sheet', task('A1', [['x']], 'Nope'), {}, 'xlsx-sheet-not-found'],
  ['on a case-mismatched sheet name', task('A1', [['x']], 'data'), {}, 'xlsx-sheet-sheet-name-case-mismatch'],
  ['on a truncated read', noMergeTask(), { readTruncated: true }, 'xlsx-observation-truncated'],
  ['when the number/text type cannot be read', noMergeTask(), { countUnreadable: true }, 'xlsx-cell-type-unreadable'],
]) {
  test(`refused BEFORE dispatch: write ${name}`, async () => {
    const store = createStore(book());
    const { result, driver } = await run(store, t, script);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0, 'no dispatch');
    assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, reason);
    assert.deepEqual(store.durable, book(), 'durable untouched');
  });
}

test('planner refuses coercible and unsupported values, bad shapes and overlaps -- before any session work', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false); assert.match(p.detail, re); };
  for (const v of ['=SUM(A1)', '+1', '-3', '@x', "'007", '12', '1,5', '3.2e4', '50%', 'TRUE', 'false', '2026-09-27', '9/27', '12:30', ' pad', '']) {
    refuse(task('D1', [[v]]), /refused before dispatch/);
  }
  refuse(task('D1', [[null]]), /unsupported value type null/);
  refuse(task('D1', [[true]]), /unsupported value type boolean/);
  refuse(task('D1', [[Number.NaN]]), /not a finite number/);
  refuse(task('D1:E1', [['a']]), /1x2 matrix/);
  refuse(task('Data!D1', [['a']]), /plain A1 range/);
  refuse(task('A1:AL30', Array.from({ length: 30 }, () => Array(38).fill(1))), /over the 1000 ceiling/);
  refuse({ id: 'x', steps: [{ intent: 'write_range', sheet: 'Data', range: 'D1', values: [['a']], formulas: [['=1']] }] }, /both a value and a formula/);
  refuse({ id: 'x', steps: [{ intent: 'write_range', sheet: 'Data', range: 'D1', values: [['a']], extra: 1 }] }, /does not support field/);
  refuse({ id: 'x', steps: [{ intent: 'add_sparkline', sheet: 'Data', range: 'D1' }] }, /intent-not-migrated-in-m6/);
  refuse({ id: 'x', retryToken: 't', steps: [] }, /task field\(s\) not supported: retryToken/);
  refuse({ id: 'x', steps: [
    { intent: 'write_range', sheet: 'Data', range: 'D1:D2', values: [['a'], ['b']] },
    { intent: 'write_range', sheet: 'Data', range: 'D2:E2', values: [['c', 'd']] },
  ] }, /overlapping-writes-refused/);
  assert.equal(planOperations({ id: 'x', steps: [
    { intent: 'write_range', sheet: 'Data', range: 'D1', values: [['a']] },
    { intent: 'write_range', sheet: 'Other', range: 'D1', values: [['b']] },
  ] }).ok, true, 'the same range on another sheet is not an overlap');
});

test('accepted plain values', () => {
  for (const v of ['Total', 'Alpha beta', 'ÁRVÍZTŰRŐ', 'x1', 'A-1', 'Q3 plan', 0, -2.5, 1e6]) assert.equal(cells.valueProblem(v), null, String(v));
});

test('the write command body compiles and writes through SetValue only, on a stand-in Api', () => {
  const op = planOperations(task('D1:E1', [['Total', 7]])).operations[0];
  const cmd = cells.lower(op);
  const writes = [];
  const ws = { GetName: () => 'Data', GetRange: a => ({ SetValue: v => writes.push([a, v]) }) };
  const out = new Function('Api', cmd.body)({ GetSheet: n => (n === 'Data' ? ws : null) });
  assert.deepEqual(out, { ok: true, written: 2 });
  assert.deepEqual(writes, [['D1', 'Total'], ['E1', 7]]);
  assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ({ GetName: () => 'data' }) }), { ok: false, stage: 'sheet-identity' }, 'exact-name check inside the command');
  assert.deepEqual([...cmd.body.matchAll(/\.([A-Z]\w*)\s*\(/g)].map(m => m[1]).sort(), ['GetName', 'GetRange', 'GetSheet', 'SetValue']);
});
