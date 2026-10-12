'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * X-L06 (behavioural LIVE probe, adc47f6): merge_range's `across: true` field. The
 * SAME public Merge(true) call splits the whole target range into ONE MERGE PER ROW, persisting
 * across a reopen. Content guard (every non-left-most cell must be blank) applies PER ROW -- the
 * engine silently empties a row's non-left-most non-blank cell otherwise (X-L07's named safety
 * boundary), so this family refuses before the write rather than weakening.
 */

const book = () => [
  { name: 'Data', cells: {
    A1: { v: 'r1' }, A2: { v: 'r2' }, A3: { v: 'r3' },
    D1: { v: 'd1' }, D2: { v: 'd2' }, E2: { v: 'poison' },
    G1: { v: 'g1' }, G2: { v: 'g2' },
  } },
  { name: 'Other', cells: { A1: { v: 'keep me' } } },
];

const task = (intent, range, extra = {}) => ({ id: 'xl06-across', steps: [{ intent, sheet: 'Data', range, ...extra }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const data = store => store.durable.find(s => s.name === 'Data');
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('merge_range across:true FIRST: APPLY -> 1 dispatch -> one merge PER ROW, each left-most cell keeps its value -> L3 PASS', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task('merge_range', 'A1:B3', { across: true }));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.writes, 1);
  assert.equal(m.commands, 1);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'PASS');
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.deepEqual(data(store).merges.slice().sort(), ['A1:B1', 'A2:B2', 'A3:B3']);
  assert.equal(data(store).cells.A1.v, 'r1');
  assert.equal(data(store).cells.A2.v, 'r2');
  assert.equal(data(store).cells.A3.v, 'r3');
});

test('merge_range across:true SECOND (reopened retry): SATISFIED, 0 write, 0 command, no barrier', async () => {
  const store = createStore(book());
  await run(store, task('merge_range', 'A1:B3', { across: true }));
  const { result, m } = await run(store, task('merge_range', 'A1:B3', { across: true }));
  assert.equal(result.ok, true);
  assert.equal(result.execution.noOp, true);
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('across: a non-left-most non-blank cell in ANY row refuses the WHOLE op, fail-closed, before any write', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task('merge_range', 'D1:E2', { across: true }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-merge-data-loss-risk');
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.cells[0].cell, 'E2');
  assert.deepEqual(data(store).merges || [], [], 'nothing merged -- the guard fired before the write');
  assert.equal(data(store).cells.E2.v, 'poison', 'untouched -- refused before any write');
});

test('across: a left-most cell having content in every row is fine (only non-left-most cells are checked)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task('merge_range', 'G1:H2', { across: true }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.deepEqual(data(store).merges.slice().sort(), ['G1:H1', 'G2:H2']);
  assert.equal(data(store).cells.G1.v, 'g1');
  assert.equal(data(store).cells.G2.v, 'g2');
});

test('unmerge_range (no across flag) dissolves an across-merge in ONE call -- every per-row merge gone, content kept', async () => {
  const store = createStore(book());
  await run(store, task('merge_range', 'A1:B3', { across: true }));
  const { result, m } = await run(store, task('unmerge_range', 'A1:B3'));
  assert.equal(result.ok, true);
  assert.equal(m.commands, 1);
  assert.deepEqual(data(store).merges || [], []);
  assert.equal(data(store).cells.A1.v, 'r1');
  assert.equal(data(store).cells.A2.v, 'r2');
  assert.equal(data(store).cells.A3.v, 'r3');
  const again = await run(store, task('unmerge_range', 'A1:B3'));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('across: a single-column range is refused at PLAN time -- nothing to merge within one cell, and its "already merged" state would be indistinguishable from "never touched"', () => {
  const p = planOperations(task('merge_range', 'A1:A3', { across: true }));
  assert.equal(p.ok, false);
  assert.match(p.detail, /at least two columns/);
});

test('plan: across is only valid on merge_range, and must be a boolean', () => {
  const p1 = planOperations(task('unmerge_range', 'A1:B3', { across: true }));
  assert.equal(p1.ok, false);
  assert.match(p1.detail, /does not support field\(s\): across/);
  const p2 = planOperations(task('merge_range', 'A1:B3', { across: 'yes' }));
  assert.equal(p2.ok, false);
  assert.match(p2.detail, /across must be true or false/);
  const p3 = planOperations(task('merge_range', 'A1:B3', { across: false }));
  assert.equal(p3.ok, true);
});

test('across: the whole-task check reports "is merged per row", not the whole-range label', async () => {
  const store = createStore(book());
  const { result } = await run(store, task('merge_range', 'G1:H2', { across: true }));
  assert.equal(result.ok, true);
  assert.equal(result.execution.wholeTaskVerification.verdict, 'PASS');
  assert.match(result.execution.wholeTaskVerification.evidence.checks[0].name, /is merged per row/);
});
