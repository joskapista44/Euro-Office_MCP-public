'use strict';

/**
 * GH #99 (owner GO 2026-10-14, owner channel): a write task's independent plain-value write_range steps share ONE
 * targeting read and ONE post-read (the executor's M3 batch strategy) instead of a fresh read per step -- measured on
 * the installed release: 43 steps on a 38k-cell workbook spent 13.2 s of 20.8 s in 44 re-observations.
 *   planner   a plain-value write_range (no formula, no typed value) carries batchSafeGroup 'xlsx-write-values';
 *             an overlap is already a dependsOn, so the group is cut there (plan.cjs: a dependency beats any proof)
 *   core      verifyPreservation hands the adapter the round's OTHER dispatched operations (siblings); each one's L1
 *             proves its own window, so an L3 must not read a sibling's write as a foreign change
 *   server    xlsx_execute_task runs strategy 'batch' (unproven steps stay singletons = the sequential behaviour)
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

const book = () => [
  { name: 'Data', cells: { A1: { v: 'Name' }, B2: { v: 3 }, C2: { v: null, f: '=B2*2' } }, merges: [] },
  { name: 'Other', cells: { A1: { v: 'keep me' } } },
];
const w = (i, extra = {}) => ({ intent: 'write_range', id: `w${i}`, sheet: 'Data', range: `H${i}`, values: [[`v${i}`]], ...extra });
const taskOf = steps => ({ id: 'gh99', steps });
const FIVE = taskOf([1, 2, 3, 4, 5].map(i => w(i)));

async function run(t, { strategy = 'batch', script = {}, adapter = xlsxAdapter } = {}) {
  const store = createStore(book());
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, adapter, createExecutor(t, { strategy, observation: { live: liveScope(t) } }));
  return { result, driver, store, statuses: result.execution.appliedOperations.map(o => o.status) };
}
const durable = store => store.durable.find(s => s.name === 'Data').cells;

const fmt = (i, range) => ({ intent: 'format_range', id: `f${i}`, sheet: 'Data', range, format: { fill: '#FFFF00' } });

test('planner: plain-value write_range steps carry the batch proof; a formula, a typed value or another intent does not', () => {
  const p = planOperations(taskOf([w(1), w(2, { formulas: [['=B2+1']], values: undefined }), w(3, { values: [[{ text: '007' }]] }), w(4), fmt(5, 'H4')]));
  assert.equal(p.ok, true, p.detail);
  assert.deepEqual(p.operations.map(o => o.batchSafeGroup || null), ['xlsx-write-values', null, null, 'xlsx-write-values', null]);
  assert.deepEqual(p.operations[4].dependsOn, ['w4'], 'a format on the written cell depends on the write');
  // two writes to one cell are refused outright (a clash), so a write_range group never holds an overlap
  assert.match(planOperations(taskOf([w(4), w(5, { range: 'H4' })])).detail, /overlapping-writes-refused/);
});

test('batch vs sequential: the same end state and statuses, 2 reads instead of steps + 1', async () => {
  const seq = await run(FIVE, { strategy: 'sequential' });
  const bat = await run(FIVE);
  for (const r of [seq, bat]) {
    assert.equal(r.result.ok, true);
    assert.deepEqual(r.statuses, Array(5).fill(OPERATION_STATUS.APPLIED));
    assert.ok(r.result.execution.preservationVerification.every(v => v.verdict === 'PASS'));
  }
  assert.deepEqual([1, 2, 3, 4, 5].map(i => durable(bat.store)[`H${i}`].v), [1, 2, 3, 4, 5].map(i => durable(seq.store)[`H${i}`].v));
  assert.equal(seq.driver.calls.readLive, 6);
  assert.equal(bat.driver.calls.readLive, 2);
  assert.equal(bat.driver.calls.dispatch, 5, 'XLSX has no batch hook: one dispatch per step, the saving is the reads');
  assert.equal(bat.result.execution.batchEvidence.groups, 1);
});

test('a formula step and an unproven intent cut the group', async () => {
  const t = taskOf([w(1), w(2), w(3, { formulas: [['=B2+1']], values: undefined }), w(4), fmt(5, 'H9')]);
  const r = await run(t);
  assert.equal(r.result.ok, true);
  assert.deepEqual(r.statuses, Array(5).fill(OPERATION_STATUS.APPLIED));
  assert.equal(r.result.execution.batchEvidence.groups, 4, '[w1,w2] | [w3] | [w4] | [f5]');
});

test('RED: a foreign change outside every window in a batch round is still caught -- not APPLIED', async () => {
  const r = await run(FIVE, { script: { fault: 'side-effect' } });
  assert.equal(r.result.ok, false);
  assert.notEqual(r.statuses[0], OPERATION_STATUS.APPLIED);
  assert.ok(r.result.execution.preservationVerification.some(v => v.verdict === 'FAIL'));
});

test('RED: without the sibling exclusion the batch round fails its own L3 -- the siblings are what makes it pass', async () => {
  const noSiblings = Object.create(xlsxAdapter);
  noSiblings.verifyPreservation = (h, op, b, a) => xlsxAdapter.verifyPreservation(h, op, b, a, {});
  const r = await run(FIVE, { adapter: noSiblings });
  assert.equal(r.result.ok, false);
  assert.ok(r.result.execution.preservationVerification.some(v => v.verdict === 'FAIL'));
});

test('a sibling that is not batch-proven never reaches a family: the adapter refuses it (planner-bug guard)', async () => {
  const p = planOperations(taskOf([w(1)]));
  const r = await xlsxAdapter.verifyPreservation(null, p.operations[0], null, null, { siblings: [{ id: 'x', intent: 'format_range' }] });
  assert.deepEqual([r.ok, r.evidence.reason], [false, 'xlsx-preservation-sibling-not-batch-proven']);
});

test('server: xlsx_execute_task runs the batch strategy', () => {
  const src = require('node:fs').readFileSync(require.resolve('../../src/server/write-tools.cjs'), 'utf8');
  assert.match(src, /createExecutor\(v\.task, \{ strategy: 'batch'/);
});
