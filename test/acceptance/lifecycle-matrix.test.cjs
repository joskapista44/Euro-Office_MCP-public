'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createFakeSession } = require('../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../fakes/fake-adapter.cjs');
const { createFakeExecutor } = require('../fakes/fake-executor.cjs');
const { hasVerifiedAppliedWrites } = require('../../src/core/persistence/verified-writes.cjs');

/**
 * M1 acceptance matrix A..L — owner requirement 20.
 *
 * Every case states its EXPECTED LIFECYCLE TRACE explicitly and asserts it, so a change that
 * silently reorders or skips a lifecycle step fails here rather than passing on outcome alone.
 */

const FULL_WRITE_TRACE = [
  'CREATED', 'OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING',
  'PERSISTING', 'CLOSING', 'CLOSED', 'FINALIZED',
];
const NO_BARRIER_TRACE = [
  'CREATED', 'OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING',
  'CLOSING', 'CLOSED', 'FINALIZED',
];
const OPEN_FAIL_TRACE = ['CREATED', 'OPENING', 'CLOSING', 'CLOSED', 'FINALIZED'];
const READY_FAIL_TRACE = ['CREATED', 'OPENING', 'READY', 'CLOSING', 'CLOSED', 'FINALIZED'];

/** No test sleeps: the injected clock and waits make every case deterministic. */
function run({ sessionScript = {}, adapterScript = {}, executorScript = {}, task }) {
  let t = 0;
  const now = () => (t += 1);
  const driver = createFakeSession(sessionScript);
  const adapter = createFakeAdapter(adapterScript);
  const options = {
    driver,
    fileId: 'fake-file',
    now,
    ready: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
    persistence: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
  };
  return withSession(options, adapter, createFakeExecutor(task, executorScript)).then(result => ({
    result,
    driver,
    adapter,
  }));
}

const oneOp = { id: 'task-1', operations: [{ id: 'op1' }] };
const twoOps = { id: 'task-2', operations: [{ id: 'op1' }, { id: 'op2' }] };

test('A. read-only success — no operations, no writes, no barrier', async () => {
  const { result, driver } = await run({ task: { id: 'read-only', operations: [] } });

  assert.equal(result.ok, true);
  assert.equal(result.documentOutcome, 'SUCCESS');
  assert.equal(result.persistenceOutcome, 'NOT_REQUIRED');
  assert.equal(result.cleanupOutcome, 'SUCCESS');
  assert.equal(result.metrics.writes, 0);
  assert.equal(driver.calls.forceSave, 0, 'no barrier may run without a verified write');
  assert.deepEqual(result.lifecycleTrace, NO_BARRIER_TRACE);
});

test('B. normal verified-write success — barrier runs, FORCESAVE_CONFIRMED, no general PERSISTED', async () => {
  const { result, driver } = await run({ task: twoOps });

  assert.equal(result.ok, true);
  assert.equal(result.documentOutcome, 'SUCCESS');
  assert.equal(result.persistenceOutcome, 'SUCCESS');
  assert.equal(result.cleanupOutcome, 'SUCCESS');
  assert.equal(result.authority, 'LIVE_VERIFY');
  assert.equal(result.persistenceState, 'FORCESAVE_CONFIRMED');
  assert.equal(result.metrics.writes, 2);
  assert.equal(driver.calls.forceSave, 1);
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('C. no-write failure — barrier does NOT run, close does', async () => {
  const { result, driver } = await run({
    task: oneOp,
    adapterScript: { primitives: { op1: 'dispatch-fail' } },
  });

  assert.equal(result.ok, false);
  assert.equal(result.documentOutcome, 'FAILED');
  assert.equal(result.persistenceOutcome, 'NOT_REQUIRED');
  assert.equal(result.cleanupOutcome, 'SUCCESS');
  assert.equal(hasVerifiedAppliedWrites(result.execution), false);
  assert.equal(driver.calls.forceSave, 0, 'F4: no verified write means no barrier');
  assert.equal(driver.calls.close, 1);
  assert.deepEqual(result.lifecycleTrace, NO_BARRIER_TRACE);
});

test('D. verified write then later failure — F5: the barrier still runs, result is ok:false', async () => {
  const { result, driver } = await run({
    task: twoOps,
    adapterScript: { primitives: { op2: 'verify-fail' } },
  });

  assert.equal(result.ok, false, 'a failing task is never reported as success');
  assert.equal(result.documentOutcome, 'FAILED');
  assert.equal(result.persistenceOutcome, 'SUCCESS', 'the applied prefix was persisted');
  assert.equal(result.cleanupOutcome, 'SUCCESS');
  assert.equal(driver.calls.forceSave, 1, 'F5: the barrier MUST run on partial failure');

  // Operation 1 is not rolled back: NON_ATOMIC stop-on-first-failure (M0 section 10).
  const applied = result.execution.appliedOperations;
  assert.equal(applied[0].operationId, 'op1');
  assert.equal(applied[0].status, 'APPLIED');
  assert.equal(applied[0].verification, 'LIVE_VERIFY');
  assert.equal(applied[1].status, 'FAILED');
  assert.equal(result.metrics.operationsApplied, 1);
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('E. persistence failure — F7: ok:false, never reported as persisted, close still attempted', async () => {
  const { result, driver } = await run({
    task: oneOp,
    sessionScript: { canSaveSequence: [true] }, // canSave never goes false -> timeout
  });

  assert.equal(result.ok, false);
  assert.equal(result.persistenceOutcome, 'FAILED');
  assert.equal(result.outcome, 'task-persistence-barrier-failed');
  assert.equal(result.authority, 'PRIMITIVE_LIVE_VERIFY_ONLY', 'PERSISTED is never claimed');
  assert.notEqual(result.authority, 'PERSISTED');
  assert.equal(result.persistence.outcome, 'forceSave-timeout');
  assert.equal(driver.calls.close, 1, 'close is attempted after a failed barrier');
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('F. close failure — F8: semantic and persistence truth kept, cleanup FAILED, top-level ok:false', async () => {
  const { result, driver } = await run({ task: oneOp, sessionScript: { closeFails: true } });

  assert.equal(result.documentOutcome, 'SUCCESS', 'a close failure does not rewrite document truth');
  assert.equal(result.persistenceOutcome, 'SUCCESS', 'nor persistence truth');
  assert.equal(result.cleanupOutcome, 'FAILED');
  assert.equal(result.ok, false, 'ok means COMPLETE lifecycle success');
  assert.equal(result.outcome, 'task-cleanup-failed', 'close failure is explicit, never metrics-only');
  assert.equal(driver.calls.close, 1);
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('G. SATISFIED no-op — zero writes, no barrier, success with noOp:true', async () => {
  const { result, driver, adapter } = await run({
    task: oneOp,
    adapterScript: { classification: 'SATISFIED' },
  });

  assert.equal(result.ok, true);
  assert.equal(result.noOp, true);
  assert.equal(result.metrics.writes, 0, 'R1: SATISFIED implies zero writes');
  assert.equal(result.metrics.operationsSatisfied, 1);
  assert.equal(adapter.calls.executePrimitive, 0, 'the core refuses a primitive on a SATISFIED step');
  assert.equal(result.persistenceOutcome, 'NOT_REQUIRED');
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(driver.calls.close, 1);
  assert.deepEqual(result.lifecycleTrace, NO_BARRIER_TRACE);
});

test('H. corrupt receipt with an authoritative verified write — the barrier still runs', async () => {
  const { result, driver } = await run({
    task: twoOps,
    adapterScript: { primitives: { op2: 'verify-fail' } },
    executorScript: { receiptCorrupt: true },
  });

  // The receipt is empty/unserialized...
  assert.equal(result.receiptEvidence.serialized, false);
  assert.equal(result.receiptEvidence.operationCount, 0);
  // ...while the AUTHORITATIVE state still carries the verified write.
  assert.equal(hasVerifiedAppliedWrites(result.execution), true);

  assert.equal(driver.calls.forceSave, 1, 'PERSISTENCE_BARRIER_CALLED');
  assert.equal(driver.calls.close, 1, 'CLOSE_CALLED');
  assert.equal(result.ok, false, 'FINAL_OK=false');
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('I. open failure — F1: no handle, no close, no barrier', async () => {
  const { result, driver } = await run({ task: oneOp, sessionScript: { openFails: true } });

  assert.equal(result.ok, false);
  assert.equal(result.documentOutcome, 'FAILED');
  assert.equal(result.execution.outcome, 'session-open-failed');
  assert.equal(result.execution.authority, 'PLAN_ONLY');
  assert.equal(result.metrics.writes, 0);
  assert.equal(driver.calls.close, 0, 'there is no session to close');
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(result.cleanupOutcome, 'SUCCESS');
  assert.equal(result.cleanup.outcome, 'not-applicable');
  assert.deepEqual(result.lifecycleTrace, OPEN_FAIL_TRACE);
});

test('J. ready failure — F2: close runs, no write, no barrier', async () => {
  const { result, driver } = await run({ task: oneOp, adapterScript: { readyFails: true } });

  assert.equal(result.ok, false);
  assert.equal(result.execution.outcome, 'session-ready-failed');
  assert.equal(result.execution.authority, 'PLAN_ONLY');
  assert.equal(result.metrics.writes, 0);
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(driver.calls.close, 1, 'the session opened, so it must close');
  assert.deepEqual(result.lifecycleTrace, READY_FAIL_TRACE);
});

test('K. executor throws after a verified write — the evidence survives and the barrier runs', async () => {
  const { result, driver } = await run({
    task: twoOps,
    executorScript: { throwAfter: 'op2' },
  });

  assert.equal(result.execution.outcome, 'execution-threw');
  assert.match(result.execution.executionError, /fake executor threw at op2/);
  assert.equal(hasVerifiedAppliedWrites(result.execution), true, 'op1 was verified before the throw');
  assert.equal(driver.calls.forceSave, 1, 'a throw must not lose a verified write');
  assert.equal(result.persistenceOutcome, 'SUCCESS');
  assert.equal(result.ok, false);
  assert.equal(driver.calls.close, 1);
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('L. deterministic close exactly once — across every outcome in the matrix', async () => {
  const cases = [
    { name: 'success', args: { task: twoOps } },
    { name: 'partial failure', args: { task: twoOps, adapterScript: { primitives: { op2: 'verify-fail' } } } },
    { name: 'persistence timeout', args: { task: oneOp, sessionScript: { canSaveSequence: [true] } } },
    { name: 'executor throw', args: { task: twoOps, executorScript: { throwAfter: 'op1' } } },
    { name: 'ready failure', args: { task: oneOp, adapterScript: { readyFails: true } } },
    { name: 'whole-task failure', args: { task: oneOp, adapterScript: { wholeTaskFails: true } } },
  ];

  for (const c of cases) {
    const { driver } = await run(c.args);
    assert.equal(driver.calls.open, 1, `OPEN_COUNT must be 1 (${c.name})`);
    assert.equal(driver.calls.close, 1, `CLOSE_COUNT must be 1 (${c.name})`);
  }
});
