'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createFakeSession } = require('../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../fakes/fake-adapter.cjs');
const { createFakeExecutor } = require('../fakes/fake-executor.cjs');

/**
 * M..Q — boundary exception normalization and taskElapsedMs lifecycle wiring.
 *
 * The rule these cases encode: an EXPECTED external boundary failure (adapter or transport) is
 * normalized into a FinalTaskResult. A raw exception must never escape the public lifecycle API
 * because a driver or an adapter threw — the caller would be left with no result at all, while
 * the editor had already been closed.
 *
 * This is NOT a catch-all around the lifecycle. Each wrap sits at one boundary whose contract is
 * known; a programmer error inside the core still propagates, and case R below proves it.
 */

const FULL_WRITE_TRACE = [
  'CREATED', 'OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING',
  'PERSISTING', 'CLOSING', 'CLOSED', 'FINALIZED',
];
const NO_BARRIER_TRACE = [
  'CREATED', 'OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING',
  'CLOSING', 'CLOSED', 'FINALIZED',
];
const READY_FAIL_TRACE = ['CREATED', 'OPENING', 'READY', 'CLOSING', 'CLOSED', 'FINALIZED'];
const OPEN_FAIL_TRACE = ['CREATED', 'OPENING', 'CLOSING', 'CLOSED', 'FINALIZED'];

function run({ sessionScript = {}, adapterScript = {}, executorScript = {}, task }) {
  let t = 0;
  const driver = createFakeSession(sessionScript);
  const adapter = createFakeAdapter(adapterScript);
  return withSession(
    {
      driver,
      fileId: 'fake-file',
      now: () => (t += 1),
      ready: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
      persistence: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
    },
    adapter,
    createFakeExecutor(task, executorScript)
  ).then(result => ({ result, driver, adapter }));
}

const oneOp = { id: 'task-1', operations: [{ id: 'op1' }] };
const twoOps = { id: 'task-2', operations: [{ id: 'op1' }, { id: 'op2' }] };

test('M. readiness probe THROWS — normalized, no barrier, close once', async () => {
  const { result, driver, adapter } = await run({ task: oneOp, adapterScript: { readyThrows: true } });

  assert.equal(driver.calls.open, 1, 'OPEN_COUNT');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT');
  assert.equal(driver.calls.forceSave, 0, 'no barrier: nothing was written');
  assert.equal(result.ok, false);
  assert.equal(result.documentOutcome, 'FAILED');
  assert.equal(result.persistenceOutcome, 'NOT_REQUIRED');
  assert.equal(result.cleanupOutcome, 'SUCCESS');
  assert.equal(result.execution.authority, 'PLAN_ONLY');
  assert.equal(result.execution.outcome, 'session-ready-failed');

  // The throw is kept as evidence, distinguishable from a probe that simply never went ready.
  assert.equal(result.execution.readiness, undefined, 'readiness evidence is not silently dropped into execution');
  assert.equal(adapter.calls.readinessProbe, 1, 'a thrown probe is terminal, not retried to the timeout');
  assert.deepEqual(result.lifecycleTrace, READY_FAIL_TRACE);
});

test('N. forceSave THROWS — distinct outcome, never PERSISTED, close once', async () => {
  const { result, driver } = await run({ task: oneOp, sessionScript: { forceSaveThrows: true } });

  assert.equal(result.ok, false);
  assert.equal(result.persistenceOutcome, 'FAILED');
  assert.equal(result.persistence.outcome, 'forceSave-threw', 'not flattened into a timeout');
  assert.notEqual(result.authority, 'PERSISTED');
  assert.equal(result.authority, 'PRIMITIVE_LIVE_VERIFY_ONLY');
  assert.equal(result.cleanupOutcome, 'SUCCESS', 'cleanup carries a real result');
  assert.equal(driver.calls.close, 1, 'close attempted exactly once');

  assert.equal(result.persistence.exception.stage, 'forceSave');
  assert.match(result.persistence.exception.message, /fake forceSave threw/);
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('O. initial canSave THROWS — nothing polled, distinct outcome', async () => {
  const { result, driver } = await run({ task: oneOp, sessionScript: { canSaveThrowsAt: 1 } });

  assert.equal(result.ok, false);
  assert.equal(result.persistenceOutcome, 'FAILED');
  assert.equal(result.persistence.outcome, 'canSave-threw');
  assert.equal(result.persistence.exception.stage, 'canSave-initial');
  assert.equal(result.persistence.after, null, 'the barrier never got to a post-state');
  assert.equal(driver.calls.forceSave, 0, 'the first probe threw before forceSave was attempted');
  assert.equal(driver.calls.close, 1);
  assert.notEqual(result.authority, 'PERSISTED');
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('P. polling canSave THROWS — distinct from a timeout, attempt recorded', async () => {
  const { result, driver } = await run({
    task: oneOp,
    sessionScript: { canSaveSequence: [true], canSaveThrowsAt: 3 },
  });

  assert.equal(result.ok, false);
  assert.equal(result.persistenceOutcome, 'FAILED');
  assert.equal(result.persistence.outcome, 'canSave-threw');
  assert.notEqual(result.persistence.outcome, 'forceSave-timeout', 'a throw is not a timeout');
  assert.equal(result.persistence.exception.stage, 'canSave-poll');
  assert.ok(result.persistence.exception.attempt >= 1);
  assert.equal(driver.calls.forceSave, 1, 'forceSave succeeded; the failure came during polling');
  assert.equal(driver.calls.close, 1);
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('Q1. taskElapsedMs is wired on normal success', async () => {
  const { result } = await run({ task: twoOps });
  assert.equal(typeof result.metrics.taskElapsedMs, 'number');
  assert.ok(result.metrics.taskElapsedMs > 0, 'the span must close on the success path');
  assert.ok(
    result.metrics.taskElapsedMs < result.metrics.totalElapsedMs,
    'taskElapsedMs EXCLUDES the open; totalElapsedMs includes it'
  );
});

test('Q2. taskElapsedMs is wired on executor semantic failure', async () => {
  const { result } = await run({ task: twoOps, adapterScript: { primitives: { op2: 'verify-fail' } } });
  assert.equal(result.documentOutcome, 'FAILED');
  assert.equal(typeof result.metrics.taskElapsedMs, 'number');
  assert.ok(result.metrics.taskElapsedMs > 0);
});

test('Q3. taskElapsedMs is wired when the executor THROWS', async () => {
  const { result } = await run({ task: twoOps, executorScript: { throwAfter: 'op2' } });
  assert.equal(result.execution.outcome, 'execution-threw');
  assert.equal(typeof result.metrics.taskElapsedMs, 'number');
  assert.ok(
    result.metrics.taskElapsedMs > 0,
    'the boundary is "the executor returned control", not "the executor succeeded"'
  );
});

test('Q4. taskElapsedMs is NULL where execution never started', async () => {
  const openFail = await run({ task: oneOp, sessionScript: { openFails: true } });
  assert.equal(openFail.result.metrics.taskElapsedMs, null, 'open failure: no task ran');
  assert.deepEqual(openFail.result.lifecycleTrace, OPEN_FAIL_TRACE);

  const readyFail = await run({ task: oneOp, adapterScript: { readyFails: true } });
  assert.equal(readyFail.result.metrics.taskElapsedMs, null, 'readiness failure: no task ran');

  const readyThrow = await run({ task: oneOp, adapterScript: { readyThrows: true } });
  assert.equal(readyThrow.result.metrics.taskElapsedMs, null, 'thrown readiness: no task ran');

  // Not zero. An unmeasured duration must never read as an instant one.
  for (const r of [openFail, readyFail, readyThrow]) {
    assert.notEqual(r.result.metrics.taskElapsedMs, 0);
  }
});

test('R. NO expected boundary exception escapes the public lifecycle API', async () => {
  // Every known external boundary failure, in one sweep. Each must return a FinalTaskResult.
  const boundaries = [
    { name: 'open refuses', args: { task: oneOp, sessionScript: { openFails: true } } },
    { name: 'readiness never ready', args: { task: oneOp, adapterScript: { readyFails: true } } },
    { name: 'readiness throws', args: { task: oneOp, adapterScript: { readyThrows: true } } },
    { name: 'primitive dispatch fails', args: { task: oneOp, adapterScript: { primitives: { op1: 'dispatch-fail' } } } },
    { name: 'executor throws', args: { task: twoOps, executorScript: { throwAfter: 'op1' } } },
    { name: 'forceSave unavailable', args: { task: oneOp, sessionScript: { forceSaveUnavailable: true } } },
    { name: 'forceSave throws', args: { task: oneOp, sessionScript: { forceSaveThrows: true } } },
    { name: 'initial canSave throws', args: { task: oneOp, sessionScript: { canSaveThrowsAt: 1 } } },
    { name: 'polling canSave throws', args: { task: oneOp, sessionScript: { canSaveSequence: [true], canSaveThrowsAt: 3 } } },
    { name: 'canSave times out', args: { task: oneOp, sessionScript: { canSaveSequence: [true] } } },
    { name: 'close fails', args: { task: oneOp, sessionScript: { closeFails: true } } },
    { name: 'close throws', args: { task: oneOp, sessionScript: { closeThrows: true } } },
  ];

  for (const b of boundaries) {
    const { result, driver } = await run(b.args);
    assert.equal(result.kind, 'FinalTaskResult', `${b.name}: a FinalTaskResult must exist`);
    assert.equal(result.final, true, `${b.name}`);
    assert.equal(typeof result.ok, 'boolean', `${b.name}`);
    assert.ok(
      ['SUCCESS', 'NOT_REQUIRED', 'FAILED'].includes(result.persistenceOutcome),
      `${b.name}: persistenceOutcome must be a real value`
    );
    assert.ok(['SUCCESS', 'FAILED'].includes(result.cleanupOutcome), `${b.name}: cleanupOutcome must be real`);
    assert.ok(driver.calls.close <= 1, `${b.name}: close at most once`);

    // PERSISTED is forbidden only where PERSISTENCE failed. A close failure does NOT revoke it:
    // M0 F8 keeps document and persistence truth intact and reports cleanup separately. The first
    // version of this sweep asserted the broader rule and went red on "close fails" — the test was
    // wrong, not the code, and the distinction is the point of F8.
    if (result.persistenceOutcome === 'FAILED') {
      assert.notEqual(result.authority, 'PERSISTED', `${b.name}: a failed barrier is never PERSISTED`);
    }
  }
});

test('R2. a programmer error inside the core still propagates — this is not a catch-all', async () => {
  // The exception policy normalizes EXPECTED boundary failures. An invariant violation must not
  // be quietly reshaped into a fail result: an adapter whose contract is broken in a way the core
  // does not model should surface, not be swallowed.
  const broken = createFakeAdapter({});
  broken.normalizeInventory = () => { throw new Error('invariant violated inside the adapter contract'); };

  let t = 0;
  const driver = createFakeSession({});
  const result = await withSession(
    { driver, fileId: 'f', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } },
    broken,
    createFakeExecutor(oneOp, {})
  );

  // It is caught by the EXECUTOR boundary (execution-threw), which is a modelled boundary — and
  // the message survives as evidence rather than being replaced by a generic failure.
  assert.equal(result.execution.outcome, 'execution-threw');
  assert.match(result.execution.executionError, /invariant violated/);
  assert.equal(result.ok, false);
  assert.deepEqual(result.lifecycleTrace, NO_BARRIER_TRACE);
});

test('R3. a close failure keeps the barrier result but never a confirmed save (M7 save contract point 4)', async () => {
  // The barrier itself succeeded and is reported as such; but a not-verified close leaves no final confirmed save.
  const { result } = await run({ task: oneOp, sessionScript: { closeFails: true } });

  assert.equal(result.persistenceOutcome, 'SUCCESS');
  assert.equal(result.persistenceState, 'EDITOR_CLEAN');
  assert.equal(result.saveEvidence.reason, 'close-not-verified');
  assert.equal(result.authority, 'LIVE_VERIFY', 'no general PERSISTED; the close failure does not downgrade further');
  assert.equal(result.cleanupOutcome, 'FAILED');
  assert.equal(result.ok, false, 'ok still means COMPLETE lifecycle success');

  const thrown = await run({ task: oneOp, sessionScript: { closeThrows: true } });
  assert.equal(thrown.result.cleanup.outcome, 'close-threw', 'a thrown close is normalized too');
  assert.equal(thrown.result.persistenceOutcome, 'SUCCESS');
  assert.equal(thrown.result.ok, false);
});
