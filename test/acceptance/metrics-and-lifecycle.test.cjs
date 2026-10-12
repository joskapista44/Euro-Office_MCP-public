'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createFakeSession } = require('../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../fakes/fake-adapter.cjs');
const { createM2Executor } = require('../fakes/fake-m2-executor.cjs');
const { CONSISTENCY } = require('../../src/core/metrics/ownership.cjs');
const { AUTHORITY } = require('../../src/core/authority.cjs');

/**
 * R and S — metric call-count consistency, and verification failure through the M1 lifecycle.
 *
 * These run the REAL observation and verification layers inside the REAL M1 lifecycle owner.
 * Nothing here is scripted at the outcome level: the numbers are whatever the wiring produces.
 */

const FULL_WRITE_TRACE = [
  'CREATED', 'OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING',
  'PERSISTING', 'CLOSING', 'CLOSED', 'FINALIZED',
];

function run({ sessionScript = {}, adapterScript = {}, executorScript = {}, task }) {
  let t = 0;
  const driver = createFakeSession(sessionScript);
  const adapter = createFakeAdapter(adapterScript);
  return withSession(
    {
      driver, fileId: 'f', now: () => (t += 1),
      ready: { wait: async () => {} }, persistence: { wait: async () => {} },
    },
    adapter,
    createM2Executor(task, executorScript)
  ).then(result => ({ result, driver, adapter }));
}

const twoOps = { id: 'm2-task', operations: [{ id: 'op1' }, { id: 'op2' }] };

test('R. metric call-count consistency — the contract of metrics/ownership.cjs, measured', async () => {
  const { result, driver, adapter } = await run({ task: twoOps, adapterScript: { enrichment: 'ok' } });
  const m = result.metrics;

  assert.equal(result.ok, true, 'the happy path, so every counter is exercised');

  // The contract, pair by pair, exactly as ownership.cjs states it.
  assert.equal(driver.calls.readLive, m.freshReads, 'driver.readLive == freshReads');
  assert.equal(driver.calls.dispatch, m.commands, 'driver.dispatch == commands');
  assert.equal(driver.calls.canSave, m.persistencePolls, 'driver.canSave == persistencePolls');
  // M4 added the common package transport. This happy path's enrichment hook reads no package, so
  // the pair is 0 == 0 here; the non-zero case is measured in test/core/observation/package-transport.
  assert.equal(driver.calls.readPackage, m.packageReads, 'driver.readPackage == packageReads');
  assert.equal(adapter.calls.enrichObservation, m.enrichmentInvocations, 'adapter.enrichObservation == enrichmentInvocations');

  const hookInvocations =
    adapter.calls.verifyPrimitive + adapter.calls.verifyWholeTask + adapter.calls.verifyPreservation;
  assert.equal(hookInvocations, m.verificationInvocations, 'hook invocations == verificationInvocations');

  // The derived totals. PHYSICAL I/O only: a verifier invocation is not a transport interaction.
  assert.equal(m.editorInteractions, m.freshReads + m.commands + m.persistencePolls);
  assert.equal(m.packageInteractions, 0);
  assert.equal(m.totalExternalInteractions, m.editorInteractions + m.packageInteractions);

  // And the numbers are not trivially zero — a consistency proof over zeros proves nothing.
  assert.ok(m.freshReads >= 3, `expected planning + one re-read per op, got ${m.freshReads}`);
  assert.equal(m.commands, 2);
  assert.equal(m.writes, 2);
  assert.ok(m.enrichmentInvocations >= 3, 'enrichment ran on every observation');
  assert.ok(m.verificationInvocations >= 3, 'two primitives plus the whole task');
  assert.ok(m.totalExternalInteractions > 0);

  // Every pair named in the contract was actually asserted above.
  assert.equal(Object.keys(CONSISTENCY).length, 6);
});

test('R2. the same identities hold on a FAILING path', async () => {
  // Consistency that only holds on the happy path is not consistency.
  const { result, driver, adapter } = await run({
    task: twoOps,
    adapterScript: { enrichment: 'ok', primitives: { op2: 'verify-fail' } },
  });
  const m = result.metrics;

  assert.equal(result.ok, false);
  assert.equal(driver.calls.readLive, m.freshReads);
  assert.equal(driver.calls.dispatch, m.commands);
  assert.equal(driver.calls.canSave, m.persistencePolls);
  assert.equal(adapter.calls.enrichObservation, m.enrichmentInvocations);
  assert.equal(
    adapter.calls.verifyPrimitive + adapter.calls.verifyWholeTask + adapter.calls.verifyPreservation,
    m.verificationInvocations
  );
  assert.equal(m.totalExternalInteractions, m.editorInteractions + m.packageInteractions);
});

test('R3. a degraded optional enrichment is still counted — the package WAS read', async () => {
  const { result, adapter } = await run({ task: twoOps, adapterScript: { enrichment: 'throws' } });

  assert.equal(result.ok, true, 'OPTIONAL degradation does not fail the task');
  assert.equal(result.metrics.enrichmentInvocations, adapter.calls.enrichObservation);
  assert.ok(result.metrics.enrichmentInvocations > 0, 'a hook that threw was still invoked, and is counted');
});

test('S. whole-task FAIL after verified writes — persists through the M1 lifecycle', async () => {
  const { result, driver } = await run({ task: twoOps, adapterScript: { wholeTaskFails: true } });

  // Every primitive was applied and live-verified...
  const applied = result.execution.appliedOperations;
  assert.equal(applied.length, 2);
  for (const e of applied) {
    assert.equal(e.status, 'APPLIED');
    assert.equal(e.verification, 'LIVE_VERIFY');
  }

  // ...and the task still fails on the terminal check.
  assert.equal(result.execution.outcome, 'whole-task-verification-failed');
  assert.equal(result.ok, false);
  assert.equal(result.documentOutcome, 'FAILED');

  // M1 rule, unchanged and NOT reimplemented in M2: verified writes must persist anyway.
  assert.equal(driver.calls.forceSave, 1, 'the barrier ran');
  assert.equal(result.persistenceOutcome, 'SUCCESS');
  assert.equal(driver.calls.close, 1, 'and close ran after it');
  assert.equal(result.persistenceState, 'FORCESAVE_CONFIRMED', 'the applied prefix reached a confirmed save');
  assert.notEqual(result.authority, AUTHORITY.PERSISTED, 'there is no general PERSISTED');
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('S2. preservation failure after a verified write takes the same route', async () => {
  const { result, driver } = await run({
    task: { id: 'p', operations: [{ id: 'op1' }] },
    adapterScript: { preservation: 'fails' },
    executorScript: { preservationRequired: true },
  });

  assert.equal(result.execution.outcome, 'preservation-failed');
  assert.equal(result.ok, false);
  assert.equal(result.execution.appliedOperations[0].verification, 'LIVE_VERIFY');
  assert.equal(driver.calls.forceSave, 1, 'a verified write persists even when preservation fails');
  assert.equal(result.persistenceOutcome, 'SUCCESS');
  assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
});

test('S3. an observation boundary failure before any write takes NO barrier', async () => {
  const { result, driver } = await run({ task: twoOps, sessionScript: { readLiveThrows: true } });

  assert.equal(result.ok, false);
  assert.equal(result.execution.outcome, 'observation-live-read-threw');
  assert.equal(result.execution.authority, AUTHORITY.PLAN_ONLY);
  assert.equal(result.metrics.writes, 0);
  assert.equal(driver.calls.forceSave, 0, 'nothing was written, so nothing is persisted');
  assert.equal(driver.calls.close, 1);
  assert.equal(result.persistenceOutcome, 'NOT_REQUIRED');
});
