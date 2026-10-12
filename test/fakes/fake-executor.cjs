'use strict';

const {
  buildExecutionResult,
  OPERATION_STATUS,
  VERIFICATION,
  AUTHORITY,
  CLASSIFICATION,
} = require('../../src/core/execution/result.cjs');
const { buildReceipt } = require('../../src/core/receipt/envelope.cjs');

/**
 * A minimal task executor used to drive the M1 lifecycle.
 *
 * The full executor with its three strategies is M3. What M1 needs is a driver that produces a
 * contract-correct ExecutionResult so the LIFECYCLE can be proven: one open, one close, the
 * persistence predicate, and the failure paths F1-F8. It lives under test/ precisely because it
 * is not the production executor.
 *
 * It implements the NON_ATOMIC stop-on-first-failure contract of M0 section 10: dispatch stops at
 * the first primitive/L1 failure, and operations already applied are NOT rolled back.
 */
function createFakeExecutor(task, script = {}) {
  return async function execute(handle, { appliedOperations, adapter }) {
    const metrics = handle.metrics;

    // Idempotency classification reads the CURRENT observation (R3), not a receipt.
    const observation = adapter.normalizeInventory(script.livePayload || { units: [] });
    const classification = adapter.classifyIdempotency(observation, task);

    if (classification.state === CLASSIFICATION.SATISFIED) {
      // R1: SATISFIED implies writes = 0. The core refuses a primitive call on a SATISFIED step.
      for (const op of task.operations || []) {
        appliedOperations.record({
          operationId: op.id,
          status: OPERATION_STATUS.SATISFIED,
          verification: VERIFICATION.NONE,
          evidence: classification.evidence,
        });
        metrics.count('operationsPlanned').count('operationsSatisfied');
      }
      return finish({ executionOk: true, outcome: 'task-satisfied', noOp: true, authority: AUTHORITY.LIVE_READ });
    }

    if (classification.state === CLASSIFICATION.FAIL_CLOSED) {
      for (const op of task.operations || []) metrics.count('operationsPlanned');
      return finish({ executionOk: false, outcome: 'task-fail-closed', authority: AUTHORITY.PLAN_ONLY });
    }

    const plan = adapter.planOperations(task, observation);
    let weakest = AUTHORITY.LIVE_VERIFY;

    for (const operation of plan.operations) {
      metrics.count('operationsPlanned');

      if (script.throwAfter === operation.id) {
        // Owner requirement 20/K: the throw happens AFTER earlier verified writes were recorded.
        throw new Error(`fake executor threw at ${operation.id}`);
      }

      const dispatched = await adapter.executePrimitive(handle, operation);
      if (!dispatched.ok) {
        appliedOperations.record({
          operationId: operation.id,
          status: OPERATION_STATUS.FAILED,
          verification: VERIFICATION.NONE,
          evidence: { stage: 'dispatch' },
        });
        // Stop-on-first-failure: no further dispatch, and no rollback of what already applied.
        return finish({ executionOk: false, outcome: 'primitive-dispatch-failed', authority: weakest });
      }

      metrics.count('verificationInvocations');
      const verified = await adapter.verifyPrimitive(handle, operation, null, null);
      if (!verified.ok) {
        appliedOperations.record({
          operationId: operation.id,
          status: OPERATION_STATUS.FAILED,
          verification: VERIFICATION.DISPATCH_ONLY,
          evidence: verified.evidence,
        });
        weakest = AUTHORITY.DISPATCH_ONLY;
        return finish({ executionOk: false, outcome: 'primitive-verify-failed', authority: weakest });
      }

      appliedOperations.record({
        operationId: operation.id,
        status: OPERATION_STATUS.APPLIED,
        verification: VERIFICATION.LIVE_VERIFY,
        evidence: verified.evidence,
      });
      metrics.count('operationsApplied').count('writes');
    }

    metrics.count('verificationInvocations');
    const wholeTask = await adapter.verifyWholeTask(handle, task, observation);
    if (!wholeTask.ok) {
      // F6: whole-task verification failed while verified writes exist -> same route as F5.
      return finish({
        executionOk: false,
        outcome: 'whole-task-verification-failed',
        authority: weakest,
        wholeTaskVerification: wholeTask,
      });
    }

    return finish({ executionOk: true, outcome: 'execution-complete', authority: weakest, wholeTaskVerification: wholeTask });

    function finish({ executionOk, outcome, authority, noOp = false, wholeTaskVerification = null }) {
      const base = buildExecutionResult({
        executionOk,
        outcome,
        authority,
        noOp,
        appliedOperations,
        classification: classification.state,
        wholeTaskVerification,
      });
      // The receipt is derived LAST, from state that already decided everything. Owner
      // requirement 9 is scripted here: a corrupt receipt must not suppress the barrier.
      const receiptEvidence = script.receiptCorrupt
        ? Object.freeze({ kind: 'ReceiptEvidence', authority: 'EVIDENCE_ONLY', serialized: false, serializationError: 'scripted corruption', operations: [], operationCount: 0 })
        : buildReceipt({ execution: base, task, sessionId: handle.id });
      return Object.freeze({ ...base, receiptEvidence });
    }
  };
}

module.exports = { createFakeExecutor };
