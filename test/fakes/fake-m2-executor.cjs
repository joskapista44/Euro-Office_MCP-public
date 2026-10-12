'use strict';

const {
  buildExecutionResult, OPERATION_STATUS, VERIFICATION, AUTHORITY, CLASSIFICATION,
} = require('../../src/core/execution/result.cjs');
const { buildReceipt } = require('../../src/core/receipt/envelope.cjs');
const { observeFresh, requireMutationAuthority } = require('../../src/core/observation/observe.cjs');
const { verifyPrimitive, verifyWholeTask, verifyPreservation, fromDispatch } =
  require('../../src/core/verification/orchestrate.cjs');
const { VERDICT, isSatisfied } = require('../../src/core/verification/layers.cjs');
const { dispatch } = require('../../src/core/session/handle.cjs');
const { weakest } = require('../../src/core/authority.cjs');

/**
 * A task executor wired to the REAL M2 observation and verification layers.
 *
 * The M1 fake executor scripted its outcomes directly; this one actually walks the contract:
 * observe -> target -> dispatch -> fresh re-read -> L1 -> (L3) -> L2. That is what makes the
 * metric consistency and the whole-task/persistence integration provable rather than asserted.
 *
 * The full executor with its three strategies is still M3. This is the smallest driver that
 * exercises M2 through the M1 lifecycle.
 */
function createM2Executor(task, script = {}) {
  return async function execute(handle, { appliedOperations, adapter }) {
    const metrics = handle.metrics;
    const enrichment = script.enrichment || undefined;

    // EVERY variable the `finish()` closure reads is declared here, before the first call site.
    // A const declared further down sits in its temporal dead zone at every early return, and
    // reading it throws. Measured twice while writing this file: an observation boundary failure
    // came back as `execution-threw`, hiding the real outcome behind a ReferenceError. The
    // lifecycle normalized it correctly — which is exactly why the bug was quiet.
    let classification = null;
    const primitiveVerification = [];

    const planning = await observeFresh(handle, adapter, enrichment ? { enrichment } : {});
    if (!planning.ok) {
      return finish({ executionOk: false, outcome: planning.outcome, authority: AUTHORITY.PLAN_ONLY });
    }

    classification = adapter.classifyIdempotency(planning.primary, task);
    if (classification.state === CLASSIFICATION.SATISFIED) {
      for (const o of task.operations || []) {
        appliedOperations.record({ operationId: o.id, status: OPERATION_STATUS.SATISFIED, evidence: classification.evidence });
        metrics.count('operationsPlanned').count('operationsSatisfied');
      }
      return finish({ executionOk: true, outcome: 'task-satisfied', noOp: true, authority: AUTHORITY.LIVE_READ });
    }

    const plan = adapter.planOperations(task, planning);
    let observation = planning;
    let lowest = AUTHORITY.LIVE_VERIFY;

    for (const operation of plan.operations) {
      metrics.count('operationsPlanned');

      // Mutation targeting binds to the CURRENT observation, never a stored one.
      const gate = requireMutationAuthority(observation, handle, 'mutation-targeting');
      if (!gate.ok) {
        appliedOperations.record({ operationId: operation.id, status: OPERATION_STATUS.FAILED, evidence: gate });
        return finish({ executionOk: false, outcome: gate.outcome, authority: AUTHORITY.PLAN_ONLY });
      }

      const before = observation;
      const dispatched = await dispatch(handle, operation); // bumps the mutation generation
      const l0 = fromDispatch(dispatched);
      if (l0.verdict !== VERDICT.PASS) {
        appliedOperations.record({ operationId: operation.id, status: OPERATION_STATUS.FAILED, evidence: l0 });
        return finish({ executionOk: false, outcome: 'primitive-dispatch-failed', authority: weakest(lowest, l0.authority) });
      }

      // The dispatch invalidated `before`; L1 needs a FRESH read, and the orchestrator enforces it.
      observation = await observeFresh(handle, adapter, enrichment ? { enrichment } : {});
      if (!observation.ok) {
        appliedOperations.record({ operationId: operation.id, status: OPERATION_STATUS.FAILED, verification: VERIFICATION.DISPATCH_ONLY, evidence: observation });
        return finish({ executionOk: false, outcome: observation.outcome, authority: AUTHORITY.DISPATCH_ONLY });
      }

      const l1 = await verifyPrimitive({ handle, adapter, operation, before, after: observation });
      primitiveVerification.push(l1);
      if (l1.verdict !== VERDICT.PASS) {
        appliedOperations.record({ operationId: operation.id, status: OPERATION_STATUS.FAILED, verification: VERIFICATION.DISPATCH_ONLY, evidence: l1 });
        return finish({ executionOk: false, outcome: `primitive-${l1.verdict.toLowerCase()}`, authority: AUTHORITY.DISPATCH_ONLY });
      }

      if (script.preservationRequired) {
        const l3 = await verifyPreservation({ handle, adapter, operation, before, after: observation, required: true });
        if (!isSatisfied(l3.verdict)) {
          // The write IS applied and live-verified; the task fails on preservation. The M1
          // lifecycle still routes it through the barrier.
          appliedOperations.record({ operationId: operation.id, status: OPERATION_STATUS.APPLIED, verification: VERIFICATION.LIVE_VERIFY, evidence: l1.evidence });
          metrics.count('operationsApplied').count('writes');
          return finish({ executionOk: false, outcome: 'preservation-failed', authority: AUTHORITY.LIVE_VERIFY, preservationVerification: l3 });
        }
      }

      appliedOperations.record({ operationId: operation.id, status: OPERATION_STATUS.APPLIED, verification: VERIFICATION.LIVE_VERIFY, evidence: l1.evidence });
      metrics.count('operationsApplied').count('writes');
    }

    const l2 = await verifyWholeTask({ handle, adapter, task, finalObservation: observation });
    if (l2.verdict !== VERDICT.PASS) {
      // Every primitive passed and the task is still wrong. Verified writes exist, so the M1
      // lifecycle must still run the barrier (owner requirements 11 and 15).
      return finish({ executionOk: false, outcome: 'whole-task-verification-failed', authority: AUTHORITY.LIVE_VERIFY, wholeTaskVerification: l2 });
    }

    return finish({ executionOk: true, outcome: 'execution-complete', authority: lowest, wholeTaskVerification: l2 });

    function finish({ executionOk, outcome, authority, noOp = false, wholeTaskVerification = null, preservationVerification = null }) {
      const base = buildExecutionResult({
        executionOk, outcome, authority, noOp, appliedOperations,
        classification: classification ? classification.state : null,
        primitiveVerification,
        wholeTaskVerification,
        preservationVerification,
      });
      return Object.freeze({ ...base, receiptEvidence: buildReceipt({ execution: base, task, sessionId: handle.id }) });
    }
  };
}

module.exports = { createM2Executor };
