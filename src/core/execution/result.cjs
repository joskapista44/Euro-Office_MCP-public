'use strict';

/**
 * ExecutionResult — M0 section 4 (`executeTask` OUTPUT) and 4.2.
 *
 * ExecutionResult reports execution/semantic state only. It is NEVER final task completion:
 * persistence and cleanup belong to the lifecycle owner and have not run yet when this exists.
 *
 * `appliedOperations[]` is the authoritative execution state. It is maintained here, directly,
 * independently of any receipt serialization — because the persistence decision reads it and must
 * not be reachable from a receipt (M0 8.1, owner requirements 7 and 8).
 */

const OPERATION_STATUS = Object.freeze({
  APPLIED: 'APPLIED',
  SATISFIED: 'SATISFIED',
  FAILED: 'FAILED',
  // M3 hybrid uncertain-mutation policy (owner decision after W3.0): a dispatch was attempted and
  // its outcome cannot be proven either way. It is NOT APPLIED, it never carries LIVE_VERIFY, and
  // nothing -- persistence included -- may promote it into a verified write.
  DISPATCHED_OUTCOME_UNKNOWN: 'DISPATCHED_OUTCOME_UNKNOWN',
  // M3: an operation that was planned but never attempted because an earlier one stopped the
  // task. Recorded explicitly so that stop-on-failure cannot make the tail of a plan disappear.
  SKIPPED_DUE_TO_PRIOR_FAILURE: 'SKIPPED_DUE_TO_PRIOR_FAILURE',
});

/**
 * Why an operation is uncertain. Only the cases the M3 audit measured; the executor must not guess
 * whether a write happened (owner decision, point 4). A refusal PROVEN to be pre-dispatch is
 * CERTAIN_NONE and is recorded as FAILED, never as one of these.
 */
const UNCERTAINTY_REASON = Object.freeze({
  DISPATCH_THREW: 'DISPATCH_THREW',
  DISPATCH_REFUSED_NOT_PROVEN_PRE_DISPATCH: 'DISPATCH_REFUSED_NOT_PROVEN_PRE_DISPATCH',
  VERIFY_CANNOT_PROVE_TERMINAL_STATE: 'VERIFY_CANNOT_PROVE_TERMINAL_STATE',
  ADAPTER_REPORTED_PARTIALLY_APPLIED: 'ADAPTER_REPORTED_PARTIALLY_APPLIED',
  // L1 passed, but the operation declared its preservation PRIMITIVE_INTEGRAL and that proof did
  // not pass: the write may have happened, and the operation's contract is not proven.
  INTEGRAL_PRESERVATION_NOT_PROVEN: 'INTEGRAL_PRESERVATION_NOT_PROVEN',
  // Recorded by the lifecycle owner, not the executor: the executor threw while the session's
  // mutation generation shows more dispatches than the register accounts for.
  EXECUTOR_THREW_WITH_UNACCOUNTED_DISPATCH: 'EXECUTOR_THREW_WITH_UNACCOUNTED_DISPATCH',
  // A batch hook broke its contract (more than one physical dispatch for one round). What reached
  // the editor is unknown per operation, so every operation of the round is uncertain.
  BATCH_HOOK_CONTRACT_VIOLATION: 'BATCH_HOOK_CONTRACT_VIOLATION',
});

class UncertainPromotionError extends Error {
  constructor(operationId) {
    super(`operation ${operationId} is DISPATCHED_OUTCOME_UNKNOWN and cannot be promoted to a verified write`);
    this.name = 'UncertainPromotionError';
    this.operationId = operationId;
  }
}

const VERIFICATION = Object.freeze({
  LIVE_VERIFY: 'LIVE_VERIFY',
  DISPATCH_ONLY: 'DISPATCH_ONLY',
  NONE: 'NONE',
});

const AUTHORITY = Object.freeze({
  PLAN_ONLY: 'PLAN_ONLY',
  DISPATCH_ONLY: 'DISPATCH_ONLY',
  LIVE_READ: 'LIVE_READ',
  LIVE_VERIFY: 'LIVE_VERIFY',
  PERSISTED: 'PERSISTED',
});

const CLASSIFICATION = Object.freeze({
  APPLY: 'APPLY',
  SATISFIED: 'SATISFIED',
  FAIL_CLOSED: 'FAIL_CLOSED',
});

/**
 * The authoritative execution-state register. The executor appends to it as operations resolve;
 * nothing downstream may rewrite an entry, because the persistence decision is derived from it.
 */
class AppliedOperations {
  constructor() {
    this._entries = [];
  }

  record({ operationId, status, verification = VERIFICATION.NONE, evidence = null, uncertainty = null }) {
    if (!Object.values(OPERATION_STATUS).includes(status)) {
      throw new Error(`unknown operation status: ${status}`);
    }
    if (!Object.values(VERIFICATION).includes(verification)) {
      throw new Error(`unknown verification level: ${verification}`);
    }
    if (status === OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE && verification !== VERIFICATION.NONE) {
      throw new Error('a skipped operation was never attempted and carries no verification');
    }
    if (status === OPERATION_STATUS.SATISFIED && verification === VERIFICATION.LIVE_VERIFY) {
      // R1: SATISFIED implies zero writes, so it can carry no live write verification.
      throw new Error('SATISFIED operation cannot carry LIVE_VERIFY write verification');
    }
    if (status === OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN) {
      if (verification === VERIFICATION.LIVE_VERIFY) {
        // NEVER fake LIVE_VERIFY: an uncertain operation has, by definition, not been verified.
        throw new Error('DISPATCHED_OUTCOME_UNKNOWN operation cannot carry LIVE_VERIFY');
      }
      uncertainty = validateUncertainty(uncertainty);
    } else if (uncertainty !== null) {
      throw new Error(`uncertainty evidence is only valid on DISPATCHED_OUTCOME_UNKNOWN, not ${status}`);
    }
    if (
      status === OPERATION_STATUS.APPLIED &&
      verification === VERIFICATION.LIVE_VERIFY &&
      this._entries.some(e => e.operationId === operationId && e.status === OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN)
    ) {
      // An uncertain operation stays uncertain for the rest of the task. A later read may inform the
      // NEXT task's classification; it may not rewrite this task's register.
      throw new UncertainPromotionError(operationId);
    }
    this._entries.push(
      Object.freeze({ operationId, status, verification, evidence, uncertainty, recordedAt: this._entries.length })
    );
    return this;
  }

  /** A copy: callers inspect the register, they never edit it. */
  entries() {
    return this._entries.slice();
  }

  countBy(status) {
    return this._entries.filter(e => e.status === status).length;
  }
}

/**
 * Operation accounting — M3 brief point 14. `planned` comes from the PLAN, the buckets from the
 * REGISTER; the identity is therefore a real check across two sources, not a derivation of one.
 * Task-level entries (operationId null, e.g. the lifecycle's unaccounted-dispatch backstop) are
 * not operations and are reported separately.
 */
function operationAccounting(planned, entries) {
  const ops = entries.filter(e => e.operationId !== null && e.operationId !== undefined);
  const by = status => ops.filter(e => e.status === status).length;
  const a = {
    planned,
    applied: by(OPERATION_STATUS.APPLIED),
    satisfied: by(OPERATION_STATUS.SATISFIED),
    failed: by(OPERATION_STATUS.FAILED),
    uncertain: by(OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN),
    skippedDueToPriorFailure: by(OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE),
    taskLevelEntries: entries.length - ops.length,
  };
  const ids = ops.map(e => e.operationId);
  a.terminalEntries = ops.length;
  a.duplicateEntries = ids.length - new Set(ids).size;
  a.balanced =
    planned === a.applied + a.satisfied + a.failed + a.uncertain + a.skippedDueToPriorFailure &&
    a.duplicateEntries === 0;
  return Object.freeze(a);
}

/**
 * Required shape of an uncertain entry (owner decision, point 3). Missing evidence is a caller
 * error and throws: an uncertain entry nobody can explain is worse than none.
 */
function validateUncertainty(u) {
  if (!u || typeof u !== 'object') throw new Error('DISPATCHED_OUTCOME_UNKNOWN requires uncertainty evidence');
  if (!Object.values(UNCERTAINTY_REASON).includes(u.reason)) {
    throw new Error(`unknown uncertainty reason: ${u && u.reason}`);
  }
  if (u.dispatchAttempted !== true) {
    throw new Error('DISPATCHED_OUTCOME_UNKNOWN requires dispatchAttempted: true (a proven non-dispatch is CERTAIN_NONE)');
  }
  if (!Number.isInteger(u.mutationGeneration)) {
    throw new Error('DISPATCHED_OUTCOME_UNKNOWN requires the mutation generation after the dispatch');
  }
  return Object.freeze({
    reason: u.reason,
    operationIndex: Number.isInteger(u.operationIndex) ? u.operationIndex : null,
    dispatchAttempted: true,
    dispatchEvidence: u.dispatchEvidence === undefined ? null : u.dispatchEvidence,
    mutationGeneration: u.mutationGeneration,
    verificationOutcome: u.verificationOutcome === undefined ? null : u.verificationOutcome,
  });
}

/**
 * Builds the frozen ExecutionResult. `executionOk` is the executor's semantic verdict and says
 * nothing about persistence or cleanup — that is the whole point of the separation.
 */
function buildExecutionResult({
  executionOk,
  outcome,
  authority = AUTHORITY.PLAN_ONLY,
  noOp = false,
  appliedOperations,
  classification = null,
  primitiveVerification = [],
  wholeTaskVerification = null,
  preservationVerification = null,
  receiptEvidence = null,
  metricsContribution = null,
  strategy = null,
  plannedCount = null,
  finalEvidence = null,
  batchEvidence = null,
  transaction = null,
}) {
  const entries = appliedOperations instanceof AppliedOperations
    ? appliedOperations.entries()
    : (appliedOperations || []).slice();

  return Object.freeze({
    kind: 'ExecutionResult',
    final: false, // never final task completion — M0 4.2
    executionOk: executionOk === true,
    outcome,
    authority,
    noOp: noOp === true,
    appliedOperations: Object.freeze(entries),
    classification,
    primitiveVerification: Object.freeze(primitiveVerification.slice()),
    wholeTaskVerification,
    preservationVerification,
    receiptEvidence,
    metricsContribution,
    // M3 fields. Additive: an M1/M2 executor that does not set them gets null, never a guess.
    strategy,
    operationAccounting: plannedCount === null ? null : operationAccounting(plannedCount, entries),
    finalEvidence,
    batchEvidence,
    transaction,
  });
}

module.exports = {
  AppliedOperations,
  buildExecutionResult,
  operationAccounting,
  OPERATION_STATUS,
  UNCERTAINTY_REASON,
  UncertainPromotionError,
  VERIFICATION,
  AUTHORITY,
  CLASSIFICATION,
};
