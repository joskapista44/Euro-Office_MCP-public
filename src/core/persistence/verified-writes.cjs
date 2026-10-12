'use strict';

const { OPERATION_STATUS, VERIFICATION } = require('../execution/result.cjs');

/**
 * hasVerifiedAppliedWrites — M0 section 8.1, verbatim:
 *
 *   execution.appliedOperations.some(
 *     e => e.status === 'APPLIED' && e.verification === 'LIVE_VERIFY'
 *   )
 *
 * The input is the AUTHORITATIVE execution state. It is deliberately not the receipt: a missing,
 * empty or failed receipt serialization must not suppress persistence (owner requirements 7-9).
 * This module therefore imports nothing from core/receipt, and the static gate proves it.
 */
function hasVerifiedAppliedWrites(execution) {
  if (!execution || !Array.isArray(execution.appliedOperations)) return false;
  return execution.appliedOperations.some(
    entry =>
      entry.status === OPERATION_STATUS.APPLIED && entry.verification === VERIFICATION.LIVE_VERIFY
  );
}

/** The entries that made the predicate true. Evidence for the barrier decision, not an input to it. */
function verifiedAppliedWrites(execution) {
  if (!execution || !Array.isArray(execution.appliedOperations)) return [];
  return execution.appliedOperations.filter(
    entry =>
      entry.status === OPERATION_STATUS.APPLIED && entry.verification === VERIFICATION.LIVE_VERIFY
  );
}

/**
 * hasUncertainWrites — the second authoritative predicate (owner decision after W3.0, point 1).
 *
 * True when the register holds a dispatch whose outcome cannot be proven. Like the first
 * predicate it reads the authoritative execution state only, never a receipt.
 */
function hasUncertainWrites(execution) {
  return uncertainWrites(execution).length > 0;
}

function uncertainWrites(execution) {
  if (!execution || !Array.isArray(execution.appliedOperations)) return [];
  return execution.appliedOperations.filter(
    entry => entry.status === OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN
  );
}

/**
 * The persistence matrix, as a pure decision. The lifecycle owner executes it; nothing else does.
 *
 *   verified  uncertain   barrier   best outcome (M7 save contract, owner GO owner channel)
 *   false     false       no        NOT_REQUIRED
 *   false     true        NO        UNCERTAIN_NOT_PERSISTED
 *   true      either      yes       FORCESAVE_CONFIRMED -- only with the driver's own paired confirmation, still
 *                                   standing at the verified close; otherwise EDITOR_CLEAN (no confirmed save)
 *
 * Mutation uncertainty stays orthogonal (FinalTaskResult.mutationCertainty). Row two is a SAFETY POLICY, not a runtime
 * assumption: the framework never forceSaves solely for a mutation whose correctness is unproven. W3.0 measured that
 * such a mutation did not reach the durable state in 5/5 short runs, but the policy does not depend on that holding.
 * The barrier saves the WHOLE editor state, so a verified prefix cannot be saved without whatever an uncertain
 * operation left behind.
 *
 * There is no general PERSISTED state: canSave false is at most EDITOR_CLEAN, and FORCESAVE_CONFIRMED means only that
 * the server reported the save of that acknowledged state as successful -- the stored content is NOT checked in the
 * product path.
 */
function decidePersistence(execution) {
  const verified = hasVerifiedAppliedWrites(execution);
  const uncertain = hasUncertainWrites(execution);
  let onSuccess;
  if (!verified && !uncertain) onSuccess = PERSISTENCE_STATE.NOT_REQUIRED;
  else if (!verified) onSuccess = PERSISTENCE_STATE.UNCERTAIN_NOT_PERSISTED;
  else onSuccess = PERSISTENCE_STATE.FORCESAVE_CONFIRMED;
  return Object.freeze({ verified, uncertain, barrierRequired: verified, onSuccess });
}

const PERSISTENCE_STATE = Object.freeze({
  NOT_REQUIRED: 'NOT_REQUIRED',
  UNCERTAIN_NOT_PERSISTED: 'UNCERTAIN_NOT_PERSISTED',
  EDITOR_CLEAN: 'EDITOR_CLEAN',
  FORCESAVE_CONFIRMED: 'FORCESAVE_CONFIRMED',
  FAILED: 'FAILED',
});

module.exports = {
  hasVerifiedAppliedWrites,
  verifiedAppliedWrites,
  hasUncertainWrites,
  uncertainWrites,
  decidePersistence,
  PERSISTENCE_STATE,
};
