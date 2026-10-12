'use strict';

const { OPERATION_STATUS, VERIFICATION, UNCERTAINTY_REASON } = require('./result.cjs');
const { VERDICT } = require('../verification/layers.cjs');
const { PRESERVATION_MODE } = require('./plan.cjs');

/**
 * When is an operation uncertain? — owner decision after W3.0, point 4.
 *
 * One pure classifier, so that every strategy (sequential, transaction, batch) turns a dispatch
 * attempt into a register entry the same way. It never guesses whether a write happened; it only
 * decides whether the evidence PROVES one of the two certain outcomes.
 *
 * THE PROOF OF "NO DISPATCH" IS STRUCTURAL. core/session/handle.cjs bumps the session's mutation
 * generation BEFORE it calls the driver, at the single point every mutation passes through. So:
 *
 *   generation unchanged  ->  the driver was never asked to mutate  ->  CERTAIN_NONE (FAILED)
 *   generation advanced   ->  the driver WAS asked; only a passing live verification makes it
 *                            certain, and everything else is DISPATCHED_OUTCOME_UNKNOWN
 *
 * A refusal reported from inside the driver call, a throw, a partial-apply report and a
 * verification that could not prove the terminal state all land in the second row. The source
 * audit found that DOCX refusals from inside the editor call and throws inside the page were
 * exactly the cases the old code reported as "no write".
 *
 * Input:
 *   generationBefore / generationAfter   the session generation around the dispatch attempt
 *   threw                                the dispatch call threw (message kept as evidence)
 *   dispatchResult                       what the dispatch returned, if it returned
 *   verification                         the L1 verdict object, if verification ran
 *   preservationMode                     the operation's declared placement (plan.cjs)
 *   integralPreservation                 the L3 verdict object, when the mode is PRIMITIVE_INTEGRAL
 *
 * PRIMITIVE_INTEGRAL: a passing L1 is not enough -- APPLIED also needs a PASSING integral
 * preservation verdict. A missing one is treated as not proven, so an executor that forgot to run
 * it cannot produce a verified write for such an operation.
 */
function classifyDispatchAttempt({
  generationBefore,
  generationAfter,
  threw = null,
  dispatchResult = null,
  verification = null,
  operationIndex = null,
  preservationMode = PRESERVATION_MODE.NONE,
  integralPreservation = null,
}) {
  if (!Number.isInteger(generationBefore) || !Number.isInteger(generationAfter)) {
    // Without the generation pair nothing can be proven either way: that is a caller error.
    throw new Error('classifyDispatchAttempt requires generationBefore and generationAfter');
  }

  const dispatched = generationAfter > generationBefore;

  if (!dispatched) {
    // Proven pre-dispatch: the mutation surface was never reached. Certain, and not a write.
    return Object.freeze({
      certainty: 'CERTAIN_NONE',
      status: OPERATION_STATUS.FAILED,
      verification: VERIFICATION.NONE,
      uncertainty: null,
    });
  }

  const uncertain = reason => Object.freeze({
    certainty: 'UNCERTAIN',
    status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN,
    verification: verification ? VERIFICATION.DISPATCH_ONLY : VERIFICATION.NONE,
    uncertainty: Object.freeze({
      reason,
      operationIndex,
      dispatchAttempted: true,
      dispatchEvidence: threw
        ? Object.freeze({ threw: true, detail: String(threw.message || threw) })
        : dispatchResult,
      mutationGeneration: generationAfter,
      verificationOutcome: verification ? verification.verdict : null,
      integralPreservationOutcome: integralPreservation ? integralPreservation.verdict : null,
    }),
  });

  if (threw) return uncertain(UNCERTAINTY_REASON.DISPATCH_THREW);
  if (!dispatchResult || dispatchResult.ok !== true) {
    return uncertain(UNCERTAINTY_REASON.DISPATCH_REFUSED_NOT_PROVEN_PRE_DISPATCH);
  }
  if (dispatchResult.partiallyApplied === true) {
    return uncertain(UNCERTAINTY_REASON.ADAPTER_REPORTED_PARTIALLY_APPLIED);
  }
  if (!verification || verification.verdict !== VERDICT.PASS) {
    return uncertain(UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE);
  }
  if (preservationMode === PRESERVATION_MODE.PRIMITIVE_INTEGRAL
      && (!integralPreservation || integralPreservation.verdict !== VERDICT.PASS)) {
    return uncertain(UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  }

  return Object.freeze({
    certainty: 'CERTAIN',
    status: OPERATION_STATUS.APPLIED,
    verification: VERIFICATION.LIVE_VERIFY,
    uncertainty: null,
  });
}

module.exports = { classifyDispatchAttempt };
