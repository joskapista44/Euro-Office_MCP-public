'use strict';

const { AUTHORITY } = require('../authority.cjs');

/**
 * The five verification layers — M0 section 9. None may substitute for another.
 *
 *   L0 DISPATCH_COMPLETE      the command returned. Proves NOTHING. Never acceptance (inv. 7).
 *   L1 PRIMITIVE_LIVE_VERIFY  a fresh post-operation read shows the intended local effect.
 *   L2 WHOLE_TASK_VERIFY      read-only terminal check of the whole task against a fresh read.
 *   L3 PRESERVATION_VERIFY    what must NOT change did not change (inv. 17).
 *   L4 PERSISTENCE_VERIFY     the barrier of M0 section 8 completed.
 *
 * L0 and L4 are core-only; an adapter cannot report them. L1-L3 are adapter hooks that the core
 * orchestrates without ever verifying content itself.
 */

const LAYER = Object.freeze({
  DISPATCH_COMPLETE: 'L0_DISPATCH_COMPLETE',
  PRIMITIVE: 'L1_PRIMITIVE_LIVE_VERIFY',
  WHOLE_TASK: 'L2_WHOLE_TASK_VERIFY',
  PRESERVATION: 'L3_PRESERVATION_VERIFY',
  PERSISTENCE: 'L4_PERSISTENCE_VERIFY',
});

/**
 * The authority a layer may claim, and the lattice layer name that enforces its ceiling.
 *
 * L0 is the entry that matters: a returned command claims DISPATCH_ONLY and nothing more. The
 * measured drift in both sources was exactly this — treating a successful dispatch as proof.
 */
const LAYER_AUTHORITY = Object.freeze({
  [LAYER.DISPATCH_COMPLETE]: { authority: AUTHORITY.DISPATCH_ONLY, ceiling: 'dispatch', adapterMayReport: false },
  [LAYER.PRIMITIVE]: { authority: AUTHORITY.LIVE_VERIFY, ceiling: 'primitiveVerifier', adapterMayReport: true },
  [LAYER.WHOLE_TASK]: { authority: AUTHORITY.LIVE_READ, ceiling: 'wholeTaskVerifier', adapterMayReport: true },
  [LAYER.PRESERVATION]: { authority: AUTHORITY.LIVE_READ, ceiling: 'preservationVerifier', adapterMayReport: true },
  [LAYER.PERSISTENCE]: { authority: AUTHORITY.PERSISTED, ceiling: 'persistence', adapterMayReport: false },
});

const VERDICT = Object.freeze({
  PASS: 'PASS',
  FAIL: 'FAIL',
  NOT_REQUIRED: 'NOT_REQUIRED',
  BOUNDARY_FAILURE: 'BOUNDARY_FAILURE',
  STALE_STATE: 'STALE_STATE',
  // A hook that returned nothing. It is NOT a pass: absence of verification evidence is not
  // verification. The first version of this module treated null as PASS and handed the layer its
  // full authority, which meant a primitive verifier could yield LIVE_VERIFY having proved
  // nothing at all -- fail-open, in the one place that must fail closed.
  NO_EVIDENCE: 'NO_EVIDENCE',
});

/**
 * NOT_REQUIRED is not PASS, and this is the distinction owner requirement 12 asks for: a task
 * with no preservation requirement has not passed preservation, it simply had none. Counting it
 * as a pass is how a missing requirement becomes an unnoticed gap.
 */
function isSatisfied(verdict) {
  return verdict === VERDICT.PASS || verdict === VERDICT.NOT_REQUIRED;
}

/** Verdicts that prove nothing and must never carry a layer's authority. */
function provesNothing(verdict) {
  return (
    verdict === VERDICT.NO_EVIDENCE ||
    verdict === VERDICT.BOUNDARY_FAILURE ||
    verdict === VERDICT.STALE_STATE ||
    verdict === VERDICT.FAIL
  );
}

function isProven(verdict) {
  return verdict === VERDICT.PASS;
}

module.exports = { LAYER, LAYER_AUTHORITY, VERDICT, isSatisfied, isProven, provesNothing };
