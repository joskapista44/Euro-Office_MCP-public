'use strict';

const { LAYER, LAYER_AUTHORITY, VERDICT } = require('./layers.cjs');
const { AUTHORITY, assertClaim } = require('../authority.cjs');
const { requireFresh } = require('../observation/freshness.cjs');

/**
 * Verification orchestration — M0 section 9, owner requirements 9-13.
 *
 * The core ORCHESTRATES and never verifies content itself: it decides which layer runs, against
 * which state, and what authority the outcome may claim. The semantics are entirely the adapter's.
 *
 * Two rules are enforced here rather than trusted:
 *
 *   1. A dispatch result can never become L1. `fromDispatch` exists precisely so that the only
 *      way to produce DISPATCH_ONLY is to ask for it explicitly, and there is no code path from
 *      a dispatch return into a LIVE_VERIFY verdict.
 *   2. Every verifier runs against a FRESH observation. Planning-time state is refused with
 *      STALE_STATE, because verifying against the state you intended to create proves nothing
 *      about the state that exists.
 */

/** L0. A returned command. Claims DISPATCH_ONLY, and the lattice enforces that ceiling. */
function fromDispatch(dispatchResult) {
  return Object.freeze({
    layer: LAYER.DISPATCH_COMPLETE,
    verdict: dispatchResult && dispatchResult.ok === true ? VERDICT.PASS : VERDICT.FAIL,
    // A dispatch PASS means "the command returned". It is not acceptance of anything.
    authority: assertClaim('dispatch', AUTHORITY.DISPATCH_ONLY),
    proves: 'nothing about document state',
    evidence: Object.freeze({ dispatched: !!(dispatchResult && dispatchResult.dispatched) }),
  });
}

/** L1. Requires a fresh observation taken AFTER the operation. */
async function verifyPrimitive({ handle, adapter, operation, before = null, after }) {
  const gate = requireFresh(after, handle, 'primitive-verification');
  if (!gate.ok) return stale(LAYER.PRIMITIVE, gate);

  return runHook({
    layer: LAYER.PRIMITIVE,
    handle,
    fn: () => adapter.verifyPrimitive(handle, operation, before, after),
    context: { operationId: operation && operation.id },
  });
}

/** L2. Read-only terminal check. It may not write to fix what it finds (M0 section 4). */
async function verifyWholeTask({ handle, adapter, task, finalObservation }) {
  const gate = requireFresh(finalObservation, handle, 'whole-task-verification');
  if (!gate.ok) return stale(LAYER.WHOLE_TASK, gate);

  const result = await runHook({
    layer: LAYER.WHOLE_TASK,
    handle,
    fn: () => adapter.verifyWholeTask(handle, task, finalObservation),
    context: { taskId: task && task.id },
  });

  // The checks[] shape is part of the M0 contract for this layer.
  const checks = (result.evidence && result.evidence.checks) || [];
  return Object.freeze({
    ...result,
    checkCount: checks.length,
    failedCount: checks.filter(c => c && c.ok !== true).length,
  });
}

/**
 * L3. OPTIONAL. An adapter that does not declare preservation yields NOT_REQUIRED — which is
 * NOT a pass (owner requirement 12).
 */
async function verifyPreservation({ handle, adapter, operation, before = null, after, required = false, siblings = [] }) {
  if (typeof adapter.verifyPreservation !== 'function') {
    if (required) {
      return Object.freeze({
        layer: LAYER.PRESERVATION,
        verdict: VERDICT.FAIL,
        authority: AUTHORITY.PLAN_ONLY,
        evidence: Object.freeze({
          reason: `adapter "${adapter.kind}" declares no preservation verifier, but the capability requires one`,
        }),
      });
    }
    return Object.freeze({
      layer: LAYER.PRESERVATION,
      verdict: VERDICT.NOT_REQUIRED,
      authority: AUTHORITY.PLAN_ONLY,
      evidence: Object.freeze({ reason: 'adapter declares no preservation verifier' }),
    });
  }

  const gate = requireFresh(after, handle, 'preservation-verification');
  if (!gate.ok) return stale(LAYER.PRESERVATION, gate);

  const result = await runHook({
    layer: LAYER.PRESERVATION,
    handle,
    // GH #99: `siblings` = the OTHER operations dispatched in the same batch round -- their own L1 proves their
    // windows, so this operation's L3 must not read their writes as a foreign change. Empty outside a batch round
    // (sequential groups are singletons), and the adapter then sees exactly the call it always saw.
    fn: () => adapter.verifyPreservation(handle, operation, before, after, { siblings }),
    context: { operationId: operation && operation.id, required },
  });

  // Preservation is the one layer where "the hook returned nothing" has a legitimate reading:
  // an OPTIONAL check with nothing to preserve. The general model already produced NO_EVIDENCE
  // (never a PASS), so this only decides which non-pass verdict fits:
  //
  //   required=true   -> FAIL          a promised check that produced no evidence has failed
  //   required=false  -> NOT_REQUIRED  there was nothing to prove, and nothing was proven
  //
  // The ordering matters: the general model never manufactures a PASS that a later layer has to
  // take back. A false PASS that gets corrected downstream is one refactor away from surviving.
  if (result.verdict === VERDICT.NO_EVIDENCE) {
    return Object.freeze({ ...result, verdict: required ? VERDICT.FAIL : VERDICT.NOT_REQUIRED });
  }
  return result;
}

/**
 * The single place an adapter verifier is called. Counts the read, enforces the ceiling, and
 * normalizes an EXPECTED boundary failure — a verifier that throws is a broken boundary, not a
 * failed verification, and conflating the two would report "the document is wrong" when the
 * truth is "we could not look".
 */
async function runHook({ layer, handle, fn, context }) {
  const spec = LAYER_AUTHORITY[layer];
  handle.metrics.count('verificationInvocations');

  let out;
  try {
    out = await fn();
  } catch (err) {
    return Object.freeze({
      layer,
      verdict: VERDICT.BOUNDARY_FAILURE,
      // A boundary failure proves nothing, so it claims the floor.
      authority: AUTHORITY.PLAN_ONLY,
      evidence: Object.freeze({ threw: true, detail: err.message, errorName: err.name, ...context }),
    });
  }

  if (out === null || out === undefined) {
    // ABSENCE OF EVIDENCE IS NOT A PASS. A hook that returned nothing proved nothing, so the
    // verdict carries the floor. Returning PASS here — as the first version did — let a primitive
    // verifier claim LIVE_VERIFY without a single observation behind it.
    return Object.freeze({
      layer,
      verdict: VERDICT.NO_EVIDENCE,
      authority: AUTHORITY.PLAN_ONLY,
      evidence: Object.freeze({ hookReturnedNull: true, ...context }),
    });
  }

  const passed = out.ok === true;
  return Object.freeze({
    layer,
    verdict: passed ? VERDICT.PASS : VERDICT.FAIL,
    // Only a PASS may claim the layer's authority. A failure claims the floor: a failed verifier
    // has not proved a live read of anything.
    authority: passed ? assertClaim(spec.ceiling, spec.authority) : AUTHORITY.PLAN_ONLY,
    evidence: Object.freeze({ ...out, ...context }),
  });
}

function stale(layer, gate) {
  return Object.freeze({
    layer,
    verdict: VERDICT.STALE_STATE,
    authority: AUTHORITY.PLAN_ONLY,
    evidence: Object.freeze({
      outcome: gate.outcome,
      reason: gate.reason,
      freshness: gate.verdict,
      purpose: gate.purpose,
    }),
  });
}

module.exports = { fromDispatch, verifyPrimitive, verifyWholeTask, verifyPreservation };
