'use strict';

const { AUTHORITY, assertClaim } = require('../authority.cjs');

/**
 * Optional package enrichment — M0 section 7 (O2), owner requirement 6.
 *
 * Enrichment may ADD semantics the live projection omits. It may never BE the live state:
 * "package-only" is not a writable target authority (owner requirement 7, invariant 4).
 *
 * Three policies, because "optional" alone cannot express a capability that genuinely needs the
 * package to be correct:
 *
 *   NONE      do not call the hook at all                                    -> SKIPPED
 *   OPTIONAL  call it; absence is fine, a failure is DEGRADED and the task continues (default)
 *   REQUIRED  call it; ANY of these is FAIL_CLOSED:
 *                the adapter has no hook        |  the hook returned null/undefined
 *                the hook reported !ok          |  the hook threw
 *
 * The null case is listed explicitly because it was the fail-open gap: "the hook ran and produced
 * nothing" is indistinguishable in shape from "this adapter has nothing to add", and only the
 * POLICY can tell them apart.
 *
 * The default is OPTIONAL rather than REQUIRED because the measured reality is asymmetric: DOCX
 * enriches from OOXML in 18 files, XLSX returns null in all of them. A REQUIRED default would
 * make the common core unusable for a format that correctly has nothing to add.
 */

const POLICY = Object.freeze({ NONE: 'NONE', OPTIONAL: 'OPTIONAL', REQUIRED: 'REQUIRED' });

const ENRICHMENT_STATE = Object.freeze({
  SKIPPED: 'SKIPPED',            // policy NONE
  NOT_IMPLEMENTED: 'NOT_IMPLEMENTED', // adapter returns null — legitimate, e.g. XLSX
  APPLIED: 'APPLIED',
  DEGRADED: 'DEGRADED',          // OPTIONAL policy, hook failed or threw
  FAILED: 'FAILED',              // REQUIRED policy, hook failed or threw -> fail closed
});

/**
 * Runs the hook under its policy and returns evidence. Never throws for an EXPECTED adapter
 * failure: a hook that reports {ok:false} or throws is a boundary failure, normalized here with
 * its detail preserved (the M1 exception policy, applied to a new boundary).
 */
async function runEnrichment(handle, adapter, observationPrimary, { policy = POLICY.OPTIONAL, readPackage = null } = {}) {
  if (!Object.values(POLICY).includes(policy)) {
    // A caller error, not a boundary failure: an unknown policy must not silently become OPTIONAL.
    throw new Error(`unknown enrichment policy: ${policy}`);
  }

  if (policy === POLICY.NONE) {
    return evidence(ENRICHMENT_STATE.SKIPPED, null, { policy });
  }

  if (typeof adapter.enrichObservation !== 'function') {
    if (policy === POLICY.REQUIRED) {
      return evidence(ENRICHMENT_STATE.FAILED, null, {
        policy,
        reason: `adapter "${adapter.kind}" does not implement enrichObservation, but the capability requires it`,
      });
    }
    return evidence(ENRICHMENT_STATE.NOT_IMPLEMENTED, null, { policy });
  }

  let result;
  try {
    handle.metrics.count('enrichmentInvocations');
    // Third argument (M4, additive): the core-mediated package read. Absent when the caller has no
    // package transport to offer; a hook that needs it must then report not-ok, not invent I/O.
    result = await adapter.enrichObservation(handle, observationPrimary, { readPackage });
  } catch (err) {
    return failOrDegrade(policy, { threw: true, detail: err.message, errorName: err.name });
  }

  if (result === null || result === undefined) {
    // An implemented hook may legitimately have nothing to add — but only under a policy that
    // tolerates absence. Under REQUIRED, "the hook ran and produced nothing" is exactly the
    // failure the policy exists to catch, and the first version of this module let it through as
    // NOT_IMPLEMENTED. That is fail-open: a capability that declared it NEEDS the package would
    // have proceeded without it.
    if (policy === POLICY.REQUIRED) {
      return evidence(ENRICHMENT_STATE.FAILED, null, {
        policy,
        threw: false,
        reason: 'hook ran and returned no enrichment, but the capability requires it',
        hookReturnedNull: true,
      });
    }
    return evidence(ENRICHMENT_STATE.NOT_IMPLEMENTED, null, { policy, reason: 'hook returned null' });
  }
  if (result.ok !== true) {
    return failOrDegrade(policy, { threw: false, detail: (result && result.detail) || 'hook reported not ok' });
  }

  return evidence(ENRICHMENT_STATE.APPLIED, result.enrichment !== undefined ? result.enrichment : result, { policy });
}

function failOrDegrade(policy, detail) {
  if (policy === POLICY.REQUIRED) {
    return evidence(ENRICHMENT_STATE.FAILED, null, { policy, ...detail });
  }
  // OPTIONAL: the task continues, but the degradation is VISIBLE in the evidence rather than
  // being indistinguishable from an adapter that simply had nothing to add.
  return evidence(ENRICHMENT_STATE.DEGRADED, null, { policy, ...detail });
}

function evidence(state, data, extra) {
  return Object.freeze({
    state,
    // Enrichment never claims more than LIVE_READ, and the lattice enforces the ceiling.
    authority: assertClaim('enrichment', AUTHORITY.LIVE_READ),
    // Marked so a verifier cannot mistake enrichment fields for LIVE evidence (M0 O2).
    isLiveEvidence: false,
    data: data === undefined ? null : data,
    ...extra,
  });
}

/** Whether this enrichment outcome must stop the observation. */
function isFailClosed(enrichmentEvidence) {
  return enrichmentEvidence.state === ENRICHMENT_STATE.FAILED;
}

module.exports = { POLICY, ENRICHMENT_STATE, runEnrichment, isFailClosed };
