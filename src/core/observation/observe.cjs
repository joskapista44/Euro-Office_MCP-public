'use strict';

const { readLive, readPackage } = require('../session/handle.cjs');
const { currentGeneration } = require('../mutation-epoch.cjs');
const { AUTHORITY, assertClaim, weakest } = require('../authority.cjs');
const { runEnrichment, isFailClosed, POLICY, ENRICHMENT_STATE } = require('./enrichment.cjs');
const { requireFresh, FRESHNESS, sessionIdentity } = require('./freshness.cjs');

/**
 * observeFresh — M0 section 4 and 7, owner requirements 3, 4, 8.
 *
 * The mandatory order, and it is an order rather than a set:
 *
 *   readLive  ->  adapter.normalizeInventory  ->  (optional) adapter.enrichObservation
 *             ->  semantic observation
 *
 * AUTHORITY is the LIVE part, always. Enrichment is additive and marked non-live, so a verifier
 * cannot mistake a package-derived field for something the editor showed us. An observation whose
 * live read failed is PLAN_ONLY and must not proceed to a write.
 *
 * Every observation carries explicit freshness/evidence metadata (owner requirement 4), including
 * the session it belongs to and the mutation generation it was taken at. That metadata is what
 * makes staleness and cross-session misuse detectable at all.
 */

let sequence = 0;

async function observeFresh(handle, adapter, request = {}) {
  const policy = request.enrichment || POLICY.OPTIONAL;
  const observedAtGeneration = currentGeneration(handle);
  const seq = ++sequence;

  // ---- 1. fresh LIVE read: the only mutation-targeting authority -----------------------------
  let live;
  try {
    live = await readLive(handle, request.live || {});
  } catch (err) {
    // Expected transport boundary failure, normalized (M1 exception policy, new boundary).
    return failed('observation-live-read-threw', {
      handle,
      seq,
      observedAtGeneration,
      liveEvidence: { ok: false, threw: true, detail: err.message, errorName: err.name },
    });
  }
  if (!live || live.ok !== true) {
    return failed('observation-live-read-failed', {
      handle,
      seq,
      observedAtGeneration,
      liveEvidence: { ok: false, threw: false, detail: (live && live.detail) || 'no live result' },
    });
  }

  // ---- 2. adapter normalization: format inventory -> common shape ----------------------------
  let primary;
  try {
    primary = adapter.normalizeInventory(live.payload);
  } catch (err) {
    // An adapter that cannot normalize its OWN live payload has broken its contract. This is
    // still treated as a boundary failure rather than a crash, because the adapter is external
    // to the core — but the outcome names normalization, so the two are never confused.
    return failed('observation-normalization-threw', {
      handle,
      seq,
      observedAtGeneration,
      liveEvidence: { ok: true, authority: AUTHORITY.LIVE_READ, source: live.authority || 'LIVE_READ' },
      normalizationEvidence: { ok: false, threw: true, detail: err.message, errorName: err.name },
    });
  }
  if (primary === null || primary === undefined) {
    return failed('observation-normalization-failed', {
      handle,
      seq,
      observedAtGeneration,
      liveEvidence: { ok: true, authority: AUTHORITY.LIVE_READ },
      normalizationEvidence: { ok: false, threw: false, detail: 'normalizeInventory returned nothing' },
    });
  }

  // ---- 3. optional enrichment: ADDS semantics, never replaces the live state ------------------
  // The enrichment hook gets a READ-ONLY, counted package capability; enrichment.cjs itself may not
  // import core/session (M2 forbidden edge), so the capability is created here.
  const enrichment = await runEnrichment(handle, adapter, primary, {
    policy,
    readPackage: packageRequest => readPackage(handle, packageRequest),
  });
  if (isFailClosed(enrichment)) {
    return failed('observation-required-enrichment-failed', {
      handle,
      seq,
      observedAtGeneration,
      liveEvidence: { ok: true, authority: AUTHORITY.LIVE_READ },
      normalizationEvidence: { ok: true },
      enrichmentEvidence: enrichment,
    });
  }

  // ---- 4. semantic observation ----------------------------------------------------------------
  return Object.freeze({
    kind: 'Observation',
    ok: true,
    // The ceiling is enforced, not assumed: an observation can never report LIVE_VERIFY.
    authority: assertClaim('observation', AUTHORITY.LIVE_READ),
    primary,
    enrichment: enrichment.state === ENRICHMENT_STATE.APPLIED ? enrichment.data : null,
    freshness: Object.freeze({
      sessionId: handle.id,
      sessionIdentity: sessionIdentity(handle),
      generation: observedAtGeneration,
      sequence: seq,
      source: 'LIVE',
      authority: AUTHORITY.LIVE_READ,
    }),
    evidence: Object.freeze({
      live: Object.freeze({ ok: true, authority: AUTHORITY.LIVE_READ, source: live.authority || 'LIVE_READ' }),
      normalization: Object.freeze({ ok: true, adapterKind: adapter.kind }),
      enrichment,
    }),
  });
}

function failed(outcome, { handle, seq, observedAtGeneration, liveEvidence, normalizationEvidence = null, enrichmentEvidence = null }) {
  return Object.freeze({
    kind: 'Observation',
    ok: false,
    outcome,
    // A failed observation is PLAN_ONLY: execution must not proceed to a write (M0 section 4).
    authority: AUTHORITY.PLAN_ONLY,
    primary: null,
    enrichment: null,
    freshness: Object.freeze({
      sessionId: handle.id,
      sessionIdentity: sessionIdentity(handle),
      generation: observedAtGeneration,
      sequence: seq,
      source: 'LIVE',
      authority: AUTHORITY.PLAN_ONLY,
    }),
    evidence: Object.freeze({
      live: Object.freeze(liveEvidence),
      normalization: normalizationEvidence ? Object.freeze(normalizationEvidence) : null,
      enrichment: enrichmentEvidence,
    }),
  });
}

/**
 * The gate for using an observation as MUTATION authority — owner requirements 5, 7, 19.
 *
 * Three separate refusals, each named, because they have different causes and the caller's next
 * step differs: re-read, abort, or fix the plumbing.
 */
function requireMutationAuthority(observation, handle, purpose = 'mutation-targeting') {
  if (!observation || observation.ok !== true) {
    return { ok: false, outcome: 'observation-not-ok', purpose, reason: (observation && observation.outcome) || 'no observation' };
  }

  // Package-only can never be a writable target authority. The observation's authority comes from
  // the LIVE part; an observation that somehow carried only enrichment is refused here.
  if (observation.evidence.live.ok !== true) {
    return { ok: false, outcome: 'observation-package-only', purpose, reason: 'no live evidence backs this observation' };
  }
  if (weakest(observation.authority, AUTHORITY.LIVE_READ) !== AUTHORITY.LIVE_READ) {
    return { ok: false, outcome: 'observation-authority-too-weak', purpose, reason: `authority ${observation.authority}` };
  }

  return requireFresh(observation, handle, purpose);
}

module.exports = { observeFresh, requireMutationAuthority, FRESHNESS, POLICY };
