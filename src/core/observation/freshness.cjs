'use strict';

const { currentGeneration } = require('../mutation-epoch.cjs');

// Session identity is object-bound, not caller-provided string identity. A transport may reuse an
// id, and tests/callers can construct handles with the same id; neither may make an observation
// from handle A authoritative in handle B.
const sessionIdentities = new WeakMap();
function sessionIdentity(handle) {
  let token = sessionIdentities.get(handle);
  if (!token) {
    token = Object.freeze({ token: Symbol('session-observation-identity') });
    sessionIdentities.set(handle, token);
  }
  return token;
}

/**
 * Freshness and reuse rules — M0 section 7.1, owner requirements 4, 5, 19, 20.
 *
 * An observation is a claim about a moment. Two things can end that moment, and they fail for
 * different reasons, so they are reported separately rather than as one "invalid":
 *
 *   STALE            the session mutated after the observation was taken  (M0 O3, O4)
 *   CROSS_SESSION    the observation belongs to a different session entirely  (owner req. 19)
 *
 * Both are FAIL_CLOSED for mutation targeting and for verification. Neither is a warning.
 */

const FRESHNESS = Object.freeze({
  FRESH: 'FRESH',
  STALE: 'STALE',
  CROSS_SESSION: 'CROSS_SESSION',
  MALFORMED: 'MALFORMED',
});

/**
 * Classifies an observation against the session it is being used in NOW.
 *
 * The check is a counter comparison and an identity comparison — no clock, no TTL, no cache
 * lookup. That is the whole model.
 */
function classify(observation, handle) {
  if (!observation || typeof observation !== 'object' || !observation.freshness) {
    return { state: FRESHNESS.MALFORMED, reason: 'observation carries no freshness metadata' };
  }
  const meta = observation.freshness;

  if (meta.sessionIdentity !== sessionIdentity(handle)) {
    // The human-readable id is evidence only. Authority uses an unforgeable per-handle token.
    return {
      state: FRESHNESS.CROSS_SESSION,
      reason: `observation belongs to a different session handle (observed id ${meta.sessionId}, used in ${handle.id})`,
      observedIn: meta.sessionId,
      usedIn: handle.id,
    };
  }

  const now = currentGeneration(handle);
  if (meta.generation !== now) {
    return {
      state: FRESHNESS.STALE,
      reason: `session mutated since this observation (generation ${meta.generation} -> ${now})`,
      observedAtGeneration: meta.generation,
      currentGeneration: now,
      mutationsSince: now - meta.generation,
    };
  }

  return { state: FRESHNESS.FRESH, generation: now };
}

function isFresh(observation, handle) {
  return classify(observation, handle).state === FRESHNESS.FRESH;
}

/**
 * The gate every mutation-targeting and verification path goes through.
 *
 * It returns a refusal rather than throwing: a stale observation is an EXPECTED state (something
 * mutated, which is the normal course of a task), not a programmer error. The caller turns it
 * into a fail-closed result with the reason attached.
 */
function requireFresh(observation, handle, purpose) {
  const verdict = classify(observation, handle);
  if (verdict.state === FRESHNESS.FRESH) return { ok: true, verdict };
  return {
    ok: false,
    verdict,
    outcome: `observation-${verdict.state.toLowerCase().replace(/_/g, '-')}`,
    purpose,
    reason: verdict.reason,
  };
}

module.exports = { FRESHNESS, classify, isFresh, requireFresh, sessionIdentity };
