'use strict';

/**
 * The authority lattice — M0 section 4 vocabulary, formalized (owner requirement 14).
 *
 *   PLAN_ONLY < DISPATCH_ONLY < LIVE_READ < LIVE_VERIFY < PERSISTED
 *
 * A layer may only CLAIM an authority it actually proved. The two rules that matter are not the
 * ordering itself but its consequences:
 *
 *   - a claim above a layer's ceiling is refused, not silently downgraded. An observation cannot
 *     report LIVE_VERIFY however confident it is, because reading is not verifying (invariant 7);
 *   - combining evidence takes the WEAKEST link, never the strongest. This is the direction the
 *     two sources drifted in: a strong local proof next to a weak one read as strong overall.
 *
 * NOTE ON THE TREE: M0 section 3 names no file for the lattice, and it cannot live in
 * core/execution/result.cjs where the constants started, because `observation -> execution` is
 * not an allowed edge (M0 3.1). It is therefore a LEAF module: it imports nothing, and any core
 * module may import it. The static gate asserts the leaf property.
 */

const AUTHORITY = Object.freeze({
  PLAN_ONLY: 'PLAN_ONLY',
  DISPATCH_ONLY: 'DISPATCH_ONLY',
  LIVE_READ: 'LIVE_READ',
  LIVE_VERIFY: 'LIVE_VERIFY',
  PERSISTED: 'PERSISTED',
});

/** Ascending strength. The index IS the ordering; nothing else encodes it. */
const ORDER = Object.freeze([
  AUTHORITY.PLAN_ONLY,
  AUTHORITY.DISPATCH_ONLY,
  AUTHORITY.LIVE_READ,
  AUTHORITY.LIVE_VERIFY,
  AUTHORITY.PERSISTED,
]);

/**
 * What each layer may claim AT MOST. A verifier cannot invent persistence; an observation cannot
 * invent verification. These ceilings are the machine-checkable form of "prove what you claim".
 */
const CEILING = Object.freeze({
  observation: AUTHORITY.LIVE_READ,
  enrichment: AUTHORITY.LIVE_READ,
  primitiveVerifier: AUTHORITY.LIVE_VERIFY,
  wholeTaskVerifier: AUTHORITY.LIVE_READ,
  preservationVerifier: AUTHORITY.LIVE_READ,
  persistence: AUTHORITY.PERSISTED,
  plan: AUTHORITY.PLAN_ONLY,
  dispatch: AUTHORITY.DISPATCH_ONLY,
});

class AuthorityViolationError extends Error {
  constructor(layer, claimed, ceiling) {
    super(`layer "${layer}" claimed ${claimed}, which exceeds its ceiling ${ceiling}`);
    this.name = 'AuthorityViolationError';
    this.layer = layer;
    this.claimed = claimed;
    this.ceiling = ceiling;
  }
}

function rank(authority) {
  const i = ORDER.indexOf(authority);
  if (i === -1) throw new Error(`unknown authority: ${authority}`);
  return i;
}

/** True when `a` is at least as strong as `b`. */
function atLeast(a, b) {
  return rank(a) >= rank(b);
}

/**
 * The weakest of the given authorities. Combining evidence never strengthens it: a task whose
 * every step but one was live-verified is not a live-verified task.
 */
function weakest(...authorities) {
  if (authorities.length === 0) return AUTHORITY.PLAN_ONLY;
  return authorities.reduce((lo, a) => (rank(a) < rank(lo) ? a : lo));
}

/**
 * Fail-closed enforcement. A claim above the ceiling THROWS rather than being clamped: silently
 * clamping would let a layer keep reporting an authority it never had, and the caller would
 * never learn that its evidence was weaker than it believed.
 */
function assertClaim(layer, claimed) {
  const ceiling = CEILING[layer];
  if (!ceiling) throw new Error(`unknown authority layer: ${layer}`);
  if (rank(claimed) > rank(ceiling)) throw new AuthorityViolationError(layer, claimed, ceiling);
  return claimed;
}

/** Non-throwing form, for places that must report rather than abort. */
function withinCeiling(layer, claimed) {
  const ceiling = CEILING[layer];
  if (!ceiling) return false;
  return rank(claimed) <= rank(ceiling);
}

module.exports = {
  AUTHORITY,
  ORDER,
  CEILING,
  rank,
  atLeast,
  weakest,
  assertClaim,
  withinCeiling,
  AuthorityViolationError,
};
