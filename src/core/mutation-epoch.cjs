'use strict';

/**
 * Mutation epoch — the freshness model (owner requirement 20).
 *
 * The simplest model that is provable, measured against the alternatives before being chosen:
 *
 *   per session, a monotonic counter. Every dispatch bumps it. An observation records the
 *   generation it was taken at. An observation is STALE when the session has moved on.
 *
 * What this deliberately is NOT:
 *
 *   NOT a wall-clock TTL. "fresh for 5 seconds" is a guess about the editor, not a fact about it:
 *        a mutation one millisecond later invalidates an observation the clock still calls fresh,
 *        and an idle session invalidates nothing while the clock expires everything.
 *   NOT a cache framework. There is no store, no eviction and no reuse policy here — only the
 *        question "has anything changed since this was read", which is a counter comparison.
 *
 * It is a LEAF module: it imports nothing, so both core/session (which bumps) and
 * core/observation (which reads) may depend on it without creating a cycle or a forbidden edge.
 */

const generations = new WeakMap();

/** Sessions start at generation 0: nothing has mutated yet. */
function currentGeneration(handle) {
  return generations.get(handle) || 0;
}

/**
 * Called by the layer that actually mutates. Bumping on DISPATCH rather than on "successful
 * mutation" is deliberate and fail-closed: a dispatch whose outcome is unknown must invalidate
 * too, because the one state we may not assume is "nothing happened".
 */
function bumpGeneration(handle) {
  const next = currentGeneration(handle) + 1;
  generations.set(handle, next);
  return next;
}

module.exports = { currentGeneration, bumpGeneration };
