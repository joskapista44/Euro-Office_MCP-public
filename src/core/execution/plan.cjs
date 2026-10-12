'use strict';

/**
 * Plan envelope, dependency projection and batch grouping — M0 section 3 (`plan.cjs`), M3 points
 * 6, 7, 10, 11 and 22.
 *
 * A plan is an ordered list of operations. Dependencies come from two places, and both are
 * EXPLICIT -- the planner never infers independence from operation shape:
 *
 *   operation.dependsOn = ['A', ...]      declared by the adapter's planner
 *   operation.selector.ref = 'A'          a logical reference to A's SEMANTIC result
 *
 * A ref is not a locator. It names an earlier operation's outcome, and the operation that uses it
 * must prove it again in the CURRENT state (core/execution/executor.cjs resolves it from a fresh
 * observation). So a ref is always also a dependency, and a dependency always forces a barrier.
 */

// The only measured ceiling in either source: the XLSX batch line refuses above 1000 operations
// (M0 section 10, measured at fedaf4a7). Nothing ABOVE it has ever been exercised, so the plan
// refuses explicitly rather than running into untested territory. It is not a performance cap:
// nothing below it is limited, and there is no 100-operation limit anywhere.
const MAX_OPERATIONS = 1000;

/**
 * Batch independence (owner review R1). NO DECLARED DEPENDENCY IS NOT PROVEN INDEPENDENCE.
 *
 * The core knows exactly two states and never infers a third:
 *
 *   PROVEN_BATCH_SAFE   the planner issued `operation.batchSafeGroup = '<key>'`: its proof that the
 *                       operations sharing that key are mutually independent (no one moves, consumes
 *                       or re-targets another's target). The proof is the planner's; the core has no
 *                       format semantics with which to check it.
 *   NOT_PROVEN          everything else -- the DEFAULT. Such an operation is never put into a physical
 *                       batch with any other operation; it runs as a singleton (the correctness baseline).
 *
 * Measured against the sources before choosing the form: the XLSX line groups at PLANNER level (one
 * `step` carries several operations, every other operation is its own step); the DOCX line has no
 * batch concept at all. A planner-issued group is therefore the form both can fill, and a key on the
 * operation is its smallest representation -- it needs no cross-referencing ids that could dangle.
 */
const BATCH_INDEPENDENCE = Object.freeze({ PROVEN_BATCH_SAFE: 'PROVEN_BATCH_SAFE', NOT_PROVEN: 'NOT_PROVEN' });

function batchIndependence(operation) {
  const key = operation && operation.batchSafeGroup;
  return typeof key === 'string' && key !== '' ? BATCH_INDEPENDENCE.PROVEN_BATCH_SAFE : BATCH_INDEPENDENCE.NOT_PROVEN;
}

/**
 * Preservation placement (owner review of the M4 first checkpoint, R1). WHERE a required
 * preservation proof sits relative to the primitive's success -- the core knows no format and
 * never decides this itself; the planner declares it per operation:
 *
 *   NONE                 no preservation contract; the preservation hook is never invoked
 *   PRIMITIVE_INTEGRAL   preservation is PART of the primitive's semantic success: the operation is
 *                        APPLIED + LIVE_VERIFY only if L1 AND the preservation proof both PASS.
 *                        L1 PASS + preservation not proven -> DISPATCHED_OUTCOME_UNKNOWN (the write
 *                        may have happened; its contract is not proven)
 *   POST_PRIMITIVE       the primitive is proven on its own; a separate invariant is checked after
 *                        APPLIED and fails the TASK (the M3 behaviour; the barrier still runs)
 *
 * `preservationRequired: true` without a mode is the M3 form and means POST_PRIMITIVE. A mode that
 * contradicts that flag is refused by validatePlan rather than resolved by precedence.
 */
const PRESERVATION_MODE = Object.freeze({
  NONE: 'NONE', PRIMITIVE_INTEGRAL: 'PRIMITIVE_INTEGRAL', POST_PRIMITIVE: 'POST_PRIMITIVE',
});

function preservationModeOf(operation) {
  const mode = operation && operation.preservationMode;
  if (mode !== undefined) return mode;
  return operation && operation.preservationRequired === true ? PRESERVATION_MODE.POST_PRIMITIVE : PRESERVATION_MODE.NONE;
}

function dependenciesOf(operation) {
  const deps = new Set(Array.isArray(operation.dependsOn) ? operation.dependsOn : []);
  const ref = operation.selector && operation.selector.ref;
  if (typeof ref === 'string') deps.add(ref);
  return deps;
}

/**
 * Validates the plan. Fail-closed and complete: every problem is returned, not just the first.
 *
 *   - operation ids unique and non-empty
 *   - a dependency or ref names an EARLIER operation of the same plan (no forward or dangling
 *     references: "the previous operation's semantic result" cannot come from the future)
 *   - a ref is a string id; an object ref (e.g. a result carried over from another session) is
 *     refused, because its authority would be the other session's
 *   - no precomputed physical target on an operation: targets are resolved from fresh state
 *   - at most MAX_OPERATIONS operations
 */
function validatePlan(operations) {
  const problems = [];
  if (!Array.isArray(operations)) return { ok: false, problems: ['plan operations is not an array'] };
  if (operations.length > MAX_OPERATIONS) {
    problems.push(`plan has ${operations.length} operations, above the explicit ceiling ${MAX_OPERATIONS}`);
  }
  const seen = new Set();
  for (const [i, op] of operations.entries()) {
    if (!op || typeof op.id !== 'string' || op.id === '') {
      problems.push(`operation ${i} has no id`);
      continue;
    }
    if (seen.has(op.id)) problems.push(`duplicate operation id ${op.id}`);
    if (op.selector && op.selector.ref !== undefined && typeof op.selector.ref !== 'string') {
      problems.push(`operation ${op.id} carries a non-local ref; refs name an earlier operation of this task`);
    }
    if (op.batchSafeGroup !== undefined && (typeof op.batchSafeGroup !== 'string' || op.batchSafeGroup === '')) {
      problems.push(`operation ${op.id} carries a malformed batchSafeGroup; the proof is a non-empty planner key`);
    }
    if (op.preservationMode !== undefined) {
      if (!Object.values(PRESERVATION_MODE).includes(op.preservationMode)) {
        problems.push(`operation ${op.id} carries an unknown preservationMode ${JSON.stringify(op.preservationMode)}`);
      } else if (op.preservationRequired === true && op.preservationMode !== PRESERVATION_MODE.POST_PRIMITIVE) {
        problems.push(`operation ${op.id} carries preservationRequired with preservationMode ${op.preservationMode}; declare one placement`);
      }
    }
    if (op.target !== undefined) {
      problems.push(`operation ${op.id} carries a precomputed target; targets are resolved from fresh state`);
    }
    for (const dep of dependenciesOf(op)) {
      if (!seen.has(dep)) problems.push(`operation ${op.id} depends on ${dep}, which is not an earlier operation`);
    }
    seen.add(op.id);
  }
  return { ok: problems.length === 0, problems };
}

/**
 * Batch grouping (M3 point 11, owner review R1). Greedy, in plan order. An operation JOINS the
 * current group only if ALL of these hold:
 *
 *   1. it is PROVEN_BATCH_SAFE, and its batchSafeGroup key equals the current group's key
 *      (a NOT_PROVEN operation always runs alone -- unknown independence is never batched)
 *   2. it does not depend on any operation already in the current group (a dependency beats any
 *      batch proof: `dependsOn` or `selector.ref` always forces a barrier)
 *   3. the group is below `maxGroupSize`
 *
 * Otherwise it starts a new group. The cut is counted by its cause:
 *
 *   dependencyBarriers  cut forced by a dependency on the current group (dispatch, re-read, verify)
 *   unprovenCuts        cut because independence was not proven (or the keys differ)
 *   sizeBreaks          cut for transport size
 *
 *   A,B,C, no proof                         -> [A] | [B] | [C]      2 unproven cuts
 *   A,B,C, same batchSafeGroup              -> [A,B,C]
 *   A,B same batchSafeGroup, B dependsOn A  -> [A] | [B]            1 dependency barrier
 */
function groupIntoBatches(operations, { maxGroupSize = Infinity } = {}) {
  const groups = [];
  const barrierBefore = [];
  let current = [];
  let currentIds = new Set();
  let nextBarrier = false;
  let dependencyBarriers = 0;
  let sizeBreaks = 0;
  let unprovenCuts = 0;
  let currentKey = null;
  for (const op of operations) {
    const deps = dependenciesOf(op);
    const dependsOnCurrent = [...deps].some(d => currentIds.has(d));
    const proven = batchIndependence(op) === BATCH_INDEPENDENCE.PROVEN_BATCH_SAFE;
    const sameProvenGroup = proven && currentKey !== null && op.batchSafeGroup === currentKey;
    if (current.length > 0 && (dependsOnCurrent || !sameProvenGroup || current.length >= maxGroupSize)) {
      groups.push(current);
      barrierBefore.push(nextBarrier);
      nextBarrier = dependsOnCurrent;
      if (dependsOnCurrent) dependencyBarriers += 1;
      else if (!sameProvenGroup) unprovenCuts += 1;
      else sizeBreaks += 1;
      current = [];
      currentIds = new Set();
    }
    if (current.length === 0) currentKey = proven ? op.batchSafeGroup : null;
    current.push(op);
    currentIds.add(op.id);
  }
  if (current.length > 0) {
    groups.push(current);
    barrierBefore.push(nextBarrier);
  }
  return { groups, barrierBefore, dependencyBarriers, sizeBreaks, unprovenCuts };
}

/** One operation per group: the sequential and transaction strategies. */
function singletonGroups(operations) {
  const barrierBefore = operations.map(op => dependenciesOf(op).size > 0);
  const dependencyBarriers = barrierBefore.filter(Boolean).length;
  return { groups: operations.map(op => [op]), barrierBefore, dependencyBarriers, sizeBreaks: 0, unprovenCuts: 0 };
}

module.exports = {
  MAX_OPERATIONS, BATCH_INDEPENDENCE, batchIndependence, PRESERVATION_MODE, preservationModeOf, validatePlan, dependenciesOf, groupIntoBatches, singletonGroups,
};
