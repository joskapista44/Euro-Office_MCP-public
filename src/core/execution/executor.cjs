'use strict';

const {
  buildExecutionResult, OPERATION_STATUS, VERIFICATION, AUTHORITY, CLASSIFICATION, UNCERTAINTY_REASON,
} = require('./result.cjs');
const { classifyDispatchAttempt } = require('./uncertainty.cjs');
const { validatePlan, PRESERVATION_MODE, preservationModeOf } = require('./plan.cjs');
const sequential = require('./strategy-sequential.cjs');
const transaction = require('./strategy-transaction.cjs');
const batch = require('./strategy-batch.cjs');
const { buildReceipt } = require('../receipt/envelope.cjs');
const { observeFresh, requireMutationAuthority } = require('../observation/observe.cjs');
const { requireFresh, sessionIdentity } = require('../observation/freshness.cjs');
const { verifyPrimitive, verifyWholeTask, verifyPreservation } = require('../verification/orchestrate.cjs');
const { VERDICT, isSatisfied } = require('../verification/layers.cjs');
const { dispatch } = require('../session/handle.cjs');
const { currentGeneration } = require('../mutation-epoch.cjs');

/**
 * executeTask — the ONE common executor (M0 section 4, M3 points 2-5).
 *
 *   per group:  targeting observation (FRESH, reused only if the generation has not moved)
 *               -> per operation: CLASSIFY -> RESOLVE (from that fresh state) -> PLAN/LOWER
 *               -> DISPATCH (one round for the group)
 *               -> FRESH READ -> L1 PRIMITIVE VERIFY -> L3 PRESERVATION by declared placement:
 *                  PRIMITIVE_INTEGRAL before the register decision, POST_PRIMITIVE after APPLIED
 *   task end:   FRESH FINAL OBSERVATION -> L2 WHOLE-TASK VERIFY -> ExecutionResult
 *
 * A strategy only decides the grouping (singletons for sequential/transaction, dependency-cut
 * groups for batch). There is no DOCX executor and no XLSX executor: the engine knows no format,
 * and the adapter supplies every semantic decision through the M0 contract.
 *
 * What this module never does, by construction and by static gate: open or close a session,
 * persist, or read a receipt as authority. It returns an ExecutionResult; the M1 lifecycle owner
 * decides persistence from the register this function writes.
 *
 * Stop rule (all strategies): the first failure stops further dispatch. Operations before it in
 * plan order are still executed and verified (so the three strategies agree on which prefix
 * lands); everything after it is recorded SKIPPED_DUE_TO_PRIOR_FAILURE.
 */

const STRATEGIES = Object.freeze({ sequential, transaction, batch });

async function executeTask({ handle, adapter, task, strategy = 'sequential', appliedOperations, options = {} }) {
  const strat = STRATEGIES[strategy];
  if (!strat) throw new Error(`unknown execution strategy: ${strategy}`);
  if (!appliedOperations) throw new Error('executeTask requires the lifecycle-owned register');

  const metrics = handle.metrics;
  const observeRequest = options.observation || {};
  const primitiveVerification = [];
  const preservationVerification = [];
  const results = new Map(); // operationId -> semantic result, the only thing a ref may name
  const identity = sessionIdentity(handle);

  let operations = [];
  let classification = null;
  let stopped = null;
  let groupsInfo = null;
  let executedGroups = 0;
  let batchDispatchRounds = 0;
  let finalEvidence = null;
  let wholeTaskVerification = null;

  const stop = (operationId, outcome, detail = null) => {
    if (!stopped) stopped = { operationId, outcome, detail };
  };

  // ---- planning observation ---------------------------------------------------------------------
  let observation = await observeFresh(handle, adapter, observeRequest);
  if (!observation.ok) {
    return finish({ executionOk: false, outcome: `planning-${observation.outcome}`, authority: AUTHORITY.PLAN_ONLY });
  }

  let plan;
  try {
    plan = adapter.planOperations(task, observation);
  } catch (err) {
    return finish({ executionOk: false, outcome: 'plan-threw', authority: AUTHORITY.PLAN_ONLY, detail: err.message });
  }
  if (!plan || plan.ok !== true || !Array.isArray(plan.operations)) {
    // #ref/#ref: `plan-invalid` and `plan-threw` already thread their own `detail` into
    // `finish()` (-> `stopDetail` on the result); this branch did not, even though
    // `planOperations()` always returns a `detail` string on failure -- silently dropping it.
    return finish({ executionOk: false, outcome: 'plan-failed', authority: AUTHORITY.PLAN_ONLY, detail: (plan && plan.detail) || null });
  }
  operations = plan.operations;
  metrics.count('operationsPlanned', operations.length);

  const validity = validatePlan(operations);
  if (!validity.ok) {
    // Nothing was attempted. Every planned operation is terminal as FAILED (certain: no dispatch),
    // so the accounting identity still holds for a refused plan.
    const evidence = Object.freeze({ stage: 'plan-validation', problems: validity.problems });
    for (const op of operations) {
      appliedOperations.record({ operationId: op && op.id, status: OPERATION_STATUS.FAILED, evidence });
    }
    return finish({ executionOk: false, outcome: 'plan-invalid', authority: AUTHORITY.PLAN_ONLY, detail: validity.problems });
  }

  // Task-level idempotency is the fallback when the adapter has no per-operation classifier.
  const perOperation = typeof adapter.classifyOperation === 'function';
  let taskClass = null;
  if (!perOperation) {
    try {
      taskClass = adapter.classifyIdempotency(observation.primary, task);
    } catch (err) {
      taskClass = { state: CLASSIFICATION.FAIL_CLOSED, evidence: { threw: true, detail: err.message } };
    }
    classification = taskClass.state;
  }

  groupsInfo = strat.groups(operations, options.batch || {});

  // ---- groups -------------------------------------------------------------------------------------
  for (const [groupIndex, group] of groupsInfo.groups.entries()) {
    if (stopped) break;
    executedGroups += 1;
    metrics.count('batches').count('batchOperations', group.length);
    if (groupsInfo.barrierBefore[groupIndex]) metrics.count('dependencyBarriers');

    // Targeting observation: the current one if the session has not moved since, else a fresh one.
    if (!requireFresh(observation, handle, 'group-targeting').ok || !observation.ok) {
      observation = await observeFresh(handle, adapter, observeRequest);
    }
    const groupObservation = observation;
    if (!groupObservation.ok) {
      appliedOperations.record({ operationId: group[0].id, status: OPERATION_STATUS.FAILED, evidence: { stage: 'targeting-observation', outcome: groupObservation.outcome } });
      stop(group[0].id, `targeting-${groupObservation.outcome}`);
      break;
    }

    // Classify and resolve, in plan order. A pre-dispatch failure stops the group at that point;
    // the operations queued before it are still dispatched and verified.
    const queue = [];
    for (const op of group) {
      const index = operations.indexOf(op);
      const state = classifyOne(adapter, groupObservation, op, task, perOperation, taskClass);

      if (state.state === CLASSIFICATION.SATISFIED) {
        // SATISFIED: dispatch 0, writes 0. The current state already carries the intended result,
        // so a dependent operation may target it.
        appliedOperations.record({ operationId: op.id, status: OPERATION_STATUS.SATISFIED, evidence: state.evidence || null });
        metrics.count('operationsSatisfied');
        results.set(op.id, Object.freeze({ operationId: op.id, status: OPERATION_STATUS.SATISFIED, evidence: state.evidence || null, identity }));
        continue;
      }
      if (state.state !== CLASSIFICATION.APPLY) {
        // FAIL_CLOSED (or an unknown state, treated the same): dispatch 0, the task stops.
        appliedOperations.record({ operationId: op.id, status: OPERATION_STATUS.FAILED, evidence: { stage: 'classification', state: state.state, evidence: state.evidence || null } });
        stop(op.id, 'operation-fail-closed');
        break;
      }

      const gate = requireMutationAuthority(groupObservation, handle, 'mutation-targeting');
      if (!gate.ok) {
        appliedOperations.record({ operationId: op.id, status: OPERATION_STATUS.FAILED, evidence: { stage: 'targeting-authority', gate } });
        stop(op.id, gate.outcome);
        break;
      }

      const refContext = refView(op, results, identity);
      if (!refContext.ok) {
        appliedOperations.record({ operationId: op.id, status: OPERATION_STATUS.FAILED, evidence: { stage: 'ref', reason: refContext.reason } });
        stop(op.id, refContext.outcome);
        break;
      }

      let resolved;
      try {
        resolved = adapter.resolveTarget(groupObservation.primary, op.selector || {}, { refs: refContext.refs, operation: op });
      } catch (err) {
        resolved = { ok: false, ambiguity: 'RESOLVE_THREW', detail: err.message };
      }
      if (!resolved || resolved.ok !== true) {
        appliedOperations.record({ operationId: op.id, status: OPERATION_STATUS.FAILED, evidence: { stage: 'resolve', ambiguity: resolved && resolved.ambiguity, detail: resolved && resolved.detail } });
        stop(op.id, 'target-resolution-failed');
        break;
      }
      queue.push({ op, index, target: resolved.target });
    }

    if (queue.length === 0) continue;

    // ---- dispatch -------------------------------------------------------------------------------
    const attempts = await dispatchQueue(handle, adapter, queue);
    if (attempts.batchRound) batchDispatchRounds += 1;

    if (attempts.violation) {
      // Fail closed: no L1 is run against a round whose own transport contract was broken. What
      // reached the editor is not attributable per operation, so every operation of the round is
      // uncertain -- or, if nothing reached the editor at all, certainly not applied.
      for (const a of attempts.list) {
        const dispatched = a.generationAfter > a.generationBefore;
        if (dispatched) {
          appliedOperations.record({
            operationId: a.op.id,
            status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN,
            evidence: attempts.violation,
            uncertainty: {
              reason: UNCERTAINTY_REASON.BATCH_HOOK_CONTRACT_VIOLATION,
              operationIndex: a.index,
              dispatchAttempted: true,
              dispatchEvidence: attempts.violation,
              mutationGeneration: a.generationAfter,
            },
          });
          metrics.count('operationsUncertain');
        } else {
          appliedOperations.record({ operationId: a.op.id, status: OPERATION_STATUS.FAILED, evidence: attempts.violation });
        }
      }
      stop(attempts.list[0].op.id, 'batch-hook-contract-violation', attempts.violation);
      if (attempts.list.some(a => a.generationAfter > a.generationBefore)) observation = await observeFresh(handle, adapter, observeRequest);
      continue;
    }
    const anyDispatched = attempts.list.some(a => a.generationAfter > a.generationBefore);

    // ---- fresh read, then L1 per operation --------------------------------------------------------
    let post = null;
    if (anyDispatched) post = await observeFresh(handle, adapter, observeRequest);

    for (const a of attempts.list) {
      let l1 = null;
      const returnedOk = !a.threw && a.dispatchResult && a.dispatchResult.ok === true && a.dispatchResult.partiallyApplied !== true;
      if (returnedOk && post && post.ok) {
        l1 = await verifyPrimitive({ handle, adapter, operation: a.op, before: groupObservation, after: post });
        primitiveVerification.push(l1);
      }
      // PRIMITIVE_INTEGRAL: the preservation proof is part of the primitive's success, so it runs
      // BEFORE the register decision and only when L1 passed (nothing to preserve otherwise).
      const mode = preservationModeOf(a.op);
      let integral = null;
      if (mode === PRESERVATION_MODE.PRIMITIVE_INTEGRAL && l1 && l1.verdict === VERDICT.PASS) {
        integral = await verifyPreservation({ handle, adapter, operation: a.op, before: groupObservation, after: post, required: true });
        preservationVerification.push(integral);
      }
      const c = classifyDispatchAttempt({
        generationBefore: a.generationBefore,
        generationAfter: a.generationAfter,
        threw: a.threw,
        dispatchResult: a.dispatchResult,
        verification: l1,
        operationIndex: a.index,
        preservationMode: mode,
        integralPreservation: integral,
      });
      appliedOperations.record({
        operationId: a.op.id,
        status: c.status,
        verification: c.verification,
        uncertainty: c.uncertainty,
        evidence: l1 ? l1.evidence : (a.threw ? { threw: true, detail: String(a.threw.message || a.threw) } : a.dispatchResult),
      });

      if (c.status === OPERATION_STATUS.APPLIED) {
        metrics.count('operationsApplied').count('writes');
        // A ref names the adapter's OWN semantic result, not the verifier layer's envelope around it.
        const semantic = l1.evidence && l1.evidence.evidence !== undefined ? l1.evidence.evidence : l1.evidence;
        results.set(a.op.id, Object.freeze({ operationId: a.op.id, status: OPERATION_STATUS.APPLIED, evidence: semantic, identity }));
        if (mode === PRESERVATION_MODE.POST_PRIMITIVE) {
          const l3 = await verifyPreservation({ handle, adapter, operation: a.op, before: groupObservation, after: post, required: true });
          preservationVerification.push(l3);
          if (!isSatisfied(l3.verdict)) stop(a.op.id, 'preservation-failed', l3.verdict);
        }
      } else if (c.status === OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN) {
        metrics.count('operationsUncertain');
        stop(a.op.id, `operation-uncertain-${c.uncertainty.reason.toLowerCase()}`);
      } else {
        stop(a.op.id, 'dispatch-refused-pre-dispatch');
      }
    }
    // Queued operations whose dispatch never started (a per-operation dispatch failed before them)
    // are left without an entry here and recorded SKIPPED below.
    if (post) observation = post;
  }

  // ---- skipped tail --------------------------------------------------------------------------------
  const recorded = new Set(appliedOperations.entries().map(e => e.operationId));
  for (const op of operations) {
    if (!recorded.has(op.id)) {
      appliedOperations.record({ operationId: op.id, status: OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE, evidence: { stoppedAt: stopped && stopped.operationId } });
      metrics.count('operationsSkipped');
    }
  }

  // ---- final evidence and whole-task verification ---------------------------------------------------
  let finalObservation = observation;
  if (!finalObservation.ok || !requireFresh(finalObservation, handle, 'final-observation').ok) {
    finalObservation = await observeFresh(handle, adapter, observeRequest);
  }
  finalEvidence = Object.freeze({
    observed: finalObservation.ok === true,
    outcome: finalObservation.ok ? 'final-observation-fresh' : finalObservation.outcome,
    generation: finalObservation.freshness.generation,
    sequence: finalObservation.freshness.sequence,
    partialFailure: stopped !== null,
  });

  if (stopped) {
    return finish({ executionOk: false, outcome: stopped.outcome, authority: failureAuthority(appliedOperations), detail: stopped });
  }
  if (!finalObservation.ok) {
    return finish({ executionOk: false, outcome: 'final-observation-failed', authority: failureAuthority(appliedOperations) });
  }

  const generationBeforeL2 = currentGeneration(handle);
  wholeTaskVerification = await verifyWholeTask({ handle, adapter, task, finalObservation });
  if (currentGeneration(handle) !== generationBeforeL2) {
    // The whole-task verifier is read-only by contract. One that dispatched has tried to repair
    // what it was supposed to judge; its verdict is void and the task fails closed.
    return finish({ executionOk: false, outcome: 'whole-task-verifier-mutated', authority: failureAuthority(appliedOperations) });
  }
  if (wholeTaskVerification.verdict !== VERDICT.PASS) {
    return finish({ executionOk: false, outcome: 'whole-task-verification-failed', authority: failureAuthority(appliedOperations) });
  }

  const applied = appliedOperations.countBy(OPERATION_STATUS.APPLIED);
  return finish({
    executionOk: true,
    outcome: applied > 0 ? 'execution-complete' : 'task-satisfied',
    authority: applied > 0 ? AUTHORITY.LIVE_VERIFY : AUTHORITY.LIVE_READ,
    noOp: applied === 0,
  });

  function finish({ executionOk, outcome, authority, noOp = false, detail = null }) {
    const base = buildExecutionResult({
      executionOk,
      outcome,
      authority,
      noOp,
      appliedOperations,
      classification,
      primitiveVerification,
      wholeTaskVerification,
      preservationVerification: preservationVerification.length ? preservationVerification : null,
      strategy: strat.name,
      plannedCount: operations.length,
      finalEvidence,
      batchEvidence: groupsInfo ? Object.freeze({
        groups: groupsInfo.groups.length,
        executedGroups,
        dependencyBarriersPlanned: groupsInfo.dependencyBarriers,
        unprovenCuts: groupsInfo.unprovenCuts,
        sizeBreaks: groupsInfo.sizeBreaks,
        batchDispatchRounds,
      }) : null,
      transaction: strat.transaction ? Object.freeze({
        ...strat.transaction,
        stoppedAt: stopped ? stopped.operationId : null,
        verifiedPrefix: Object.freeze(appliedOperations.entries()
          .filter(e => e.status === OPERATION_STATUS.APPLIED && e.verification === VERIFICATION.LIVE_VERIFY)
          .map(e => e.operationId)),
      }) : null,
    });
    const withDetail = detail === null ? base : Object.freeze({ ...base, stopDetail: detail });
    // The receipt is derived LAST, from state that already decided everything.
    return Object.freeze({ ...withDetail, receiptEvidence: buildReceipt({ execution: base, task, sessionId: handle.id }) });
  }
}

function classifyOne(adapter, observation, op, task, perOperation, taskClass) {
  if (!perOperation) return taskClass;
  try {
    const out = adapter.classifyOperation(observation.primary, op, { task });
    return out && typeof out.state === 'string' ? out : { state: CLASSIFICATION.FAIL_CLOSED, evidence: { reason: 'classifier returned no state' } };
  } catch (err) {
    return { state: CLASSIFICATION.FAIL_CLOSED, evidence: { threw: true, detail: err.message } };
  }
}

/**
 * The refs an operation may use: only results of EARLIER operations of THIS task in THIS session.
 * The value handed to the adapter is the semantic result (verification evidence), never a locator;
 * the adapter must find it again in the fresh observation it is resolving against.
 */
function refView(op, results, identity) {
  const ref = op.selector && op.selector.ref;
  if (ref === undefined) return { ok: true, refs: Object.freeze({}) };
  const result = results.get(ref);
  if (!result) {
    return { ok: false, outcome: 'ref-unresolved', reason: `ref ${ref} has no result in this task` };
  }
  if (result.identity !== identity) {
    return { ok: false, outcome: 'ref-cross-session', reason: `ref ${ref} belongs to another session` };
  }
  return { ok: true, refs: Object.freeze({ [ref]: Object.freeze({ operationId: result.operationId, status: result.status, evidence: result.evidence }) }) };
}

/**
 * One dispatch round for a queue. With `executePrimitiveBatch` and more than one operation, the
 * whole queue is ONE physical round; otherwise each operation is dispatched in order and the
 * round stops at the first failure. The adapter reaches the editor only through the `dispatch`
 * function it is handed, which is core/session/handle.cjs `dispatch` -- so every physical command
 * bumps the mutation generation, and the generation is the evidence the uncertainty classifier uses.
 */
async function dispatchQueue(handle, adapter, queue) {
  const dispatchFn = command => dispatch(handle, command);
  const list = [];

  if (queue.length > 1 && typeof adapter.executePrimitiveBatch === 'function') {
    // THE BATCH HOOK CONTRACT: one logical round = exactly ONE physical dispatch. It is enforced at
    // runtime, not trusted: a second call is refused BEFORE it reaches the editor (so a faulty hook
    // cannot mutate further), a call after the round is over is refused too, and zero calls with a
    // success report is a violation as well. `commands` stays the true physical count either way,
    // because the only calls counted are the ones that went through core dispatch.
    let calls = 0;
    let refused = 0;
    let open = true;
    const guardedDispatch = command => {
      if (!open) { refused += 1; throw new Error('batch dispatch capability expired: the round is over'); }
      if (calls >= 1) { refused += 1; throw new Error('batch hook contract: one physical dispatch per round'); }
      calls += 1;
      return dispatch(handle, command);
    };
    const generationBefore = currentGeneration(handle);
    let threw = null;
    let result = null;
    try {
      result = await adapter.executePrimitiveBatch(
        handle,
        queue.map(q => ({ operation: q.op, target: q.target })),
        { dispatch: guardedDispatch }
      );
    } catch (err) {
      threw = err;
    }
    open = false;
    const generationAfter = currentGeneration(handle);
    const reportedOk = !threw && result && result.ok === true;
    const violation = refused > 0 || (calls === 0 && reportedOk)
      ? Object.freeze({ contract: 'one physical dispatch per batch round', dispatchCalls: calls, refusedCalls: refused, reportedOk: !!reportedOk })
      : null;
    if (violation) {
      return {
        list: queue.map(q => ({ op: q.op, index: q.index, generationBefore, generationAfter, threw, dispatchResult: result })),
        batchRound: true,
        violation,
      };
    }
    for (const [i, q] of queue.entries()) {
      const per = result && Array.isArray(result.perOperation) ? result.perOperation[i] : result;
      list.push({ op: q.op, index: q.index, generationBefore, generationAfter, threw, dispatchResult: per || null });
    }
    return { list, batchRound: true };
  }

  for (const q of queue) {
    const generationBefore = currentGeneration(handle);
    let threw = null;
    let result = null;
    try {
      result = await adapter.executePrimitive(handle, q.op, { dispatch: dispatchFn, target: q.target });
    } catch (err) {
      threw = err;
    }
    const generationAfter = currentGeneration(handle);
    list.push({ op: q.op, index: q.index, generationBefore, generationAfter, threw, dispatchResult: result });
    if (threw || !result || result.ok !== true) break;
  }
  return { list, batchRound: false };
}

function failureAuthority(appliedOperations) {
  const entries = appliedOperations.entries();
  if (entries.some(e => e.status === OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN)) return AUTHORITY.DISPATCH_ONLY;
  if (entries.some(e => e.status === OPERATION_STATUS.APPLIED)) return AUTHORITY.LIVE_VERIFY;
  return AUTHORITY.PLAN_ONLY;
}

/** The lifecycle-facing form: `withSession(options, adapter, createExecutor(task, {strategy}))`. */
function createExecutor(task, { strategy = 'sequential', ...options } = {}) {
  return (handle, { appliedOperations, adapter }) =>
    executeTask({ handle, adapter, task, strategy, appliedOperations, options });
}

module.exports = { executeTask, createExecutor, STRATEGIES };
