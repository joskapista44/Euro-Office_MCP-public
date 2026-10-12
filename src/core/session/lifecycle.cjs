'use strict';

const { openSession, awaitReady, closeSession } = require('./open.cjs');
const { STATES, LifecycleState } = require('./state.cjs');
const { Metrics, collectMetrics } = require('../metrics/collect.cjs');
const {
  AppliedOperations, buildExecutionResult, AUTHORITY, OPERATION_STATUS, VERIFICATION, UNCERTAINTY_REASON,
} = require('../execution/result.cjs');
const {
  decidePersistence, hasUncertainWrites, uncertainWrites, PERSISTENCE_STATE,
} = require('../persistence/verified-writes.cjs');
const { persist } = require('../persistence/barrier.cjs');
const { currentGeneration } = require('../mutation-epoch.cjs');

/**
 * withSession — the ONLY owner of open and close in the codebase (M0 4.1, owner requirement 16).
 *
 *   OPEN -> READY -> OBSERVE -> EXECUTE -> VERIFY -> (PERSIST if required) -> CLOSE -> FINALIZE
 *
 * One task = exactly one owned editor session. There is no per-operation reopen: the executor
 * receives the handle and works inside the session it was given.
 *
 * The lifecycle owner, not the executor, holds the authoritative AppliedOperations register. That
 * is what makes owner requirement 20/K provable: if the executor THROWS after a verified write,
 * the register still exists, the predicate still fires, and the barrier still runs.
 */
async function withSession(options, adapter, fn) {
  const metrics = new Metrics({ now: options.now });
  const lifecycle = new LifecycleState();
  const appliedOperations = new AppliedOperations();

  let handle = null;
  let execution = null;
  let persistence = null;
  let cleanup = null;

  lifecycle.transition(STATES.OPENING);
  try {
    handle = await openSession({
      driver: options.driver,
      url: options.url,
      user: options.user,
      pass: options.pass,
      fileId: options.fileId,
      adapterKind: adapter.kind,
      metrics,
    });
  } catch (err) {
    // F1: OPENING failed. No handle exists, so there is nothing to close and no barrier to run.
    lifecycle.transition(STATES.CLOSING).transition(STATES.CLOSED);
    execution = buildExecutionResult({
      executionOk: false,
      outcome: 'session-open-failed',
      authority: AUTHORITY.PLAN_ONLY,
      appliedOperations,
    });
    metrics.clock.mark('closed').span('totalElapsedMs', 'openEntry', 'closed');
    return finalize({
      lifecycle,
      execution,
      persistence: { ok: true, outcome: 'not-required', authority: null },
      cleanup: { ok: true, outcome: 'not-applicable', detail: 'no session was opened' },
      metrics,
      openFailure: err.openFailure || { ok: false, detail: err.message },
    });
  }

  try {
    lifecycle.transition(STATES.READY);
    const ready = await awaitReady(handle, adapter, options.ready || {});
    if (!ready.ok) {
      // F2: READY never reached -> close without any write, no barrier.
      execution = buildExecutionResult({
        executionOk: false,
        outcome: 'session-ready-failed',
        authority: AUTHORITY.PLAN_ONLY,
        appliedOperations,
      });
    } else {
      lifecycle.transition(STATES.OBSERVING).transition(STATES.EXECUTING);
      try {
        execution = await fn(handle, { appliedOperations, adapter, lifecycle });
      } catch (err) {
        // Owner requirement 20/K: an executor that throws AFTER a verified write must not lose
        // that write. The register is the lifecycle owner's, so the evidence survives the throw.
        //
        // And the M3 uncertain-mutation policy: a throw may also have interrupted a dispatch the
        // executor never got to record. The session's mutation generation counts every dispatch;
        // if it shows more than the register accounts for, the difference is recorded as
        // uncertain HERE, by the register's owner. Fail-closed toward uncertainty: the one thing
        // we may not assume after an unaccounted dispatch is that nothing happened.
        recordUnaccountedDispatches(handle, appliedOperations, err);
        execution = buildExecutionResult({
          executionOk: false,
          outcome: 'execution-threw',
          authority: appliedOperations.countBy('APPLIED') > 0 ? AUTHORITY.LIVE_VERIFY : AUTHORITY.PLAN_ONLY,
          appliedOperations,
          receiptEvidence: null,
        });
        execution = Object.freeze({ ...execution, executionError: err.message });
      }
      // taskElapsedMs closes HERE — M0 12.1: READY reached -> executeTask returned. It is marked
      // on every path where the executor actually ran, including the one where it threw, because
      // the boundary is "the executor returned control", not "the executor succeeded".
      // Paths where execution never started (open failure, readiness failure) never reach this
      // line, so taskElapsedMs stays null rather than reading as an instant task.
      metrics.clock.mark('taskDone').span('taskElapsedMs', 'ready', 'taskDone');

      if (lifecycle.state === STATES.EXECUTING) lifecycle.transition(STATES.VERIFYING);
    }

    // The persistence decision reads the AUTHORITATIVE execution state, never a receipt. The
    // barrier runs ONLY for verified writes; an uncertain-only task is deliberately not saved.
    const decision = decidePersistence(execution);
    if (decision.barrierRequired) {
      lifecycle.transition(STATES.PERSISTING);
      persistence = await persist(handle, options.persistence || {});
    } else if (decision.uncertain) {
      persistence = { ok: true, outcome: 'uncertain-not-persisted', authority: null };
    } else {
      persistence = { ok: true, outcome: 'not-required', authority: null };
    }
  } finally {
    // Close always runs, after any required persistence (invariant 10). Its outcome is captured,
    // never swallowed: a close failure is explicit in FinalTaskResult, not metrics-only (F8).
    if (lifecycle.state !== STATES.CLOSING && lifecycle.state !== STATES.CLOSED) {
      lifecycle.transition(STATES.CLOSING);
    }
    cleanup = await closeSession(handle);
    lifecycle.transition(STATES.CLOSED);
    metrics.clock.mark('closed').span('totalElapsedMs', 'openEntry', 'closed');
  }

  return finalize({ lifecycle, execution, persistence, cleanup, metrics });
}

/** The barrier's outcome and the save claim it supports (M7 save contract), with the reason for that claim. */
function saveStateOf(persistence, cleanup) {
  if (persistence.outcome === 'not-required') {
    return { persistenceOutcome: 'NOT_REQUIRED', persistenceState: PERSISTENCE_STATE.NOT_REQUIRED, saveReason: 'no-verified-write' };
  }
  if (persistence.outcome === 'uncertain-not-persisted') {
    // Never reported as NOT_REQUIRED: that is exactly the collapse the policy exists to prevent.
    return { persistenceOutcome: 'UNCERTAIN_NOT_PERSISTED', persistenceState: PERSISTENCE_STATE.UNCERTAIN_NOT_PERSISTED, saveReason: 'only-uncertain-writes' };
  }
  if (persistence.ok !== true) {
    return { persistenceOutcome: 'FAILED', persistenceState: PERSISTENCE_STATE.FAILED, saveReason: persistence.outcome || 'barrier-failed' };
  }
  const saveReason = confirmedSaveReason(persistence, cleanup);
  const persistenceState = saveReason === 'forcesave-confirmed' ? PERSISTENCE_STATE.FORCESAVE_CONFIRMED : PERSISTENCE_STATE.EDITOR_CLEAN;
  return { persistenceOutcome: 'SUCCESS', persistenceState, saveReason };
}

/**
 * Why a barrier that succeeded is (not) a confirmed save: the barrier must carry the driver's own confirmation, the
 * close must be verified, and the confirmation re-evaluated at that close must still stand (no later send, no
 * anomaly). The first failing condition is the reason; a driver without a save observer never passes the first.
 */
function confirmedSaveReason(persistence, cleanup) {
  const atBarrier = persistence.confirmation || { state: 'unconfirmed', reason: 'driver-provides-no-save-confirmation' };
  const atClose = cleanup && cleanup.confirmation;
  let reason;
  if (atBarrier.state !== 'confirmed') reason = `barrier:${atBarrier.reason}`;
  else if (!cleanup || cleanup.ok !== true) reason = 'close-not-verified';
  else if (!atClose) reason = 'close:no-confirmation-reported';
  else if (atClose.state !== 'confirmed') reason = `close:${atClose.reason}`;
  else reason = 'forcesave-confirmed';
  return reason;
}

/**
 * finalizeTask — M0 4.2. The ONLY final task-completion result, and it can exist only after the
 * lifecycle has finished: execution, persistence and cleanup are all already decided here.
 */
function finalize({ lifecycle, execution, persistence, cleanup, metrics, openFailure = null }) {
  const documentOutcome = execution.executionOk ? 'SUCCESS' : 'FAILED';

  // Mutation certainty is ORTHOGONAL to the save state (owner decision, point 7): a task can reach a confirmed save
  // and still hold an operation whose outcome nobody proved.
  const uncertain = hasUncertainWrites(execution);
  const mutationCertainty = uncertain ? 'UNCERTAIN' : 'CERTAIN';
  const cleanupOutcome = cleanup.ok === true ? 'SUCCESS' : 'FAILED';

  // M7 save contract (owner GO, owner channel). persistenceOutcome is the BARRIER's result; persistenceState is the
  // save claim. canSave false alone is at most EDITOR_CLEAN. FORCESAVE_CONFIRMED needs the driver's own paired
  // confirmation at the barrier AND the same confirmation still standing when re-evaluated at a verified close.
  const { persistenceOutcome, persistenceState, saveReason } = saveStateOf(persistence, cleanup);
  const saveSettled = persistenceState === PERSISTENCE_STATE.NOT_REQUIRED || persistenceState === PERSISTENCE_STATE.FORCESAVE_CONFIRMED;

  // editOk: the EDIT itself -- executed, verified and certain. ok: the WHOLE task -- the edit, a settled save state
  // (nothing to save, or a confirmed save) and a clean close. A task that needed a save is never complete on
  // EDITOR_CLEAN alone; a write-free SATISFIED retry needs no save and starts none.
  const editOk = documentOutcome === 'SUCCESS' && mutationCertainty === 'CERTAIN';
  const ok = editOk && saveSettled && cleanupOutcome === 'SUCCESS';

  let outcome;
  if (persistenceOutcome === 'FAILED') outcome = 'task-persistence-barrier-failed';
  else if (cleanupOutcome === 'FAILED') outcome = 'task-cleanup-failed';
  else if (documentOutcome === 'FAILED') outcome = execution.outcome;
  else if (persistenceState === PERSISTENCE_STATE.EDITOR_CLEAN) outcome = 'task-save-not-confirmed';
  else outcome = 'task-complete';

  // F7: a failed barrier downgrades the reported authority. No save state raises it: there is no general PERSISTED.
  const authority = persistenceOutcome === 'FAILED' ? 'PRIMITIVE_LIVE_VERIFY_ONLY' : execution.authority;

  // What the save claim means, for any consumer: the server's report, never a check of the stored content.
  const saveEvidence = Object.freeze({
    state: persistenceState,
    reason: saveReason,
    meaning: persistenceState === PERSISTENCE_STATE.FORCESAVE_CONFIRMED
      ? 'the server reported the save of this acknowledged editor state as successful'
      : null,
    storedContentVerified: false,
  });

  if (lifecycle.state === STATES.CLOSED) lifecycle.transition(STATES.FINALIZED);

  return Object.freeze({
    kind: 'FinalTaskResult',
    final: true,
    ok,
    editOk,
    outcome,
    documentOutcome,
    persistenceOutcome,
    persistenceState,
    saveEvidence,
    mutationCertainty,
    uncertainOperations: Object.freeze(uncertainWrites(execution)),
    cleanupOutcome,
    noOp: execution.noOp === true,
    authority,
    execution,
    persistence,
    cleanup,
    openFailure,
    receiptEvidence: execution.receiptEvidence || null,
    metrics: collectMetrics(metrics),
    lifecycleTrace: lifecycle.trace,
  });
}

/**
 * Dispatches the register does not account for, after the executor threw. Counted, not guessed:
 * every dispatch bumps the session's mutation generation, and every register entry that could
 * only exist after a dispatch accounts for one. A batch dispatch that covers several operations
 * over-accounts, which errs toward NOT adding an entry -- the executor is the primary recorder;
 * this is the owner's backstop for the case where the executor never got the chance.
 */
function recordUnaccountedDispatches(handle, appliedOperations, err) {
  const generation = currentGeneration(handle);
  const accounted = appliedOperations.entries().filter(e =>
    e.status === OPERATION_STATUS.APPLIED ||
    e.status === OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN ||
    (e.status === OPERATION_STATUS.FAILED && e.verification === VERIFICATION.DISPATCH_ONLY)
  ).length;
  if (generation <= accounted) return;
  appliedOperations.record({
    operationId: null, // task-level: which operation it belonged to is exactly what is unknown
    status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN,
    evidence: { executorError: err && err.message },
    uncertainty: {
      reason: UNCERTAINTY_REASON.EXECUTOR_THREW_WITH_UNACCOUNTED_DISPATCH,
      dispatchAttempted: true,
      dispatchEvidence: { generation, accounted, unaccounted: generation - accounted },
      mutationGeneration: generation,
    },
  });
}

module.exports = { withSession, finalize };
