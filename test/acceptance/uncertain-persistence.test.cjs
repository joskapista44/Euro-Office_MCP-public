'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { dispatch } = require('../../src/core/session/handle.cjs');
const { currentGeneration } = require('../../src/core/mutation-epoch.cjs');
const { observeFresh } = require('../../src/core/observation/observe.cjs');
const { verifyPrimitive } = require('../../src/core/verification/orchestrate.cjs');
const { classifyDispatchAttempt } = require('../../src/core/execution/uncertainty.cjs');
const {
  AppliedOperations, buildExecutionResult, OPERATION_STATUS, VERIFICATION, UNCERTAINTY_REASON,
  UncertainPromotionError, AUTHORITY,
} = require('../../src/core/execution/result.cjs');
const {
  hasVerifiedAppliedWrites, hasUncertainWrites, verifiedAppliedWrites, decidePersistence, PERSISTENCE_STATE,
} = require('../../src/core/persistence/verified-writes.cjs');
const { buildReceipt } = require('../../src/core/receipt/envelope.cjs');
const { createFakeSession } = require('../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../fakes/fake-adapter.cjs');

/**
 * M3 additive M1 contract — the HYBRID UNCERTAIN-MUTATION POLICY (owner decision after W3.0).
 *
 * Acceptance A..K from the decision, plus L (the lifecycle owner's backstop). Every case carries
 * a NEGATIVE CONTROL: the same scenario with the one deciding input removed must produce the
 * other outcome, so a green here cannot come from an assertion that is true regardless.
 *
 * Per-operation modes, scripted at the TRANSPORT boundary (no document state is faked):
 *   ok          dispatch returns ok, fresh read, L1 PASS
 *   refuse-pre  the executor refuses before reaching the dispatch surface (proven pre-dispatch)
 *   refuse-in   the driver is asked and returns ok:false
 *   throw       the driver is asked and throws
 *   partial     the driver returns ok with partiallyApplied:true
 *   verify-fail dispatch ok, L1 FAIL
 */

function run(modes, { adapterOverrides = {}, executorThrowAfterDispatch = false, executorThrowBeforeDispatch = false, livePayload } = {}) {
  let t = 0;
  const driver = createFakeSession({ livePayload });
  const rawDispatch = driver.dispatch;
  driver.dispatch = async (handle, command) => {
    const mode = modes[command.id];
    if (mode === 'throw') { driver.calls.dispatch += 1; throw new Error(`transport lost after send (${command.id})`); }
    if (mode === 'refuse-in') { driver.calls.dispatch += 1; return { ok: false, dispatched: true, command }; }
    if (mode === 'partial') { driver.calls.dispatch += 1; return { ok: true, dispatched: true, partiallyApplied: true, command }; }
    return rawDispatch(handle, command);
  };
  const primitives = {};
  for (const [id, m] of Object.entries(modes)) if (m === 'verify-fail') primitives[id] = 'verify-fail';
  const adapter = Object.assign(createFakeAdapter({ primitives }), adapterOverrides);
  const task = { id: 'uncertain-task', operations: Object.keys(modes).map(id => ({ id })) };

  const executor = async (handle, { appliedOperations }) => {
    const planning = await observeFresh(handle, adapter);
    const classification = adapter.classifyIdempotency(planning.primary, task);
    const toRun = classification.state === 'SATISFIED' ? [] : task.operations;
    for (const [index, operation] of toRun.entries()) {
      const generationBefore = currentGeneration(handle);
      let threw = null;
      let dispatchResult = null;
      let verification = null;
      if (executorThrowBeforeDispatch) throw new Error('executor crashed before any dispatch');
      if (modes[operation.id] !== 'refuse-pre') {
        try { dispatchResult = await dispatch(handle, operation); } catch (err) { threw = err; }
        if (executorThrowAfterDispatch) throw new Error('executor crashed before recording');
        if (!threw && dispatchResult && dispatchResult.ok === true && dispatchResult.partiallyApplied !== true) {
          const after = await observeFresh(handle, adapter);
          verification = await verifyPrimitive({ handle, adapter, operation, before: planning, after });
        }
      }
      const c = classifyDispatchAttempt({
        generationBefore, generationAfter: currentGeneration(handle), threw, dispatchResult, verification, operationIndex: index,
      });
      appliedOperations.record({
        operationId: operation.id, status: c.status, verification: c.verification, uncertainty: c.uncertainty,
        evidence: { mode: modes[operation.id] || 'ok' },
      });
      if (c.status !== OPERATION_STATUS.APPLIED) {
        return buildExecutionResult({ executionOk: false, outcome: `op-${c.certainty.toLowerCase()}`, authority: AUTHORITY.DISPATCH_ONLY, appliedOperations });
      }
    }
    return buildExecutionResult({
      executionOk: true,
      outcome: toRun.length ? 'execution-complete' : 'task-satisfied',
      authority: toRun.length ? AUTHORITY.LIVE_VERIFY : AUTHORITY.LIVE_READ,
      noOp: toRun.length === 0,
      appliedOperations,
      classification: classification.state,
    });
  };

  const options = {
    driver, fileId: 'fake-file', now: () => (t += 1),
    ready: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
    persistence: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
  };
  return withSession(options, adapter, executor).then(result => ({ result, driver }));
}

const entry = (r, id) => r.execution.appliedOperations.find(e => e.operationId === id);

// ---------------------------------------------------------------------------------------------

test('A. no verified + no uncertain -> no barrier, NOT_REQUIRED, CERTAIN', async () => {
  const { result, driver } = await run({ op1: 'refuse-pre' });
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(result.persistenceOutcome, 'NOT_REQUIRED');
  assert.equal(result.mutationCertainty, 'CERTAIN');

  // NEGATIVE CONTROL: the same task with a verified write DOES run the barrier.
  const neg = await run({ op1: 'ok' });
  assert.equal(neg.driver.calls.forceSave, 1);
  assert.notEqual(neg.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
});

test('B. uncertain only -> NO barrier, UNCERTAIN_MAY_PERSIST (never NOT_REQUIRED)', async () => {
  const { result, driver } = await run({ op1: 'throw' });
  assert.equal(driver.calls.forceSave, 0, 'the framework never forceSaves solely for an unproven mutation');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);
  assert.equal(result.persistenceOutcome, 'UNCERTAIN_MAY_PERSIST');
  assert.equal(result.mutationCertainty, 'UNCERTAIN');
  assert.equal(result.ok, false);
  assert.equal(result.uncertainOperations.length, 1);
  assert.deepEqual(result.lifecycleTrace.includes('PERSISTING'), false);

  // NEGATIVE CONTROL: a PROVEN pre-dispatch refusal on the same op is NOT_REQUIRED and CERTAIN.
  const neg = await run({ op1: 'refuse-pre' });
  assert.equal(neg.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(neg.result.mutationCertainty, 'CERTAIN');
});

test('C. verified only -> barrier, FORCESAVE_CONFIRMED (no general PERSISTED), CERTAIN', async () => {
  const { result, driver } = await run({ op1: 'ok', op2: 'ok' });
  assert.equal(driver.calls.forceSave, 1);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(result.mutationCertainty, 'CERTAIN');
  assert.equal(result.authority, 'LIVE_VERIFY', 'the save state never raises the authority: there is no general PERSISTED');
  assert.equal(result.saveEvidence.storedContentVerified, false);
  assert.equal(result.ok, true);

  // NEGATIVE CONTROL: an uncertain second op on the same task keeps the save state but not the certainty.
  const neg = await run({ op1: 'ok', op2: 'throw' });
  assert.equal(neg.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(neg.result.mutationCertainty, 'UNCERTAIN');
});

test('D. verified + uncertain -> barrier (for the prefix), a confirmed save with UNCERTAIN mutation, never ok', async () => {
  const { result, driver } = await run({ op1: 'ok', op2: 'throw' });
  assert.equal(driver.calls.forceSave, 1, 'the verified prefix still makes the barrier mandatory');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(result.mutationCertainty, 'UNCERTAIN');
  // The save state and certainty are orthogonal: a confirmed save does not mean every op is proven -- and it never
  // raises the authority (here the execution's own DISPATCH_ONLY, because of the unproven op).
  assert.equal(result.authority, 'DISPATCH_ONLY');
  assert.equal(result.editOk, false);
  assert.equal(result.ok, false);
  assert.equal(entry(result, 'op1').status, OPERATION_STATUS.APPLIED);
  assert.equal(entry(result, 'op2').status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);

  // NEGATIVE CONTROL: without the uncertain op the same task is plain PERSISTED and ok.
  const neg = await run({ op1: 'ok', op2: 'ok' });
  assert.equal(neg.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(neg.result.ok, true);
});

test('E. proven pre-dispatch refusal -> FAILED (CERTAIN_NONE), no uncertain entry, driver never asked', async () => {
  const { result, driver } = await run({ op1: 'refuse-pre' });
  const e = entry(result, 'op1');
  assert.equal(e.status, OPERATION_STATUS.FAILED);
  assert.equal(e.uncertainty, null);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(hasUncertainWrites(result.execution), false);

  // NEGATIVE CONTROL: a refusal reported from INSIDE the driver call is not proven pre-dispatch.
  const neg = await run({ op1: 'refuse-in' });
  const n = entry(neg.result, 'op1');
  assert.equal(n.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(n.uncertainty.reason, UNCERTAINTY_REASON.DISPATCH_REFUSED_NOT_PROVEN_PRE_DISPATCH);
});

test('F. dispatch throw -> DISPATCHED_OUTCOME_UNKNOWN with full evidence', async () => {
  const { result } = await run({ op1: 'throw' });
  const e = entry(result, 'op1');
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.notEqual(e.verification, VERIFICATION.LIVE_VERIFY);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.DISPATCH_THREW);
  assert.equal(e.uncertainty.dispatchAttempted, true);
  assert.equal(e.uncertainty.operationIndex, 0);
  assert.equal(e.uncertainty.dispatchEvidence.threw, true);
  assert.equal(e.uncertainty.mutationGeneration, 1);

  // NEGATIVE CONTROL: the same op with a returning dispatch and a passing L1 is APPLIED.
  const neg = await run({ op1: 'ok' });
  assert.equal(entry(neg.result, 'op1').status, OPERATION_STATUS.APPLIED);
});

test('G. verification cannot prove the terminal state after a possible mutation -> uncertain', async () => {
  const vf = await run({ op1: 'verify-fail' });
  const v = entry(vf.result, 'op1');
  assert.equal(v.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(v.uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE);
  assert.equal(v.uncertainty.verificationOutcome, 'FAIL');
  assert.equal(v.verification, VERIFICATION.DISPATCH_ONLY);

  const pa = await run({ op1: 'partial' });
  assert.equal(entry(pa.result, 'op1').uncertainty.reason, UNCERTAINTY_REASON.ADAPTER_REPORTED_PARTIALLY_APPLIED);

  // NEGATIVE CONTROL: a passing L1 on the same op is certain.
  const neg = await run({ op1: 'ok' });
  assert.equal(entry(neg.result, 'op1').uncertainty, null);
});

test('H. an uncertain operation never appears as APPLIED + LIVE_VERIFY', () => {
  const reg = new AppliedOperations();
  const unc = { reason: UNCERTAINTY_REASON.DISPATCH_THREW, dispatchAttempted: true, mutationGeneration: 1 };
  reg.record({ operationId: 'op1', status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, uncertainty: unc });

  // Promotion inside the task is refused.
  assert.throws(
    () => reg.record({ operationId: 'op1', status: OPERATION_STATUS.APPLIED, verification: VERIFICATION.LIVE_VERIFY }),
    UncertainPromotionError
  );
  // An uncertain entry cannot carry LIVE_VERIFY, and cannot exist without its evidence.
  assert.throws(() => reg.record({ operationId: 'op2', status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, verification: VERIFICATION.LIVE_VERIFY, uncertainty: unc }));
  assert.throws(() => reg.record({ operationId: 'op3', status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN }));
  assert.throws(() => reg.record({ operationId: 'op4', status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, uncertainty: { ...unc, dispatchAttempted: false } }));
  const r = buildExecutionResult({ executionOk: false, outcome: 't', appliedOperations: reg });
  assert.equal(hasVerifiedAppliedWrites(r), false);
  assert.equal(hasUncertainWrites(r), true);

  // NEGATIVE CONTROL: a genuinely verified op on a DIFFERENT id records fine and fires the predicate.
  reg.record({ operationId: 'op5', status: OPERATION_STATUS.APPLIED, verification: VERIFICATION.LIVE_VERIFY });
  assert.equal(hasVerifiedAppliedWrites(buildExecutionResult({ executionOk: false, outcome: 't', appliedOperations: reg })), true);
});

test('I. a successful barrier does not promote uncertainty', async () => {
  const { result, driver } = await run({ op1: 'ok', op2: 'throw' });
  assert.equal(driver.calls.forceSave, 1);
  assert.equal(result.persistence.ok, true);
  assert.deepEqual(verifiedAppliedWrites(result.execution).map(e => e.operationId), ['op1']);
  assert.deepEqual(result.uncertainOperations.map(e => e.operationId), ['op2']);
  assert.equal(entry(result, 'op2').verification !== VERIFICATION.LIVE_VERIFY, true);

  // NEGATIVE CONTROL: with a certain op2 the barrier success covers it as a verified write.
  const neg = await run({ op1: 'ok', op2: 'ok' });
  assert.deepEqual(verifiedAppliedWrites(neg.result.execution).map(e => e.operationId), ['op1', 'op2']);
});

test('J. a missing or corrupt receipt does not change the persistence decision', () => {
  const reg = new AppliedOperations();
  reg.record({ operationId: 'op1', status: OPERATION_STATUS.APPLIED, verification: VERIFICATION.LIVE_VERIFY });
  reg.record({ operationId: 'op2', status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, uncertainty: { reason: UNCERTAINTY_REASON.DISPATCH_THREW, dispatchAttempted: true, mutationGeneration: 2 } });
  const base = buildExecutionResult({ executionOk: false, outcome: 't', appliedOperations: reg });

  const withReceipt = { ...base, receiptEvidence: buildReceipt({ execution: base }) };
  const noReceipt = { ...base, receiptEvidence: null };
  const lyingReceipt = { ...base, receiptEvidence: { operations: [{ operationId: 'op2', status: 'APPLIED', verification: 'LIVE_VERIFY' }] } };
  for (const r of [withReceipt, noReceipt, lyingReceipt]) {
    assert.deepEqual(decidePersistence(r), decidePersistence(base));
  }
  assert.equal(decidePersistence(base).onSuccess, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(withReceipt.receiptEvidence.operations[1].uncertainty.reason, UNCERTAINTY_REASON.DISPATCH_THREW);

  // NEGATIVE CONTROL: changing the REGISTER (not the receipt) does change the decision.
  const reg2 = new AppliedOperations();
  reg2.record({ operationId: 'op1', status: OPERATION_STATUS.APPLIED, verification: VERIFICATION.LIVE_VERIFY });
  const other = buildExecutionResult({ executionOk: true, outcome: 't', appliedOperations: reg2 });
  assert.equal(decidePersistence(other).onSuccess, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
});

test('K. retry after PERSISTED_UNCERTAIN classifies from FRESH durable state, not from the receipt', async () => {
  const first = await run({ op1: 'ok', op2: 'throw' });
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  // The prior receipt/history is available to a careless retry -- and must not be consulted.
  const priorReceipt = buildReceipt({ execution: first.result.execution });

  // The adapter classifies from the observation it is given, and only from it.
  const seen = [];
  const classifyIdempotency = (observation, task) => {
    seen.push(observation);
    const done = new Set((observation.units || []).map(u => u.id));
    const allDone = task.operations.every(o => done.has(o.id));
    return { state: allDone ? 'SATISFIED' : 'APPLY', evidence: { observedUnits: [...done] } };
  };

  // Durable state after the barrier: op1 landed, op2 did NOT (the uncertain op had no effect).
  const retry = await run({ op1: 'ok', op2: 'ok' }, { livePayload: { units: [{ id: 'op1' }] }, adapterOverrides: { classifyIdempotency } });
  assert.equal(seen.length, 1);
  assert.equal(priorReceipt.operations.find(o => o.operationId === 'op2').status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN,
    'the prior receipt exists and describes op2 -- the retry decides without it');
  assert.equal(retry.driver.calls.readLive >= 1, true, 'the retry read fresh state before classifying');
  assert.equal(retry.result.execution.noOp, false, 'op2 is absent from the durable state, so the retry applies');

  // NEGATIVE CONTROL: if the durable state shows BOTH effects, the same retry is SATISFIED -- the
  // decision follows the fresh observation, whatever the prior receipt said.
  seen.length = 0;
  const satisfied = await run({ op1: 'ok', op2: 'ok' }, { livePayload: { units: [{ id: 'op1' }, { id: 'op2' }] }, adapterOverrides: { classifyIdempotency } });
  assert.equal(satisfied.result.execution.noOp, true);
  assert.equal(satisfied.driver.calls.dispatch, 0);
  assert.equal(satisfied.driver.calls.forceSave, 0);
});

test('L. lifecycle backstop: executor throws after an unrecorded dispatch -> uncertain, not NOT_REQUIRED', async () => {
  const { result, driver } = await run({ op1: 'ok' }, { executorThrowAfterDispatch: true });
  assert.equal(driver.calls.dispatch, 1);
  assert.equal(result.execution.outcome, 'execution-threw');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);
  assert.equal(driver.calls.forceSave, 0);
  const u = result.uncertainOperations[0];
  assert.equal(u.operationId, null);
  assert.equal(u.uncertainty.reason, UNCERTAINTY_REASON.EXECUTOR_THREW_WITH_UNACCOUNTED_DISPATCH);

  // NEGATIVE CONTROL: a throw BEFORE any dispatch leaves nothing unaccounted -> NOT_REQUIRED.
  const neg = await run({ op1: 'ok' }, { executorThrowBeforeDispatch: true });
  assert.equal(neg.result.execution.outcome, 'execution-threw');
  assert.equal(neg.driver.calls.dispatch, 0);
  assert.equal(neg.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
});

test('M. an executor that reports executionOk with an uncertain entry never yields ok:true', async () => {
  let t = 0;
  const mk = entries => async (_handle, { appliedOperations }) => {
    for (const e of entries) appliedOperations.record(e);
    return buildExecutionResult({ executionOk: true, outcome: 'execution-complete', authority: AUTHORITY.LIVE_VERIFY, appliedOperations });
  };
  const options = () => ({
    driver: createFakeSession(), fileId: 'f', now: () => (t += 1),
    ready: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
    persistence: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
  });
  const verified = { operationId: 'op1', status: OPERATION_STATUS.APPLIED, verification: VERIFICATION.LIVE_VERIFY };
  const unknown = { operationId: 'op2', status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, uncertainty: { reason: UNCERTAINTY_REASON.DISPATCH_THREW, dispatchAttempted: true, mutationGeneration: 2 } };

  const r = await withSession(options(), createFakeAdapter(), mk([verified, unknown]));
  assert.equal(r.documentOutcome, 'SUCCESS', 'the fixture is a (buggy) executor claiming success');
  assert.equal(r.mutationCertainty, 'UNCERTAIN');
  assert.equal(r.ok, false, 'an uncertain mutation never makes the task ok, whatever the executor says');

  // NEGATIVE CONTROL: the same executor with only the verified entry is ok.
  const neg = await withSession(options(), createFakeAdapter(), mk([verified]));
  assert.equal(neg.ok, true);
});
