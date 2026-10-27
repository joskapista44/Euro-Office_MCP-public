'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { MAX_OPERATIONS } = require('../../src/core/execution/plan.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { createDocument, createDocumentDriver, createDocumentAdapter, snapshot } = require('../fakes/fake-document.cjs');

/**
 * M3 acceptance matrix A..V (owner brief point 23), on the ONE common executor through the M1
 * lifecycle. The document is a fake with real state and physical-index locators, so wrong
 * targeting produces a genuinely wrong document rather than a scripted failure.
 *
 * Every case asserts OPEN_COUNT=1 / CLOSE_COUNT=1 and the operation accounting identity; most carry
 * an explicit NEGATIVE CONTROL (the same scenario without the deciding input).
 */

const base = () => [
  { key: 'u1', text: 'one' },
  { key: 'u2', text: 'two' },
  { key: 'u3', text: 'three' },
];

async function runTask({ units = base(), operations, expect = [], strategy = 'sequential', driverScript = {}, adapterScript = {}, batch = {} }) {
  let t = 0;
  const doc = createDocument(units);
  const driver = createDocumentDriver(doc, driverScript);
  const adapter = createDocumentAdapter(adapterScript);
  const options = {
    driver, fileId: 'doc', now: () => (t += 1),
    ready: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
    persistence: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
  };
  const task = { id: 'm3', operations, expect };
  const result = await withSession(options, adapter, createExecutor(task, { strategy, batch }));
  // Lifecycle invariants every case must hold, whatever its outcome.
  assert.equal(driver.calls.open, 1, 'OPEN_COUNT=1');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT=1');
  const acc = result.execution.operationAccounting;
  if (acc) assert.equal(acc.balanced, true, `accounting unbalanced: ${JSON.stringify(acc)}`);
  assert.equal(result.metrics.commands, driver.calls.dispatch, 'commands == physical dispatches');
  return { result, driver, adapter, doc, acc, exec: result.execution };
}

const status = (r, id) => r.execution.appliedOperations.find(e => e.operationId === id).status;
const set = (id, key, text, extra = {}) => ({ id, kind: 'set', selector: { key }, text, ...extra });

// ---------------------------------------------------------------------------------------------

test('A. single APPLY success -> APPLIED+LIVE_VERIFY, barrier, PERSISTED', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'ONE')], expect: [{ key: 'u1', text: 'ONE' }] });
  assert.equal(r.result.ok, true);
  assert.equal(r.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(status(r.result, 'a'), OPERATION_STATUS.APPLIED);
  assert.equal(r.result.metrics.writes, 1);
  assert.equal(r.result.metrics.commands, 1);
  assert.deepEqual(snapshot(r.doc), ['u1=ONE', 'u2=two', 'u3=three']);
});

test('B. single SATISFIED -> zero dispatch, zero write, no barrier', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'one')], expect: [{ key: 'u1', text: 'one' }] });
  assert.equal(r.result.ok, true);
  assert.equal(r.exec.noOp, true);
  assert.equal(r.result.metrics.commands, 0);
  assert.equal(r.result.metrics.writes, 0);
  assert.equal(r.driver.calls.forceSave, 0);
  assert.equal(r.acc.satisfied, 1);
  // NEGATIVE CONTROL: a different target text is APPLY and dispatches.
  const neg = await runTask({ operations: [set('a', 'u1', 'ONE')] });
  assert.equal(neg.result.metrics.commands, 1);
});

test('C. single FAIL_CLOSED -> zero dispatch, task fails, no barrier', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'ONE', { failClosed: true })] });
  assert.equal(r.result.ok, false);
  assert.equal(r.exec.outcome, 'operation-fail-closed');
  assert.equal(r.driver.calls.dispatch, 0);
  assert.equal(r.driver.calls.forceSave, 0);
  assert.equal(status(r.result, 'a'), OPERATION_STATUS.FAILED);
  assert.deepEqual(snapshot(r.doc), ['u1=one', 'u2=two', 'u3=three']);
});

test('D. three independent sequential operations', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'A'), set('b', 'u2', 'B'), set('c', 'u3', 'C')], expect: [{ key: 'u3', text: 'C' }] });
  assert.equal(r.result.ok, true);
  assert.equal(r.result.metrics.commands, 3);
  assert.equal(r.result.metrics.writes, 3);
  assert.equal(r.result.metrics.batches, 3, 'sequential: one group per operation');
  assert.deepEqual(snapshot(r.doc), ['u1=A', 'u2=B', 'u3=C']);
});

const chain = () => [
  { id: 'A', kind: 'insertAfter', selector: { key: 'u1' }, newKey: 'X', text: 'x' },
  { id: 'B', kind: 'append', selector: { ref: 'A' }, expectKey: 'X', suffix: '+b' },
  { id: 'C', kind: 'append', selector: { ref: 'B' }, expectKey: 'X', suffix: '+c' },
];

test('E. A -> B -> C fresh dependency chain, targets from the state each dependency produced', async () => {
  for (const strategy of ['sequential', 'transaction', 'batch']) {
    const r = await runTask({ operations: chain(), strategy, expect: [{ key: 'X', text: 'x+b+c' }] });
    assert.equal(r.result.ok, true, strategy);
    assert.deepEqual(snapshot(r.doc), ['u1=one', 'X=x+b+c', 'u2=two', 'u3=three'], strategy);
    assert.equal(r.result.metrics.dependencyBarriers, 2, `${strategy}: two dependency barriers`);
    // Every group was targeted from a read taken after the previous group's dispatch.
    assert.equal(r.result.metrics.freshReads, 4, `${strategy}: planning + one post-read per group`);
  }
});

test('F. operation 2 fails after operation 1 verified write -> stop, prefix persists, tail skipped', async () => {
  // Certain failure (proven pre-dispatch refusal).
  const r = await runTask({ operations: [set('a', 'u1', 'A'), set('b', 'u2', 'B', { lowerFails: true }), set('c', 'u3', 'C')], strategy: 'transaction' });
  assert.equal(r.result.ok, false);
  assert.deepEqual([status(r.result, 'a'), status(r.result, 'b'), status(r.result, 'c')],
    [OPERATION_STATUS.APPLIED, OPERATION_STATUS.FAILED, OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE]);
  assert.equal(r.driver.calls.dispatch, 1, 'no dispatch after the failure');
  assert.equal(r.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED, 'verified prefix persists (M1 barrier)');
  assert.equal(r.exec.transaction.atomicity, 'NON_ATOMIC');
  assert.equal(r.exec.transaction.rollback, false);
  assert.deepEqual(r.exec.transaction.verifiedPrefix, ['a']);
  assert.deepEqual(snapshot(r.doc), ['u1=A', 'u2=two', 'u3=three']);
  // Uncertain failure: the prefix still persists, but as PERSISTED_UNCERTAIN.
  const u = await runTask({ operations: [set('a', 'u1', 'A'), set('b', 'u2', 'B'), set('c', 'u3', 'C')], driverScript: { faults: { 2: 'throw-after-apply' } } });
  assert.equal(status(u.result, 'b'), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(u.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(status(u.result, 'c'), OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE);
});

test('G. required preservation fails after a verified write -> task FAIL, barrier runs', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'A', { preservationRequired: true }), set('b', 'u2', 'B')], adapterScript: { preservationFailsFor: ['a'] } });
  assert.equal(r.result.ok, false);
  assert.equal(r.exec.outcome, 'preservation-failed');
  assert.equal(status(r.result, 'a'), OPERATION_STATUS.APPLIED, 'the write itself was verified');
  assert.equal(status(r.result, 'b'), OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE);
  assert.equal(r.driver.calls.forceSave, 1);
  // NEGATIVE CONTROL: preservation passing -> task ok.
  const neg = await runTask({ operations: [set('a', 'u1', 'A', { preservationRequired: true })], adapterScript: { preservationFailsFor: [] } });
  assert.equal(neg.result.ok, true);
});

test('H. whole-task failure after all primitives pass -> FAIL, barrier still runs, verifier cannot repair', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'A')], expect: [{ key: 'u1', text: 'SOMETHING ELSE' }] });
  assert.equal(r.result.ok, false);
  assert.equal(r.exec.outcome, 'whole-task-verification-failed');
  assert.equal(status(r.result, 'a'), OPERATION_STATUS.APPLIED);
  assert.equal(r.driver.calls.forceSave, 1);
  assert.equal(r.result.metrics.commands, 1, 'the whole-task verifier dispatched nothing');
  // NEGATIVE CONTROL: the right expectation passes.
  const neg = await runTask({ operations: [set('a', 'u1', 'A')], expect: [{ key: 'u1', text: 'A' }] });
  assert.equal(neg.result.ok, true);
});

test('I. dispatch refuses before mutation -> CERTAIN_NONE, driver never asked', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'A', { lowerFails: true })] });
  assert.equal(status(r.result, 'a'), OPERATION_STATUS.FAILED);
  assert.equal(r.driver.calls.dispatch, 0);
  assert.equal(r.result.mutationCertainty, 'CERTAIN');
  assert.equal(r.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
});

test('J. dispatch throws -> uncertain, and the document shows why it must not be called "no write"', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'A')], driverScript: { faults: { 1: 'throw-after-apply' } } });
  const e = r.exec.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.DISPATCH_THREW);
  assert.equal(r.result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);
  assert.equal(r.driver.calls.forceSave, 0);
  assert.deepEqual(snapshot(r.doc), ['u1=A', 'u2=two', 'u3=three'], 'the editor DID change');
  // And a throw before the editor applied it looks identical to the caller -- hence UNKNOWN.
  const b = await runTask({ operations: [set('a', 'u1', 'A')], driverScript: { faults: { 1: 'throw-before-apply' } } });
  assert.equal(b.exec.appliedOperations[0].status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.deepEqual(snapshot(b.doc), ['u1=one', 'u2=two', 'u3=three']);
});

test('K. primitive verification fails after dispatch -> uncertain, never APPLIED+LIVE_VERIFY', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'A', { forceVerifyFail: true }), set('b', 'u2', 'B')] });
  const e = r.exec.appliedOperations.find(x => x.operationId === 'a');
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE);
  assert.notEqual(e.verification, 'LIVE_VERIFY');
  assert.equal(status(r.result, 'b'), OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE);
  assert.equal(r.driver.calls.dispatch, 1);
  // A driver refusal after it was asked is uncertain too.
  const ref = await runTask({ operations: [set('a', 'u1', 'A')], driverScript: { faults: { 1: 'refuse' } } });
  assert.equal(ref.exec.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.DISPATCH_REFUSED_NOT_PROVEN_PRE_DISPATCH);
});

const proven = (ops, key = 'g') => ops.map(o => ({ ...o, batchSafeGroup: key }));

test('L. batch of PROVEN-independent operations -> one dispatch round, per-operation L1', async () => {
  const ops = proven([set('a', 'u1', 'A'), set('b', 'u2', 'B'), set('c', 'u3', 'C')]);
  const r = await runTask({ operations: ops, strategy: 'batch', expect: [{ key: 'u2', text: 'B' }] });
  assert.equal(r.result.ok, true);
  assert.equal(r.result.metrics.commands, 1, 'one physical batch dispatch = ONE command');
  assert.equal(r.result.metrics.writes, 3, 'writes are proven APPLIED operations, not commands');
  assert.equal(r.result.metrics.batches, 1);
  assert.equal(r.result.metrics.batchOperations, 3);
  assert.equal(r.adapter.calls.verifyPrimitive, 3, 'batching never skips L1');
  assert.equal(r.result.metrics.freshReads, 2, 'planning read + one post-dispatch read');
  assert.deepEqual(snapshot(r.doc), ['u1=A', 'u2=B', 'u3=C']);
});

const shifting = ({ declareDependency, proof = false }) => [
  // A inserts before u2: every later physical index moves by one.
  { id: 'A', kind: 'insertAfter', selector: { key: 'u1' }, newKey: 'N', text: 'new', ...(proof ? { batchSafeGroup: 'g' } : {}) },
  // B edits u2. Its PHYSICAL index depends on A.
  { id: 'B', kind: 'set', selector: { key: 'u2' }, text: 'TWO', ...(declareDependency ? { dependsOn: ['A'] } : {}), ...(proof ? { batchSafeGroup: 'g' } : {}) },
  { id: 'C', kind: 'set', selector: { key: 'u3' }, text: 'THREE', ...(declareDependency ? { dependsOn: ['A'] } : {}), ...(proof ? { batchSafeGroup: 'g' } : {}) },
];
const RIGHT = ['u1=one', 'N=new', 'u2=TWO', 'u3=THREE'];

test('M. batch dependency split: the planner cuts at the dependency, and the result is right', async () => {
  const r = await runTask({ operations: shifting({ declareDependency: true, proof: true }), strategy: 'batch', expect: [{ key: 'u2', text: 'TWO' }, { key: 'u3', text: 'THREE' }] });
  assert.equal(r.result.ok, true);
  assert.equal(r.exec.batchEvidence.groups, 2, '[A] | [B, C]');
  assert.equal(r.result.metrics.dependencyBarriers, 1);
  assert.deepEqual(snapshot(r.doc), RIGHT);
});

test('M-NEG (regression, R1). the former fail-open case -- undeclared dependency, no proof -- now yields the RIGHT document', async () => {
  // Before R1 this exact scenario was grouped into ONE physical batch: B and C were targeted from
  // the pre-A observation, their stale indexes overwrote the wrong units, and the document was WRONG
  // (L1 caught it after the damage). Unknown independence is no longer batched, so each operation
  // is targeted from the state the previous one produced.
  const r = await runTask({ operations: shifting({ declareDependency: false }), strategy: 'batch', expect: [{ key: 'u2', text: 'TWO' }, { key: 'u3', text: 'THREE' }] });
  assert.equal(r.exec.batchEvidence.groups, 3, 'no physical batch without a proof');
  assert.equal(r.exec.batchEvidence.unprovenCuts, 2);
  assert.deepEqual(snapshot(r.doc), RIGHT);
  assert.equal(r.result.ok, true);
});

test('N. batch partial failure: one op in a group fails L1 -> others verified, later groups skipped', async () => {
  const ops = [...proven([set('a', 'u1', 'A'), set('b', 'u2', 'B', { forceVerifyFail: true }), set('c', 'u3', 'C')]),
    { id: 'd', kind: 'append', selector: { key: 'u1' }, suffix: '!', dependsOn: ['a'] }];
  const r = await runTask({ operations: ops, strategy: 'batch' });
  assert.equal(status(r.result, 'a'), OPERATION_STATUS.APPLIED);
  assert.equal(status(r.result, 'b'), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(status(r.result, 'c'), OPERATION_STATUS.APPLIED);
  assert.equal(status(r.result, 'd'), OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE);
  assert.equal(r.result.metrics.commands, 1, 'no dispatch after the failed group');
  assert.equal(r.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  // A whole-batch throw makes every operation of the round uncertain.
  const t = await runTask({ operations: ops.slice(0, 3).map(o => ({ ...o, forceVerifyFail: false })), strategy: 'batch', driverScript: { faults: { 1: 'throw-after-apply' } } });
  assert.deepEqual(t.exec.appliedOperations.map(e => e.status), Array(3).fill(OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN));
  assert.equal(t.result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);
});

function scale(n) {
  const units = Array.from({ length: n }, (_, i) => ({ key: `k${i}`, text: `v${i}` }));
  // Every operation targets a distinct unit, so the test planner issues the independence proof.
  const operations = units.map((u, i) => set(`op${i}`, u.key, `w${i}`, { batchSafeGroup: 'distinct-units' }));
  const expect = [{ key: `k${n - 1}`, text: `w${n - 1}` }];
  return { units, operations, expect };
}

for (const n of [1, 19, 108]) {
  test(`O/P. ${n} operations -> all strategies correct, one open, one close, accounting balanced`, async () => {
    for (const strategy of ['sequential', 'transaction', 'batch']) {
      const s = scale(n);
      const r = await runTask({ ...s, strategy });
      assert.equal(r.result.ok, true, `${strategy} ${n}`);
      assert.equal(r.acc.applied, n);
      assert.equal(r.result.metrics.writes, n);
      assert.equal(r.result.metrics.commands, strategy === 'batch' ? 1 : n);
      assert.equal(r.doc.units[n - 1].text, `w${n - 1}`);
    }
  });
}

test('Q. 1000-operation boundary: 1000 plans and runs, 1001 is refused explicitly', async () => {
  const s = scale(MAX_OPERATIONS);
  const heapBefore = process.memoryUsage().heapUsed;
  const r = await runTask({ ...s, strategy: 'batch' });
  const heapDeltaMb = (process.memoryUsage().heapUsed - heapBefore) / 1048576;
  assert.equal(r.result.ok, true);
  assert.equal(r.acc.planned, 1000);
  assert.equal(r.acc.applied, 1000);
  assert.equal(r.result.metrics.commands, 1);
  assert.ok(heapDeltaMb < 200, `heap delta ${heapDeltaMb.toFixed(1)} MB`);
  // The sequential strategy at the same scale: no recursion, so no stack growth with n.
  const seq = await runTask({ ...scale(MAX_OPERATIONS), strategy: 'sequential' });
  assert.equal(seq.result.ok, true);
  assert.equal(seq.result.metrics.commands, 1000);

  const over = scale(MAX_OPERATIONS + 1);
  const o = await runTask({ ...over, strategy: 'batch' });
  assert.equal(o.exec.outcome, 'plan-invalid');
  assert.equal(o.driver.calls.dispatch, 0);
  assert.equal(o.acc.failed, 1001);
  assert.match(JSON.stringify(o.exec.stopDetail), /above the explicit ceiling 1000/);
});

test('R. strategy equivalence: same independent task, same final state; only interactions differ', async () => {
  const ops = () => proven([set('a', 'u1', 'A'), set('b', 'u2', 'B'), { id: 'c', kind: 'append', selector: { key: 'u3' }, suffix: '!' }]);
  const out = {};
  for (const strategy of ['sequential', 'transaction', 'batch']) {
    const r = await runTask({ operations: ops(), strategy });
    out[strategy] = { state: snapshot(r.doc), commands: r.result.metrics.commands, reads: r.result.metrics.freshReads, statuses: r.exec.appliedOperations.map(e => e.status) };
  }
  assert.deepEqual(out.sequential.state, out.transaction.state);
  assert.deepEqual(out.sequential.state, out.batch.state);
  assert.deepEqual(out.sequential.statuses, out.batch.statuses);
  assert.ok(out.batch.commands < out.sequential.commands);
  assert.ok(out.batch.reads < out.sequential.reads);
});

test('S. retry after success -> every operation SATISFIED, zero write, zero dispatch, no barrier', async () => {
  const ops = [set('a', 'u1', 'A'), set('b', 'u2', 'B')];
  const first = await runTask({ operations: ops });
  const retry = await runTask({ units: first.doc.units, operations: ops });
  assert.equal(retry.result.ok, true);
  assert.equal(retry.acc.satisfied, 2);
  assert.equal(retry.result.metrics.commands, 0);
  assert.equal(retry.result.metrics.writes, 0);
  assert.equal(retry.driver.calls.forceSave, 0);
});

test('T. mixed SATISFIED / APPLY in one task', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'one'), set('b', 'u2', 'B'), set('c', 'u3', 'three')] });
  assert.deepEqual(r.exec.appliedOperations.map(e => e.status),
    [OPERATION_STATUS.SATISFIED, OPERATION_STATUS.APPLIED, OPERATION_STATUS.SATISFIED]);
  assert.equal(r.result.metrics.commands, 1);
  assert.equal(r.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
});

test('U. stale dependency target: a precomputed physical target is refused; the executor re-reads', async () => {
  const r = await runTask({ operations: [{ ...set('a', 'u1', 'A'), target: { index: 0 } }] });
  assert.equal(r.exec.outcome, 'plan-invalid');
  assert.equal(r.driver.calls.dispatch, 0);
  // And a dependent operation is always resolved from a read taken after its dependency.
  const chainRun = await runTask({ operations: chain(), strategy: 'sequential' });
  assert.equal(chainRun.adapter.calls.resolveTarget, 3);
  assert.equal(chainRun.result.metrics.freshReads, 4);
});

test('V. cross-session / non-local ref is refused before any dispatch', async () => {
  const foreign = { operationId: 'A', status: 'APPLIED', evidence: { key: 'u1' } }; // a result from elsewhere
  const r = await runTask({ operations: [{ id: 'b', kind: 'set', selector: { ref: foreign }, expectKey: 'u1', text: 'B' }] });
  assert.equal(r.exec.outcome, 'plan-invalid');
  assert.equal(r.driver.calls.dispatch, 0);
  // A ref to an operation that is not an earlier operation of THIS task is refused too.
  const dangling = await runTask({ operations: [{ id: 'b', kind: 'set', selector: { ref: 'somewhere-else' }, expectKey: 'u1', text: 'B' }] });
  assert.equal(dangling.exec.outcome, 'plan-invalid');
  // NEGATIVE CONTROL: a local ref to an earlier operation of this task is accepted.
  const ok = await runTask({ operations: chain() });
  assert.equal(ok.result.ok, true);
});

test('W. a whole-task verifier that dispatches (tries to repair) voids its own verdict -> FAIL', async () => {
  let t = 0;
  const doc = createDocument(base());
  const driver = createDocumentDriver(doc);
  const adapter = createDocumentAdapter();
  let kept = null;
  const origExec = adapter.executePrimitive;
  adapter.executePrimitive = async (h, op, ctx) => { kept = ctx.dispatch; return origExec(h, op, ctx); };
  const origWhole = adapter.verifyWholeTask;
  adapter.verifyWholeTask = async (h, task, obs) => {
    await kept({ kind: 'set', index: 1, text: 'repaired' }); // a write smuggled into a read-only layer
    return origWhole(h, task, obs);
  };
  const options = { driver, fileId: 'doc', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const r = await withSession(options, adapter, createExecutor({ id: 'w', operations: [set('a', 'u1', 'A')], expect: [] }));
  assert.equal(r.execution.outcome, 'whole-task-verifier-mutated');
  assert.equal(r.ok, false);
  // NEGATIVE CONTROL: the untouched verifier on the same task passes.
  const clean = await runTask({ operations: [set('a', 'u1', 'A')] });
  assert.equal(clean.result.ok, true);
});

// ---------------------------------------------------------------------------------------------
// R1 -- UNDECLARED DEPENDENCY MUST NOT MEAN INDEPENDENT (owner review)
// ---------------------------------------------------------------------------------------------

test('R1-M1. unknown independence: three operations, no dependency, no proof -> NOT one physical batch', async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'A'), set('b', 'u2', 'B'), set('c', 'u3', 'C')], strategy: 'batch' });
  assert.equal(r.result.ok, true);
  assert.equal(r.exec.batchEvidence.groups, 3);
  assert.equal(r.result.metrics.commands, 3, 'singleton execution, one physical dispatch each');
  assert.equal(r.adapter.calls.executePrimitiveBatch, 0, 'the batch hook is never used without a proof');
  assert.deepEqual(snapshot(r.doc), ['u1=A', 'u2=B', 'u3=C']);
});

test('R1-M2. explicit proven independence -> one batch is allowed', async () => {
  const r = await runTask({ operations: proven([set('a', 'u1', 'A'), set('b', 'u2', 'B'), set('c', 'u3', 'C')]), strategy: 'batch' });
  assert.equal(r.exec.batchEvidence.groups, 1);
  assert.equal(r.result.metrics.commands, 1);
  assert.equal(r.adapter.calls.executePrimitiveBatch, 1);
  // NEGATIVE CONTROL: DIFFERENT proof keys are not one group.
  const split = await runTask({ operations: [set('a', 'u1', 'A', { batchSafeGroup: 'g1' }), set('b', 'u2', 'B', { batchSafeGroup: 'g2' })], strategy: 'batch' });
  assert.equal(split.exec.batchEvidence.groups, 2);
});

test('R1-M3. explicit batch-safe + dependency conflict -> the dependency wins: [A] | [B]', async () => {
  const ops = [set('a', 'u1', 'A', { batchSafeGroup: 'g' }), { id: 'b', kind: 'append', selector: { key: 'u1' }, suffix: '!', dependsOn: ['a'], batchSafeGroup: 'g' }];
  const r = await runTask({ operations: ops, strategy: 'batch', expect: [{ key: 'u1', text: 'A!' }] });
  assert.equal(r.exec.batchEvidence.groups, 2);
  assert.equal(r.result.metrics.dependencyBarriers, 1);
  assert.equal(r.result.ok, true);
  // A ref is a dependency too, and beats the proof the same way.
  const refOps = [{ id: 'A', kind: 'insertAfter', selector: { key: 'u1' }, newKey: 'X', text: 'x', batchSafeGroup: 'g' },
    { id: 'B', kind: 'append', selector: { ref: 'A' }, expectKey: 'X', suffix: '+b', batchSafeGroup: 'g' }];
  const rr = await runTask({ operations: refOps, strategy: 'batch' });
  assert.equal(rr.exec.batchEvidence.groups, 2);
});

test('R1-M4. the old M-NEG scenario (no declaration, no proof) -> RIGHT document', async () => {
  const r = await runTask({ operations: shifting({ declareDependency: false }), strategy: 'batch' });
  assert.deepEqual(snapshot(r.doc), RIGHT);
  // NEGATIVE CONTROL: the damage is real when a planner FALSELY proves independence -- which is
  // why the proof is the planner's responsibility and the core never manufactures one.
  const lying = await runTask({ operations: shifting({ declareDependency: false, proof: true }), strategy: 'batch' });
  assert.equal(lying.exec.batchEvidence.groups, 1);
  assert.notDeepEqual(snapshot(lying.doc), RIGHT);
  assert.notEqual(status(lying.result, 'B'), OPERATION_STATUS.APPLIED, 'L1 still refuses to call it applied');
});

test('R1-HOOK. a batch hook that dispatches more than once breaks the contract -> fail closed', async () => {
  const ops = proven([set('a', 'u1', 'A'), set('b', 'u2', 'B'), set('c', 'u3', 'C')]);
  let t = 0;
  const doc = createDocument(base());
  const driver = createDocumentDriver(doc);
  const adapter = createDocumentAdapter();
  // A faulty hook: one dispatch per operation instead of one per round.
  adapter.executePrimitiveBatch = async (_h, items, { dispatch }) => {
    let last = null;
    for (const it of items) last = await dispatch({ kind: 'set', index: it.target.index, text: it.operation.text });
    return last;
  };
  const options = { driver, fileId: 'doc', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const r = await withSession(options, adapter, createExecutor({ id: 'h', operations: ops }, { strategy: 'batch' }));
  assert.equal(r.execution.outcome, 'batch-hook-contract-violation');
  assert.equal(r.ok, false);
  assert.equal(driver.calls.dispatch, 1, 'the second call is refused BEFORE it reaches the editor');
  assert.equal(r.metrics.commands, 1, 'commands is the true physical count');
  assert.deepEqual(r.execution.appliedOperations.map(e => e.status), Array(3).fill(OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN));
  assert.equal(r.execution.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.BATCH_HOOK_CONTRACT_VIOLATION);
  assert.equal(r.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);

  // A hook that reports success WITHOUT dispatching is a violation too -- and certain: nothing reached the editor.
  const silent = createDocumentAdapter();
  silent.executePrimitiveBatch = async () => ({ ok: true, dispatched: true });
  const d2 = createDocumentDriver(createDocument(base()));
  const r2 = await withSession({ ...options, driver: d2 }, silent, createExecutor({ id: 'h2', operations: ops }, { strategy: 'batch' }));
  assert.equal(r2.execution.outcome, 'batch-hook-contract-violation');
  assert.equal(d2.calls.dispatch, 0);
  assert.deepEqual(r2.execution.appliedOperations.map(e => e.status), Array(3).fill(OPERATION_STATUS.FAILED));
  assert.equal(r2.mutationCertainty, 'CERTAIN');

  // NEGATIVE CONTROL: the contract-conforming hook on the same task is one dispatch and PASS.
  const ok = await runTask({ operations: ops, strategy: 'batch' });
  assert.equal(ok.result.ok, true);
  assert.equal(ok.driver.calls.dispatch, 1);
});

// #ref/#ref: the executor's own 'plan-failed' outcome (adapter.planOperations() itself
// returning { ok: false, detail }, NOT the plan()-succeeded-but-invalid case covered by test Q's
// 'plan-invalid' assertion) previously dropped `detail` -- `finish()` was called without it,
// unlike the 'plan-threw'/'plan-invalid' branches, which already thread their own detail through.
test("plan-failed threads the adapter's own detail into stopDetail, the same way plan-invalid/plan-threw already do", async () => {
  const r = await runTask({ operations: [set('a', 'u1', 'ONE')], adapterScript: { planFailDetail: 'scripted-plan-detail-57762' } });
  assert.equal(r.exec.outcome, 'plan-failed');
  assert.equal(r.driver.calls.dispatch, 0);
  assert.equal(r.exec.stopDetail, 'scripted-plan-detail-57762');
});
