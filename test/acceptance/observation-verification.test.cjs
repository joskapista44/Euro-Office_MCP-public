'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { observeFresh, requireMutationAuthority } = require('../../src/core/observation/observe.cjs');
const { POLICY, ENRICHMENT_STATE } = require('../../src/core/observation/enrichment.cjs');
const { FRESHNESS } = require('../../src/core/observation/freshness.cjs');
const { fromDispatch, verifyPrimitive, verifyWholeTask, verifyPreservation } =
  require('../../src/core/verification/orchestrate.cjs');
const { VERDICT, LAYER } = require('../../src/core/verification/layers.cjs');
const { AUTHORITY } = require('../../src/core/authority.cjs');
const { dispatch } = require('../../src/core/session/handle.cjs');
const { SessionHandle } = require('../../src/core/session/handle.cjs');
const { bindDriver } = require('../../src/core/session/internal.cjs');
const { Metrics } = require('../../src/core/metrics/collect.cjs');
const { createFakeSession } = require('../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../fakes/fake-adapter.cjs');

/**
 * M2 acceptance matrix A..S — owner requirement 21.
 * Every relevant case asserts an EXPLICIT authority, not just an outcome.
 *
 * A..Q live here. R (metric call-count consistency) and S (whole-task failure persisting through
 * the M1 lifecycle) need the whole lifecycle, so they live in metrics-and-lifecycle.test.cjs.
 * Q2/Q3 are the verification-freshness cases of owner requirement 13.
 */

function session({ sessionScript = {}, adapterScript = {}, id = 'sess-1' } = {}) {
  let t = 0;
  const metrics = new Metrics({ now: () => (t += 1) });
  const handle = new SessionHandle({ id, fileId: 'f', adapterKind: 'fake', metrics });
  const driver = createFakeSession(sessionScript);
  bindDriver(handle, driver);
  return { handle, driver, adapter: createFakeAdapter(adapterScript), metrics };
}

const op = { id: 'op1' };
const task = { id: 'task-1', operations: [op] };

test('A. fresh LIVE observation without enrichment — authority LIVE_READ', async () => {
  const { handle, adapter, driver } = session({ adapterScript: { enrichment: 'absent' } });
  const obs = await observeFresh(handle, adapter);

  assert.equal(obs.ok, true);
  assert.equal(obs.authority, AUTHORITY.LIVE_READ, 'an observation may never claim more');
  assert.equal(obs.evidence.live.ok, true);
  assert.equal(obs.evidence.enrichment.state, ENRICHMENT_STATE.NOT_IMPLEMENTED);
  assert.equal(obs.enrichment, null);
  assert.equal(driver.calls.readLive, 1);
  assert.equal(handle.metrics.freshReads, 1);
  assert.equal(handle.metrics.enrichmentInvocations, 0, 'no hook, no package read');

  assert.equal(obs.freshness.sessionId, 'sess-1');
  assert.equal(obs.freshness.generation, 0);
  assert.equal(typeof obs.freshness.sequence, 'number');
});

test('B. fresh LIVE + optional enrichment success — enrichment is additive, not authority', async () => {
  const { handle, adapter } = session({ adapterScript: { enrichment: 'ok' } });
  const obs = await observeFresh(handle, adapter, { enrichment: POLICY.OPTIONAL });

  assert.equal(obs.ok, true);
  assert.equal(obs.authority, AUTHORITY.LIVE_READ, 'enrichment does not raise authority');
  assert.equal(obs.evidence.enrichment.state, ENRICHMENT_STATE.APPLIED);
  assert.equal(obs.evidence.enrichment.isLiveEvidence, false, 'marked so a verifier cannot mistake it');
  assert.equal(obs.evidence.enrichment.authority, AUTHORITY.LIVE_READ);
  assert.deepEqual(obs.enrichment, { extra: 'fake-package-semantics' });
  assert.equal(handle.metrics.enrichmentInvocations, 1);
});

test('C. required enrichment success', async () => {
  const { handle, adapter } = session({ adapterScript: { enrichment: 'ok' } });
  const obs = await observeFresh(handle, adapter, { enrichment: POLICY.REQUIRED });

  assert.equal(obs.ok, true);
  assert.equal(obs.evidence.enrichment.state, ENRICHMENT_STATE.APPLIED);
  assert.equal(obs.evidence.enrichment.policy, POLICY.REQUIRED);
  assert.equal(obs.authority, AUTHORITY.LIVE_READ);
});

test('D. required enrichment failure — FAIL CLOSED, authority drops to PLAN_ONLY', async () => {
  for (const mode of ['fails', 'throws', 'absent']) {
    const { handle, adapter } = session({ adapterScript: { enrichment: mode } });
    const obs = await observeFresh(handle, adapter, { enrichment: POLICY.REQUIRED });

    assert.equal(obs.ok, false, `mode ${mode}`);
    assert.equal(obs.outcome, 'observation-required-enrichment-failed', `mode ${mode}`);
    assert.equal(obs.authority, AUTHORITY.PLAN_ONLY, `mode ${mode}: must not proceed to a write`);
    assert.equal(obs.evidence.enrichment.state, ENRICHMENT_STATE.FAILED, `mode ${mode}`);
  }
});

test('E. optional enrichment failure — continues, but the degradation is VISIBLE', async () => {
  for (const mode of ['fails', 'throws']) {
    const { handle, adapter } = session({ adapterScript: { enrichment: mode } });
    const obs = await observeFresh(handle, adapter, { enrichment: POLICY.OPTIONAL });

    assert.equal(obs.ok, true, `mode ${mode}: OPTIONAL does not stop the task`);
    assert.equal(obs.authority, AUTHORITY.LIVE_READ, `mode ${mode}: the LIVE part still stands`);
    assert.equal(obs.evidence.enrichment.state, ENRICHMENT_STATE.DEGRADED, `mode ${mode}`);
    assert.notEqual(
      obs.evidence.enrichment.state,
      ENRICHMENT_STATE.NOT_IMPLEMENTED,
      `mode ${mode}: a failure must be distinguishable from an adapter that had nothing to add`
    );
    assert.equal(obs.enrichment, null, `mode ${mode}: no enrichment data was produced`);
  }
});

test('F. package-only authority rejected — enrichment can never back a write', async () => {
  const { handle, adapter } = session({ adapterScript: { enrichment: 'ok' } });
  const real = await observeFresh(handle, adapter);

  // Construct the forbidden shape: enrichment present, live evidence absent.
  const packageOnly = Object.freeze({
    ...real,
    evidence: Object.freeze({ ...real.evidence, live: Object.freeze({ ok: false, detail: 'no live read' }) }),
  });

  const gate = requireMutationAuthority(packageOnly, handle);
  assert.equal(gate.ok, false);
  assert.equal(gate.outcome, 'observation-package-only');
  assert.match(gate.reason, /no live evidence/);

  // Positive control: the real observation passes the same gate.
  assert.equal(requireMutationAuthority(real, handle).ok, true);
});

test('G. stale pre-mutation observation rejected — a dispatch invalidates it', async () => {
  const { handle, adapter } = session();
  const obs = await observeFresh(handle, adapter);
  assert.equal(requireMutationAuthority(obs, handle).ok, true, 'fresh before the mutation');

  await dispatch(handle, { id: 'op1', kind: 'fake-write' });

  const gate = requireMutationAuthority(obs, handle);
  assert.equal(gate.ok, false);
  assert.equal(gate.verdict.state, FRESHNESS.STALE);
  assert.equal(gate.outcome, 'observation-stale');
  assert.equal(gate.verdict.mutationsSince, 1);
  assert.equal(gate.verdict.observedAtGeneration, 0);
  assert.equal(gate.verdict.currentGeneration, 1);

  // And a fresh re-read is accepted again — staleness is not a dead end, it is a re-read.
  const reread = await observeFresh(handle, adapter);
  assert.equal(requireMutationAuthority(reread, handle).ok, true);
});

test('H. cross-session observation rejected', async () => {
  const a = session({ id: 'sess-A' });
  const b = session({ id: 'sess-B' });
  const obsA = await observeFresh(a.handle, a.adapter);

  const gate = requireMutationAuthority(obsA, b.handle);
  assert.equal(gate.ok, false);
  assert.equal(gate.verdict.state, FRESHNESS.CROSS_SESSION);
  assert.equal(gate.verdict.observedIn, 'sess-A');
  assert.equal(gate.verdict.usedIn, 'sess-B');

  // Both sessions are at generation 0, so a generation-only check would have PASSED this.
  assert.equal(obsA.freshness.generation, 0);
  assert.equal(requireMutationAuthority(obsA, a.handle).ok, true, 'it is fresh in its own session');
});

test('I. dispatch-only cannot become LIVE_VERIFY', async () => {
  const l0 = fromDispatch({ ok: true, dispatched: true });

  assert.equal(l0.layer, LAYER.DISPATCH_COMPLETE);
  assert.equal(l0.verdict, VERDICT.PASS, 'the command did return');
  assert.equal(l0.authority, AUTHORITY.DISPATCH_ONLY);
  assert.notEqual(l0.authority, AUTHORITY.LIVE_VERIFY);
  assert.equal(l0.proves, 'nothing about document state');

  // There is no code path from a dispatch result into an L1 verdict: the only producer of
  // LIVE_VERIFY is verifyPrimitive, and it requires a fresh observation.
  assert.equal(typeof fromDispatch({ ok: true }).evidence.dispatched, 'boolean');
});

test('J. primitive verification against a FRESH observation — LIVE_VERIFY', async () => {
  const { handle, adapter } = session();
  await dispatch(handle, { id: 'op1' });
  const after = await observeFresh(handle, adapter);

  const l1 = await verifyPrimitive({ handle, adapter, operation: op, after });
  assert.equal(l1.verdict, VERDICT.PASS);
  assert.equal(l1.authority, AUTHORITY.LIVE_VERIFY);
  assert.equal(l1.layer, LAYER.PRIMITIVE);
  assert.equal(handle.metrics.verificationInvocations, 1);
});

test('K. primitive verification FAIL — claims the floor, not LIVE_VERIFY', async () => {
  const { handle, adapter } = session({ adapterScript: { primitives: { op1: 'verify-fail' } } });
  await dispatch(handle, { id: 'op1' });
  const after = await observeFresh(handle, adapter);

  const l1 = await verifyPrimitive({ handle, adapter, operation: op, after });
  assert.equal(l1.verdict, VERDICT.FAIL);
  assert.equal(l1.authority, AUTHORITY.PLAN_ONLY, 'a failed verifier proved nothing');
  assert.notEqual(l1.authority, AUTHORITY.LIVE_VERIFY);
});

test('L. primitives PASS, whole-task FAIL — the layers do not substitute', async () => {
  const { handle, adapter } = session({ adapterScript: { wholeTaskFails: true } });
  await dispatch(handle, { id: 'op1' });
  const after = await observeFresh(handle, adapter);

  const l1 = await verifyPrimitive({ handle, adapter, operation: op, after });
  assert.equal(l1.verdict, VERDICT.PASS, 'every primitive is green');
  assert.equal(l1.authority, AUTHORITY.LIVE_VERIFY);

  const l2 = await verifyWholeTask({ handle, adapter, task, finalObservation: after });
  assert.equal(l2.verdict, VERDICT.FAIL, 'and the task is still wrong');
  assert.equal(l2.failedCount, 1);
  assert.equal(l2.checkCount, 1);
  assert.equal(l2.authority, AUTHORITY.PLAN_ONLY);
});

test('M. preservation required PASS', async () => {
  const { handle, adapter } = session({ adapterScript: { preservation: 'ok' } });
  await dispatch(handle, { id: 'op1' });
  const after = await observeFresh(handle, adapter);

  const l3 = await verifyPreservation({ handle, adapter, operation: op, after, required: true });
  assert.equal(l3.verdict, VERDICT.PASS);
  assert.equal(l3.authority, AUTHORITY.LIVE_READ);
  assert.equal(l3.layer, LAYER.PRESERVATION);
});

test('N. preservation required FAIL — and a missing verifier is also a FAIL when required', async () => {
  const failing = session({ adapterScript: { preservation: 'fails' } });
  await dispatch(failing.handle, { id: 'op1' });
  const afterF = await observeFresh(failing.handle, failing.adapter);
  const l3 = await verifyPreservation({ handle: failing.handle, adapter: failing.adapter, operation: op, after: afterF, required: true });
  assert.equal(l3.verdict, VERDICT.FAIL);
  assert.equal(l3.authority, AUTHORITY.PLAN_ONLY);

  const absent = session({ adapterScript: { preservation: 'absent' } });
  const afterA = await observeFresh(absent.handle, absent.adapter);
  const missing = await verifyPreservation({ handle: absent.handle, adapter: absent.adapter, operation: op, after: afterA, required: true });
  assert.equal(missing.verdict, VERDICT.FAIL, 'required but undeclared is a failure, not NOT_REQUIRED');
  assert.match(missing.evidence.reason, /requires one/);
});

test('O. preservation not required — NOT_REQUIRED, and that is NOT a pass', async () => {
  const { handle, adapter } = session({ adapterScript: { preservation: 'absent' } });
  const after = await observeFresh(handle, adapter);

  const l3 = await verifyPreservation({ handle, adapter, operation: op, after, required: false });
  assert.equal(l3.verdict, VERDICT.NOT_REQUIRED);
  assert.notEqual(l3.verdict, VERDICT.PASS, 'having no requirement is not evidence of preservation');

  const { isSatisfied, isProven } = require('../../src/core/verification/layers.cjs');
  assert.equal(isSatisfied(l3.verdict), true, 'it does not block the task');
  assert.equal(isProven(l3.verdict), false, 'but nothing was proven');
});

test('P. live read boundary failure — normalized, PLAN_ONLY, no raw escape', async () => {
  for (const mode of [{ readLiveThrows: true }, { readLiveFails: true }]) {
    const { handle, adapter } = session({ sessionScript: mode });
    const obs = await observeFresh(handle, adapter);

    assert.equal(obs.ok, false, JSON.stringify(mode));
    assert.equal(obs.authority, AUTHORITY.PLAN_ONLY, JSON.stringify(mode));
    assert.match(obs.outcome, /^observation-live-read-(threw|failed)$/);
    assert.equal(obs.evidence.live.ok, false);
  }

  // Normalization is its own boundary, named separately.
  const norm = session({ adapterScript: { normalizeThrows: true } });
  const obsN = await observeFresh(norm.handle, norm.adapter);
  assert.equal(obsN.outcome, 'observation-normalization-threw');
  assert.equal(obsN.evidence.live.ok, true, 'the live read succeeded; normalization did not');
  assert.equal(obsN.evidence.normalization.threw, true);
});

test('Q. verifier boundary failure — BOUNDARY_FAILURE, not a FAIL verdict', async () => {
  const cases = [
    ['primitive', { primitives: { op1: 'verify-throws' } }, (h, a, o) => verifyPrimitive({ handle: h, adapter: a, operation: op, after: o })],
    ['whole-task', { wholeTaskThrows: true }, (h, a, o) => verifyWholeTask({ handle: h, adapter: a, task, finalObservation: o })],
    ['preservation', { preservation: 'throws' }, (h, a, o) => verifyPreservation({ handle: h, adapter: a, operation: op, after: o, required: true })],
  ];

  for (const [name, script, run] of cases) {
    const { handle, adapter } = session({ adapterScript: script });
    await dispatch(handle, { id: 'op1' });
    const after = await observeFresh(handle, adapter);
    const v = await run(handle, adapter, after);

    assert.equal(v.verdict, VERDICT.BOUNDARY_FAILURE, `${name}: "we could not look" is not "the document is wrong"`);
    assert.notEqual(v.verdict, VERDICT.FAIL, name);
    assert.equal(v.authority, AUTHORITY.PLAN_ONLY, name);
    assert.equal(v.evidence.threw, true, name);
    assert.ok(v.evidence.detail, `${name}: the detail survives as evidence`);
  }
});

test('Q2. verification against a STALE observation is refused', async () => {
  const { handle, adapter } = session();
  const planningTime = await observeFresh(handle, adapter);
  await dispatch(handle, { id: 'op1' });

  // Owner requirement 13: verifying against planning-time state proves nothing about what exists.
  const l1 = await verifyPrimitive({ handle, adapter, operation: op, after: planningTime });
  assert.equal(l1.verdict, VERDICT.STALE_STATE);
  assert.equal(l1.authority, AUTHORITY.PLAN_ONLY);
  assert.equal(l1.evidence.outcome, 'observation-stale');
  assert.equal(adapter.calls.verifyPrimitive, 0, 'the hook is never even called on stale state');

  const l2 = await verifyWholeTask({ handle, adapter, task, finalObservation: planningTime });
  assert.equal(l2.verdict, VERDICT.STALE_STATE);

  // Positive control: a fresh re-read verifies normally.
  const fresh = await observeFresh(handle, adapter);
  const ok = await verifyPrimitive({ handle, adapter, operation: op, after: fresh });
  assert.equal(ok.verdict, VERDICT.PASS);
  assert.equal(adapter.calls.verifyPrimitive, 1);
});

test('Q3. cross-session state is refused by the verifiers too', async () => {
  const a = session({ id: 'sess-A' });
  const b = session({ id: 'sess-B' });
  const obsA = await observeFresh(a.handle, a.adapter);

  const l1 = await verifyPrimitive({ handle: b.handle, adapter: b.adapter, operation: op, after: obsA });
  assert.equal(l1.verdict, VERDICT.STALE_STATE);
  assert.equal(l1.evidence.outcome, 'observation-cross-session');
  assert.equal(b.adapter.calls.verifyPrimitive, 0);
});
