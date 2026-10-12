'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { observeFresh } = require('../../src/core/observation/observe.cjs');
const { POLICY, ENRICHMENT_STATE, runEnrichment, isFailClosed } =
  require('../../src/core/observation/enrichment.cjs');
const { verifyPrimitive, verifyWholeTask, verifyPreservation } =
  require('../../src/core/verification/orchestrate.cjs');
const { VERDICT, provesNothing } = require('../../src/core/verification/layers.cjs');
const { AUTHORITY } = require('../../src/core/authority.cjs');
const { SessionHandle, dispatch, readLive } = require('../../src/core/session/handle.cjs');
const { bindDriver } = require('../../src/core/session/internal.cjs');
const { Metrics, collectMetrics } = require('../../src/core/metrics/collect.cjs');
const { createFakeSession } = require('../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../fakes/fake-adapter.cjs');

/**
 * T, U, V — the M2 final-review findings.
 *
 * All three were fail-open, and all three were reproduced before being fixed:
 *
 *   T  a verifier returning null/undefined yielded PASS and the layer's full authority, so a
 *      primitive verifier could report LIVE_VERIFY having proved nothing
 *   U  REQUIRED enrichment whose hook RAN and returned null was reported NOT_IMPLEMENTED, so a
 *      capability that declared it needs the package proceeded without it
 *   V  editorInteractions included verifier invocations, double-counting a verifier that reads
 *      through readLive(): three physical reads were reported as five interactions
 */

function session({ sessionScript = {}, adapterScript = {}, id = 'rev-1' } = {}) {
  let t = 0;
  const metrics = new Metrics({ now: () => (t += 1) });
  const handle = new SessionHandle({ id, fileId: 'f', adapterKind: 'fake', metrics });
  const driver = createFakeSession(sessionScript);
  bindDriver(handle, driver);
  return { handle, driver, adapter: createFakeAdapter(adapterScript), metrics };
}

const op = { id: 'op1' };
const task = { id: 't', operations: [op] };

// ---------------------------------------------------------------------------------------------
// T — absence of verification evidence is not verification
// ---------------------------------------------------------------------------------------------

test('T1. a primitive verifier returning null NEVER yields LIVE_VERIFY', async () => {
  for (const nothing of [null, undefined]) {
    const { handle, adapter } = session();
    adapter.verifyPrimitive = async () => nothing;
    await dispatch(handle, op);
    const after = await observeFresh(handle, adapter);

    const l1 = await verifyPrimitive({ handle, adapter, operation: op, after });
    assert.equal(l1.verdict, VERDICT.NO_EVIDENCE, `returned ${String(nothing)}`);
    assert.notEqual(l1.verdict, VERDICT.PASS, 'absence of evidence is not a pass');
    assert.equal(l1.authority, AUTHORITY.PLAN_ONLY);
    assert.notEqual(l1.authority, AUTHORITY.LIVE_VERIFY, 'the fail-open case: no evidence, full authority');
    assert.equal(l1.evidence.hookReturnedNull, true, 'and the reason is explicit');
    assert.equal(provesNothing(l1.verdict), true);
  }
});

test('T2. a whole-task verifier returning null cannot be a PASS', async () => {
  const { handle, adapter } = session();
  adapter.verifyWholeTask = async () => null;
  await dispatch(handle, op);
  const after = await observeFresh(handle, adapter);

  const l2 = await verifyWholeTask({ handle, adapter, task, finalObservation: after });
  assert.equal(l2.verdict, VERDICT.NO_EVIDENCE);
  assert.notEqual(l2.verdict, VERDICT.PASS);
  assert.equal(l2.authority, AUTHORITY.PLAN_ONLY);
});

test('T3. preservation null is harmonised with the general model, not special-cased into a PASS', async () => {
  // required -> FAIL: a promised check that produced no evidence has failed.
  const req = session({ adapterScript: { preservation: 'ok' } });
  req.adapter.verifyPreservation = async () => null;
  await dispatch(req.handle, op);
  const afterR = await observeFresh(req.handle, req.adapter);
  const l3r = await verifyPreservation({ handle: req.handle, adapter: req.adapter, operation: op, after: afterR, required: true });
  assert.equal(l3r.verdict, VERDICT.FAIL);
  assert.equal(l3r.authority, AUTHORITY.PLAN_ONLY);

  // not required -> NOT_REQUIRED: nothing to prove, and nothing proven.
  const opt = session({ adapterScript: { preservation: 'ok' } });
  opt.adapter.verifyPreservation = async () => null;
  await dispatch(opt.handle, op);
  const afterO = await observeFresh(opt.handle, opt.adapter);
  const l3o = await verifyPreservation({ handle: opt.handle, adapter: opt.adapter, operation: op, after: afterO, required: false });
  assert.equal(l3o.verdict, VERDICT.NOT_REQUIRED);
  assert.notEqual(l3o.verdict, VERDICT.PASS);

  // Neither path ever produced a PASS that a later layer had to take back.
  for (const v of [l3r.verdict, l3o.verdict]) assert.notEqual(v, VERDICT.PASS);
});

// ---------------------------------------------------------------------------------------------
// U — REQUIRED enrichment and a null return
// ---------------------------------------------------------------------------------------------

test('U1. a hook that EXISTS and returns null is not the same case as a hook that is ABSENT', async () => {
  const h = () => ({ id: 'h', metrics: new Metrics({ now: () => 1 }) });

  const returnsNull = await runEnrichment(h(), { kind: 'fake', enrichObservation: async () => null }, {}, { policy: POLICY.REQUIRED });
  const absent = await runEnrichment(h(), { kind: 'fake', enrichObservation: null }, {}, { policy: POLICY.REQUIRED });

  // Both FAIL under REQUIRED...
  assert.equal(returnsNull.state, ENRICHMENT_STATE.FAILED);
  assert.equal(absent.state, ENRICHMENT_STATE.FAILED);
  assert.equal(isFailClosed(returnsNull), true);
  assert.equal(isFailClosed(absent), true);

  // ...but the reasons stay distinct, because the fix is different: implement the hook, versus
  // find out why the hook produced nothing.
  assert.equal(returnsNull.hookReturnedNull, true);
  assert.notEqual(returnsNull.reason, absent.reason);
  assert.match(returnsNull.reason, /ran and returned no enrichment/);
  assert.match(absent.reason, /does not implement/);
});

test('U2. the full policy matrix — every REQUIRED failure mode is fail-closed', async () => {
  const h = () => ({ id: 'h', metrics: new Metrics({ now: () => 1 }) });
  const cases = [
    [POLICY.NONE, async () => ({ ok: true, enrichment: {} }), ENRICHMENT_STATE.SKIPPED, false],
    [POLICY.OPTIONAL, async () => null, ENRICHMENT_STATE.NOT_IMPLEMENTED, false],
    [POLICY.OPTIONAL, null, ENRICHMENT_STATE.NOT_IMPLEMENTED, false],
    [POLICY.OPTIONAL, async () => ({ ok: false }), ENRICHMENT_STATE.DEGRADED, false],
    [POLICY.OPTIONAL, async () => { throw new Error('x'); }, ENRICHMENT_STATE.DEGRADED, false],
    [POLICY.REQUIRED, null, ENRICHMENT_STATE.FAILED, true],
    [POLICY.REQUIRED, async () => null, ENRICHMENT_STATE.FAILED, true],
    [POLICY.REQUIRED, async () => undefined, ENRICHMENT_STATE.FAILED, true],
    [POLICY.REQUIRED, async () => ({ ok: false }), ENRICHMENT_STATE.FAILED, true],
    [POLICY.REQUIRED, async () => { throw new Error('x'); }, ENRICHMENT_STATE.FAILED, true],
    [POLICY.REQUIRED, async () => ({ ok: true, enrichment: { a: 1 } }), ENRICHMENT_STATE.APPLIED, false],
  ];

  for (const [policy, hook, expected, failClosed] of cases) {
    const e = await runEnrichment(h(), { kind: 'fake', enrichObservation: hook }, {}, { policy });
    assert.equal(e.state, expected, `${policy} + ${hook ? 'hook' : 'absent'}`);
    assert.equal(isFailClosed(e), failClosed, `${policy} + ${hook ? 'hook' : 'absent'}`);
  }
});

test('U3. observeFresh fails closed to PLAN_ONLY when REQUIRED enrichment returns null', async () => {
  const { handle, adapter } = session({ adapterScript: { enrichment: 'null' } });
  const obs = await observeFresh(handle, adapter, { enrichment: POLICY.REQUIRED });

  assert.equal(obs.ok, false);
  assert.equal(obs.outcome, 'observation-required-enrichment-failed');
  assert.equal(obs.authority, AUTHORITY.PLAN_ONLY);
  assert.equal(obs.evidence.enrichment.state, ENRICHMENT_STATE.FAILED);
  assert.equal(obs.evidence.enrichment.hookReturnedNull, true);

  // Positive control: the same adapter under OPTIONAL proceeds.
  const opt = await observeFresh(handle, adapter, { enrichment: POLICY.OPTIONAL });
  assert.equal(opt.ok, true);
  assert.equal(opt.authority, AUTHORITY.LIVE_READ);
});

// ---------------------------------------------------------------------------------------------
// V — no interaction double-counting
// ---------------------------------------------------------------------------------------------

/** A verifier that performs a REAL physical read inside its hook — the case that double-counted. */
function readingVerifierAdapter(reads = 1) {
  const adapter = createFakeAdapter({});
  adapter.verifyPrimitive = async (handle, operation) => {
    for (let i = 0; i < reads; i++) await readLive(handle, { purpose: 'verify' });
    return { ok: true, evidence: { unit: operation.id, readsPerformed: reads } };
  };
  return adapter;
}

test('V1. a verifier that reads through readLive is counted ONCE physically and ONCE semantically', async () => {
  const { handle, driver, metrics } = session();
  const adapter = readingVerifierAdapter(1);

  await dispatch(handle, op);
  const after = await observeFresh(handle, adapter);          // 1 physical read
  const readsBeforeVerify = driver.calls.readLive;

  const l1 = await verifyPrimitive({ handle, adapter, operation: op, after }); // 1 more physical read
  assert.equal(l1.verdict, VERDICT.PASS);
  assert.equal(l1.authority, AUTHORITY.LIVE_VERIFY);

  const m = collectMetrics(metrics);

  // The physical read inside the hook IS counted, exactly once, by the handle.
  assert.equal(driver.calls.readLive, readsBeforeVerify + 1, 'the hook really did read');
  assert.equal(m.freshReads, driver.calls.readLive, 'one physical counter, one increment site');

  // The invocation is counted exactly once, separately.
  assert.equal(m.verificationInvocations, 1);

  // And the physical total does NOT include it.
  assert.equal(m.editorInteractions, m.freshReads + m.commands + m.persistencePolls);
  assert.equal(m.totalExternalInteractions, m.editorInteractions + m.packageInteractions);

  // THE REGRESSION THIS PINS: the old formula would have counted the hook's read twice.
  const oldFormula = m.freshReads + m.verificationInvocations + m.commands + m.persistencePolls;
  assert.ok(oldFormula > m.editorInteractions, 'the old formula inflated the total');
  assert.equal(oldFormula - m.editorInteractions, m.verificationInvocations, 'by exactly the invocations');
});

test('V2. the identity holds for a verifier performing SEVERAL reads, and for one performing none', async () => {
  // Several reads in one invocation: physical count rises, invocation count does not.
  const many = session();
  const manyAdapter = readingVerifierAdapter(3);
  await dispatch(many.handle, op);
  const afterM = await observeFresh(many.handle, manyAdapter);
  await verifyPrimitive({ handle: many.handle, adapter: manyAdapter, operation: op, after: afterM });
  const mm = collectMetrics(many.metrics);

  assert.equal(mm.verificationInvocations, 1, 'one invocation, whatever it did inside');
  assert.equal(mm.freshReads, many.driver.calls.readLive, 'and every physical read is accounted for');
  assert.equal(mm.freshReads, 4, 'one observation plus three reads inside the hook');
  assert.equal(mm.editorInteractions, mm.freshReads + mm.commands + mm.persistencePolls);

  // NEGATIVE CONTROL: a verifier that performs NO read. The invocation is still counted, and the
  // physical total is unchanged by it — so the separation is not an artefact of reading hooks.
  const none = session();
  const noneAdapter = createFakeAdapter({});
  await dispatch(none.handle, op);
  const afterN = await observeFresh(none.handle, noneAdapter);
  const physicalBefore = none.driver.calls.readLive;
  await verifyPrimitive({ handle: none.handle, adapter: noneAdapter, operation: op, after: afterN });
  const mn = collectMetrics(none.metrics);

  assert.equal(none.driver.calls.readLive, physicalBefore, 'no physical read happened');
  assert.equal(mn.verificationInvocations, 1, 'but the invocation is still counted');
  assert.equal(mn.freshReads, physicalBefore);
  assert.equal(mn.editorInteractions, mn.freshReads + mn.commands + mn.persistencePolls);
});

test('V3. the retired names are gone, not redefined', () => {
  const m = new Metrics({ now: () => 1 });
  const s = collectMetrics(m);

  assert.equal(s.totalInteractions, undefined, 'a consumer of the old name gets undefined, loudly');
  assert.equal(s.verificationReads, undefined);
  assert.equal(typeof s.totalExternalInteractions, 'number');
  assert.equal(typeof s.verificationInvocations, 'number');

  // And the counter name is validated, so a stale `count('verificationReads')` fails loudly too.
  assert.throws(() => m.count('verificationReads'), /unknown metric counter/);
});
