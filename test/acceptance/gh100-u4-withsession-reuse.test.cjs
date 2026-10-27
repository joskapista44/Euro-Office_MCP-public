'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createFakeSession } = require('../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../fakes/fake-adapter.cjs');
const { createM2Executor } = require('../fakes/fake-m2-executor.cjs');
const { buildExecutionResult, AUTHORITY, OPERATION_STATUS, VERIFICATION, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { createSessionRegistry, sessionKey } = require('../../src/server/session-registry.cjs');
const { decideReuse } = require('../../src/server/session-reuse-rule.cjs');

/**
 * GH #100 U4: the withSession reuse path. The REAL lifecycle, the REAL registry (U2) and the REAL rule (U3) on a scripted driver: every decideReuse
 * reason leads to "warm session closed + fresh open", a clean call keeps its session through release(), anything else closes it as before, and with the
 * switch OFF the lifecycle never touches health / documentKey / release / the config fetch.
 */

const CFG = { enabled: true, idleMs: 60000, maxAgeMs: 600000, maxCalls: 50 };
const KEY = sessionKey({ fileId: '42', principal: 'u1', configSignature: 'sig' });
const GOOD_HEALTH = () => ({ alive: true, canSave: false, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 });
const OPS = { id: 'u4-task', operations: [{ id: 'op1' }] };

/** A scripted session driver: the fake transport plus the U1/U3 surface (release / resetForNextCall / health / documentKey). */
function warmDriver(script = {}) {
  const base = createFakeSession(script);
  const extra = { release: 0, reset: 0, health: 0, documentKey: 0 };
  return Object.assign(base, {
    extra,
    async release() {
      extra.release += 1;
      if (script.releaseThrows) throw new Error('release threw');
      if (script.releaseFails) return { ok: false, detail: 'release refused', confirmation: null };
      const confirmation = script.releaseConfirmation !== undefined ? script.releaseConfirmation : (base.calls.forceSave > 0 ? { state: 'confirmed', reason: 'forcesave-confirmed' } : null);
      return { ok: true, detail: 'session-kept', confirmation };
    },
    resetForNextCall() { extra.reset += 1; return { ok: script.resetFails !== true }; },
    async health() { extra.health += 1; if (script.healthThrows) throw new Error('health threw'); return script.health || GOOD_HEALTH(); },
    documentKey() { extra.documentKey += 1; return script.openedKey === undefined ? 'key-A' : script.openedKey; },
  });
}

function reuseOptions(registry, over = {}) {
  return { registry, key: KEY, fetchFreshKey: async () => 'key-A', decide: decideReuse, ...over };
}

function call({ driver, registry = null, reuse = {}, task = OPS, executor = null, adapterScript = {} }) {
  let t = 0;
  const options = { driver, fileId: '42', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  if (registry) options.reuse = reuseOptions(registry, reuse);
  return withSession(options, createFakeAdapter(adapterScript), executor || createM2Executor(task));
}

const seedWarm = async (registry, driver, lastRelease = { ok: true, confirmation: { state: 'confirmed' } }, opened = { id: 'warm-1', frame: 'f', apiWhere: 'a', operation: 'WRITE' }) =>
  registry.hold(KEY, driver, { openedAt: Date.now(), calls: 1, lastRelease, opened });

const writeFreeExecutor = () => async (_handle, { appliedOperations }) =>
  buildExecutionResult({ executionOk: true, outcome: 'task-satisfied', noOp: true, authority: AUTHORITY.LIVE_READ, appliedOperations });

// ---- switch OFF: the product behaves exactly as before -------------------------------------------------------------------------------------

test('OFF: no reuse option, and a disabled registry, are byte-identical to today -- and never touch health / documentKey / release / the config fetch', async () => {
  const touched = [];
  const guarded = () => {
    const d = warmDriver();
    for (const m of ['health', 'documentKey', 'release', 'resetForNextCall']) d[m] = () => { touched.push(m); throw new Error(`${m} must not be called`); };
    return d;
  };
  const a = await call({ driver: guarded() });
  const off = createSessionRegistry({ enabled: false, reason: 'switch-off' });
  let fetches = 0;
  const b = await call({ driver: guarded(), registry: off, reuse: { fetchFreshKey: async () => { fetches += 1; return 'k'; } } });
  assert.equal(a.ok, true);
  assert.deepEqual(touched, []);
  assert.equal(fetches, 0);
  assert.equal('reuse' in a, false);
  assert.equal('reuse' in b, false);
  assert.equal(JSON.stringify(a), JSON.stringify(b));
});

test('OFF: the session closes at the call end (close 1, no release)', async () => {
  const d = warmDriver();
  const r = await call({ driver: d });
  assert.deepEqual([d.calls.open, d.calls.close, d.extra.release], [1, 1, 0]);
  assert.equal(r.cleanup.outcome, 'closed');
});

// ---- ON: keep, then reuse --------------------------------------------------------------------------------------------------------------------

test('ON: a clean WRITE call opens fresh, RELEASES (no close) and the registry holds the session', async () => {
  const registry = createSessionRegistry(CFG);
  const d = warmDriver();
  const r = await call({ driver: d, registry });
  assert.equal(r.ok, true);
  assert.equal(r.persistenceState, 'FORCESAVE_CONFIRMED');
  assert.deepEqual([d.calls.open, d.calls.close, d.extra.release], [1, 0, 1]);
  assert.equal(r.cleanup.outcome, 'released');
  assert.deepEqual({ ...r.reuse }, { attempted: true, reused: false, reason: 'no-warm-session', kept: true, keepReason: 'held' });
  assert.equal(registry.status().held, true);
  assert.equal(registry.status().calls, 1);
});

test('ON: the next call TAKES the kept session -- no open, one reset, the same session id, calls counted -- and keeps it again', async () => {
  const registry = createSessionRegistry(CFG);
  const d = warmDriver();
  await call({ driver: d, registry });
  const fresh = warmDriver();
  const r2 = await call({ driver: fresh, registry });
  assert.equal(r2.ok, true);
  assert.equal(r2.reuse.reused, true);
  assert.equal(r2.reuse.reason, 'all-checks-passed');
  assert.equal(r2.reuse.kept, true);
  assert.deepEqual([fresh.calls.open, fresh.calls.close, fresh.extra.release], [0, 0, 0], 'the fresh driver of the second call is never used');
  assert.deepEqual([d.calls.open, d.calls.close, d.extra.release, d.extra.reset], [1, 0, 2, 1]);
  assert.equal(registry.status().calls, 2);
  assert.equal(r2.lifecycleTrace[r2.lifecycleTrace.length - 1], 'FINALIZED');
});

test('ON: a write-free call (no confirmation at the release) is kept too', async () => {
  const registry = createSessionRegistry(CFG);
  const d = warmDriver();
  const r = await call({ driver: d, registry, executor: writeFreeExecutor() });
  assert.equal(r.reuse.kept, true);
  assert.equal(registry.status().held, true);
  assert.equal(d.calls.close, 0);
});

// ---- ON: every decideReuse reason -> warm closed, fresh open --------------------------------------------------------------------------------

const REASONS = [
  ['editor-not-alive', w => { w.health = { ...GOOD_HEALTH(), alive: false }; }],
  ['unsaved-changes-in-editor', w => { w.health = { ...GOOD_HEALTH(), canSave: true }; }],
  ['cansave-unreadable', w => { w.health = { ...GOOD_HEALTH(), canSave: null }; }],
  ['ds-socket-not-open', w => { w.health = { ...GOOD_HEALTH(), ds: { open: false, sockets: 0, busy: false } }; }],
  ['save-in-flight', w => { w.health = { ...GOOD_HEALTH(), ds: { open: true, sockets: 1, busy: true } }; }],
  ['observer-anomaly', w => { w.health = { ...GOOD_HEALTH(), anomalies: 2 }; }],
  ['no-recorded-document-key', w => { w.openedKey = null; }],
  ['document-session-changed', (w, ctx) => { ctx.fresh = 'key-ROLLED'; }],
  ['fresh-config-key-unavailable', (w, ctx) => { ctx.fresh = null; }],
  ['last-call-end-not-clean', (w, ctx) => { ctx.lastRelease = { ok: true, confirmation: { state: 'unconfirmed' } }; }],
];

for (const [reason, arrange] of REASONS) {
  test(`ON: ${reason} -> the warm session is CLOSED, never served, and the call opens fresh`, async () => {
    const registry = createSessionRegistry(CFG);
    const warmScript = {};
    const ctx = { fresh: 'key-A', lastRelease: { ok: true, confirmation: { state: 'confirmed' } } };
    arrange(warmScript, ctx);
    const warm = warmDriver(warmScript);
    await seedWarm(registry, warm, ctx.lastRelease);
    const fresh = warmDriver();
    const r = await call({ driver: fresh, registry, reuse: { fetchFreshKey: async () => ctx.fresh } });
    assert.equal(r.reuse.reused, false);
    assert.equal(r.reuse.reason, reason);
    assert.equal(r.reuse.warmClosed, true);
    assert.equal(warm.calls.close, 1, 'the warm session was torn down');
    assert.equal(warm.extra.reset, 0, 'a refused session is never reset for use');
    assert.equal(fresh.calls.open, 1, 'the call ran on a FRESH session');
    assert.equal(r.ok, true, 'a failed reuse costs time, never correctness');
  });
}

test('ON: a throwing fetch, a throwing health(), a failing reset and a missing open descriptor are each a "no" with the warm session closed', async () => {
  const cases = [
    ['fresh-config-key-unavailable', {}, { fetchFreshKey: async () => { throw new Error('broker down'); } }, undefined],
    ['reuse-check-threw', { healthThrows: true }, {}, undefined],
    ['reset-failed', { resetFails: true }, {}, undefined],
    ['no-recorded-open-descriptor', {}, {}, null],
  ];
  for (const [reason, warmScript, reuse, opened] of cases) {
    const registry = createSessionRegistry(CFG);
    const warm = warmDriver(warmScript);
    await seedWarm(registry, warm, undefined, opened === undefined ? undefined : opened);
    const fresh = warmDriver();
    const r = await call({ driver: fresh, registry, reuse });
    assert.equal(r.reuse.reason, reason, reason);
    assert.equal(warm.calls.close, 1, reason);
    assert.equal(fresh.calls.open, 1, reason);
  }
});

test('ON: a warm session of ANOTHER key is closed by the registry and the call opens fresh', async () => {
  const registry = createSessionRegistry(CFG);
  const warm = warmDriver();
  await registry.hold(sessionKey({ fileId: '99', principal: 'u1', configSignature: 'sig' }), warm, { openedAt: Date.now(), calls: 1, lastRelease: { ok: true }, opened: { id: 'x' } });
  const fresh = warmDriver();
  const r = await call({ driver: fresh, registry });
  assert.equal(r.reuse.reason, 'no-warm-session');
  assert.equal(warm.calls.close, 1);
  assert.equal(fresh.calls.open, 1);
});

// ---- ON: the call end that must NOT keep the session -----------------------------------------------------------------------------------------

test('ON: an executor that THROWS closes the session (no release, nothing held)', async () => {
  const registry = createSessionRegistry(CFG);
  const d = warmDriver();
  const r = await call({ driver: d, registry, executor: async () => { throw new Error('boom'); } });
  assert.equal(r.reuse.kept, false);
  assert.equal(r.reuse.keepReason, 'call-not-clean');
  assert.deepEqual([d.calls.close, d.extra.release], [1, 0]);
  assert.equal(registry.status().held, false);
});

test('ON: a call whose execution FAILED, or holds an UNCERTAIN write, closes its session (nothing is kept)', async () => {
  const failed = async (_h, { appliedOperations }) => buildExecutionResult({ executionOk: false, outcome: 'primitive-dispatch-failed', authority: AUTHORITY.PLAN_ONLY, appliedOperations });
  // a verified write (so the barrier runs and the save IS confirmed) next to an unproven dispatch: only the uncertainty can stop the keep
  const uncertain = async (_h, { appliedOperations }) => {
    appliedOperations.record({ operationId: 'op0', status: OPERATION_STATUS.APPLIED, verification: VERIFICATION.LIVE_VERIFY, evidence: {} });
    appliedOperations.record({ operationId: 'op1', status: OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, evidence: {},
      uncertainty: { reason: UNCERTAINTY_REASON.EXECUTOR_THREW_WITH_UNACCOUNTED_DISPATCH, dispatchAttempted: true, dispatchEvidence: {}, mutationGeneration: 1 } });
    return buildExecutionResult({ executionOk: true, outcome: 'task-complete', authority: AUTHORITY.LIVE_VERIFY, appliedOperations });
  };
  for (const [name, executor] of [['failed', failed], ['uncertain', uncertain]]) {
    const registry = createSessionRegistry(CFG);
    const d = warmDriver();
    const r = await call({ driver: d, registry, executor });
    assert.equal(r.reuse.kept, false, name);
    assert.equal(r.reuse.keepReason, 'call-not-clean', name);
    assert.deepEqual([d.calls.close, d.extra.release], [1, 0], name);
    assert.equal(registry.status().held, false, name);
  }
});

test('ON: a failed session open is the old F1 path (nothing to close, nothing held)', async () => {
  const registry = createSessionRegistry(CFG);
  const d = warmDriver({ openFails: true });
  const r = await call({ driver: d, registry });
  assert.equal(r.outcome, 'session-open-failed');
  assert.equal(r.reuse.kept, false);
  assert.deepEqual([d.calls.close, d.extra.release], [0, 0]);
  assert.equal(registry.status().held, false);
});

test('ON: a release that is not ok, throws, or leaves a write unconfirmed -> a REAL close, and that close result is the cleanup', async () => {
  const cases = [
    ['release-not-ok', { releaseFails: true }],
    ['release-not-ok', { releaseThrows: true }],
    ['release-claim-not-confirmed', { releaseConfirmation: { state: 'unconfirmed', reason: 'x' } }],
    ['release-claim-not-confirmed', { releaseConfirmation: null }], // a WRITE call with no confirmation is not a kept-worthy end
  ];
  for (const [keepReason, script] of cases) {
    const registry = createSessionRegistry(CFG);
    const d = warmDriver(script);
    const r = await call({ driver: d, registry });
    assert.equal(r.reuse.kept, false, JSON.stringify(script));
    assert.equal(r.reuse.keepReason, keepReason, JSON.stringify(script));
    assert.equal(d.calls.close, 1, JSON.stringify(script));
    assert.equal(r.cleanup.outcome, 'closed', JSON.stringify(script));
    assert.equal(registry.status().held, false, JSON.stringify(script));
  }
});

test('ON: a registry that refuses the hold (max calls reached) -> the session is closed, the reason is reported', async () => {
  const registry = createSessionRegistry({ ...CFG, maxCalls: 1 });
  const d = warmDriver();
  const r = await call({ driver: d, registry });
  assert.equal(r.reuse.kept, false);
  assert.equal(r.reuse.keepReason, 'max-calls-reached');
  assert.equal(d.calls.close, 1);
  assert.equal(registry.status().held, false);
});

test('ON: a READY failure closes the session (nothing was executed, nothing is kept)', async () => {
  const registry = createSessionRegistry(CFG);
  const d = warmDriver();
  const r = await call({ driver: d, registry, adapterScript: { readyThrows: true } });
  assert.equal(r.outcome, 'session-ready-failed');
  assert.equal(r.reuse.kept, false);
  assert.equal(d.calls.close, 1);
  assert.equal(registry.status().held, false);
});

test('ON: a registry whose take() throws is a "no" -- the call opens fresh', async () => {
  const registry = createSessionRegistry(CFG);
  registry.take = async () => { throw new Error('registry down'); };
  const d = warmDriver();
  const r = await call({ driver: d, registry });
  assert.equal(r.reuse.reason, 'registry-take-threw');
  assert.equal(d.calls.open, 1);
  assert.equal(r.ok, true);
});
