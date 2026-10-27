'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { ENV, CEILING, sessionReuseConfig, sessionKey, createSessionRegistry } = require('../../src/server/session-registry.cjs');
const { buildServerTools } = require('../../src/server/server-tools.cjs');

/**
 * GH #100 U2: the warm-session registry, its switch and its limits. Nothing in the product holds or takes a session before U4; with the switch OFF
 * (the default) hold() refuses and take() answers null, so the product is exactly as before -- proved here, not assumed.
 */

const ON = { [ENV.SWITCH]: '1', [ENV.IDLE]: '5000', [ENV.AGE]: '60000', [ENV.CALLS]: '10' };
const KEY = sessionKey({ fileId: '42', principal: 'alice', configSignature: 'cfg-1' });

/** Deterministic clock + timers + drivers whose close() is observable. */
function rig(env = ON, extra = {}) {
  const clock = { t: 1000 };
  const timers = [];
  const registry = createSessionRegistry(sessionReuseConfig(env), {
    now: () => clock.t,
    setTimer: (fn, ms) => { const h = { fn, ms, cleared: false, unref() {} }; timers.push(h); return h; },
    clearTimer: h => { h.cleared = true; },
    ...extra,
  });
  const driver = name => { const d = { name, closed: 0, close: async () => { d.closed += 1; return { ok: true }; } }; return d; };
  const fire = () => { for (const h of timers) if (!h.cleared) { h.cleared = true; h.fn(); } };
  return { registry, clock, timers, driver, fire };
}
const settle = () => new Promise(r => setImmediate(r));

test('config: OFF by default and for anything unclear; ON only with the switch AND all three limits, within the hard ceilings', () => {
  assert.deepEqual(sessionReuseConfig({}), { enabled: false, reason: 'switch-off' });
  assert.deepEqual(sessionReuseConfig({ [ENV.SWITCH]: '0' }), { enabled: false, reason: 'switch-off' });
  assert.equal(sessionReuseConfig({ [ENV.SWITCH]: 'yes' }).reason, 'switch-invalid');
  assert.equal(sessionReuseConfig({ [ENV.SWITCH]: '1' }).reason, 'idle-ms-required-or-invalid');
  assert.equal(sessionReuseConfig({ [ENV.SWITCH]: '1', [ENV.IDLE]: '5000' }).reason, 'max-age-ms-required-or-invalid');
  assert.equal(sessionReuseConfig({ [ENV.SWITCH]: '1', [ENV.IDLE]: '5000', [ENV.AGE]: '60000' }).reason, 'max-calls-required-or-invalid');
  for (const bad of ['0', '-5', '1.5', '5e3', ' 5', 'abc', '', String(CEILING.idleMs + 1)]) {
    assert.equal(sessionReuseConfig({ ...ON, [ENV.IDLE]: bad }).enabled, false, `idle ${JSON.stringify(bad)}`);
  }
  assert.equal(sessionReuseConfig({ ...ON, [ENV.AGE]: String(CEILING.maxAgeMs + 1) }).enabled, false);
  assert.equal(sessionReuseConfig({ ...ON, [ENV.CALLS]: String(CEILING.maxCalls + 1) }).enabled, false);
  assert.deepEqual(sessionReuseConfig(ON), { enabled: true, idleMs: 5000, maxAgeMs: 60000, maxCalls: 10 });
});

test('OFF: hold refuses (the caller keeps ownership and closes), take answers null, no timer is ever created, status says disabled', async () => {
  const { registry, timers, driver } = rig({});
  const d = driver('a');
  assert.deepEqual(await registry.hold(KEY, d), { held: false, reason: 'reuse-disabled' });
  assert.equal(d.closed, 0, 'the registry never closes a session it did not accept');
  assert.equal(await registry.take(KEY), null);
  assert.equal(timers.length, 0);
  assert.deepEqual(registry.status(), { enabled: false, reason: 'switch-off', held: false });
  assert.deepEqual(await registry.shutdown({ timeoutMs: 10 }), { sessions: 0, cleanup: 'NONE', ms: 0 });
});

test('OFF product: buildServerTools with the default environment serves the same tools, and its registry is disabled', () => {
  const off = buildServerTools({ config: { broker: { url: 'http://127.0.0.1:1', instance: 'x' } }, configIdentity: 'id', env: {} });
  assert.equal(off.sessions.enabled, false);
  assert.deepEqual(Object.keys(off.tools), Object.keys(buildServerTools({ config: { broker: { url: 'http://127.0.0.1:1', instance: 'x' } }, configIdentity: 'id', env: ON }).tools), 'the switch changes no tool');
  assert.equal(off.gate.max, 1, 'MAX_ACTIVE_SESSIONS_PER_PROCESS stays 1');
});

test('key: only (fileId, principal, config signature) strings; anything missing is no key', () => {
  assert.equal(typeof KEY, 'string');
  for (const bad of [{ fileId: '', principal: 'a', configSignature: 'c' }, { fileId: '1', principal: null, configSignature: 'c' }, { fileId: '1', principal: 'a' }, { fileId: 1, principal: 'a', configSignature: 'c' }]) {
    assert.equal(sessionKey(bad), null);
  }
});

test('ON: hold then take with the SAME key returns the driver, once; the registry no longer owns it', async () => {
  const { registry, driver } = rig();
  const d = driver('a');
  assert.deepEqual(await registry.hold(KEY, d, { openedAt: 1000, calls: 1 }), { held: true });
  assert.equal(registry.status().held, true);
  const got = await registry.take(KEY);
  assert.equal(got.driver, d);
  assert.equal(d.closed, 0);
  assert.equal(await registry.take(KEY), null, 'a taken session cannot be taken twice');
  assert.equal(registry.status().held, false);
});

test('a different file, requester or config NEVER gets the held session: it is closed (one browser per process) and null comes back', async () => {
  for (const other of [{ fileId: '43', principal: 'alice', configSignature: 'cfg-1' }, { fileId: '42', principal: 'bob', configSignature: 'cfg-1' }, { fileId: '42', principal: 'alice', configSignature: 'cfg-2' }]) {
    const { registry, driver } = rig();
    const d = driver('a');
    await registry.hold(KEY, d);
    assert.equal(await registry.take(sessionKey(other)), null);
    assert.equal(d.closed, 1, JSON.stringify(other));
    assert.equal(registry.status().held, false);
  }
});

test('hold while another is held: the older is closed FIRST, only the newer remains', async () => {
  const { registry, driver } = rig();
  const a = driver('a'); const b = driver('b');
  await registry.hold(KEY, a);
  await registry.hold(sessionKey({ fileId: '7', principal: 'alice', configSignature: 'cfg-1' }), b);
  assert.deepEqual([a.closed, b.closed], [1, 0]);
});

test('idle timeout: the timer closes the held session; an unfired-timer mutant would leave it held', async () => {
  const { registry, driver, timers, fire } = rig();
  const d = driver('a');
  await registry.hold(KEY, d);
  assert.equal(timers.length, 1);
  assert.equal(timers[0].ms, 5000);
  fire(); await settle();
  assert.equal(d.closed, 1);
  assert.equal(registry.status().held, false);
  assert.equal(await registry.take(KEY), null);
});

test('take clears the idle timer (a taken session is not closed under the running call)', async () => {
  const { registry, driver, timers, fire } = rig();
  const d = driver('a');
  await registry.hold(KEY, d);
  await registry.take(KEY);
  assert.equal(timers[0].cleared, true);
  fire(); await settle();
  assert.equal(d.closed, 0);
});

test('an EXPIRED session is never served even if its timer did not run: idle, age and call limits are checked at take()', async () => {
  for (const [mutate, why] of [[c => { c.t += 5000; }, 'idle'], [c => { c.t += 60000; }, 'age']]) {
    const { registry, driver, clock } = rig();
    const d = driver('a');
    await registry.hold(KEY, d, { openedAt: clock.t, calls: 1 });
    mutate(clock);
    assert.equal(await registry.take(KEY), null, why);
    assert.equal(d.closed, 1, why);
  }
});

test('limits at hold(): a session at its maximum age or call count is refused (the caller closes it), never held', async () => {
  const { registry, driver, clock, timers } = rig();
  const old = driver('old'); const busy = driver('busy');
  assert.deepEqual(await registry.hold(KEY, old, { openedAt: clock.t - 60000, calls: 1 }), { held: false, reason: 'max-age-reached' });
  assert.deepEqual(await registry.hold(KEY, busy, { openedAt: clock.t, calls: 10 }), { held: false, reason: 'max-calls-reached' });
  assert.deepEqual([old.closed, busy.closed], [0, 0]);
  assert.equal(timers.length, 0);
  assert.equal((await registry.hold(KEY, { name: 'no-close' })).reason, 'invalid-offer');
  assert.equal((await registry.hold(null, driver('k'))).reason, 'invalid-offer');
});

test('a close() that throws or fails is reported, never thrown, and the registry is empty afterwards', async () => {
  const { registry } = rig();
  await registry.hold(KEY, { close: async () => { throw new Error('browser gone'); } });
  assert.deepEqual(await registry.discard('test'), { discarded: true, reason: 'test', closeOk: false });
  assert.equal(registry.status().held, false);
  await registry.hold(KEY, { close: async () => ({ ok: false, detail: 'x' }) });
  assert.equal((await registry.discard('test')).closeOk, false);
});

test('shutdown closes the held session; a close that never returns is cut at the timeout (TIMEOUT), the registry still empties', async () => {
  const a = rig();
  const d = a.driver('a');
  await a.registry.hold(KEY, d);
  const r = await a.registry.shutdown({ timeoutMs: 1000 });
  assert.deepEqual([r.sessions, r.cleanup, d.closed], [1, 'CLOSED', 1]);
  const b = rig();
  await b.registry.hold(KEY, { close: () => new Promise(() => {}) });
  const t = await b.registry.shutdown({ timeoutMs: 30 });
  assert.deepEqual([t.sessions, t.cleanup], [1, 'TIMEOUT']);
  assert.equal(b.registry.status().held, false);
});

test('status exposes a DIGEST of the key, never the key (no uid, no file id)', async () => {
  const { registry, driver } = rig();
  await registry.hold(KEY, driver('a'));
  const s = JSON.stringify(registry.status());
  assert.match(s, /keyDigest/);
  assert.equal(s.includes('alice'), false);
});

test('bin: the registry is drained on SIGTERM AND when the server stops on its own (stdin closed), both bounded', () => {
  const src = fs.readFileSync(path.join(__dirname, '../../bin/euro-office-mcp.cjs'), 'utf8');
  assert.equal((src.match(/sessions\.shutdown\(\{ timeoutMs: SHUTDOWN_TIMEOUT_MS \}\)/g) || []).length, 2);
  assert.match(src, /const \{ tools, gate, sessions \} = buildServerTools\(/);
});
