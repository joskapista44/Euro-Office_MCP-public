'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { book, FILE } = require('../fakes/write-wiring.cjs');
const { createWriteTool } = require('../../src/server/write-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { createSessionRegistry } = require('../../src/server/session-registry.cjs');
const { buildServerTools } = require('../../src/server/server-tools.cjs');

/**
 * GH #100 U4b: the product wiring of the warm-session path in the WRITE tool. The REAL write tool, executor, lifecycle and registry on the fake XLSX
 * transport (given the U1/U3 driver surface by a thin wrapper): switch OFF changes nothing (no config fetch), switch ON turns two calls on one file
 * into ONE open, another file closes the warm session, and shutdown closes a held one.
 */

const ON = { enabled: true, idleMs: 60000, maxAgeMs: 600000, maxCalls: 50 };
const OFF = { enabled: false, reason: 'switch-off' };
const step = (v, id = 'w1') => ({ intent: 'write_range', id, sheet: 'Budget', range: 'F2', values: [[v]] });

/** The fake driver plus the surface the reuse path needs. release = the call end of a kept session (a write call has force-saved and confirmed). */
function withReuseSurface(d, ctl) {
  return Object.assign(d, {
    async release() { ctl.releases += 1; return { ok: true, detail: 'session-kept', confirmation: d.calls.forceSave > 0 ? { state: 'confirmed', reason: 'forcesave-confirmed' } : null }; },
    resetForNextCall() { return { ok: true }; },
    async health() { return { alive: true, canSave: false, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 }; },
    documentKey() { return ctl.openKey; },
  });
}

function setup({ config, instance = 'production' } = {}) {
  const store = createStore(book());
  const ctl = { openKey: 'KEY-1', freshKey: 'KEY-1', fetches: [], releases: 0 };
  const registry = createSessionRegistry(config || OFF);
  const drivers = [];
  const configSource = {
    kind: 'BROKER',
    async getEditorConfig({ fileId, operation }) { ctl.fetches.push({ fileId, operation }); return { document: { key: ctl.freshKey } }; },
  };
  const createDriver = () => { const d = withReuseSurface(createFakeXlsxDriver(store, {}), ctl); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  let t = 0;
  const session = { ready: { wait: async () => {} }, persistence: { wait: async () => {} }, now: () => (t += 1) };
  const gate = createSessionGate();
  const reuse = { registry, principal: '1000', configSignature: `${instance}|http://127.0.0.1:1/b|WRITE` };
  const tool = createWriteTool('xlsx', { configSource, gate, runtime, createDriver, session, reuse });
  const run = (v, fileId = FILE) => tool.handler({ fileId, task: { id: 't', steps: [step(v)] } });
  const sum = k => drivers.reduce((a, d) => a + d.calls[k], 0);
  return { run, registry, ctl, drivers, gate, sum };
}

test('OFF: the write call is the pre-reuse one -- no config fetch, the session is closed at the end', async () => {
  const s = setup();
  const r = await s.run(1);
  assert.equal(r.ok, true);
  assert.deepEqual(s.ctl.fetches, []);
  assert.deepEqual([s.sum('open'), s.sum('close')], [1, 1]);
  assert.equal(s.ctl.releases, 0, 'no release() on the OFF path');
  assert.equal(s.registry.status().held, false);
});

test('ON: two consecutive calls on the same file open ONCE; the second fetches the fresh key and runs on the kept session', async () => {
  const s = setup({ config: ON });
  const r1 = await s.run(1);
  assert.equal(r1.ok, true);
  assert.equal(s.registry.status().held, true);
  assert.deepEqual(s.ctl.fetches, [], 'the first call has nothing to compare');
  const r2 = await s.run(2);
  assert.equal(r2.ok, true);
  assert.equal(s.sum('open'), 1, 'one openSession for two calls');
  assert.equal(s.sum('close'), 0);
  assert.deepEqual(s.ctl.fetches, [{ fileId: FILE, operation: 'WRITE' }], 'the fresh key comes from the same config source, in WRITE mode');
  assert.equal(s.drivers[1].calls.open, 0, 'the second call\'s own driver is never opened');
  assert.equal(s.registry.status().calls, 2);
});

test('ON: a rolled document key -> the kept session is closed and the call opens fresh (correctness unchanged)', async () => {
  const s = setup({ config: ON });
  await s.run(1);
  s.ctl.freshKey = 'KEY-ROLLED';
  const r2 = await s.run(2);
  assert.equal(r2.ok, true);
  assert.equal(s.drivers[0].calls.close, 1, 'the warm session was torn down');
  assert.equal(s.drivers[1].calls.open, 1, 'the call ran on its own fresh session');
});

test('ON: a failed config fetch is a "no", never "unchanged"', async () => {
  const s = setup({ config: ON });
  await s.run(1);
  s.ctl.freshKey = null;
  const r2 = await s.run(2);
  assert.equal(r2.ok, true);
  assert.equal(s.drivers[0].calls.close, 1);
  assert.equal(s.drivers[1].calls.open, 1);
});

test('ON: a call on ANOTHER file closes the warm session (one browser per process)', async () => {
  const s = setup({ config: ON });
  await s.run(1);
  const r2 = await s.run(2, '777');
  assert.equal(s.drivers[0].calls.close, 1);
  assert.equal(s.drivers[1].calls.open, 1);
  assert.equal(r2.ok === true || r2.ok === false, true);
});

test('ON: shutdown closes a held session (the process end)', async () => {
  const s = setup({ config: ON });
  await s.run(1);
  const r = await s.registry.shutdown({ timeoutMs: 1000 });
  assert.deepEqual([r.sessions, r.cleanup], [1, 'CLOSED']);
  assert.equal(s.drivers[0].calls.close, 1);
  assert.equal(s.registry.status().held, false);
});

test('ON: while a reused call runs, the gate lease points at the session that is ACTIVE (a shutdown closes the right browser)', async () => {
  const s = setup({ config: ON });
  await s.run(1);
  const leases = [];
  const realTry = s.gate.tryEnter;
  s.gate.tryEnter = () => { const l = realTry(); if (l) leases.push(l); return l; };
  await s.run(2);
  assert.equal(leases.length, 1);
  assert.equal(leases[0].driver.calls, s.drivers[0].calls, 'the lease tracks the taken (warm) driver, not the never-opened fresh one');
});

test('product wiring: buildServerTools hands the write tool the process registry; the switch alone (no limits) keeps it OFF', () => {
  const off = buildServerTools({ config: { broker: { url: 'http://127.0.0.1:1/b' } }, env: {} });
  assert.equal(off.sessions.enabled, false);
  const partial = buildServerTools({ config: { broker: { url: 'http://127.0.0.1:1/b' } }, env: { EURO_MCP_SESSION_REUSE: '1' } });
  assert.equal(partial.sessions.enabled, false);
  const on = buildServerTools({ config: { broker: { url: 'http://127.0.0.1:1/b' } }, env: { EURO_MCP_SESSION_REUSE: '1', EURO_MCP_SESSION_IDLE_MS: '1000', EURO_MCP_SESSION_MAX_AGE_MS: '5000', EURO_MCP_SESSION_MAX_CALLS: '5' } });
  assert.equal(on.sessions.enabled, true);
});
