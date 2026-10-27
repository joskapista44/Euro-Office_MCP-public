'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionRegistry, sessionKey } = require('../../src/server/session-registry.cjs');

/**
 * GH #100 U4c: xlsx_read warm-session reuse. docs/GH100-U4C-READ-REUSE-DESIGN.md.
 *
 * RED-first baseline this unit changes: on today's main, a read NEVER touches the registry -- `reuse` is not even
 * an accepted constructor option, and two consecutive reads of the same file always open twice. Every test below
 * fails on that baseline (no `take`/`hold` would ever be called) and passes once read-tools.cjs is wired.
 */

const CFG = { enabled: true, idleMs: 60000, maxAgeMs: 600000, maxCalls: 50 };
const FILE = '42';
const GOOD_HEALTH = () => ({ alive: true, canSave: false, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 });

/** A minimal fake XLSX driver: only what readOnce + the U4c admission/release logic touch. */
function fakeDriver(script = {}) {
  const calls = { open: 0, close: 0, readLive: 0, health: 0, documentKey: 0, resetForNextCall: 0, release: 0 };
  return {
    calls,
    async open(args) { calls.open += 1; if (script.openFails) return { ok: false, code: 'OPEN_FAILED', detail: 'boom' }; return { ok: true, operation: args.operation, id: `fake-${calls.open}` }; },
    async close() { calls.close += 1; return { ok: true }; },
    async readLive() {
      calls.readLive += 1;
      if (script.readFails) return { ok: false, detail: 'boom' };
      return { ok: true, payload: { sheets: [{ name: 'Sheet1', index: 0 }] } };
    },
    async health() { calls.health += 1; if (script.healthThrows) throw new Error('health threw'); return script.health || GOOD_HEALTH(); },
    documentKey() { calls.documentKey += 1; return script.openedKey === undefined ? 'key-A' : script.openedKey; },
    resetForNextCall() { calls.resetForNextCall += 1; return { ok: script.resetFails !== true }; },
    async release() {
      calls.release += 1;
      if (script.releaseThrows) throw new Error('release threw');
      if (script.releaseFails) return { ok: false, detail: 'release refused', confirmation: null };
      return { ok: true, detail: 'session-kept', confirmation: script.releaseConfirmation !== undefined ? script.releaseConfirmation : null };
    },
  };
}

function fakeConfigSource(key = 'key-A') {
  return { kind: 'BROKER', async getEditorConfig() { return { document: { key } }; } };
}

function makeTool({ registry, drivers, configSource = fakeConfigSource(), principal = 'u1', configSignature = 'sig' }) {
  let i = 0;
  const createDriver = () => drivers[Math.min(i++, drivers.length - 1)];
  return createReadTool('xlsx', { configSource, createDriver, loadPlaywright: () => {}, reuse: { registry, principal, configSignature } });
}

const read = t => t.handler({ fileId: FILE, operation: 'workbook_overview' });

test('a clean read with reuse wired holds the session instead of closing it, and a second call TAKES it (no second open)', async () => {
  const registry = createSessionRegistry(CFG);
  const d1 = fakeDriver();
  const t1 = makeTool({ registry, drivers: [d1] });
  const out1 = await read(t1);
  assert.equal(out1.ok, true);
  assert.equal(d1.calls.open, 1);
  assert.equal(d1.calls.close, 0, 'a clean read must be HELD, not closed');
  assert.equal(out1.close, null);

  const d2 = fakeDriver();
  const t2 = makeTool({ registry, drivers: [d2] });
  const out2 = await read(t2);
  assert.equal(out2.ok, true);
  assert.equal(d2.calls.open, 0, 'the second call must take the warm session, not open a fresh one');
  assert.equal(d1.calls.health, 1);
  assert.equal(d1.calls.documentKey, 1);
  assert.equal(d1.calls.resetForNextCall, 1);
});

test('without `reuse` wired, two reads behave exactly as before: open, read, close -- every time', async () => {
  const d1 = fakeDriver();
  const d2 = fakeDriver();
  const t = createReadTool('xlsx', { configSource: fakeConfigSource(), createDriver: (() => { let i = 0; const ds = [d1, d2]; return () => ds[i++]; })(), loadPlaywright: () => {} });
  await read(t);
  await read(t);
  assert.equal(d1.calls.open, 1); assert.equal(d1.calls.close, 1);
  assert.equal(d2.calls.open, 1); assert.equal(d2.calls.close, 1);
});

test('a held session that is no longer alive is closed and the call opens fresh', async () => {
  const registry = createSessionRegistry(CFG);
  const stale = fakeDriver();
  stale.health = async () => { stale.calls.health += 1; return { alive: false }; };
  await registry.hold(sessionKey({ fileId: FILE, principal: 'u1', configSignature: 'sig' }), stale, { openedAt: Date.now(), calls: 1, lastRelease: { ok: true, confirmation: null }, opened: { operation: 'READ' } });
  const fresh = fakeDriver();
  const t = makeTool({ registry, drivers: [fresh] });
  const out = await read(t);
  assert.equal(out.ok, true);
  assert.equal(fresh.calls.open, 1, 'a dead held session must not be served -- a fresh one opens');
  assert.equal(stale.calls.close, 1);
});

test('operation axis: a held READ session is taken by a second READ (same key, no eviction)', async () => {
  const registry = createSessionRegistry(CFG);
  const d1 = fakeDriver();
  await registry.hold(sessionKey({ fileId: FILE, principal: 'u1', configSignature: 'sig' }), d1, { openedAt: Date.now(), calls: 1, lastRelease: { ok: true, confirmation: null }, opened: { operation: 'READ' } });
  const fresh = fakeDriver();
  const t = makeTool({ registry, drivers: [fresh] });
  const out = await read(t);
  assert.equal(out.ok, true);
  assert.equal(fresh.calls.open, 0, 'a READ-opened held session must serve a READ request');
});

test('operation axis: a held WRITE-opened session (operation absent from this unit, written by write-tools.cjs) is also taken by a READ', async () => {
  const registry = createSessionRegistry(CFG);
  const dWrite = fakeDriver();
  await registry.hold(sessionKey({ fileId: FILE, principal: 'u1', configSignature: 'sig' }), dWrite, { openedAt: Date.now(), calls: 1, lastRelease: { ok: true, confirmation: { state: 'confirmed' } }, opened: { operation: 'WRITE' } });
  const fresh = fakeDriver();
  const t = makeTool({ registry, drivers: [fresh] });
  const out = await read(t);
  assert.equal(out.ok, true);
  assert.equal(fresh.calls.open, 0, 'a WRITE-opened session is edit-capable, which always serves a READ');
});

test('a stale document key (a foreign writer rolled the session) closes the held driver and opens fresh', async () => {
  const registry = createSessionRegistry(CFG);
  const d1 = fakeDriver({ openedKey: 'key-OLD' });
  const t1 = makeTool({ registry, drivers: [d1], configSource: fakeConfigSource('key-OLD') });
  await read(t1);
  assert.equal(d1.calls.close, 0);

  const fresh = fakeDriver();
  const t2 = makeTool({ registry, drivers: [fresh], configSource: fakeConfigSource('key-NEW') });
  const out = await read(t2);
  assert.equal(out.ok, true);
  assert.equal(d1.calls.close, 1, 'the stale-keyed warm session must be closed, never served');
  assert.equal(fresh.calls.open, 1);
});

test('a key mismatch (different fileId/principal/configSignature) never takes the held session', async () => {
  const registry = createSessionRegistry(CFG);
  const dOther = fakeDriver();
  const t1 = makeTool({ registry, drivers: [dOther], principal: 'OTHER-USER' });
  await read(t1);

  const fresh = fakeDriver();
  const t2 = makeTool({ registry, drivers: [fresh], principal: 'u1' });
  const out = await read(t2);
  assert.equal(out.ok, true);
  assert.equal(fresh.calls.open, 1, 'a different principal must never take another principal\'s warm session');
  assert.equal(dOther.calls.close, 1, 'the registry itself closes the mismatched held session');
});

test('an OPEN failure on a fresh call never holds anything, and reports exactly as before reuse existed', async () => {
  const registry = createSessionRegistry(CFG);
  const failing = fakeDriver({ openFails: true });
  const t = makeTool({ registry, drivers: [failing] });
  const out = await read(t);
  assert.equal(out.ok, false);
  assert.equal(out.stage, 'OPEN');
});

// ---- [ATNEZES: VISSZA], 4 required RED-first tests -----------------------------------------------------------------------------

test('zero-1: a FAILED read is never offered back -- the next call opens fresh, and the failed session is closed, not held', async () => {
  const registry = createSessionRegistry(CFG);
  const d1 = fakeDriver({ readFails: true });
  const t1 = makeTool({ registry, drivers: [d1] });
  const out1 = await read(t1);
  assert.equal(out1.ok, false);
  assert.equal(d1.calls.close, 1, 'a failed read must be CLOSED');
  assert.equal(d1.calls.release, 0, 'a failed read must never attempt release() -- it was never a candidate to keep');
  assert.equal(registry.status().held, false);

  const d2 = fakeDriver();
  const t2 = makeTool({ registry, drivers: [d2] });
  const out2 = await read(t2);
  assert.equal(out2.ok, true);
  assert.equal(d2.calls.open, 1, 'the next call must open fresh -- nothing was held to take');
});

test('zero-2: a release() that is not ok, or not confirmed/null, is never held -- the driver is closed instead', async () => {
  for (const script of [{ releaseFails: true }, { releaseThrows: true }, { releaseConfirmation: { state: 'unconfirmed' } }]) {
    const registry = createSessionRegistry(CFG);
    const d1 = fakeDriver(script);
    const t1 = makeTool({ registry, drivers: [d1] });
    const out1 = await read(t1);
    assert.equal(out1.ok, true, JSON.stringify(script));
    assert.equal(d1.calls.release, 1, JSON.stringify(script));
    assert.equal(d1.calls.close, 1, `a non-clean release must still CLOSE the driver: ${JSON.stringify(script)}`);
    assert.equal(registry.status().held, false, JSON.stringify(script));
  }
  // a CONFIRMED release is accepted (the write-free acceptance rule), exactly like a null one
  const registry = createSessionRegistry(CFG);
  const d2 = fakeDriver({ releaseConfirmation: { state: 'confirmed' } });
  const t2 = makeTool({ registry, drivers: [d2] });
  await read(t2);
  assert.equal(d2.calls.close, 0, 'a confirmed release is clean -- the session is held');
  assert.equal(registry.status().held, true);
});

test('zero-3: the ORIGINAL openedAt/calls carry through a reused chain -- R6 (max-age/max-calls) still fires on a read-only chain', async () => {
  const registry = createSessionRegistry({ enabled: true, idleMs: 60000, maxAgeMs: 600000, maxCalls: 3 });
  const drivers = [fakeDriver(), fakeDriver(), fakeDriver(), fakeDriver()];
  for (let i = 0; i < 3; i += 1) {
    const t = makeTool({ registry, drivers: [drivers[i]] });
    const out = await read(t);
    assert.equal(out.ok, true, `call ${i}`);
    assert.equal(drivers[i].calls.open, i === 0 ? 1 : 0, `call ${i}: only the first call opens`);
  }
  // the 3rd held offer hit calls===maxCalls -> the registry refuses to hold it, so it closed instead
  assert.equal(drivers[0].calls.close, 1, 'the 3rd call`s hold must be refused by max-calls and close the chain');
  assert.equal(registry.status().held, false);
  // a 4th call therefore opens fresh -- the chain did NOT silently keep aging past the limit
  const t4 = makeTool({ registry, drivers: [drivers[3]] });
  await read(t4);
  assert.equal(drivers[3].calls.open, 1, 'the 4th call must open fresh: the chain was correctly discarded at max-calls');
});

test('zero-4: write -> read -> write, ONE shared registry/key -- the SECOND write REUSES the session the read took and gave back', async () => {
  const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
  const { book, FILE: WFILE } = require('../fakes/write-wiring.cjs');
  const { createWriteTool } = require('../../src/server/write-tools.cjs');
  const store = createStore(book());
  const registry = createSessionRegistry(CFG);
  const ctl = { key: 'KEY-1' };
  const wrapDriver = d => Object.assign(d, {
    resetForNextCall() { return { ok: true }; },
    async health() { return { alive: true, canSave: false, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 }; },
    documentKey() { return ctl.key; },
    async release() { return { ok: true, detail: 'kept', confirmation: d.calls.forceSave > 0 ? { state: 'confirmed', reason: 'forcesave-confirmed' } : null }; },
  });
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: ctl.key } }; } };
  const reuse = { registry, principal: 'u1', configSignature: 'sig' };
  const writeDrivers = [];
  const writeTool = createWriteTool('xlsx', {
    configSource, reuse, runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) },
    createDriver: () => { const d = wrapDriver(createFakeXlsxDriver(store, {})); writeDrivers.push(d); return d; },
    session: { ready: { wait: async () => {} }, persistence: { wait: async () => {} } },
  });
  const readDrivers = [];
  const readTool = createReadTool('xlsx', {
    configSource, reuse, loadPlaywright: () => {},
    createDriver: () => { const d = wrapDriver(createFakeXlsxDriver(store, {})); readDrivers.push(d); return d; },
  });
  const step = v => ({ intent: 'write_range', id: 'w', sheet: 'Budget', range: 'F2', values: [[v]] });

  // the production createDriver() factory runs on EVERY call regardless of reuse outcome (write-tools.cjs
  // builds a driver before knowing whether admitWarmSession will take an existing one -- it is simply never
  // opened when reuse succeeds). The real reuse signal is therefore open()/close() CALLS on the underlying
  // driver objects, summed across every constructed instance -- the same measure the U4b write-wiring suite
  // uses, not how many driver objects got constructed.
  const sumOpen = ds => ds.reduce((a, d) => a + d.calls.open, 0);
  const sumClose = ds => ds.reduce((a, d) => a + d.calls.close, 0);

  const w1 = await writeTool.handler({ fileId: WFILE, task: { id: 't1', steps: [step(1)] } });
  assert.equal(w1.ok, true, JSON.stringify(w1));
  assert.equal(sumOpen(writeDrivers), 1);
  assert.equal(sumClose(writeDrivers), 0, 'a clean write with reuse ON must be HELD, not closed');

  const r1 = await readTool.handler({ fileId: WFILE, operation: 'read_range', sheet: 'Budget', range: 'F2:F2' });
  assert.equal(r1.ok, true, JSON.stringify(r1));
  assert.equal(sumOpen(readDrivers), 0, 'the read must TAKE the write-opened session, never open its own driver');
  assert.equal(sumClose(readDrivers), 0, 'nothing on the read side was ever opened, so nothing closes');

  const w2 = await writeTool.handler({ fileId: WFILE, task: { id: 't2', steps: [step(2)] } });
  assert.equal(w2.ok, true, JSON.stringify(w2));
  assert.equal(sumOpen(writeDrivers), 1, 'the second write must REUSE: no driver anywhere opens a second time');
  assert.equal(sumClose(writeDrivers), 0, 'the chain is still held, not closed, after the second write');
});

test('a disabled registry (the switch OFF) never calls take/hold, and the result is byte-identical to no reuse wired at all', async () => {
  const registry = createSessionRegistry({ enabled: false, reason: 'switch-off' });
  let took = 0; let held = 0;
  const wrapped = { ...registry, take: async k => { took += 1; return registry.take(k); }, hold: async (...a) => { held += 1; return registry.hold(...a); } };
  const d1 = fakeDriver();
  const t = makeTool({ registry: wrapped, drivers: [d1] });
  const out = await read(t);
  assert.equal(out.ok, true);
  assert.equal(took, 0);
  assert.equal(held, 0);
  assert.equal(d1.calls.close, 1, 'disabled registry: a read closes exactly as before U4c');
});
