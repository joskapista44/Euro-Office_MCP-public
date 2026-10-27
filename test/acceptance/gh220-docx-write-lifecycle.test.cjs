'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { createSessionRegistry } = require('../../src/server/session-registry.cjs');

/**
 * GH #220 1/2 (DOCX P0) -- the GH #100 warm-session wiring of createDocxWriteExecution, offline. Docx mirror
 * of test/acceptance/gh100-u4b-write-wiring.test.cjs: the REAL docx adapter, executor, lifecycle and registry
 * on the fake DOCX transport's OWN GH #100 U1/U3 surface (release / resetForNextCall / health / documentKey,
 * added to test/fakes/fake-docx-driver.cjs by GH #220 1b, mirroring the REAL driver's semantics added to
 * src/drivers/eurooffice/docx-driver.cjs by that same unit -- NOT a thin per-test mock).
 *
 * Covers the GH #220 issue's own required measurements, offline-translatable half: retry SATISFIED / 0 writes
 * ON a warm session, forced reuse rejection on a changed document key, a failed session never held, and the
 * reuse-key identity fields (requester/file/config). The LIVE latency measurements (fresh/warm WRITE,
 * WRITE->READ, READ->WRITE, multi-step) are test/live/docx-write-lifecycle.live.cjs, NOT run here.
 */

const ON = { enabled: true, idleMs: 60000, maxAgeMs: 600000, maxCalls: 50 };
const OFF = { enabled: false, reason: 'switch-off' };
const FILE = '1303626';

const RICH = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'en-US' };
const A = 'A paragraph.';
const B = 'B paragraph.';
const doc = (first = A) => [
  { text: 'Heading', level: 1, props: { ...RICH } },
  { text: first, props: { ...RICH } },
  { text: 'Closing paragraph.', props: { ...RICH } },
];

/** The selector a caller builds from a FRESH inventory (same discipline as docx-rewrite-slice.test.cjs). */
function selectorFrom(paragraphs, target) {
  const primary = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  const b = primary.blocks.find(x => x.text === target);
  return { text: target, kind: b.kind, before: b.locator.context.before, after: b.locator.context.after };
}
const rewriteTask = (from, to, id = 'w1') => ({ id, steps: [{ intent: 'rewrite_paragraph', selector: selectorFrom(doc(from), from), text: to }] });

function setup({ config = OFF, principal = '1000', instance = 'production' } = {}) {
  const store = createStore(doc());
  const ctl = { openKey: 'KEY-1', freshKey: 'KEY-1', fetches: [] };
  // the fake driver's own GH #100 surface reads `script.saveObserver` / `script.documentKey` /
  // `script.releaseFails` / `script.healthThrows` dynamically on every call -- one shared, mutable
  // object lets a test flip a flag BETWEEN two run() calls on the SAME (warm) driver.
  const script = { saveObserver: true, documentKey: () => ctl.openKey };
  const registry = createSessionRegistry(config);
  const drivers = [];
  const configSource = {
    kind: 'BROKER',
    async getEditorConfig({ fileId, operation }) { ctl.fetches.push({ fileId, operation }); return { document: { key: ctl.freshKey } }; },
  };
  const createDriver = () => { const d = createFakeDocxDriver(store, script); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  let t = 0;
  const gate = createSessionGate();
  const reuse = { registry, principal, configSignature: `${instance}|http://127.0.0.1:1/b|WRITE` };
  const execute = D.createDocxWriteExecution({ configSource, gate, runtime, createDriver, reuse, session: { now: () => (t += 1) } });
  const run = (task, fileId = FILE) => execute({ fileId, task });
  return { run, registry, ctl, script, drivers, gate, store };
}

test('OFF: the write call is the pre-#220 shape -- no config fetch, closes every time, zero reuse attempt', async () => {
  const s = setup({ config: OFF });
  const r = await s.run(rewriteTask(A, B));
  assert.equal(r.ok, true, r.reason);
  assert.deepEqual(s.ctl.fetches, []);
  assert.equal(s.drivers[0].calls.open, 1);
  assert.equal(s.drivers[0].calls.close, 1);
  assert.equal(s.drivers[0].calls.release, 0, 'no release() on the OFF path');
  assert.equal(s.registry.status().held, false);
});

test('ON: a clean write is HELD; the next call on the same file+principal+config takes it (one open for two calls)', async () => {
  const s = setup({ config: ON });
  const r1 = await s.run(rewriteTask(A, B));
  assert.equal(r1.ok, true, r1.reason);
  assert.equal(s.registry.status().held, true);
  assert.deepEqual(s.ctl.fetches, [], 'the first call has nothing to compare yet');
  const r2 = await s.run(rewriteTask(B, A, 'w2'));
  assert.equal(r2.ok, true, r2.reason);
  // createDocxWriteExecution always constructs a driver object before the warm check (the SAME shape as
  // createWriteExecution for xlsx) -- the proof is that the SECOND driver is never opened, not that it
  // does not exist.
  assert.equal(s.drivers[1].calls.open, 0, 'the second call\'s own driver is never opened');
  assert.equal(s.drivers[0].calls.open, 1, 'one openSession served both calls');
  assert.equal(s.drivers[0].calls.close, 0);
  assert.deepEqual(s.ctl.fetches, [{ fileId: FILE, operation: 'WRITE' }], 'the fresh key comes from the SAME configSource, WRITE mode');
  assert.equal(s.registry.status().calls, 2);
  assert.equal(s.drivers[0].calls.release, 2, 'BOTH calls ended by release(), not close() -- each left the session warm again');
  assert.equal(s.drivers[0].calls.reset, 1, 'resetForNextCall() ran once, at the SECOND call\'s admission (the first had nothing to take)');
});

test('RETRY on a warm session: resubmitting the FIRST task again is SATISFIED, 0 dispatch, 0 forceSave -- never a blind repeat of an APPLIED mutation', async () => {
  const s = setup({ config: ON });
  const task = rewriteTask(A, B);
  const r1 = await s.run(task);
  assert.equal(r1.ok, true, r1.reason);
  assert.equal(s.drivers[0].calls.forceSave, 1);
  const retry = await s.run(task); // the SAME logical task, same warm session, no process memory re-sent
  assert.equal(retry.ok, true, retry.reason);
  assert.equal(retry.noOp, true, 'RETRY_ZERO_DUPLICATE_WRITE');
  assert.equal(retry.operations[0].status, 'SATISFIED');
  assert.equal(s.drivers[1].calls.open, 0, 'the retry ran on the SAME (warm) session, never opening a fresh one');
  assert.equal(s.drivers[0].calls.forceSave, 1, 'still 1 -- the retry wrote nothing');
  assert.equal(s.registry.status().held, true, 'a write-free, clean retry call stays warm-eligible too');
});

test('a FORCED document-key rotation closes the warm session and the next call opens fresh (never serves a stale model)', async () => {
  const s = setup({ config: ON });
  await s.run(rewriteTask(A, B));
  s.ctl.freshKey = 'KEY-ROLLED';
  const r2 = await s.run(rewriteTask(B, A, 'w2'));
  assert.equal(r2.ok, true, r2.reason);
  assert.equal(s.drivers[0].calls.close, 1, 'the stale warm session was torn down, never served');
  assert.equal(s.drivers[1].calls.open, 1, 'the call ran on its OWN fresh session');
});

test('a call on a DIFFERENT file closes the warm session (one browser per process, file is part of the key)', async () => {
  const s = setup({ config: ON });
  await s.run(rewriteTask(A, B));
  await s.run(rewriteTask(A, B, 'w2'), '9999999');
  assert.equal(s.drivers[0].calls.close, 1);
  assert.equal(s.drivers[1].calls.open, 1);
});

test('a call under a DIFFERENT principal (same file/config) does not take the warm session -- the key includes the requester', async () => {
  const store = createStore(doc());
  const ctl = { freshKey: 'KEY-1' };
  const script = { saveObserver: true };
  const registry = createSessionRegistry(ON);
  const drivers = [];
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: ctl.freshKey } }; } };
  const createDriver = () => { const d = createFakeDocxDriver(store, script); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const execA = D.createDocxWriteExecution({ configSource, gate: createSessionGate(), runtime, createDriver, reuse: { registry, principal: 'user-A', configSignature: 'sig' } });
  const execB = D.createDocxWriteExecution({ configSource, gate: createSessionGate(), runtime, createDriver, reuse: { registry, principal: 'user-B', configSignature: 'sig' } });
  const r1 = await execA({ fileId: FILE, task: rewriteTask(A, B) });
  assert.equal(r1.ok, true, r1.reason);
  assert.equal(registry.status().held, true);
  const r2 = await execB({ fileId: FILE, task: rewriteTask(A, B, 'w2') });
  assert.equal(r2.ok, true, r2.reason);
  assert.equal(drivers[0].calls.close, 1, 'user-B\'s call could not take user-A\'s held session -- it was discarded');
  assert.equal(drivers[1].calls.open, 1, 'user-B opened its own session');
});

test('a call under a DIFFERENT config signature (same file/principal) does not take the warm session', async () => {
  const s = setup({ config: ON });
  await s.run(rewriteTask(A, B));
  const reuse2 = { registry: s.registry, principal: '1000', configSignature: 'a-different-config-signature' };
  const execDifferent = D.createDocxWriteExecution({
    configSource: { kind: 'BROKER', async getEditorConfig() { return { document: { key: s.ctl.freshKey } }; } },
    gate: createSessionGate(), runtime: { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) },
    createDriver: () => { const d = createFakeDocxDriver(s.store, s.script); s.drivers.push(d); return d; },
    reuse: reuse2,
  });
  const r2 = await execDifferent({ fileId: FILE, task: rewriteTask(A, B, 'w2') });
  assert.equal(r2.ok, true, r2.reason);
  assert.equal(s.drivers[0].calls.close, 1, 'the config-signature mismatch discarded the held session');
  assert.equal(s.drivers[1].calls.open, 1);
});

test('a FAILED release (the driver cannot confirm the call ended clean) is never held -- the next call opens fresh, not on a discarded one', async () => {
  const s = setup({ config: ON });
  s.script.releaseFails = true;
  const r1 = await s.run(rewriteTask(A, B));
  assert.equal(r1.ok, true, r1.reason, 'the WRITE itself still succeeded -- only the call-end keep attempt failed');
  assert.equal(s.registry.status().held, false, 'FAILED_SESSION_REUSE=0: a release that cannot confirm is never held');
  assert.equal(s.drivers[0].calls.close, 1, 'the session was closed (the old path), not left dangling');
  s.script.releaseFails = false;
  const r2 = await s.run(rewriteTask(B, A, 'w2'));
  assert.equal(r2.ok, true, r2.reason);
  assert.equal(s.drivers[1].calls.open, 1, 'the second call opened its OWN fresh session -- nothing stale was served');
});

test('health() throwing mid-reuse is treated as a "no" (reuse-check-threw) -- discarded, never held open across the failure', async () => {
  const s = setup({ config: ON });
  await s.run(rewriteTask(A, B));
  s.script.healthThrows = true;
  const r2 = await s.run(rewriteTask(B, A, 'w2'));
  assert.equal(r2.ok, true, r2.reason);
  assert.equal(s.drivers[0].calls.close, 1, 'the held session was torn down after health() threw');
  assert.equal(s.drivers[1].calls.open, 1, 'the call fell back to its own fresh session');
});

test('shutdown closes a held docx session (process end, bounded)', async () => {
  const s = setup({ config: ON });
  await s.run(rewriteTask(A, B));
  const r = await s.registry.shutdown({ timeoutMs: 1000 });
  assert.deepEqual([r.sessions, r.cleanup], [1, 'CLOSED']);
  assert.equal(s.drivers[0].calls.close, 1);
  assert.equal(s.registry.status().held, false);
});

test('while a reused call runs, the gate lease points at the TAKEN (warm) driver, not a never-opened fresh one', async () => {
  const s = setup({ config: ON });
  await s.run(rewriteTask(A, B));
  const leases = [];
  const realTry = s.gate.tryEnter;
  s.gate.tryEnter = () => { const l = realTry(); if (l) leases.push(l); return l; };
  await s.run(rewriteTask(B, A, 'w2'));
  assert.equal(leases.length, 1);
  assert.equal(leases[0].driver.calls, s.drivers[0].calls);
});

test('RED: createDocxWriteExecution with reuse OMITTED is BYTE-IDENTICAL to passing reuse:null -- the default never fetches config or calls release/health', async () => {
  const store = createStore(doc());
  const ctl = { fetches: 0 };
  const drivers = [];
  const configSource = { kind: 'BROKER', async getEditorConfig() { ctl.fetches += 1; return { document: { key: 'K' } }; } };
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, documentKey: () => 'K' }); drivers.push(d); return d; };
  const run = D.createDocxWriteExecution({ configSource, gate: createSessionGate(), runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver });
  const r = await run({ fileId: FILE, task: rewriteTask(A, B) });
  assert.equal(r.ok, true, r.reason);
  assert.equal(ctl.fetches, 0);
  assert.equal(drivers[0].calls.release, 0);
  assert.equal(drivers[0].calls.health, 0);
  assert.equal(drivers[0].calls.close, 1, 'the old, pre-#220 close() path ran instead');
});
