'use strict';

const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');

const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { decideReuse } = require('../../src/server/session-reuse-rule.cjs');
const { fakeDriverEnv, openFakeDriver, saveCycle, answerOk, write } = require('../fakes/fake-xlsx-session-env.cjs');

/**
 * GH #100 U3: the reuse rule (pure) and the driver's health() / documentKey(). The rule's every check, with a mutant that removes it, and the
 * "unknown is a no" discipline. The measurement behind it: docs/GH100-U3-FRESHNESS-PLAN.md.
 */

const GOOD = () => ({
  health: { alive: true, canSave: false, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 },
  openedKey: 'key-A', freshKey: 'key-A', lastRelease: { ok: true, confirmation: { state: 'confirmed' } },
});

test('all checks pass: reuse, with a write-free last call (no confirmation) as well', () => {
  assert.deepEqual(decideReuse(GOOD()), { reuse: true, reason: 'all-checks-passed' });
  assert.equal(decideReuse({ ...GOOD(), lastRelease: { ok: true, confirmation: null } }).reuse, true);
  assert.equal(decideReuse({ ...GOOD(), lastRelease: { ok: true } }).reuse, true);
});

const CASES = [
  ['editor not alive', g => { g.health.alive = false; }, 'editor-not-alive'],
  ['alive unknown', g => { delete g.health.alive; }, 'editor-not-alive'],
  ['unsaved changes', g => { g.health.canSave = true; }, 'unsaved-changes-in-editor'],
  ['canSave unreadable (null)', g => { g.health.canSave = null; }, 'cansave-unreadable'],
  ['canSave missing', g => { delete g.health.canSave; }, 'cansave-unreadable'],
  ['DS socket closed', g => { g.health.ds.open = false; }, 'ds-socket-not-open'],
  ['DS state missing', g => { delete g.health.ds; }, 'ds-socket-not-open'],
  ['a save in flight', g => { g.health.ds.busy = true; }, 'save-in-flight'],
  ['busy unknown', g => { delete g.health.ds.busy; }, 'save-in-flight'],
  ['observer anomaly', g => { g.health.anomalies = 1; }, 'observer-anomaly'],
  ['anomalies unknown (null)', g => { g.health.anomalies = null; }, 'observer-anomaly'],
  ['last call end failed', g => { g.lastRelease = { ok: false, confirmation: { state: 'confirmed' } }; }, 'last-call-end-not-clean'],
  ['last call end unconfirmed write', g => { g.lastRelease = { ok: true, confirmation: { state: 'unconfirmed' } }; }, 'last-call-end-not-clean'],
  ['no last call end', g => { g.lastRelease = undefined; }, 'last-call-end-not-clean'],
  ['no recorded key', g => { g.openedKey = null; }, 'no-recorded-document-key'],
  ['empty recorded key', g => { g.openedKey = ''; }, 'no-recorded-document-key'],
  ['fresh config key unavailable', g => { g.freshKey = null; }, 'fresh-config-key-unavailable'],
  ['fresh config key empty', g => { g.freshKey = ''; }, 'fresh-config-key-unavailable'],
  ['the document session rolled (key changed)', g => { g.freshKey = 'key-B'; }, 'document-session-changed'],
];
for (const [name, mutate, reason] of CASES) {
  test(`no reuse: ${name} -> ${reason}`, () => {
    const g = GOOD(); mutate(g);
    assert.deepEqual(decideReuse(g), { reuse: false, reason });
  });
}

// ---- GH #100 U4c: the operation axis (docs/GH100-U4C-READ-REUSE-DESIGN.md sec. 4) ----

test('operation axis: a WRITE request may only take a session opened for WRITE', () => {
  assert.deepEqual(
    decideReuse({ ...GOOD(), requestedOperation: 'WRITE', openedOperation: 'READ' }),
    { reuse: false, reason: 'held-session-not-write-capable' },
  );
  assert.deepEqual(
    decideReuse({ ...GOOD(), requestedOperation: 'WRITE', openedOperation: undefined }),
    { reuse: false, reason: 'held-session-not-write-capable' },
  );
  assert.equal(decideReuse({ ...GOOD(), requestedOperation: 'WRITE', openedOperation: 'WRITE' }).reuse, true);
});

test('operation axis: a READ request accepts a session opened for either operation', () => {
  assert.equal(decideReuse({ ...GOOD(), requestedOperation: 'READ', openedOperation: 'WRITE' }).reuse, true);
  assert.equal(decideReuse({ ...GOOD(), requestedOperation: 'READ', openedOperation: 'READ' }).reuse, true);
});

test('operation axis: checked FIRST, before any health/freshness I/O (a WRITE-incompatible session never pays the rest of the chain)', () => {
  const broken = { ...GOOD(), requestedOperation: 'WRITE', openedOperation: 'READ' };
  broken.health.alive = false; // would ALSO fail on health -- the operation reason must still win, proving check order
  assert.equal(decideReuse(broken).reason, 'held-session-not-write-capable');
});

test('no input at all, or an empty object, is a no (never a default yes)', () => {
  assert.equal(decideReuse().reuse, false);
  assert.equal(decideReuse({}).reuse, false);
  assert.equal(decideReuse({ health: {} }).reuse, false);
});

test('the key compare is exact: a prefix or a different case is a different document session', () => {
  assert.equal(decideReuse({ ...GOOD(), freshKey: 'key-' }).reuse, false);
  assert.equal(decideReuse({ ...GOOD(), freshKey: 'KEY-A' }).reuse, false);
});

test('the rule never reads the key into its answer (the reason names the check, not the value)', () => {
  const r = decideReuse({ ...GOOD(), freshKey: 'secret-key-value' });
  assert.equal(JSON.stringify(r).includes('secret-key-value'), false);
  assert.equal(JSON.stringify(r).includes('key-A'), false);
});

// ---- the driver side: health() and documentKey() on the fake editor ----

const opened = [];
after(async () => { for (const d of opened) await d.close(); });
const open = async (env, o) => { const d = await openFakeDriver(drv, env, o); opened.push(d); return d; };

test('health(): a healthy kept session after an acked, confirmed write call', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  await write(d);
  saveCycle(env.ctl.ws, 0, 4004, { time: 4704 });
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 4704);
  assert.equal((await d.forceSave()).confirmed, true);
  const h = await d.health();
  assert.deepEqual(h, { alive: true, canSave: false, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 });
  const rel = await d.release();
  assert.equal(decideReuse({ health: h, openedKey: d.documentKey(), freshKey: 'KEY1', lastRelease: rel }).reuse, true);
});

test('health(): the DS socket closing, a dead frame and unreadable canSave are each visible', async () => {
  const a = fakeDriverEnv();
  const d1 = await open(a);
  a.ctl.ws.handlers.close();
  assert.equal((await d1.health()).ds.open, false);
  const b = fakeDriverEnv();
  const d2 = await open(b);
  b.ctl.frameDead = true;
  assert.deepEqual(await d2.health(), { alive: false, canSave: null, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 });
  const c = fakeDriverEnv({ canSave: true });
  const d3 = await open(c);
  assert.equal((await d3.health()).canSave, true);
});

test('documentKey(): the key of the config the session was opened with; null before open', async () => {
  const d0 = drv.createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false, detail: 'none' }) });
  assert.equal(d0.documentKey(), null);
  assert.deepEqual(await d0.health(), { alive: false, canSave: null, ds: { open: false, sockets: 0, busy: false }, anomalies: null });
  const env = fakeDriverEnv();
  const d = await open(env);
  assert.equal(d.documentKey(), 'KEY1');
});
