'use strict';

const { HANDSHAKE_BODY } = require('../../src/core/runtime-identity.cjs');
const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');

const drv = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { decideReuse } = require('../../src/server/session-reuse-rule.cjs');

/**
 * GH #220 1b (DOCX P0): the GH #100 U1/U3 warm-session driver surface on the REAL docx driver --
 * release() / resetForNextCall() / health() / documentKey(), mirroring xlsx-driver.cjs:1480-1529 method
 * for method, but exercised against docx-driver.cjs's OWN state shape through its OWN fake Playwright
 * (the same low-level env test/acceptance/m7-docx-save-ack-driver.test.cjs already proved close()/
 * forceSave() against -- duplicated here, NOT imported, because that env is module-local there and this
 * is a separate unit's test file, the same convention test/acceptance/gh100-u3-reuse-rule.test.cjs follows
 * on the xlsx side (its OWN driver-side section, health()/documentKey(), reusing the SHARED
 * test/fakes/fake-xlsx-session-env.cjs -- docx has no such shared module yet; extracting one is a
 * refactor of a working, unrelated file, out of this card's scope, named here rather than done silently).
 *
 * open()/close()/forceSave() themselves are UNCHANGED (m7-docx-save-ack-driver.test.cjs stays green,
 * unmodified) -- this file tests only the FOUR new methods.
 */

const frame = msg => '42' + JSON.stringify(['message', msg]);
const auth = () => frame({ type: 'auth', result: 1 });
const isSaveLock = sync => frame({ type: 'isSaveLock', syncChangesIndex: sync });
const saveLock = granted => frame({ type: 'saveLock', saveLock: !granted });
const saveChanges = (start, end) => frame({ type: 'saveChanges', changes: '[]', startSaveChanges: start, endSaveChanges: end });
const unSaveLock = (sync, time = 1) => frame({ type: 'unSaveLock', index: sync, time, syncChangesIndex: sync });
const fsStartReq = () => frame({ type: 'forceSaveStart' });
const fsStart = (code, time) => frame({ type: 'forceSaveStart', messages: { code, time } });
const fsDone = (time, success = true, type = 1) => frame({ type: 'forceSave', messages: { type, time, success } });

const DS_URL = 'wss://ds.example/9.3.4-x/doc/42/c/?shardkey=KEY1&EIO=4&transport=websocket';
const CONFIG = { documentServerUrl: 'https://ds.example/', document: { key: 'KEY1', url: 'https://ds.example/doc.docx' } };

function fakeWs(url) {
  return { url: () => url, handlers: {}, on(ev, cb) { this.handlers[ev] = cb; }, emit(ev, payload) { this.handlers[ev] && this.handlers[ev]({ payload }); } };
}

function fakeDriverEnv({ canSave = false } = {}) {
  const log = [];
  const ctl = { canSave, log, wsHandler: null };
  const frameObj = {
    url: () => 'https://ds.example/9.3.4-x/web-apps/apps/documenteditor/main/index.html',
    evaluate: async (fn, arg) => {
      if (ctl.frameDead) throw new Error('frame is dead');
      if (arg && arg.body === HANDSHAKE_BODY) { log.push('handshake'); return { ok: true, outcome: 'command-completed', result: null }; }
      if (arg && arg.body) { log.push('command'); return { ok: true, outcome: 'command-completed', result: null }; }
      const src = String(fn);
      if (src.includes('asc_isDocumentCanSave')) { log.push('canSave'); return { ok: true, canSave: ctl.canSave }; }
      if (src.includes('e.forceSave()')) {
        log.push('forceSave');
        const sent = ctl.editorSends !== false;
        if (sent) { ctl.ws.emit('framesent', fsStartReq()); if (ctl.onForceSave) ctl.onForceSave(); }
        return { ok: true, sent };
      }
      if (src.includes('true')) return true; // the aliveNow() probe (`() => true`)
      return true; // the open condition: a non-empty live projection
    },
  };
  const page = {
    on(ev, cb) { if (ev === 'websocket') ctl.wsHandler = cb; },
    goto: async () => {}, waitForFunction: async () => {}, waitForTimeout: async () => {},
    frames: () => [frameObj],
    evaluate: async fn => {
      if (String(fn).includes('requestClose')) {
        log.push('requestClose');
        if (ctl.onRequestClose) ctl.onRequestClose();
        if (ctl.requestCloseFails) return { ok: false, detail: 'requestClose-threw' };
      }
      return { ok: true, detail: 'requestClose-called' };
    },
  };
  const browser = { newPage: async () => page, close: async () => {
    log.push('browser-close');
    if (ctl.browserCloseRejects) throw new Error('browser.close rejected');
    if (ctl.onTeardown) ctl.onTeardown();
    if (ctl.ws && ctl.ws.handlers.close) ctl.ws.handlers.close();
  } };
  return { ctl, loadPlaywright: () => ({ ok: true, pw: { chromium: { launch: async () => browser } } }) };
}

const opened = [];
after(async () => { for (const d of opened) await d.close().catch(() => {}); });

async function openFakeDriver(env, opts = {}, { socket = DS_URL } = {}) {
  const realFetch = globalThis.fetch;
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: CONFIG } }) });
  try {
    const d = drv.createEuroOfficeDocxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5,
      closeAckTimeoutMs: 200, forceSaveAckTimeoutMs: 300, forceSaveConfirmTimeoutMs: 200, ...opts });
    opened.push(d);
    const r = await d.open({ fileId: '1' });
    assert.equal(r.ok, true, r.detail);
    if (socket && env.ctl.wsHandler) { env.ctl.ws = fakeWs(socket); env.ctl.wsHandler(env.ctl.ws); env.ctl.ws.emit('framereceived', auth()); }
    return d;
  } finally { globalThis.fetch = realFetch; }
}
const open = async (env, o) => openFakeDriver(env, o);

const write = d => d.dispatch(null, { kind: 'callCommand', body: 'return 1;' });
function ackedCycle(ws, from, to, time) {
  ws.emit('framesent', isSaveLock(from)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true));
  ws.emit('framereceived', unSaveLock(to, time));
}
const answerOk = (ws, time) => setTimeout(() => { ws.emit('framereceived', fsStart(0, time)); ws.emit('framereceived', fsDone(time)); }, 5);

// ---- release() ---------------------------------------------------------------------------------------------------

test('release after an acked, confirmed write: ok, verified, confirmed -- and NOTHING was closed (no requestClose, no browser-close)', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  await write(d);
  ackedCycle(env.ctl.ws, 0, 4004, 700);
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  const fs = await d.forceSave();
  assert.equal(fs.confirmed, true, fs.detail);
  const r = await d.release();
  assert.equal(r.ok, true, r.detail);
  assert.equal(r.saveAck.state, 'verified');
  assert.equal(r.confirmation.state, 'confirmed');
  assert.equal(env.ctl.log.includes('requestClose'), false, 'release never calls requestClose');
  assert.equal(env.ctl.log.includes('browser-close'), false, 'release never tears down');
  assert.equal(await d.health().then(h => h.alive), true, 'the editor still answers after release');
});

test('release with an unacked write: not ok (a failed call end), still nothing closed -- the caller must close(), never reuse', async () => {
  const env = fakeDriverEnv();
  const d = await open(env, { closeAckTimeoutMs: 60 });
  await write(d);
  env.ctl.ws.emit('framesent', isSaveLock(0)); env.ctl.ws.emit('framereceived', saveLock(true)); env.ctl.ws.emit('framesent', saveChanges(true, true));
  const r = await d.release();
  assert.equal(r.ok, false);
  assert.match(r.saveAck.reason, /^timeout:save-outstanding/);
  assert.equal(env.ctl.log.includes('browser-close'), false);
});

test('release vs close on the SAME events: the same save claim -- release only skips the teardown (the shared save-ack.cjs logic, unchanged by this unit)', async () => {
  const claims = [];
  for (const how of ['release', 'close']) {
    const env = fakeDriverEnv();
    const d = await open(env);
    await write(d);
    ackedCycle(env.ctl.ws, 0, 4004, 700);
    env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
    assert.equal((await d.forceSave()).confirmed, true);
    const r = await d[how]();
    claims.push({ ok: r.ok, ack: [r.saveAck.state, r.saveAck.reason], confirmation: r.confirmation && r.confirmation.state });
    assert.equal(env.ctl.log.includes('browser-close'), how === 'close');
  }
  assert.deepEqual(claims[0], claims[1]);
});

// ---- resetForNextCall() -------------------------------------------------------------------------------------------

test('resetForNextCall(): ONE session, TWO full write+release cycles -- each gets its own, independent confirmation (no state bleeds across the reset, GH #100 design doc rule R4)', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  let sync = 0;
  for (let call = 1; call <= 2; call += 1) {
    await write(d);
    ackedCycle(env.ctl.ws, sync, sync + 4004, 700 + sync);
    env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700 + sync);
    const fs = await d.forceSave();
    assert.deepEqual([fs.ok, fs.confirmed], [true, true], `call ${call}`);
    const r = await d.release();
    assert.deepEqual([r.ok, r.confirmation && r.confirmation.state], [true, 'confirmed'], `release ${call}`);
    assert.equal(d.resetForNextCall().ok, true, `reset ${call}`);
    sync += 4004;
  }
  assert.equal(env.ctl.log.filter(x => x === 'forceSave').length, 2, 'two independent forceSave requests, one browser');
  assert.equal(env.ctl.log.includes('browser-close'), false, 'neither call ever closed the session');
});

test('resetForNextCall(): ok:false once the page/frame are gone -- never claims a reset that has nothing to reset', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  await d.close();
  assert.equal(d.resetForNextCall().ok, false);
});

// ---- health() ------------------------------------------------------------------------------------------------------

test('health(): a healthy kept session after an acked, confirmed write call -- and its own shape feeds decideReuse to IGEN', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  await write(d);
  ackedCycle(env.ctl.ws, 0, 4004, 700);
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  assert.equal((await d.forceSave()).confirmed, true);
  const h = await d.health();
  assert.deepEqual(h, { alive: true, canSave: false, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 });
  const rel = await d.release();
  assert.equal(decideReuse({ health: h, openedKey: d.documentKey(), freshKey: 'KEY1', lastRelease: rel, requestedOperation: 'WRITE', openedOperation: 'WRITE' }).reuse, true);
});

test('health(): the DS socket closing, a dead frame and unreadable canSave are each visible -- never guessed healthy', async () => {
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

test('health(): never throws -- a driver that never opened answers the SAME fail-closed shape health() returns on xlsx', async () => {
  const d0 = drv.createEuroOfficeDocxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false, detail: 'none' }) });
  assert.deepEqual(await d0.health(), { alive: false, canSave: null, ds: { open: false, sockets: 0, busy: false }, anomalies: null });
});

// ---- documentKey() -------------------------------------------------------------------------------------------------

test('documentKey(): the key of the config the session was opened with; null before open; never invented', async () => {
  const d0 = drv.createEuroOfficeDocxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false, detail: 'none' }) });
  assert.equal(d0.documentKey(), null);
  const env = fakeDriverEnv();
  const d = await open(env);
  assert.equal(d.documentKey(), 'KEY1');
});

test('documentKey(): stable across a write + forceSave + release (release never touches the session identity)', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  const before = d.documentKey();
  await write(d);
  ackedCycle(env.ctl.ws, 0, 4004, 700);
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  await d.forceSave();
  await d.release();
  assert.equal(d.documentKey(), before);
});
