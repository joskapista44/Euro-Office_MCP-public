'use strict';

const { HANDSHAKE_BODY } = require('../../src/core/runtime-identity.cjs');
const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');

const drv = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * M7 DOCX save confirmation, variant A (owner GO 2026-10-03, owner channel): the DOCX driver gets the XLSX driver's
 * M7 save sequence, reusing save-ack.cjs AS IS. Offline controls, on a fake Playwright:
 *   - the observer from the first frame, keyed on the session config (DS host + document key);
 *   - forceSave: the last save acked FIRST, then ONE request, then its own confirmation; a write-free session (the
 *     retry) starts no forceSave;
 *   - close: the ack awaited BEFORE requestClose; the FINAL save-ack and confirmation evaluated AFTER requestClose and
 *     the teardown; a late send, a late anomaly or a failed close leaves no confirmation standing;
 *   - the teardown failure REPORTED, never swallowed.
 * The LIVE DOCX probe (docs/evidence/M7-DOCX-SAVE-ACK-PROBE.md) measured the same frames on the real editor.
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
      // the open's runtime handshake (core/runtime-identity.cjs) is not a write command: logged on its own
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

const write = d => d.dispatch(null, { kind: 'callCommand', body: 'return 1;' });
function ackedCycle(ws, time, from = 0, to = 300) {
  ws.emit('framesent', isSaveLock(from)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true));
  ws.emit('framereceived', unSaveLock(to, time));
}
const answerOk = (ws, time) => setTimeout(() => { ws.emit('framereceived', fsStart(0, time)); ws.emit('framereceived', fsDone(time)); }, 5);
const lateSave = ws => { ws.emit('framesent', isSaveLock(300)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true)); };

async function confirmedSession(opts = {}) {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, opts);
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  const fs = await d.forceSave();
  assert.equal(fs.confirmed, true, fs.detail);
  return { env, d };
}

// ---- forceSave ----------------------------------------------------------------------------------------------------

test('forceSave, positive control: the acked save, then ONE request, its own code 0 + success with that save time -> confirmed', async () => {
  const { env } = await confirmedSession();
  assert.equal(env.ctl.log.filter(x => x === 'forceSave').length, 1);
  assert.ok(env.ctl.log.indexOf('command') < env.ctl.log.indexOf('forceSave'));
});

test('forceSave: no ack within the bound -> ok:false and NO request (never too early)', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  await write(d);
  env.ctl.ws.emit('framesent', isSaveLock(0)); env.ctl.ws.emit('framereceived', saveLock(true)); env.ctl.ws.emit('framesent', saveChanges(true, true));
  const r = await d.forceSave();
  assert.deepEqual([r.ok, r.confirmed], [false, false]);
  assert.match(r.detail, /precondition:timeout:save-outstanding/);
  assert.equal(env.ctl.log.includes('forceSave'), false);
});

test('forceSave: code 4 / success:false / no answer -> confirmed:false; nothing is re-requested', async () => {
  for (const answer of [ws => ws.emit('framereceived', fsStart(4, 700)), ws => { ws.emit('framereceived', fsStart(0, 700)); ws.emit('framereceived', fsDone(700, false)); }, () => {}]) {
    const env = fakeDriverEnv();
    const d = await openFakeDriver(env);
    await write(d);
    ackedCycle(env.ctl.ws, 700);
    env.ctl.onForceSave = () => answer(env.ctl.ws);
    const r = await d.forceSave();
    assert.deepEqual([r.ok, r.confirmed], [true, false], r.detail);
    assert.equal(env.ctl.log.filter(x => x === 'forceSave').length, 1);
  }
});

test('forceSave: a write-free session (the retry) starts NO forceSave -- not-required, no request', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  const r = await d.forceSave();
  assert.deepEqual([r.ok, r.confirmed, r.forceSave.state], [true, false, 'not-required']);
  assert.equal(env.ctl.log.includes('forceSave'), false);
  const c = await d.close();
  assert.deepEqual([c.ok, c.saveAck.state, c.saveAck.reason], [true, 'verified', 'no-save-traffic']);
});

test('forceSave: the editor declines to send the request -> unconfirmed at once', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.editorSends = false;
  const r = await d.forceSave();
  assert.deepEqual([r.ok, r.confirmed, r.forceSave.reason], [true, false, 'request-not-sent-by-editor']);
});

// ---- close ---------------------------------------------------------------------------------------------------------

test('close, positive control: confirmed, clean requestClose + teardown -> ok, verified, confirmed, teardown ok', async () => {
  const { env, d } = await confirmedSession();
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.state, r.confirmation && r.confirmation.state, r.teardown.ok], [true, 'verified', 'confirmed', true]);
  assert.deepEqual(env.ctl.log.filter(x => x !== 'canSave').slice(-2), ['requestClose', 'browser-close']);
});

test('close: the ack is awaited BEFORE requestClose -- an ack arriving later still verifies', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { closeAckTimeoutMs: 1000 });
  await write(d);
  const { ws } = env.ctl;
  ws.emit('framesent', isSaveLock(0)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true));
  setTimeout(() => { env.ctl.log.push('ack'); ws.emit('framereceived', unSaveLock(300, 700)); }, 30);
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.state], [true, 'verified']);
  assert.ok(env.ctl.log.indexOf('ack') < env.ctl.log.indexOf('requestClose'));
});

test('close: no ack within the bound -> ok:false (not-verified close), the browser still closed', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  await write(d);
  env.ctl.ws.emit('framesent', isSaveLock(0)); env.ctl.ws.emit('framereceived', saveLock(true)); env.ctl.ws.emit('framesent', saveChanges(true, true));
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.match(r.saveAck.reason, /save-outstanding/);
  assert.ok(env.ctl.log.includes('browser-close'));
});

test('close: a write with no observed save -> never "nothing to save" (mutated-without-observed-save)', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  await write(d);
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [false, 'mutated-without-observed-save']);
});

test('close: a socket with ANOTHER key is not this session -> not-verified close (fail-closed)', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, {}, { socket: DS_URL.replace('KEY1', 'OTHER') });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [false, 'no-ds-socket-observed']);
});

test('close-end: a new saveChanges DURING requestClose -> not verified, the confirmation superseded', async () => {
  const { env, d } = await confirmedSession();
  env.ctl.onRequestClose = () => lateSave(env.ctl.ws);
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.equal(r.saveAck.state, 'unverified');
  assert.equal(r.confirmation.state, 'unconfirmed');
});

test('close-end: an anomaly DURING the teardown -> not verified, no confirmation', async () => {
  const { env, d } = await confirmedSession();
  env.ctl.onTeardown = () => env.ctl.ws.emit('framereceived', '42[broken');
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.deepEqual(r.confirmation, { state: 'unconfirmed', reason: 'anomaly:uninterpreted-ds-frame:text' });
});

test('close-end: requestClose fails -> ok:false and no confirmation remains', async () => {
  const { env, d } = await confirmedSession();
  env.ctl.requestCloseFails = true;
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.deepEqual(r.confirmation, { state: 'unconfirmed', reason: 'close-failed' });
});

test('teardown: browser.close() rejects -> REPORTED (ok:false, the error in the result), no confirmation', async () => {
  const { env, d } = await confirmedSession();
  env.ctl.browserCloseRejects = true;
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.match(r.detail, /teardown-failed: browser\.close rejected/);
  assert.deepEqual(r.teardown, { ok: false, detail: 'browser-close-failed: browser.close rejected' });
  assert.deepEqual(r.confirmation, { state: 'unconfirmed', reason: 'close-failed' });
});

// ---- the barrier reads the driver's confirmation (the core path, unchanged) --------------------------------------

test('the barrier shape: forceSave returns { ok, confirmed, forceSave } -- what core/persistence/barrier.cjs reads', async () => {
  const { d } = await confirmedSession();
  const fs = await d.forceSave(); // a second operation in the same session: nothing new was saved
  assert.equal(typeof fs.confirmed, 'boolean');
  assert.ok(fs.forceSave && typeof fs.forceSave.state === 'string');
});

// ---- the DOCX save bound (owner GO 2026-10-03, after the 7a1d5ea review): the last-ack precondition waits up to
// 120 s by DEFAULT (a 67-part / ~111 MB save acked ~84 s after the change), condition-based; the confirmation and the
// close bounds stay 30 s. A simulated clock (node:test mock timers) drives the waits -- no real 45 s sleep.

const { mock } = test;

test('the DOCX defaults: last-ack precondition 120 s; confirmation 30 s; close 30 s', () => {
  assert.deepEqual(drv.DOCX_SAVE_BOUNDS, { forceSaveAckTimeoutMs: 120000, forceSaveConfirmTimeoutMs: 30000, closeAckTimeoutMs: 30000 });
});

/** Opens a driver with the DEFAULT bounds, makes a write and a final send without its ack yet; returns the env. */
async function slowSaveSession() {
  const env = fakeDriverEnv();
  const realFetch = globalThis.fetch;
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: CONFIG } }) });
  let d;
  try {
    d = drv.createEuroOfficeDocxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 100 });
    opened.push(d);
    assert.equal((await d.open({ fileId: '1' })).ok, true);
  } finally { globalThis.fetch = realFetch; }
  env.ctl.ws = fakeWs(DS_URL); env.ctl.wsHandler(env.ctl.ws); env.ctl.ws.emit('framereceived', auth());
  await write(d);
  env.ctl.ws.emit('framesent', isSaveLock(0)); env.ctl.ws.emit('framereceived', saveLock(true)); env.ctl.ws.emit('framesent', saveChanges(true, true));
  return { env, d };
}

/** Advances the simulated clock in poll-sized steps, letting the driver's promise chain run between them. */
async function advance(ms, step = 100) {
  for (let t = 0; t < ms; t += step) { mock.timers.tick(step); await new Promise(r => setImmediate(r)); }
}

test('late but in bound: the own final ack arrives at 45 s (beyond the old 30 s) -> the request goes, confirmed, ok:true', async () => {
  const { env, d } = await slowSaveSession();
  mock.timers.enable({ apis: ['setTimeout', 'Date'], now: Date.now() });
  try {
    const p = d.forceSave();
    await advance(45000);
    assert.equal(env.ctl.log.includes('forceSave'), false, 'no request before the last ack');
    env.ctl.onForceSave = () => { env.ctl.ws.emit('framereceived', fsStart(0, 700)); env.ctl.ws.emit('framereceived', fsDone(700)); };
    env.ctl.ws.emit('framereceived', unSaveLock(300, 700));
    await advance(1000);
    const r = await p;
    assert.deepEqual([r.ok, r.confirmed], [true, true], r.detail);
    assert.equal(env.ctl.log.filter(x => x === 'forceSave').length, 1);
  } finally { mock.timers.reset(); }
});

test('beyond the bound: no ack within 120 s -> ok:false (precondition timeout), NO forceSave request; a later ack changes nothing', async () => {
  const { env, d } = await slowSaveSession();
  mock.timers.enable({ apis: ['setTimeout', 'Date'], now: Date.now() });
  try {
    const p = d.forceSave();
    await advance(121000);
    const r = await p;
    assert.deepEqual([r.ok, r.confirmed], [false, false]);
    assert.match(r.detail, /precondition:timeout:save-outstanding/);
    env.ctl.ws.emit('framereceived', unSaveLock(300, 700)); // the ack after the bound
    assert.equal(env.ctl.log.includes('forceSave'), false, 'no request, ever');
  } finally { mock.timers.reset(); }
});

test('a FOREIGN ack (another DS socket) inside the bound never satisfies the precondition -> ok:false, no request', async () => {
  const { env, d } = await slowSaveSession();
  const other = fakeWs(DS_URL.replace('KEY1', 'OTHER'));
  env.ctl.wsHandler(other); other.emit('framereceived', auth());
  mock.timers.enable({ apis: ['setTimeout', 'Date'], now: Date.now() });
  try {
    const p = d.forceSave();
    await advance(10000);
    other.emit('framereceived', unSaveLock(300, 700)); // the right shape, the wrong connection
    await advance(111000);
    const r = await p;
    assert.deepEqual([r.ok, r.confirmed], [false, false]);
    assert.equal(env.ctl.log.includes('forceSave'), false);
  } finally { mock.timers.reset(); }
});
