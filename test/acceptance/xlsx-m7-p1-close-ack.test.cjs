'use strict';

const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');

const { readFrame, isDsSocket, createSaveAckObserver, awaitSaveAck, awaitForceSave } = require('../../src/drivers/eurooffice/save-ack.cjs');
const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

/**
 * M7 P1 (owner GO 2026-10-02, owner channel): the XLSX driver close waits for the DS unSaveLock acknowledgement of
 * the editor's LAST saveChanges before requestClose and the browser teardown. Offline controls: ack before close,
 * ack after close, stale ack, several saveChanges, missing ack, timeout, a session with no write; plus the owner's
 * three review points: (1) only this session's DS socket, per-connection state; (2) verified = the last send is a
 * final chunk with its own ack, sync advance only from an observed start; (3) "no save traffic" only without a
 * mutation, unreadable DS traffic never reads as "nothing saved". A verified ack is not storage proof.
 */

// socket.io text frames as the DS client sends / receives them
const frame = msg => '42' + JSON.stringify(['message', msg]);
const auth = () => frame({ type: 'auth', result: 1 });
const isSaveLock = sync => frame({ type: 'isSaveLock', syncChangesIndex: sync });
const saveLock = granted => frame({ type: 'saveLock', saveLock: !granted });
const saveChanges = (start, end) => frame({ type: 'saveChanges', changes: '[]', startSaveChanges: start, endSaveChanges: end, isExcel: true });
const savePart = sync => frame({ type: 'savePartChanges', changesIndex: sync, syncChangesIndex: sync });
const unSaveLock = (sync, time = 1) => frame({ type: 'unSaveLock', index: sync, time, syncChangesIndex: sync });
const fsStartReq = () => frame({ type: 'forceSaveStart' });
const fsStart = (code, time) => frame({ type: 'forceSaveStart', messages: { code, time } });
const fsDone = (time, success = true, type = 1) => frame({ type: 'forceSave', messages: { type, time, success } });

const IDS = { dsHost: 'ds.example', key: 'KEY1' };
const DS_URL = 'wss://ds.example/9.3.4-x/doc/42/c/?shardkey=KEY1&EIO=4&transport=websocket';

/** An observer with this session's DS socket opened and authenticated. */
function open() {
  const obs = createSaveAckObserver(IDS);
  const s = obs.onSocketOpen(DS_URL);
  s.onFrame('recv', auth());
  return { obs, s };
}
// one full, single-chunk save cycle from -> to on socket s
function cycle(s, from, to, { ack = true, time = 1 } = {}) {
  s.onFrame('sent', isSaveLock(from));
  s.onFrame('recv', saveLock(true));
  s.onFrame('sent', saveChanges(true, true));
  if (ack) s.onFrame('recv', unSaveLock(to, time));
}
const W = { mutated: true };

test('frame reader: message events are read; engine.io control is skipped; binary and unparsable text are unreadable', () => {
  assert.equal(readFrame(frame({ type: 'unSaveLock', syncChangesIndex: 3 })).msg.type, 'unSaveLock');
  assert.equal(readFrame('42' + JSON.stringify(['message', JSON.stringify({ type: 'saveLock', saveLock: false })])).msg.type, 'saveLock');
  assert.deepEqual(readFrame('2'), { skip: true });
  assert.deepEqual(readFrame('40'), { skip: true });
  assert.deepEqual(readFrame('0{"sid":"x"}'), { skip: true });
  assert.deepEqual(readFrame(Buffer.from('42[]')), { unreadable: 'binary' });
  assert.deepEqual(readFrame('42' + JSON.stringify(['other', { type: 'x' }])), { unreadable: 'text' });
  assert.deepEqual(readFrame('42[broken'), { unreadable: 'text' });
});

test('DS socket identity: DS host + /doc/ socket + this session key; anything else is not the DS connection', () => {
  assert.equal(isDsSocket(DS_URL, IDS), true);
  assert.equal(isDsSocket(DS_URL.replace('KEY1', 'OTHER'), IDS), false);
  assert.equal(isDsSocket(DS_URL.replace('ds.example', 'evil.example'), IDS), false);
  assert.equal(isDsSocket(DS_URL.replace('/doc/', '/x/'), IDS), false);
  assert.equal(isDsSocket('wss://ds.example/doc/42/c/?EIO=4', IDS), false);
  assert.equal(isDsSocket(DS_URL, { dsHost: 'ds.example', key: null }), false);
  assert.equal(isDsSocket('not a url', IDS), false);
});

test('ack BEFORE close: the last send is a final chunk with its own ack -> verified at once', () => {
  const { obs, s } = open();
  cycle(s, 0, 4004);
  assert.deepEqual(obs.verdict(false, W), { state: 'verified', reason: 'last-save-acked' });
  assert.equal(obs.evidence().lastSend.syncTo, 4004);
});

test('ack AFTER close is called: pending until the ack frame, then verified (no fixed sleep)', async () => {
  const { obs, s } = open();
  cycle(s, 0, 4004, { ack: false });
  assert.equal(obs.verdict(false, W).state, 'pending');
  let t = 0; let polls = 0;
  const r = await awaitSaveAck(obs, { readCanSave: async () => false, mutated: true, timeoutMs: 1000, pollMs: 10, now: () => t,
    wait: async ms => { t += ms; if (++polls === 3) s.onFrame('recv', unSaveLock(4004)); } });
  assert.equal(r.state, 'verified');
  assert.equal(r.waitedMs, 30);
});

test('a STALE ack never covers a later send: cycle 1 acked, cycle 2 sent without ack -> not verified', () => {
  const { obs, s } = open();
  cycle(s, 0, 10);
  cycle(s, 10, 20, { ack: false });
  assert.equal(obs.verdict(false, W).state, 'pending');
  s.onClose();
  assert.equal(obs.verdict(false, W).state, 'unverified');
});

test('an unsolicited ack (nothing outstanding) does not verify a send that comes after it', () => {
  const { obs, s } = open();
  s.onFrame('recv', unSaveLock(5));
  cycle(s, 5, 0, { ack: false });
  assert.equal(obs.verdict(false, W).state, 'pending');
});

test('several saveChanges: chunked save (part ack, then final ack) and two cycles -> verified only after the last', () => {
  const { obs, s } = open();
  s.onFrame('sent', isSaveLock(0)); s.onFrame('recv', saveLock(true));
  s.onFrame('sent', saveChanges(true, false));
  assert.equal(obs.verdict(false, W).state, 'pending');
  s.onFrame('recv', savePart(2000));
  s.onFrame('sent', saveChanges(false, true));
  assert.equal(obs.verdict(false, W).state, 'pending');
  s.onFrame('recv', unSaveLock(4004));
  assert.equal(obs.verdict(false, W).state, 'verified');
  cycle(s, 4004, 5000);
  assert.deepEqual(obs.verdict(false, W), { state: 'verified', reason: 'last-save-acked' });
});

// ---- owner review point 1: the editor <-> DS connection, per-connection state

test('R1: a socket that is not this session DS connection never feeds the state (its ack cannot verify)', () => {
  const { obs, s } = open();
  cycle(s, 0, 0, { ack: false });
  const foreign = obs.onSocketOpen(DS_URL.replace('KEY1', 'OTHER'));
  foreign.onFrame('recv', unSaveLock(4004));
  assert.equal(obs.verdict(false, W).state, 'pending');
  assert.equal(obs.evidence().foreignSockets, 1);
});

test('R1: connection change -- an ack on the NEW DS socket does not pair with the send on the OLD one', () => {
  const { obs, s } = open();
  cycle(s, 0, 0, { ack: false });
  const s2 = obs.onSocketOpen(DS_URL);
  s2.onFrame('recv', auth());
  s2.onFrame('recv', unSaveLock(4004)); // nothing outstanding on s2: no pairing
  assert.equal(obs.verdict(false, W).state, 'pending');
  s.onClose(); // the old connection is gone with its send unanswered
  assert.deepEqual(obs.verdict(false, W), { state: 'unverified', reason: 'socket-closed-while-saving' });
});

test('R1: after a reconnect, the last send is the one that must be acked -- on its own socket', () => {
  const { obs, s } = open();
  cycle(s, 0, 10);
  s.onClose();
  const s2 = obs.onSocketOpen(DS_URL);
  s2.onFrame('recv', auth());
  cycle(s2, 10, 20, { ack: false });
  assert.equal(obs.verdict(false, W).state, 'pending');
  s2.onFrame('recv', unSaveLock(20));
  assert.deepEqual(obs.verdict(false, W), { state: 'verified', reason: 'last-save-acked' });
  assert.deepEqual(obs.evidence().dsSockets.map(x => x.sends), [1, 1]);
});

test('R1: a socket seen but its DS handshake never observed is not a DS observation', () => {
  const obs = createSaveAckObserver(IDS);
  const s = obs.onSocketOpen(DS_URL);
  assert.deepEqual(obs.verdict(false, { mutated: false }), { state: 'unverified', reason: 'ds-handshake-not-observed' });
  s.onFrame('recv', frame({ type: 'auth', result: 0 }));
  assert.equal(obs.verdict(false, { mutated: false }).reason, 'ds-handshake-not-observed');
  const none = createSaveAckObserver(IDS);
  none.onSocketOpen('wss://elsewhere.example/socket');
  assert.equal(none.verdict(false, { mutated: false }).reason, 'no-ds-socket-observed');
});

// ---- owner review point 2: explicit last-send-is-an-acked-final, observed sync start

test('R2: an acked cycle, then an unfinished part cycle -> never verified (the last send is a part)', () => {
  const { obs, s } = open();
  cycle(s, 0, 10);
  s.onFrame('sent', isSaveLock(10)); s.onFrame('recv', saveLock(true));
  s.onFrame('sent', saveChanges(true, false));
  s.onFrame('recv', savePart(15));
  s.onFrame('recv', unSaveLock(15)); // an answer with nothing outstanding closes the cycle
  assert.deepEqual(obs.verdict(false, W), { state: 'unverified', reason: 'last-send-not-an-acked-final' });
});

test('R2: part chunks only, no final chunk -> never verified', () => {
  const { obs, s } = open();
  s.onFrame('sent', isSaveLock(0)); s.onFrame('recv', saveLock(true));
  s.onFrame('sent', saveChanges(true, false)); s.onFrame('recv', savePart(10));
  s.onFrame('recv', unSaveLock(10));
  assert.deepEqual(obs.verdict(false, W), { state: 'unverified', reason: 'last-send-not-an-acked-final' });
});

test('R2: the sync advance is only judged from an OBSERVED start index; a missing or invalid start is not verified', () => {
  for (const from of [undefined, null, -1, 1.5, 'x']) {
    const { obs, s } = open();
    s.onFrame('sent', frame({ type: 'isSaveLock', syncChangesIndex: from }));
    s.onFrame('recv', saveLock(true)); s.onFrame('sent', saveChanges(true, true)); s.onFrame('recv', unSaveLock(4004));
    assert.equal(obs.verdict(false, W).reason, 'sync-from-unobserved', String(from));
  }
  const noIndex = open();
  noIndex.s.onFrame('sent', isSaveLock(0)); noIndex.s.onFrame('recv', saveLock(true)); noIndex.s.onFrame('sent', saveChanges(true, true));
  noIndex.s.onFrame('recv', frame({ type: 'unSaveLock', index: -1, time: 1 }));
  assert.equal(noIndex.obs.verdict(false, W).reason, 'ack-sync-index-unpairable');
  const notAhead = open();
  cycle(notAhead.s, 50, 50);
  assert.equal(notAhead.obs.verdict(false, W).reason, 'ack-sync-index-unpairable');
});

test('R2: other unpairable shapes are never verified', () => {
  const twice = open();
  twice.s.onFrame('sent', isSaveLock(0)); twice.s.onFrame('recv', saveLock(true));
  twice.s.onFrame('sent', saveChanges(true, true)); twice.s.onFrame('sent', saveChanges(true, true));
  twice.s.onFrame('recv', unSaveLock(9));
  assert.equal(twice.obs.verdict(false, W).reason, 'save-sent-while-outstanding');
  const noLock = open();
  noLock.s.onFrame('sent', saveChanges(true, true)); noLock.s.onFrame('recv', unSaveLock(9));
  assert.equal(noLock.obs.verdict(false, W).reason, 'save-without-granted-lock');
  const partAsFinal = open();
  partAsFinal.s.onFrame('sent', isSaveLock(0)); partAsFinal.s.onFrame('recv', saveLock(true)); partAsFinal.s.onFrame('sent', saveChanges(true, true));
  partAsFinal.s.onFrame('recv', savePart(10));
  assert.equal(partAsFinal.obs.verdict(false, W).reason, 'part-ack-unpaired');
});

// ---- owner review point 3: "no save traffic" only without a mutation

test('R3: no save traffic in a session with NO mutation -> verified without waiting', async () => {
  const { obs } = open();
  let waits = 0;
  const r = await awaitSaveAck(obs, { readCanSave: async () => false, mutated: false, timeoutMs: 1000, pollMs: 10, now: () => 0, wait: async () => { waits++; } });
  assert.deepEqual({ state: r.state, reason: r.reason }, { state: 'verified', reason: 'no-save-traffic' });
  assert.equal(waits, 0);
});

test('R3: no save traffic AFTER a mutation -> never "nothing to save"; unknown mutation state likewise', () => {
  const { obs } = open();
  assert.deepEqual(obs.verdict(false, { mutated: true }), { state: 'unverified', reason: 'mutated-without-observed-save' });
  assert.deepEqual(obs.verdict(false, {}), { state: 'unverified', reason: 'mutation-unknown' });
});

test('R3: unreadable DS traffic (binary, unparsable, another event) is never read as "nothing saved"', () => {
  for (const bad of [Buffer.from([1, 2, 3]), '42[broken', '42' + JSON.stringify(['other', { type: 'saveChanges' }])]) {
    const { obs, s } = open();
    s.onFrame('sent', bad);
    assert.match(obs.verdict(false, { mutated: false }).reason, /^uninterpreted-ds-frame:/);
  }
});

// ---- FORCESAVE_CONFIRMED (owner GO, owner channel): on the P1 observer

test('FS: the right order -- acked save, then the request, its own code 0 and success with that save time -> confirmed', () => {
  const { obs, s } = open();
  cycle(s, 0, 4004, { time: 700 });
  s.onFrame('sent', fsStartReq());
  assert.deepEqual(obs.forceSaveVerdict(1), { state: 'pending', reason: 'awaiting-start' });
  s.onFrame('recv', fsStart(0, 700));
  assert.deepEqual(obs.forceSaveVerdict(1), { state: 'pending', reason: 'awaiting-completion' });
  s.onFrame('recv', fsDone(700, true));
  assert.deepEqual(obs.forceSaveVerdict(1), { state: 'confirmed', reason: 'forcesave-confirmed' });
  assert.equal(obs.verdict(false, W).state, 'verified'); // the close verdict is untouched by the forcesave
});

test('FS: a request sent before the last save is acked never confirms, whatever answers follow', () => {
  const { obs, s } = open();
  cycle(s, 0, 4004, { ack: false });
  s.onFrame('sent', fsStartReq());
  s.onFrame('recv', unSaveLock(4004, 700));
  s.onFrame('recv', fsStart(0, 700)); s.onFrame('recv', fsDone(700));
  assert.equal(obs.forceSaveVerdict(1).reason, 'request-before-acked-save');
});

test('FS: a STALE answer (an earlier save time) and a FOREIGN answer (another DS socket) never confirm', () => {
  const stale = open();
  cycle(stale.s, 0, 10, { time: 500 });
  cycle(stale.s, 10, 20, { time: 600 });
  stale.s.onFrame('sent', fsStartReq());
  stale.s.onFrame('recv', fsStart(0, 500)); // answers the OLD save time
  assert.equal(stale.obs.forceSaveVerdict(1).reason, 'start-time-not-the-acked-save');
  const late = open();
  cycle(late.s, 0, 10, { time: 600 });
  late.s.onFrame('sent', fsStartReq()); late.s.onFrame('recv', fsStart(0, 600));
  late.s.onFrame('recv', fsDone(500)); // a completion for another request's time
  assert.equal(late.obs.forceSaveVerdict(1).reason, 'awaiting-completion');
  assert.equal(late.obs.evidence().strayForceSave, 1);
  const foreign = open();
  cycle(foreign.s, 0, 10, { time: 600 });
  foreign.s.onFrame('sent', fsStartReq());
  const s2 = foreign.obs.onSocketOpen(DS_URL); s2.onFrame('recv', auth());
  s2.onFrame('recv', fsStart(0, 600)); s2.onFrame('recv', fsDone(600));
  assert.equal(foreign.obs.forceSaveVerdict(1).reason, 'awaiting-start');
});

test('FS: an error code, success:false, a non-button completion type -> never confirmed', () => {
  for (const [code, reason] of [[4, 'start-code-4'], [5, 'start-code-5'], ['x', 'start-code--1']]) {
    const { obs, s } = open();
    cycle(s, 0, 10, { time: 600 });
    s.onFrame('sent', fsStartReq()); s.onFrame('recv', fsStart(code, 600));
    assert.equal(obs.forceSaveVerdict(1).reason, reason);
  }
  const failed = open();
  cycle(failed.s, 0, 10, { time: 600 });
  failed.s.onFrame('sent', fsStartReq()); failed.s.onFrame('recv', fsStart(0, 600)); failed.s.onFrame('recv', fsDone(600, false));
  assert.equal(failed.obs.forceSaveVerdict(1).reason, 'completion-not-successful');
  const timer = open();
  cycle(timer.s, 0, 10, { time: 600 });
  timer.s.onFrame('sent', fsStartReq()); timer.s.onFrame('recv', fsStart(0, 600)); timer.s.onFrame('recv', fsDone(600, true, 2));
  assert.equal(timer.obs.forceSaveVerdict(1).reason, 'awaiting-completion');
});

test('FS: a missing completion -> unconfirmed at the fixed bound; the socket gone -> unconfirmed at once', async () => {
  const { obs, s } = open();
  cycle(s, 0, 10, { time: 600 });
  s.onFrame('sent', fsStartReq()); s.onFrame('recv', fsStart(0, 600));
  let t = 0;
  const r = await awaitForceSave(obs, 1, { timeoutMs: 30000, pollMs: 100, now: () => t, wait: async ms => { t += ms; } });
  assert.deepEqual({ state: r.state, reason: r.reason, waitedMs: r.waitedMs }, { state: 'unconfirmed', reason: 'timeout:awaiting-completion', waitedMs: 30000 });
  s.onClose();
  assert.equal(obs.forceSaveVerdict(1).state, 'unconfirmed');
});

test('FS: a NEW save cycle after the request -> the confirmation does not cover the new state', () => {
  const { obs, s } = open();
  cycle(s, 0, 10, { time: 600 });
  s.onFrame('sent', fsStartReq()); s.onFrame('recv', fsStart(0, 600)); s.onFrame('recv', fsDone(600));
  assert.equal(obs.forceSaveVerdict(1).state, 'confirmed');
  cycle(s, 10, 20, { time: 700 });
  assert.deepEqual(obs.forceSaveVerdict(1), { state: 'unconfirmed', reason: 'superseded-by-a-newer-save' });
});

// ---- the remaining controls

test('ack received, but the editor still has a change to send -> pending, never verified on the old ack', () => {
  const { obs, s } = open();
  cycle(s, 0, 4004);
  assert.deepEqual(obs.verdict(true, W), { state: 'pending', reason: 'editor-has-unsent-changes' });
  assert.equal(obs.verdict(null, W).state, 'pending');
});

test('MISSING ack + the socket gone: unverified at once, no wait for an event that cannot come', async () => {
  const { obs, s } = open();
  cycle(s, 0, 4004, { ack: false });
  s.onClose();
  let waits = 0;
  const r = await awaitSaveAck(obs, { readCanSave: async () => false, mutated: true, timeoutMs: 1000, pollMs: 10, now: () => 0, wait: async () => { waits++; } });
  assert.equal(r.state, 'unverified');
  assert.equal(waits, 0);
  assert.deepEqual(obs.evidence().anomalies, ['socket-closed-while-saving']);
});

test('TIMEOUT: the ack never comes -> unverified at the fixed bound, the reason names what was pending', async () => {
  const { obs, s } = open();
  cycle(s, 0, 4004, { ack: false });
  let t = 0;
  const r = await awaitSaveAck(obs, { readCanSave: async () => false, mutated: true, timeoutMs: 15000, pollMs: 100, now: () => t, wait: async ms => { t += ms; } });
  assert.deepEqual({ state: r.state, reason: r.reason, waitedMs: r.waitedMs }, { state: 'unverified', reason: 'timeout:save-outstanding', waitedMs: 15000 });
});

test('a granted lock released without any change (the client own unSaveLock) is no save', () => {
  const { obs, s } = open();
  s.onFrame('sent', isSaveLock(0)); s.onFrame('recv', saveLock(true)); s.onFrame('sent', frame({ type: 'unSaveLock' }));
  assert.equal(obs.verdict(false, { mutated: false }).state, 'pending');
  s.onFrame('recv', unSaveLock(0));
  assert.deepEqual(obs.verdict(false, { mutated: false }), { state: 'verified', reason: 'no-save-traffic' });
});

// ---- the driver close path, on a fake Playwright (no network beyond the driver's own loopback host)

function fakeWs(url) {
  return { url: () => url, handlers: {}, on(ev, cb) { this.handlers[ev] = cb; }, emit(ev, payload) { this.handlers[ev] && this.handlers[ev]({ payload }); } };
}

function fakeDriverEnv({ canSave = false } = {}) {
  const log = [];
  const ctl = { canSave, log, wsHandler: null };
  const frameObj = {
    url: () => 'https://ds.example/9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b/web-apps/apps/spreadsheeteditor/main/index.html',
    evaluate: async (fn, arg) => {
      if (arg && arg.body) return { ok: true, outcome: 'command-completed', result: 'Sheet1' };
      if (String(fn).includes('asc_isDocumentCanSave')) { log.push('canSave'); return { ok: true, canSave: ctl.canSave }; }
      if (String(fn).includes('e.forceSave()')) {
        log.push('forceSave');
        const sent = ctl.editorSends !== false;
        if (sent && ctl.beforeOwnRequest) ctl.beforeOwnRequest();
        if (sent) { ctl.ws.emit('framesent', fsStartReq()); if (ctl.onForceSave) ctl.onForceSave(); }
        return { ok: true, sent };
      }
      return { ok: true };
    },
  };
  const page = {
    on(ev, cb) { if (ev === 'websocket') ctl.wsHandler = cb; },
    goto: async () => {}, waitForFunction: async () => {}, waitForTimeout: async () => {},
    frames: () => [frameObj],
    evaluate: async fn => {
      if (String(fn).includes('requestClose')) {
        log.push('requestClose');
        if (ctl.onRequestClose) ctl.onRequestClose(); // late DS events WHILE requestClose runs
        if (ctl.requestCloseFails) return { ok: false, detail: 'requestClose-threw' };
      }
      return { ok: true, detail: 'requestClose-called' };
    },
  };
  const browser = { newPage: async () => page, close: async () => {
    log.push('browser-close');
    if (ctl.browserCloseRejects) throw new Error('browser.close rejected'); // the DS socket stays OPEN
    if (ctl.onTeardown) ctl.onTeardown(); // late DS events WHILE the browser closes
    if (ctl.ws && ctl.ws.handlers.close) ctl.ws.handlers.close(); // the DS socket closes with the page
  } };
  const pw = { chromium: { launch: async () => browser } };
  return { ctl, loadPlaywright: () => ({ ok: true, pw }) };
}

const CONFIG = { documentServerUrl: 'https://ds.example/', document: { key: 'KEY1' } };

const opened = [];
after(async () => { for (const d of opened) await d.close(); }); // a driver left open keeps its loopback host alive

async function openFakeDriver(env, opts = {}, { socket = DS_URL } = {}) {
  const realFetch = globalThis.fetch;
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: CONFIG } }) });
  try {
    const d = drv.createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5, closeAckTimeoutMs: 50, ...opts });
    opened.push(d);
    const r = await d.open({ fileId: '1' });
    assert.equal(r.ok, true);
    if (socket && env.ctl.wsHandler) { // the editor's DS socket opens and authenticates
      env.ctl.ws = fakeWs(socket);
      env.ctl.wsHandler(env.ctl.ws);
      env.ctl.ws.emit('framereceived', auth());
    }
    return d;
  } finally { globalThis.fetch = realFetch; }
}

const write = d => d.dispatch(null, { kind: 'callCommand', body: 'return 1;' });

test('driver: the ack is awaited BEFORE requestClose and the teardown; acked save -> ok:true', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { closeAckTimeoutMs: 2000 });
  await write(d);
  const { ws } = env.ctl;
  ws.emit('framesent', isSaveLock(0)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true));
  setTimeout(() => { env.ctl.log.push('ack'); ws.emit('framereceived', unSaveLock(4004)); }, 40);
  const r = await d.close();
  assert.equal(r.ok, true);
  assert.equal(r.saveAck.state, 'verified');
  assert.equal(r.saveAck.lastSend.syncTo, 4004);
  assert.deepEqual(env.ctl.log.filter(x => x !== 'canSave'), ['ack', 'requestClose', 'browser-close']);
  assert.match(r.detail, /save-ack verified:last-save-acked/);
});

test('driver: no ack within the bound -> ok:false (not-verified close), browser still closed, nothing rewritten', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { closeAckTimeoutMs: 60 });
  await write(d);
  const { ws } = env.ctl;
  ws.emit('framesent', isSaveLock(0)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true));
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.match(r.saveAck.reason, /^timeout:save-outstanding/);
  assert.deepEqual(env.ctl.log.filter(x => x !== 'canSave'), ['requestClose', 'browser-close']);
  assert.ok(r.saveAck.waitedMs >= 60 && r.saveAck.waitedMs < 1000);
});

test('driver R3: no write -> closes at once; a write with no observed save -> not verified', async () => {
  const quiet = fakeDriverEnv();
  const d = await openFakeDriver(quiet, { closeAckTimeoutMs: 5000 });
  const t0 = Date.now();
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [true, 'no-save-traffic']);
  assert.ok(Date.now() - t0 < 1000);
  const wrote = fakeDriverEnv();
  const d2 = await openFakeDriver(wrote, { closeAckTimeoutMs: 5000 });
  await write(d2);
  const r2 = await d2.close();
  assert.deepEqual([r2.ok, r2.saveAck.reason], [false, 'mutated-without-observed-save']);
});

test('driver R1: the identity comes from the session config -- a socket with another key is not observed', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { closeAckTimeoutMs: 5000 }, { socket: DS_URL.replace('KEY1', 'OTHER') });
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason, r.saveAck.foreignSockets], [false, 'no-ds-socket-observed', 1]);
  assert.ok(env.ctl.log.includes('browser-close'));
});

test('driver: the DS socket is never observed -> not-verified close (fail-closed), the browser is still closed', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { closeAckTimeoutMs: 5000 }, { socket: null });
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [false, 'no-ds-socket-observed']);
  assert.ok(env.ctl.log.includes('browser-close'));
});

// ---- the driver forceSave path

function ackedCycle(ws, time) {
  ws.emit('framesent', isSaveLock(0)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true));
  ws.emit('framereceived', unSaveLock(4004, time));
}

test('driver FS: waits for the ack FIRST, then requests, then waits for its own confirmation -> confirmed', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 2000, forceSaveConfirmTimeoutMs: 2000 });
  await write(d);
  const { ws } = env.ctl;
  ws.emit('framesent', isSaveLock(0)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true));
  setTimeout(() => { env.ctl.log.push('ack'); ws.emit('framereceived', unSaveLock(4004, 700)); }, 30);
  env.ctl.onForceSave = () => setTimeout(() => { ws.emit('framereceived', fsStart(0, 700)); ws.emit('framereceived', fsDone(700)); }, 20);
  const r = await d.forceSave();
  assert.deepEqual([r.ok, r.confirmed, r.forceSave.state, r.forceSave.reason], [true, true, 'confirmed', 'forcesave-confirmed']);
  const order = env.ctl.log.filter(x => x !== 'canSave');
  assert.deepEqual(order, ['ack', 'forceSave']);
  assert.equal(r.forceSave.request.saveTime, 700);
});

test('driver FS: too early is prevented -- no ack within the bound -> ok:false and NO forceSave request', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 60 });
  await write(d);
  const { ws } = env.ctl;
  ws.emit('framesent', isSaveLock(0)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true));
  const r = await d.forceSave();
  assert.deepEqual([r.ok, r.confirmed, r.forceSave.reason], [false, false, 'precondition:timeout:save-outstanding']);
  assert.equal(env.ctl.log.includes('forceSave'), false);
});

test('driver FS: code 4 / success:false / no answer -> confirmed:false, ok:true (requested), no rewrite', async () => {
  for (const [answer, reason] of [[ws => ws.emit('framereceived', fsStart(4, 700)), 'start-code-4'],
    [ws => { ws.emit('framereceived', fsStart(0, 700)); ws.emit('framereceived', fsDone(700, false)); }, 'completion-not-successful'],
    [() => {}, 'timeout:awaiting-start']]) {
    const env = fakeDriverEnv();
    const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 60 });
    await write(d);
    ackedCycle(env.ctl.ws, 700);
    env.ctl.onForceSave = () => answer(env.ctl.ws);
    const r = await d.forceSave();
    assert.deepEqual([r.ok, r.confirmed, r.forceSave.reason], [true, false, reason]);
    assert.equal(env.ctl.log.filter(x => x === 'forceSave').length, 1);
  }
});

test('driver FS: the editor declines to send the request -> unconfirmed at once', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 5000 });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.editorSends = false;
  const t0 = Date.now();
  const r = await d.forceSave();
  assert.deepEqual([r.ok, r.confirmed, r.forceSave.reason], [true, false, 'request-not-sent-by-editor']);
  assert.ok(Date.now() - t0 < 1000);
});

test('driver FS: a write-free session (the retry) is not forced to save just to get a confirmation', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000 });
  const r = await d.forceSave();
  assert.deepEqual([r.ok, r.confirmed, r.forceSave.state, r.forceSave.reason], [true, false, 'not-required', 'nothing-saved-in-session']);
  assert.equal(env.ctl.log.includes('forceSave'), false);
});

// ---- the request-identification gap (owner, owner channel): exactly ONE request per forceSave operation

const answerOk = (ws, time) => setTimeout(() => { ws.emit('framereceived', fsStart(0, time)); ws.emit('framereceived', fsDone(time)); }, 10);

test('driver FS-ID: an INTERVENING second request inside the operation -> ambiguous, never confirmed', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 500 });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.onForceSave = () => { env.ctl.ws.emit('framesent', fsStartReq()); answerOk(env.ctl.ws, 700); }; // a second request after ours
  const r = await d.forceSave();
  assert.deepEqual([r.ok, r.confirmed, r.forceSave.reason], [true, false, 'ambiguous:2-requests-in-the-operation']);
});

test('driver FS-ID: a PARALLEL request leaving just before ours, inside the window -> ambiguous, never confirmed', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 500 });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.beforeOwnRequest = () => env.ctl.ws.emit('framesent', fsStartReq()); // another source's request first
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  const r = await d.forceSave();
  assert.deepEqual([r.ok, r.confirmed, r.forceSave.reason], [true, false, 'ambiguous:2-requests-in-the-operation']);
});

test('driver FS-ID: an EARLIER request still open when the operation starts -> ambiguous (the answers carry no id)', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 500 });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.ws.emit('framesent', fsStartReq()); // open, unanswered, before our operation
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  const r = await d.forceSave();
  assert.deepEqual([r.confirmed, r.forceSave.reason], [false, 'ambiguous:earlier-request-still-open']);
});

test('driver FS-ID: an unmatched (stray) forcesave answer during the operation -> ambiguous', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 500 });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.onForceSave = () => { env.ctl.ws.emit('framereceived', fsDone(123)); answerOk(env.ctl.ws, 700); };
  const r = await d.forceSave();
  assert.deepEqual([r.confirmed, r.forceSave.reason], [false, 'ambiguous:unmatched-forcesave-answer']);
});

test('driver FS-ID: exactly one request, its own answers -> confirmed (the positive control of the four above)', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 500 });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  const r = await d.forceSave();
  assert.deepEqual([r.confirmed, r.forceSave.reason], [true, 'forcesave-confirmed']);
});

// ---- core contract point 3: the confirmation is re-evaluated at close, on the observer's state then

test('driver CLOSE: confirmed, nothing after -> the close carries the confirmation; a later send -> superseded at close', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 500, closeAckTimeoutMs: 500 });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  assert.equal((await d.forceSave()).confirmed, true);
  const clean = await d.close();
  assert.deepEqual([clean.ok, clean.confirmation], [true, { state: 'confirmed', reason: 'forcesave-confirmed' }]);

  const env2 = fakeDriverEnv();
  const d2 = await openFakeDriver(env2, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 500, closeAckTimeoutMs: 500 });
  await write(d2);
  ackedCycle(env2.ctl.ws, 700);
  env2.ctl.onForceSave = () => answerOk(env2.ctl.ws, 700);
  assert.equal((await d2.forceSave()).confirmed, true); // returned confirmed:true ...
  env2.ctl.ws.emit('framesent', isSaveLock(4004)); env2.ctl.ws.emit('framereceived', saveLock(true)); env2.ctl.ws.emit('framesent', saveChanges(true, true));
  env2.ctl.ws.emit('framereceived', unSaveLock(5000, 800)); // ... then a later save
  const later = await d2.close();
  assert.deepEqual([later.ok, later.confirmation], [true, { state: 'unconfirmed', reason: 'superseded-by-a-newer-save' }]);
});

test('driver CLOSE: an anomaly after the confirmation -> no confirmation at close', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 500, closeAckTimeoutMs: 100 });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  assert.equal((await d.forceSave()).confirmed, true);
  env.ctl.ws.emit('framesent', '42[broken');
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.deepEqual(r.confirmation, { state: 'unconfirmed', reason: 'anomaly:uninterpreted-ds-frame:text' });
});

// ---- the close-end gap (owner review of 2ab3aca, owner channel): late events DURING requestClose and the teardown

async function confirmedSession(opts = {}) {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env, { forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 500, closeAckTimeoutMs: 500, ...opts });
  await write(d);
  ackedCycle(env.ctl.ws, 700);
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700);
  assert.equal((await d.forceSave()).confirmed, true);
  return { env, d };
}
const lateSave = ws => { ws.emit('framesent', isSaveLock(4004)); ws.emit('framereceived', saveLock(true)); ws.emit('framesent', saveChanges(true, true)); };

test('CLOSE-END: a clean close (no late event) -> ok, verified, confirmed (the positive control)', async () => {
  const { d } = await confirmedSession();
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.state, r.confirmation && r.confirmation.state], [true, 'verified', 'confirmed']);
});

test('CLOSE-END: a new saveChanges DURING requestClose -> not verified, confirmation superseded', async () => {
  const { env, d } = await confirmedSession();
  env.ctl.onRequestClose = () => lateSave(env.ctl.ws);
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.deepEqual([r.saveAck.state, r.saveAck.lateDuringClose], ['unverified', 'late:socket-closed-while-saving']);
  assert.equal(r.confirmation.state, 'unconfirmed');
});

test('CLOSE-END: an anomaly DURING the teardown -> not verified, no confirmation', async () => {
  const { env, d } = await confirmedSession();
  env.ctl.onTeardown = () => env.ctl.ws.emit('framereceived', '42[broken');
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.deepEqual([r.saveAck.state, r.saveAck.lateDuringClose], ['unverified', 'late:uninterpreted-ds-frame:text']);
  assert.deepEqual(r.confirmation, { state: 'unconfirmed', reason: 'anomaly:uninterpreted-ds-frame:text' });
});

test('CLOSE-END: a new saveChanges DURING the teardown (the socket closes with it in flight) -> not verified, superseded', async () => {
  const { env, d } = await confirmedSession();
  env.ctl.onTeardown = () => lateSave(env.ctl.ws);
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.equal(r.saveAck.state, 'unverified');
  assert.equal(r.confirmation.state, 'unconfirmed');
});

test('CLOSE-END: a close failure (requestClose failed) -> ok:false and no confirmation remains', async () => {
  const { env, d } = await confirmedSession();
  env.ctl.requestCloseFails = true;
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.deepEqual(r.confirmation, { state: 'unconfirmed', reason: 'close-failed' });
});

// ---- the teardown failure (owner review of 582a6e8, owner channel): browser.close() rejects, the DS socket stays open

test('TEARDOWN: browser.close() rejects (the DS socket stays open) -> ok:false, the error in the result, no confirmation', async () => {
  const { env, d } = await confirmedSession();
  env.ctl.browserCloseRejects = true;
  const r = await d.close();
  assert.equal(r.ok, false);
  assert.match(r.detail, /teardown-failed: browser\.close rejected/);
  assert.deepEqual(r.teardown, { ok: false, detail: 'browser-close-failed: browser.close rejected' });
  assert.deepEqual(r.confirmation, { state: 'unconfirmed', reason: 'close-failed' });
  assert.ok(r.saveAck.dsSockets.some(x => x.open), 'the observed DS socket is still open');
  assert.deepEqual(env.ctl.log.filter(x => x !== 'canSave').slice(-2), ['requestClose', 'browser-close']);
});

test('TEARDOWN: a clean teardown -> ok, verified, confirmed, teardown ok (the positive control)', async () => {
  const { d } = await confirmedSession();
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.state, r.confirmation.state, r.teardown.ok], [true, 'verified', 'confirmed', true]);
});
