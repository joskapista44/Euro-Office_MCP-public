'use strict';

/**
 * A fake Playwright + DS socket for the XLSX driver's session paths (open / forceSave / close / release), for tests that need a driver
 * that stays open across several operations. The same stand-ins the P1 close-ack tests use, kept in one importable place; those tests
 * keep their own copy (they are not touched).
 */

const frame = msg => '42' + JSON.stringify(['message', msg]);
const F = {
  auth: () => frame({ type: 'auth', result: 1 }),
  isSaveLock: sync => frame({ type: 'isSaveLock', syncChangesIndex: sync }),
  saveLock: granted => frame({ type: 'saveLock', saveLock: !granted }),
  saveChanges: (start, end) => frame({ type: 'saveChanges', changes: '[]', startSaveChanges: start, endSaveChanges: end, isExcel: true }),
  unSaveLock: (sync, time = 1) => frame({ type: 'unSaveLock', index: sync, time, syncChangesIndex: sync }),
  fsStartReq: () => frame({ type: 'forceSaveStart' }),
  fsStart: (code, time) => frame({ type: 'forceSaveStart', messages: { code, time } }),
  fsDone: (time, success = true, type = 1) => frame({ type: 'forceSave', messages: { type, time, success } }),
};

const DS_URL = 'wss://ds.example/9.3.4-x/doc/42/c/?shardkey=KEY1&EIO=4&transport=websocket';
const CONFIG = { documentServerUrl: 'https://ds.example/', document: { key: 'KEY1' } };

function fakeWs(url) {
  return { url: () => url, handlers: {}, on(ev, cb) { this.handlers[ev] = cb; }, emit(ev, payload) { this.handlers[ev] && this.handlers[ev]({ payload }); } };
}

function fakeDriverEnv({ canSave = false } = {}) {
  const log = [];
  const ctl = { canSave, log, wsHandler: null, frameDead: false };
  const frameObj = {
    url: () => 'https://ds.example/9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b/web-apps/apps/spreadsheeteditor/main/index.html',
    evaluate: async (fn, arg) => {
      if (ctl.frameDead) throw new Error('frame was detached');
      if (arg && arg.body) return { ok: true, outcome: 'command-completed', result: 'Sheet1' };
      if (String(fn).includes('asc_isDocumentCanSave')) { log.push('canSave'); return { ok: true, canSave: ctl.canSave }; }
      if (String(fn).includes('e.forceSave()')) {
        log.push('forceSave');
        ctl.ws.emit('framesent', F.fsStartReq());
        if (ctl.onForceSave) ctl.onForceSave();
        return { ok: true, sent: true };
      }
      return { ok: true };
    },
  };
  const page = {
    on(ev, cb) { if (ev === 'websocket') ctl.wsHandler = cb; },
    goto: async () => {}, waitForFunction: async () => {}, waitForTimeout: async () => {},
    frames: () => [frameObj],
    evaluate: async fn => { if (String(fn).includes('requestClose')) log.push('requestClose'); return { ok: true, detail: 'requestClose-called' }; },
  };
  const browser = { newPage: async () => page, close: async () => { log.push('browser-close'); if (ctl.ws && ctl.ws.handlers.close) ctl.ws.handlers.close(); } };
  const pw = { chromium: { launch: async () => { log.push('launch'); return browser; } } };
  return { ctl, loadPlaywright: () => ({ ok: true, pw }) };
}

/** Opens a driver on the fake env with the editor's DS socket opened and authenticated. Restores fetch. */
async function openFakeDriver(drv, env, opts = {}) {
  const realFetch = globalThis.fetch;
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: CONFIG } }) });
  try {
    const d = drv.createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5, closeAckTimeoutMs: 50, forceSaveAckTimeoutMs: 1000, forceSaveConfirmTimeoutMs: 1000, ...opts });
    const r = await d.open({ fileId: '1' });
    if (!r.ok) throw new Error(`fake open failed: ${r.detail}`);
    env.ctl.ws = fakeWs(DS_URL);
    env.ctl.wsHandler(env.ctl.ws);
    env.ctl.ws.emit('framereceived', F.auth());
    return d;
  } finally { globalThis.fetch = realFetch; }
}

/** One full single-chunk save cycle from -> to on the session's socket, acked (or not). */
function saveCycle(ws, from, to, { time = 700, ack = true } = {}) {
  ws.emit('framesent', F.isSaveLock(from)); ws.emit('framereceived', F.saveLock(true)); ws.emit('framesent', F.saveChanges(true, true));
  if (ack) ws.emit('framereceived', F.unSaveLock(to, time));
}

const answerOk = (ws, time) => setTimeout(() => { ws.emit('framereceived', F.fsStart(0, time)); ws.emit('framereceived', F.fsDone(time)); }, 10);
const write = d => d.dispatch(null, { kind: 'callCommand', body: 'return 1;' });

module.exports = { F, DS_URL, CONFIG, fakeWs, fakeDriverEnv, openFakeDriver, saveCycle, answerOk, write };
