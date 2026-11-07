'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createEuroOfficeXlsxDriver } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

/**
 * The close -> product-OPEN race's OPEN_FAILED 'no-ds-event' is now retried ONCE inside open() -- a short first
 * attempt (shortOpenMs, default 15000, well above the measured normal-open ~3-3.5s and the measured
 * close-transition window's own end at 6.6-8.1s), then ONE retry with a fresh browser/page at the FULL original
 * timeoutMs if (and only if) the first attempt's failure is the EXACT no-ds-event shape. Every other failure
 * (onError, a config-source refusal, a navigation error, ...) is returned after the first attempt, unretried.
 *
 * Drives open() on a FAKE Playwright whose behaviour is picked per chromium.launch() call, in sequence -- this
 * proves the PRODUCT driver code retries (a fresh launch() per attempt), not a hand-written copy of its logic.
 */

const CONFIG = { documentServerUrl: 'https://ds.example/', document: { key: 'KEY1' } };

async function withFakeFetch(fn) {
  const real = globalThis.fetch;
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: CONFIG } }) });
  try { return await fn(); } finally { globalThis.fetch = real; }
}

/**
 * outcomes[i] picks what the i-th chromium.launch() call's page does:
 *   'no-ds-event' -- waitForFunction times out, DS reports nothing (-> openWaitFailureDetail's 'no-ds-event:' text)
 *   'onError'     -- waitForFunction resolves (DS already set __error), the editor frame never appears
 *                    (-> the model-ready loop's own wrapped "live spreadsheet model not ready (DS onError: ...)")
 *   'ok'          -- waitForFunction resolves, a spreadsheeteditor frame appears with a ready active sheet
 * Records every `timeout` seen by page.goto(), in call order, so a test can prove which attempt got which bound.
 */
function fakeSequencedEnv(outcomes) {
  const timeoutsSeen = [];
  let call = -1;
  const loadPlaywright = () => ({
    ok: true,
    pw: {
      chromium: {
        launch: async () => {
          call += 1;
          const outcome = outcomes[call];
          const frameObj = {
            url: () => 'https://ds.example/9.3.4/web-apps/apps/spreadsheeteditor/main/index.html',
            evaluate: async (fn, arg) => (arg && arg.body ? { ok: true, outcome: 'command-completed', result: 'Sheet1' } : { ok: true }),
          };
          const page = {
            on() {},
            goto: async (_u, opts) => { timeoutsSeen.push(opts && opts.timeout); },
            waitForFunction: async (_fn, _arg, opts) => {
              if (outcome === 'no-ds-event') throw new Error(`page.waitForFunction: Timeout ${opts && opts.timeout}ms exceeded.`);
              // 'onError' and 'ok' both resolve, as a real wait does once DS sets __ready or __error
            },
            waitForTimeout: async () => {},
            frames: () => (outcome === 'ok' ? [frameObj] : []),
            evaluate: async () => (outcome === 'onError'
              ? { error: { kind: 'onError', code: -20, description: 'Error document VKey' }, warnings: [], info: null }
              : { error: null, warnings: [], info: null }),
          };
          return { newPage: async () => page, close: async () => {} };
        },
      },
    },
  });
  return { loadPlaywright, timeoutsSeen, calls: () => call + 1 };
}

test('open(): first attempt no-ds-event, second attempt OK -- exactly one retry, succeeds', () => withFakeFetch(async () => {
  const env = fakeSequencedEnv(['no-ds-event', 'ok']);
  const d = createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5, shortOpenMs: 20, timeoutMs: 40, closeAckTimeoutMs: 50 });
  const r = await d.open({ fileId: '1' });
  assert.equal(r.ok, true);
  assert.equal(env.calls(), 2, 'exactly one retry -- a fresh browser per attempt');
  await d.close(); // a driver left open keeps its loopback host (a real listener) alive
}));

test('open(): first attempt no-ds-event, second attempt ALSO no-ds-event -- named "after the second attempt", no third attempt', () => withFakeFetch(async () => {
  const env = fakeSequencedEnv(['no-ds-event', 'no-ds-event']);
  const d = createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5, shortOpenMs: 20, timeoutMs: 40 });
  const r = await d.open({ fileId: '1' });
  assert.equal(r.ok, false);
  assert.equal(r.detail, 'no-ds-event: no onDocumentReady/onError/onWarning/onInfo observed in the wait window (after the second attempt)');
  assert.equal(env.calls(), 2, 'no third attempt');
}));

test('open(): first attempt onError (not the no-ds-event shape) -- NO retry at all', () => withFakeFetch(async () => {
  const env = fakeSequencedEnv(['onError', 'ok']);
  const d = createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5, shortOpenMs: 20, timeoutMs: 40 });
  const r = await d.open({ fileId: '1' });
  assert.equal(r.ok, false);
  assert.match(r.detail, /DS onError: code -20/);
  assert.equal(env.calls(), 1, 'the second outcome ("ok") must never be reached -- onError is not retried');
}));

test('open(): the retry gets the FULL original timeoutMs, never the short one', () => withFakeFetch(async () => {
  const env = fakeSequencedEnv(['no-ds-event', 'no-ds-event']);
  const d = createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5, shortOpenMs: 20, timeoutMs: 9000 });
  await d.open({ fileId: '1' });
  assert.deepEqual(env.timeoutsSeen, [20, 9000], 'first attempt: min(shortOpenMs, timeoutMs) == 20; retry: the full timeoutMs == 9000, unshortened');
}));

test('open(): with the default shortOpenMs (15000) above a small configured timeoutMs, the first attempt is not lengthened', () => withFakeFetch(async () => {
  // a caller configuring a small timeoutMs (e.g. a fast test elsewhere in this suite) must not have its first
  // attempt silently stretched to the 15s default -- min(shortOpenMs, timeoutMs) protects that
  const env = fakeSequencedEnv(['no-ds-event', 'ok']);
  const d = createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5, timeoutMs: 20, closeAckTimeoutMs: 50 });
  await d.open({ fileId: '1' });
  assert.deepEqual(env.timeoutsSeen, [20, 20], 'both attempts bounded by the small configured timeoutMs');
  await d.close(); // a driver left open keeps its loopback host (a real listener) alive
}));
