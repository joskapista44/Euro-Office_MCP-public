'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createEuroOfficeXlsxDriver } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createEuroOfficeDocxDriver } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * GitHub #77: the open() wait (`window.__ready === true || window.__error`) could time out after 90s with NO
 * reason if DS never flips __ready and never reports anything at all -- or, when it DID report something
 * (onWarning/onInfo, which deliberately do not stop the wait), the generic message swallowed it too. These tests
 * drive open() on a FAKE Playwright (no browser, no network) whose top-level page.evaluate() answers as
 * config-host.cjs's readDsOpenState() expects -- proving the PRODUCT driver code (not a hand-written copy of its
 * logic) reports the DS's own reason.
 *
 * TWO different fakes model TWO different real shapes of waitForFunction(`__ready === true || __error`):
 *   fakeTimeoutEnv       -- the predicate never becomes true (no __ready, no __error) -> Playwright itself times
 *                           out and the wait REJECTS. Real for: no-ds-event, ds-state-unreadable, and
 *                           onWarning/onInfo-only (neither of those touches __ready/__error).
 *   fakeOnErrorStallEnv  -- onError sets __error, which the predicate's `|| window.__error` already satisfies ->
 *                           the wait RESOLVES normally, same as a real onDocumentReady would. An onError must not
 *                           fail-fast (whether onDocumentReady still arrives afterward is NEM-MERT), so the
 *                           driver falls through into the model-ready polling loop below -- which here never
 *                           finds its frame and exhausts for real, and THAT exhaustion is where the DS detail is
 *                           read back and reported.
 */

const CONFIG = { documentServerUrl: 'https://ds.example/', document: { key: 'KEY1' } };

/** A fake Playwright whose waitForFunction always times out; `dsState` is what page.evaluate() answers afterwards. */
function fakeTimeoutEnv(dsState) {
  const page = {
    on() {},
    goto: async () => {},
    waitForFunction: async () => { throw new Error('page.waitForFunction: Timeout 90000ms exceeded.'); },
    waitForTimeout: async () => {},
    frames: () => [],
    evaluate: async () => { if (dsState === 'throws') throw new Error('Target closed'); return dsState; },
  };
  const browser = { newPage: async () => page, close: async () => {} };
  const pw = { chromium: { launch: async () => browser } };
  return { loadPlaywright: () => ({ ok: true, pw }) };
}

/**
 * A fake Playwright whose waitForFunction RESOLVES (as it does for real once DS sets window.__error) and whose
 * editor/document frame never appears, so the model-ready loop below exhausts for real; `dsState` is what
 * page.evaluate() answers at that point.
 */
function fakeOnErrorStallEnv(dsState) {
  const page = {
    on() {},
    goto: async () => {},
    waitForFunction: async () => {},
    waitForTimeout: async () => {},
    frames: () => [],
    evaluate: async () => dsState,
  };
  const browser = { newPage: async () => page, close: async () => {} };
  const pw = { chromium: { launch: async () => browser } };
  return { loadPlaywright: () => ({ ok: true, pw }) };
}

async function withFakeFetch(fn) {
  const real = globalThis.fetch;
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: CONFIG } }) });
  try { return await fn(); } finally { globalThis.fetch = real; }
}

const MODEL_NOT_READY = { XLSX: 'live spreadsheet model not ready', DOCX: 'live editor model not ready' };

for (const [label, create] of [['XLSX', createEuroOfficeXlsxDriver], ['DOCX', createEuroOfficeDocxDriver]]) {
  test(`${label}: an onError does not fail-fast the ready/error wait -- OPEN_FAILED still names its code/description, from the model-ready loop's own exhaustion`, () => withFakeFetch(async () => {
    const env = fakeOnErrorStallEnv({ error: { kind: 'onError', code: -20, description: 'Error document VKey' }, warnings: [], info: null });
    const d = create({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, timeoutMs: 20, pollMs: 5 });
    const r = await d.open({ fileId: '1' });
    assert.equal(r.ok, false);
    assert.match(r.detail, new RegExp(`^${MODEL_NOT_READY[label]} \\(`), 'wrapped in the generic message, not replacing it');
    assert.match(r.detail, /DS onError: code -20, Error document VKey/);
  }));

  test(`${label}: OPEN_FAILED names onWarning/onInfo even though neither ever flips __ready/__error`, () => withFakeFetch(async () => {
    const env = fakeTimeoutEnv({ error: null, warnings: [{ code: 1, description: 'Force save error' }], info: { mode: 'view' } });
    const d = create({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5 });
    const r = await d.open({ fileId: '1' });
    assert.equal(r.ok, false);
    assert.match(r.detail, /no-ds-ready-or-error/);
    assert.match(r.detail, /Force save error/);
    assert.match(r.detail, /"mode":"view"/);
  }));

  test(`${label}: OPEN_FAILED says no-ds-event when literally nothing arrived in the whole wait window`, () => withFakeFetch(async () => {
    const env = fakeTimeoutEnv({ error: null, warnings: [], info: null });
    const d = create({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5 });
    const r = await d.open({ fileId: '1' });
    assert.equal(r.ok, false);
    assert.equal(r.detail, 'no-ds-event: no onDocumentReady/onError/onWarning/onInfo observed in the wait window');
  }));

  test(`${label}: OPEN_FAILED says ds-state-unreadable when the page cannot even be read back (distinct from no-ds-event)`, () => withFakeFetch(async () => {
    const env = fakeTimeoutEnv('throws');
    const d = create({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5 });
    const r = await d.open({ fileId: '1' });
    assert.equal(r.ok, false);
    assert.match(r.detail, /^ds-state-unreadable/);
  }));

  test(`${label}: the onError detail from the model-ready loop's exhaustion is REDACTED like every other open() detail`, () => withFakeFetch(async () => {
    let hostUrl = null;
    const page = {
      on() {},
      goto: async u => { hostUrl = u; },
      waitForFunction: async () => {},
      waitForTimeout: async () => {},
      frames: () => [],
      // the DS echoes the page URL (it carries the session token) in its own error description
      evaluate: async () => ({ error: { kind: 'onError', code: -4, description: `download failed: ${hostUrl}` }, warnings: [], info: null }),
    };
    const browser = { newPage: async () => page, close: async () => {} };
    const loadPlaywright = () => ({ ok: true, pw: { chromium: { launch: async () => browser } } });
    const d = create({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright, timeoutMs: 20, pollMs: 5 });
    const r = await d.open({ fileId: '1' });
    assert.equal(r.ok, false);
    assert.ok(hostUrl, 'the fake saw the host URL');
    assert.match(r.detail, /DS onError: code -4, download failed: /);
    assert.equal(r.detail.includes(hostUrl), false, 'the host URL (session token) must not leave open()');
    assert.match(r.detail, /\[redacted\]/);
  }));

  test(`${label}: a navigation-time failure (before the wait even starts) is UNCHANGED -- still the generic redacted message`, () => withFakeFetch(async () => {
    const page = { on() {}, goto: async () => { throw new Error('net::ERR_CONNECTION_REFUSED'); } };
    const browser = { newPage: async () => page, close: async () => {} };
    const loadPlaywright = () => ({ ok: true, pw: { chromium: { launch: async () => browser } } });
    const d = create({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright, pollMs: 5 });
    const r = await d.open({ fileId: '1' });
    assert.equal(r.ok, false);
    assert.match(r.detail, /ERR_CONNECTION_REFUSED/);
    assert.doesNotMatch(r.detail, /no-ds-event|DS onError|ds-state-unreadable/, 'a goto() failure is not reclassified as a DS-event failure');
  }));
}
