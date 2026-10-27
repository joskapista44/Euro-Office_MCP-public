'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');

const { editorPage, readDsOpenState, openWaitFailureDetail } = require('../../src/drivers/eurooffice/config-host.cjs');

/**
 * GitHub #77: before this, config-host.cjs's injected page only wired onDocumentReady -- a DS onError / onWarning /
 * onInfo event never touched window.__error, so a driver's open() wait (`window.__ready === true || window.__error`)
 * could sit blind for the whole 90s even when DS reported something. These tests extract the SAME inline script
 * editorPage() builds for the real product path and run it against a stub DocsAPI.DocEditor (no browser), so a
 * regression in the generated JS itself -- not just in a hand-written copy of it -- fails here.
 */

const CONFIG = { documentServerUrl: 'https://ds.example/', document: { key: 'KEY1' } };

/** Run editorPage()'s own inline <script> body against a stub window + DocsAPI, return the sandbox's window. */
function runEditorPageScript(constructorThrows = null) {
  const html = editorPage(CONFIG);
  const m = /<script>(?!src=)(window\.__ready[\s\S]*)<\/script><\/body>/.exec(html);
  assert.ok(m, 'the inline (non-src) <script> block must be present and extractable');
  const window = {};
  let capturedConfig = null;
  const DocsAPI = {
    DocEditor: function DocEditor(id, config) {
      capturedConfig = config;
      if (constructorThrows) throw constructorThrows;
      this.id = id;
      this.config = config;
    },
  };
  const sandbox = { window, DocsAPI, console };
  vm.createContext(sandbox);
  vm.runInContext(m[1], sandbox);
  // the sandbox's objects are cross-realm (a different Object.prototype) -- assert.deepStrictEqual would report
  // "same structure but not reference-equal" even on genuinely equal plain data; normalize through JSON once
  const snapshot = () => JSON.parse(JSON.stringify({ ready: sandbox.window.__ready, error: sandbox.window.__error, warnings: sandbox.window.__warnings, info: sandbox.window.__info }));
  return { window: sandbox.window, capturedConfig, snapshot };
}

test('the inline script is present, extractable, and wires all four events (not just onDocumentReady)', () => {
  const html = editorPage(CONFIG);
  assert.ok(html.includes('web-apps/apps/api/documents/api.js'), 'the api.js <script src> tag is untouched');
  const { capturedConfig, snapshot } = runEditorPageScript();
  assert.deepEqual(snapshot(), { ready: false, error: null, warnings: [], info: null }, 'initial state, before any event fires');
  for (const k of ['onDocumentReady', 'onError', 'onWarning', 'onInfo']) assert.equal(typeof capturedConfig.events[k], 'function', k);
  capturedConfig.events.onDocumentReady();
  assert.equal(snapshot().ready, true, 'onDocumentReady still flips __ready (unchanged contract)');
});

test('onError sets window.__error = { kind, code, description } -- the wait predicate (__ready || __error) is satisfied by it', () => {
  const { capturedConfig, snapshot } = runEditorPageScript();
  capturedConfig.events.onError({ data: { errorCode: -20, errorDescription: 'Error document VKey' } });
  const s = snapshot();
  assert.deepEqual(s.error, { kind: 'onError', code: -20, description: 'Error document VKey' });
  assert.equal(s.ready, false, 'onError does not fabricate readiness');
});

test('onWarning pushes to window.__warnings[] and does NOT touch __ready or __error (the wait is unaffected on purpose)', () => {
  const { capturedConfig, snapshot } = runEditorPageScript();
  capturedConfig.events.onWarning({ data: { errorCode: 1, errorDescription: 'Force save error' } });
  capturedConfig.events.onWarning({ data: { errorCode: 2, errorDescription: 'Warning 2' } });
  const s = snapshot();
  assert.deepEqual(s.warnings, [{ code: 1, description: 'Force save error' }, { code: 2, description: 'Warning 2' }]);
  assert.equal(s.ready, false);
  assert.equal(s.error, null);
});

test('onInfo sets window.__info = { mode } and does NOT touch __ready or __error', () => {
  const { capturedConfig, snapshot } = runEditorPageScript();
  capturedConfig.events.onInfo({ data: { mode: 'view' } });
  const s = snapshot();
  assert.deepEqual(s.info, { mode: 'view' });
  assert.equal(s.ready, false);
  assert.equal(s.error, null);
});

test('the constructor throwing still sets window.__error to a plain string (the pre-existing contract, unchanged)', () => {
  const { window } = runEditorPageScript(new Error('boom'));
  assert.match(window.__error, /^Error: boom/);
});

test('the injected config JSON is unaffected by the events object: document.key still reaches DocsAPI.DocEditor', () => {
  const { capturedConfig } = runEditorPageScript();
  assert.equal(capturedConfig.document.key, 'KEY1');
  assert.equal(capturedConfig.documentServerUrl, CONFIG.documentServerUrl);
});

// ---- openWaitFailureDetail: pure function, the three (+ one unreadable) shapes a timed-out wait can see ----

test('openWaitFailureDetail: an onError event names the DS code and description', () => {
  const d = openWaitFailureDetail({ error: { kind: 'onError', code: -20, description: 'Error document VKey' }, warnings: [], info: null });
  assert.match(d, /DS onError: code -20, Error document VKey/);
});

test('openWaitFailureDetail: a constructor-thrown string error is named as such, distinct from an onError event', () => {
  const d = openWaitFailureDetail({ error: 'Error: boom', warnings: [], info: null });
  assert.match(d, /the editor constructor threw: Error: boom/);
});

test('openWaitFailureDetail: no ready/error, but onWarning/onInfo fired -- named, not silently dropped', () => {
  const d = openWaitFailureDetail({ error: null, warnings: [{ code: 1, description: 'w' }], info: { mode: 'view' } });
  assert.match(d, /no-ds-ready-or-error/);
  assert.match(d, /"warnings":\[\{"code":1,"description":"w"\}\]/);
  assert.match(d, /"mode":"view"/);
});

test('openWaitFailureDetail: literally nothing happened -> no-ds-event', () => {
  assert.equal(openWaitFailureDetail({ error: null, warnings: [], info: null }), 'no-ds-event: no onDocumentReady/onError/onWarning/onInfo observed in the wait window');
});

test("openWaitFailureDetail: the page could not be read at all -> 'ds-state-unreadable', NOT the same fact as no-ds-event", () => {
  const d = openWaitFailureDetail(null);
  assert.match(d, /^ds-state-unreadable/);
  assert.doesNotMatch(d, /no-ds-event/);
});

// ---- readDsOpenState: the one evaluate() call a driver's open() makes after the wait times out ----

test('readDsOpenState reads error/warnings/info from the page in one evaluate() call', async () => {
  const page = { evaluate: async fn => fn.toString().includes('__error') ? { error: { kind: 'onError', code: -20, description: 'x' }, warnings: [{ code: 1, description: 'w' }], info: { mode: 'edit' } } : null };
  const s = await readDsOpenState(page);
  assert.deepEqual(s, { error: { kind: 'onError', code: -20, description: 'x' }, warnings: [{ code: 1, description: 'w' }], info: { mode: 'edit' } });
});

test('readDsOpenState -> null when the page itself cannot be evaluated (e.g. already closed/crashed)', async () => {
  const page = { evaluate: async () => { throw new Error('Target closed'); } };
  assert.equal(await readDsOpenState(page), null);
});
