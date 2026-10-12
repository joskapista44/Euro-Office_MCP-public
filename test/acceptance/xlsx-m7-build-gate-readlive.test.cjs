'use strict';

const test = require('node:test');
const { after } = require('node:test');
const assert = require('node:assert/strict');

const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { installBridge, BRIDGE_RUNTIME_ID } = require('../fakes/ds-bridge-runtime.cjs');

/**
 * M7 D-B build gate through the REAL driver (owner 2026-10-04, owner channel): the two internal reads -- the shape
 * text's Drawing.getDocContent() and the hydration's WorkbookView.getWorksheet() -- are switched on ONLY when readLive
 * finds the session's DS runtime identity in their lists. Since owner decision (B) (internal tracker) that identity
 * is the served bridge manifest (documentServerVersion | patchVersion) read by the open's handshake -- NOT the editor
 * URL's cache key. The driver is opened on a fake Playwright whose editor frame names a cache key in its URL and runs
 * every command body against an in-memory Api carrying the REAL served bridge block (test/fakes/ds-bridge-runtime.cjs):
 * on an unaccepted runtime neither internal call happens and the structural step stops before the write; on the
 * accepted one both run (the positive control) -- whatever the cache key.
 */

const KEY_A = '9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b';
const KEY_B = '9.3.4-94bc577ecccd9583be1b130bd5a6b407';
const ACCEPTED = 'euro-office-bridge-v1/9.3.4.60/5';
const UNKNOWN = 'euro-office-bridge-v1/9.3.4.60/2';
const CONFIG = { documentServerUrl: 'https://ds.example/', document: { key: 'KEY1' } };

/** One sheet 'S' (A1:A20 'r1'..'r20') with one shape; the calls into the two internal getters are logged. */
function editorApi(log, patchVersion) {
  const lane = Array.from({ length: 20 }, (_, i) => `r${i + 1}`);
  const cell = a => {
    const r = Number(a.slice(1));
    const v = a[0] === 'A' && r >= 1 && r <= 20 ? lane[r - 1] : '';
    return { GetAddress: () => a, GetValue: () => v, GetFormula: () => v, MergeArea: { GetAddress: () => a }, GetNumberFormat: () => 'General', GetFillColor: () => null };
  };
  const range = a1 => ({ GetAddress: () => a1, GetValue: () => cell(a1.split(':')[0]).GetValue(), ForEach: fn => { const m = /^A(\d+):A(\d+)$/.exec(a1); if (m) for (let r = Number(m[1]); r <= Number(m[2]); r += 1) fn(cell(`A${r}`)); else fn(cell(a1)); } });
  const ws = { GetName: () => 'S', GetIndex: () => 0, GetUsedRange: () => range('A1:A20'), GetRange: a1 => range(a1) };
  ws.worksheet = { getIndex: () => 0, workbook: { oApi: { wb: { getWorksheet: i => { log.push(`getWorksheet(${i})`); return {}; } } } } };
  const shape = {
    GetClassType: () => 'shape', GetName: () => 'SHP', GetParentSheet: () => ws, GetPosition: () => ({ fromCol: 7, colOffset: 0, fromRow: 2, rowOffset: 0 }),
    GetWidth: () => 914400, GetHeight: () => 457200, GetRotation: () => 0, GetFlipH: () => false, GetFlipV: () => false,
    GetContent: () => { log.push('GetContent'); return null; },
    Drawing: { getDocContent: () => { log.push('getDocContent'); return null; }, drawingBase: { Type: 1, worksheet: {}, boundsFromTo: { to: { col: 8, colOff: 0, row: 4, rowOff: 0 } } } },
  };
  ws.GetAllDrawings = () => [shape];
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'S' ? ws : null), GetActiveSheet: () => ws, WorksheetFunction: { COUNT: () => 0 }, private_CreateApiDocContent: () => ({ GetElementsCount: () => 0, GetElement: () => null }) };
  installBridge(api);
  const served = api.GetEuroOfficeBridge;
  // the served bundle's own patch identity (a different patch revision = an unaccepted runtime)
  api.GetEuroOfficeBridge = () => { const b = served(); const m = b.manifest(); return { ...b, manifest: () => ({ ...m, patchVersion }) }; };
  return api;
}

function fakeEnv({ cacheKey, patchVersion }, log) {
  const Api = editorApi(log, patchVersion);
  const frame = {
    url: () => `https://ds.example/${cacheKey}/web-apps/apps/spreadsheeteditor/main/index.html`,
    evaluate: async (fn, arg) => {
      if (arg && typeof arg.body === 'string') {
        try { return { ok: true, outcome: 'command-completed', result: new Function('Api', arg.body)(Api) ?? null }; } catch (err) { return { ok: false, outcome: 'call-threw', error: String(err) }; }
      }
      if (String(fn).includes('asc_isDocumentCanSave')) return { ok: true, canSave: false };
      return { ok: true };
    },
  };
  const page = { on() {}, goto: async () => {}, waitForFunction: async () => {}, waitForTimeout: async () => {}, frames: () => [frame], evaluate: async () => ({ ok: true }) };
  const browser = { newPage: async () => page, close: async () => {} };
  return () => ({ ok: true, pw: { chromium: { launch: async () => browser } } });
}

const drivers = [];
after(async () => { for (const d of drivers) await d.close().catch(() => null); });

async function withFetch(fn) {
  const real = globalThis.fetch;
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: CONFIG } }) });
  try { return await fn(); } finally { globalThis.fetch = real; }
}

async function liveRead(env) {
  const log = [];
  const d = drv.createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: fakeEnv(env, log), pollMs: 5, closeAckTimeoutMs: 50 });
  drivers.push(d);
  return withFetch(async () => {
    assert.equal((await d.open({ fileId: '1' })).ok, true);
    const r = await d.readLive(null, { windows: [], drawingSheets: ['*'], drawingAnchor: true });
    return { r, log, shape: r.ok ? r.payload.drawingReads.S[0] : null };
  });
}

test('the served bridge manifest is the accepted runtime (the precondition of the positive control)', () => {
  assert.equal(BRIDGE_RUNTIME_ID, `9.3.4|${ACCEPTED}`);
});

for (const cacheKey of [KEY_A, KEY_B]) {
  test(`the ACCEPTED runtime (positive control, cache key ${cacheKey.slice(6, 14)}): readLive runs getWorksheet before the drawing read and getDocContent for the text; GetContent never`, async () => {
    const { r, log, shape } = await liveRead({ cacheKey, patchVersion: ACCEPTED });
    assert.equal(r.ok, true);
    assert.equal(r.payload.runtime.identity, `9.3.4|${ACCEPTED}`);
    assert.equal(r.payload.runtime.cacheKey, cacheKey, 'the cache key is recorded -- for diagnostics only');
    assert.deepEqual(log, ['getWorksheet(0)', 'getDocContent']);
    assert.equal(shape.text, '', 'the measured null: no text body');
    assert.deepEqual(shape.anchor, { type: 1, to: { col: 8, colOff: 0, row: 4, rowOff: 0 } });
  });
}

test('an UNACCEPTED runtime (same cache key): neither internal call runs; the text and the anchor are refused by the bridge identity gate', async () => {
  const { r, log } = await liveRead({ cacheKey: KEY_A, patchVersion: UNKNOWN });
  assert.equal(r.ok, true);
  assert.equal(r.payload.runtime.identity, null);
  assert.deepEqual(log, [], 'no getWorksheet, no getDocContent, no GetContent');
  // B15 /5: the hydration is a bridge step now, so the FIRST gate to refuse is the hydration client's own identity gate --
  // the whole sheet is UNOBSERVABLE before any drawing getter (before the cutover: per drawing, from the READ gate)
  assert.equal(r.payload.drawingReads.S, 'UNOBSERVABLE:hydration-PATCH_VERSION_MISMATCH');
});

test('an UNACCEPTED runtime: the structural step stops BEFORE the write, the internal getters untouched', async () => {
  const log = [];
  const task = { id: 'gate', steps: [{ intent: 'insert_rows', sheet: 'S', at: 15, count: 1, anchor: { column: 'A', value: 'r15' } }] };
  const d = drv.createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: fakeEnv({ cacheKey: KEY_A, patchVersion: UNKNOWN }, log), pollMs: 5, closeAckTimeoutMs: 50 });
  let dispatched = 0;
  const dispatch = d.dispatch.bind(d);
  d.dispatch = async (...a) => { dispatched += 1; return dispatch(...a); };
  const result = await withFetch(() => withSession({ driver: d, fileId: '1' }, xlsxAdapter, createExecutor(task, { observation: { live: liveScope(task) } })));
  assert.equal(result.ok, false);
  assert.equal(dispatched, 0, 'nothing dispatched');
  assert.equal(result.metrics.writes, 0);
  assert.deepEqual(log, []);
});
