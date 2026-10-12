'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { readQuery } = drv;
const { installBridge, withPatch, BRIDGE_RUNTIME_ID } = require('../fakes/ds-bridge-runtime.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 D-B hydration (owner GO 2026-10-04, owner channel; docs/evidence/M7-DB8-PROBE.md on the experiments branch):
 * before the structure family's anchor read, every drawing sheet's view is created -- since the B15 /5 cutover through
 * the DS Bridge sheet.model.ensureReady/v1 (RUNTIME_HYDRATION, the REAL served block's hydrate(), which runs the
 * editor's own view creator) -- so a NON-ACTIVE sheet's drawings get the editor's own geometry. A missing or
 * throwing step, or a drawing the view's init did not reach, is UNOBSERVABLE and refuses the structural step before
 * the write. The editor body runs here against a stand-in (shape only -- the LIVE acceptance proves the editor).
 * DS Bridge cutover (internal tracker point 8): the anchor itself is read through the REAL served bridge block
 * (drawing.anchor.read/v1, test/fakes/ds-bridge-runtime.cjs); the hydration is a bridge step before it, no direct fallback.
 */

const BUILD = '9.3.4|euro-office-bridge-v1/9.3.4.60/5';

/** A one-sheet stand-in: `wb` is the view container the hydration reaches through ws.worksheet.workbook.oApi.wb. */
function standIn({ wb, worksheetSet = true, log = [] } = {}) {
  const ws = { GetName: () => 'D', GetIndex: () => 0, GetUsedRange: () => null };
  const base = { Type: 1, boundsFromTo: { to: { col: 4, colOff: 1, row: 8, rowOff: 2 } }, ...(worksheetSet ? { worksheet: {} } : {}) };
  const getter = (name, v) => () => { log.push(name); return v; };
  const d = { GetClassType: getter('GetClassType', 'image'), GetName: () => 'I', GetParentSheet: () => ws, GetPosition: () => ({ fromCol: 2, colOffset: 0, fromRow: 3, rowOffset: 0 }), GetWidth: getter('GetWidth', 1), GetHeight: () => 1, GetRotation: () => 0, GetFlipH: () => null, GetFlipV: () => null, Drawing: { drawingBase: base } };
  ws.GetAllDrawings = () => { log.push('GetAllDrawings'); return [d]; };
  ws.worksheet = { getIndex: () => 0, workbook: { oApi: { wb } } };
  const api = { GetSheets: () => [ws], GetSheet: () => ws, WorksheetFunction: { COUNT: () => 0 } };
  installBridge(api);
  return api;
}
const read = (api, opts = { drawingSheets: ['*'], drawingAnchor: true, hydrateSheets: true }) => new Function('Api', readQuery([], opts))(api).drawingReads.D;

test('the hydration is generated only with hydrateSheets; without it the anchor query is byte-identical', () => {
  assert.equal(readQuery([], { drawingSheets: ['*'], drawingAnchor: true }), readQuery([], { drawingSheets: ['*'], drawingAnchor: true, hydrateSheets: false }));
  assert.doesNotMatch(readQuery([], { drawingSheets: ['*'], drawingAnchor: true }), /getWorksheet|geometry-not-hydrated/);
  const hq = readQuery([], { drawingSheets: ['*'], drawingAnchor: true, hydrateSheets: true });
  assert.match(hq, /__eoBridgeHydrate\("sheet\.model\.ensureReady\/v1", 1, /);
  assert.doesNotMatch(hq, /getWorksheet|\.worksheet\b|oApi/, 'B15 cutover: no direct hydration, so no fallback can hide in the query');
  assert.equal('EO_INTERNAL_SHEET_HYDRATION_RUNTIMES' in drv, false, 'MIGRATED to the DS Bridge (B15 /5)');
  assert.equal(BRIDGE_RUNTIME_ID, BUILD, 'the bridge manifest names the same runtime');
  for (const gone of ['EO_INTERNAL_DRAWING_ANCHOR_BUILDS', 'EO_INTERNAL_SHAPE_TEXT_BUILDS']) assert.equal(gone in drv, false, `${gone}: MIGRATED to the DS Bridge`);
});

test('the view is created BEFORE any drawing getter, by the editor\'s own getWorksheet(index); then the anchor is read', () => {
  const log = [];
  const wb = { getWorksheet: i => { log.push(`getWorksheet(${i})`); return {}; } };
  const out = read(standIn({ wb, log }));
  assert.deepEqual(log.slice(0, 3), ['getWorksheet(0)', 'GetAllDrawings', 'GetClassType']);
  assert.ok(log.indexOf('getWorksheet(0)') < log.indexOf('GetWidth'));
  assert.deepEqual(out[0].anchor, { type: 1, to: { col: 4, colOff: 1, row: 8, rowOff: 2 } });
});

test('UNOBSERVABLE, never a fallback: getWorksheet missing / throwing / no view, the model unreachable, a drawing the init did not reach', () => {
  assert.equal(read(standIn({ wb: {} })), 'UNOBSERVABLE:hydration-BRIDGE:view-creator-missing');
  assert.equal(read(standIn({ wb: null })), 'UNOBSERVABLE:hydration-BRIDGE:view-creator-missing');
  assert.equal(read(standIn({ wb: { getWorksheet: () => { throw new Error('boom'); } } })), 'UNOBSERVABLE:hydration-BRIDGE:view-creator-threw');
  assert.equal(read(standIn({ wb: { getWorksheet: () => null } })), 'UNOBSERVABLE:hydration-BRIDGE:no-view');
  const api = standIn({ wb: { getWorksheet: () => ({}) } });
  Object.defineProperty(api.GetSheet(), 'worksheet', { get() { throw new Error('x'); } });
  assert.equal(read(api), 'UNOBSERVABLE:hydration-BRIDGE:model-threw');
  // the runtime / capability gates of the hydration client: refused BEFORE the view creator, no direct fallback
  const calls = [];
  const gatedApi = withPatch(standIn({ wb: { getWorksheet: i => { calls.push(i); return {}; } } }));
  assert.equal(read(gatedApi), 'UNOBSERVABLE:hydration-PATCH_VERSION_MISMATCH');
  assert.deepEqual(calls, [], 'an unaccepted runtime never reaches the view creator');
  const notReached = read(standIn({ wb: { getWorksheet: () => ({}) }, worksheetSet: false }));
  assert.equal(notReached[0].anchor, 'UNOBSERVABLE:BRIDGE:geometry-not-hydrated', 'the bridge refuses an un-hydrated geometry itself');
});

// ---- the product path (fake transport) --------------------------------------------------------------------------

const cells = () => { const c = {}; for (let i = 1; i <= 20; i += 1) c[`A${i}`] = { v: `r${i}` }; return c; };
const img = { kind: 'image', name: 'IMG1', pos: { fromCol: 2, colOffset: 0, fromRow: 3, rowOffset: 0 }, width: 1828800, height: 914400 };
const book = () => [{ name: 'S', cells: cells(), drawings: [img] }, { name: 'O', cells: { A1: { v: 'o' } } }];
const insRows = at => ({ intent: 'insert_rows', sheet: 'S', at, count: 1, anchor: { column: 'A', value: `r${at}` } });

async function run(store, step, script = {}) {
  let n = 0;
  const t = { id: 'm7-db-h', steps: [step] };
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics, driver };
}
const ev = r => r.execution.appliedOperations[0].evidence.evidence;

test('product: a hydration that cannot run, or a geometry it did not reach, refuses the structural step before the write', async () => {
  for (const [hydration, reason, detail] of [['missing', 'xlsx-structure-drawings-unobservable', /hydration-BRIDGE:view-creator-missing/], ['not-hydrated', 'xlsx-structure-drawing-geometry-unobservable', /geometry-not-hydrated/]]) {
    const store = createStore(book());
    const before = JSON.stringify(store.durable);
    const { result, m, driver } = await run(store, insRows(15), { hydration });
    assert.equal(result.ok, false, hydration);
    assert.deepEqual([m.writes, driver.calls.dispatch], [0, 0], hydration);
    assert.equal(ev(result).reason, reason, hydration);
    assert.match(JSON.stringify(ev(result)), detail, hydration);
    assert.equal(JSON.stringify(store.durable), before, `${hydration}: nothing changed`);
  }
  const store = createStore(book());
  const ok = await run(store, insRows(15));
  assert.equal(ok.result.ok, true, 'the control: the same step with the geometry hydrated is APPLIED');
  assert.equal(ok.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
});
