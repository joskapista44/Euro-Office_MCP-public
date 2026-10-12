'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { installBridge } = require('../fakes/ds-bridge-runtime.cjs');
const drawings = require('../../src/adapters/xlsx/drawings.cjs');
const { colName, colNum } = require('../../src/adapters/xlsx/a1.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 D-B1 / D-B2 (owner GO 2026-10-03, owner channel point 2; docs/evidence/M7-DB-PROBE.md): a structural step
 * on a sheet hosting drawings, when every drawing's WHOLE extent is proven before the cut by the build-gated anchor read
 * (Type 1 + boundsFromTo.to), and every chart series reaching the sheet ends before the cut. The fake's anchor: the
 * probe's cell metrics (col 609600, row 181000 EMU) -- the image at (2, 3), 1828800 x 914400, ends at col 5 / row 8
 * (0-based) = column F / row 9.
 */

const cells = () => { const c = {}; for (let i = 1; i <= 20; i += 1) c[`A${i}`] = { v: `r${i}` }; for (let k = 2; k <= 12; k += 1) c[`${colName(k)}1`] = { v: `c${colName(k)}` }; return c; };
const img = (over = {}) => ({ kind: 'image', name: 'IMG1', pos: { fromCol: 2, colOffset: 0, fromRow: 3, rowOffset: 0 }, width: 1828800, height: 914400, ...over });
const shp = (over = {}) => ({ kind: 'shape', name: 'SHP1', pos: { fromCol: 7, colOffset: 0, fromRow: 3, rowOffset: 0 }, width: 1828800, height: 914400, text: 'x', ...over });
const ser = (s, col, r1, r2) => [{ name: `=${s}!$${col}$1`, values: `=${s}!$${col}$${r1}:$${col}$${r2}`, cat: `=${s}!$A$${r1}:$A$${r2}`, x: '' }];
const chart = (series, over = {}) => ({ name: 'CH1', type: 'bar', title: 't', series, pos: { fromCol: 1, colOffset: 0, fromRow: 1, rowOffset: 0 }, w: 3600000, h: 900000, ...over });
const book = ({ s = {}, o = {} } = {}) => [{ name: 'S', cells: cells(), ...s }, { name: 'O', cells: { A1: { v: 'o' } }, ...o }];
const objs = charts => ({ objects: { charts, pivots: [], names: [] } });

const insRows = at => ({ intent: 'insert_rows', sheet: 'S', at, count: 1, anchor: { column: 'A', value: `r${at}` } });
const delRows = at => ({ intent: 'delete_rows', sheet: 'S', at, count: 1, anchor: { column: 'A', value: `r${at + 1}` } });
const insCols = at => ({ intent: 'insert_columns', sheet: 'S', at, count: 1, anchor: { row: 1, value: `c${at}` } });
const delCols = at => ({ intent: 'delete_columns', sheet: 'S', at, count: 1, anchor: { row: 1, value: `c${colName(colNum(at) + 1)}` } });

async function run(store, step, script = {}) {
  let n = 0;
  const t = { id: 'm7-db', steps: [step] };
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics, driver };
}
const op0 = r => r.execution.appliedOperations[0];
const ev = r => (op0(r) && op0(r).evidence ? op0(r).evidence.evidence : null);

async function applied(b, step) {
  const store = createStore(b);
  const first = await run(store, step);
  assert.equal(first.result.ok, true, JSON.stringify([ev(first.result), first.result.execution.preservationVerification]));
  assert.equal(op0(first.result).status, OPERATION_STATUS.APPLIED);
  const again = await run(store, step);
  assert.equal(again.m.commands, 0, 'the reopened retry 0/0');
  assert.equal(op0(again.result).status, OPERATION_STATUS.SATISFIED);
  return store;
}
async function refused(b, step, want, why = null) {
  const { result, m } = await run(createStore(b), step);
  assert.equal(result.ok, false);
  assert.equal(ev(result).reason, want, JSON.stringify(ev(result)));
  if (why) assert.match(JSON.stringify(ev(result)), why);
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0);
}

test('driver: the anchor read is generated only when asked; the default drawing read is byte-identical without it', () => {
  assert.doesNotMatch(readQuery([], { drawingSheets: ['*'] }), /drawingBase/);
  assert.equal(readQuery([], { drawingSheets: ['*'] }), readQuery([], { drawingSheets: ['*'], drawingAnchor: false }));
  const ws = { GetName: () => 'D', GetIndex: () => 0, GetUsedRange: () => null };
  const base = { Type: 1, boundsFromTo: { to: { col: 4, colOff: 16.93, row: 8, rowOff: 0.26, extra: 1 } } };
  const d = { GetClassType: () => 'image', GetName: () => 'I', GetParentSheet: () => ws, GetPosition: () => ({ fromCol: 2, colOffset: 0, fromRow: 3, rowOffset: 0 }), GetWidth: () => 1, GetHeight: () => 1, GetRotation: () => 0, GetFlipH: () => null, GetFlipV: () => null, Drawing: { drawingBase: base } };
  const odd = { ...d, GetName: () => 'J', Drawing: {} };
  ws.GetAllDrawings = () => [d, odd];
  base.worksheet = {};
  const api = { GetSheets: () => [ws], GetSheet: () => ws, WorksheetFunction: { COUNT: () => 0 } };
  installBridge(api);
  const out = new Function('Api', readQuery([], { drawingSheets: ['*'], drawingAnchor: true }))(api).drawingReads.D;
  assert.deepEqual(out[0].anchor, { type: 1, to: { col: 4, colOff: 16.93, row: 8, rowOff: 0.26 } }, 'through the DS Bridge drawing.anchor.read/v1');
  assert.equal(out[1].anchor, 'UNOBSERVABLE:BRIDGE:no-drawingBase');
});

test('liveScope: the structure family asks for the anchor read, the others do not', () => {
  assert.equal(liveScope({ id: 'x', steps: [insRows(15)] }).drawingAnchor, true);
  assert.equal(liveScope({ id: 'x', steps: [{ intent: 'rename_sheet', sheet: 'S', name: 'R' }] }).drawingAnchor, undefined);
});

test('anchorExtent: only the measured mode with a complete geometry', () => {
  const r = a => drawings.anchorExtent({ ...img(), parent: 'S', rotation: 0, flipH: null, flipV: null, anchor: a }, 'S');
  const to = { col: 5, colOff: 0, row: 8, rowOff: 100 };
  assert.deepEqual(r({ type: 1, to }), { lastCol: 5, lastRow: 8 });
  assert.match(r({ type: 2, to }).why, /anchor-mode-not-measured/);
  assert.match(r({ type: 0, to }).why, /anchor-mode-not-measured/);
  assert.match(r({ type: 1, to: { ...to, row: undefined } }).why, /anchor-to-incomplete/);
  assert.match(r({ type: 1, to: { ...to, colOff: -1 } }).why, /anchor-to-incomplete/);
  assert.match(r({ type: 1, to: { ...to, row: 2 } }).why, /anchor-to-before-from/);
  assert.match(r(undefined).why, /anchor-not-read/);
  assert.match(r('UNOBSERVABLE:no-drawingBase').why, /no-drawingBase/);
  assert.match(drawings.anchorExtent({ ...img(), parent: 'S', rotation: 90, flipH: null, flipV: null, anchor: { type: 1, to } }, 'S').why, /rotation-not-measured/);
  assert.match(drawings.anchorExtent({ ...img({ width: 0, height: 0 }), parent: 'S', rotation: 0, flipH: null, flipV: null, anchor: { type: 1, to } }, 'S').why, /size-not-observed/);
});

test('APPLIED + reopened 0/0: an image / a shape before the cut, through insert rows, delete rows and insert columns; kept exactly', async () => {
  for (const step of [insRows(15), delRows(15), insRows(10), insCols('L')]) {
    const store = await applied(book({ s: { drawings: [img(), shp()] } }), step);
    assert.deepEqual(store.durable.find(x => x.name === 'S').drawings, [img(), shp()], step.intent);
  }
  // the image alone ends at column F: a column insert at G is after it
  await applied(book({ s: { drawings: [img()] } }), insCols('G'));
});

test('APPLIED + reopened 0/0: a chart before the cut whose series are before the cut; a chart on another sheet over cells before the cut', async () => {
  await applied(book({ s: objs([chart(ser('S', 'B', 2, 4))]) }), insRows(15));
  await applied(book({ s: objs([chart(ser('S', 'B', 2, 4))]) }), delRows(15));
  await applied(book({ o: objs([chart(ser('S', 'B', 2, 4))]) }), insRows(15));
  await applied(book({ o: objs([chart(ser('O', 'B', 2, 4))]) }), insRows(2));
});

test('refused before the write: a drawing at or after the cut (the boundary row / column included)', async () => {
  await refused(book({ s: { drawings: [img()] } }), insRows(9), 'xlsx-structure-drawing-at-or-after-the-cut');
  await refused(book({ s: { drawings: [img()] } }), insRows(2), 'xlsx-structure-drawing-at-or-after-the-cut');
  await refused(book({ s: { drawings: [img()] } }), delRows(9), 'xlsx-structure-drawing-at-or-after-the-cut');
  await refused(book({ s: { drawings: [img()] } }), insCols('F'), 'xlsx-structure-drawing-at-or-after-the-cut');
  await refused(book({ s: objs([chart(ser('O', 'B', 2, 4), { pos: { fromCol: 1, colOffset: 0, fromRow: 12, rowOffset: 0 } })]) }), insRows(15), 'xlsx-structure-drawing-at-or-after-the-cut');
});

test('refused before the write: delete_columns over a drawing or a series is not measured', async () => {
  await refused(book({ s: { drawings: [img()] } }), delCols('K'), 'xlsx-structure-drawing-delete-columns-not-measured');
  await refused(book({ o: objs([chart(ser('S', 'B', 2, 4))]) }), delCols('K'), 'xlsx-structure-chart-series-not-before-the-cut', /delete-columns-over-a-series-not-measured/);
});

test('refused before the write: an unmeasured anchor mode, an incomplete geometry, a rotated drawing', async () => {
  const to = { col: 5, colOff: 0, row: 8, rowOff: 0 };
  for (const [d, why] of [[img({ anchor: { type: 2, to } }), /anchor-mode-not-measured:2/], [img({ anchor: { type: 0, to } }), /anchor-mode-not-measured:0/],
    [img({ anchor: { type: 1, to: { col: 5, row: 8 } } }), /anchor-to-incomplete/], [img({ anchor: 'UNOBSERVABLE:no-drawingBase' }), /no-drawingBase/], [img({ rotation: 30 }), /rotation-not-measured/]]) {
    await refused(book({ s: { drawings: [d] } }), insRows(15), 'xlsx-structure-drawing-geometry-unobservable', why);
  }
});

test('refused before the write: a series reaching the cut, after it, or in an unmeasured shape', async () => {
  await refused(book({ o: objs([chart(ser('S', 'B', 10, 20))]) }), insRows(15), 'xlsx-structure-chart-series-not-before-the-cut', /series-at-or-after-the-cut/);
  await refused(book({ o: objs([chart(ser('S', 'B', 16, 18))]) }), insRows(15), 'xlsx-structure-chart-series-not-before-the-cut', /series-at-or-after-the-cut/);
  await refused(book({ o: objs([chart(ser('S', 'H', 2, 4))]) }), insCols('E'), 'xlsx-structure-chart-series-not-before-the-cut', /series-at-or-after-the-cut/);
  const twoAreas = [{ name: '', values: '=S!$B$2:$B$4,S!$C$2:$C$4', cat: '', x: '' }];
  const { result, m } = await run(createStore(book({ o: objs([chart(twoAreas)]) })), insRows(15));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
});

test('L3 RED: an anchor-only change, a series change, a moved drawing -> not APPLIED, no save', async () => {
  for (const [b, script, rx] of [
    [book({ s: { drawings: [img()] } }), { fault: 'drawing-anchor', drawingSheet: 'S' }, /"anchorsKept":false/],
    [book({ o: { drawings: [img()] } }), { fault: 'drawing-anchor', drawingSheet: 'O' }, /"anchorsKept":false/],
    [book({ o: objs([chart([...ser('O', 'B', 2, 4), ...ser('O', 'C', 2, 4)])]) }), { fault: 'chart-series-other', drawingSheet: 'O' }, /"seriesKept":false/],
    [book({ s: { drawings: [img()] } }), { fault: 'drawing-moves', drawingSheet: 'S' }, /"drawingsKept":false/],
  ]) {
    const { result, driver } = await run(createStore(b), insRows(15), script);
    assert.equal(result.ok, false, script.fault);
    assert.notEqual(op0(result).status, OPERATION_STATUS.APPLIED, script.fault);
    assert.match(JSON.stringify(result.execution.preservationVerification), rx, script.fault);
    assert.equal(driver.calls.forceSave, 0, script.fault);
  }
});
