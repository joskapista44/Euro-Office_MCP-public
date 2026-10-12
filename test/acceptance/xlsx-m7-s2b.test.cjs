'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const drawings = require('../../src/adapters/xlsx/drawings.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { colName } = require('../../src/adapters/xlsx/a1.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { installBridge, withPatch, BRIDGE_RUNTIME_ID } = require('../fakes/ds-bridge-runtime.cjs');

/**
 * M7 S2b (owner GO 2026-10-01, owner channel; decision A on the 0 x 0 size): the drawings (image / shape / chart) of
 * every sheet through the public surface -- an explicit per-kind required field list; the structure family refuses a
 * target sheet hosting ANY drawing (a safety refusal: EO moves the cells and not the drawings, M7-S2B-DRAWING-PROBE.md);
 * copy_sheet refuses a source with an image or a shape; a workbook-wide, order-free, multiplicity-keeping L3 in the
 * structure / sheets / charts families. A chart reading EXACTLY 0 x 0 only makes its size unobserved.
 */

// ---- the pure layer ----
const POS = { fromCol: 2, colOffset: 0, fromRow: 3, rowOffset: 0 };
const img = over => ({ kind: 'image', name: 'IMG', parent: 'D', pos: { ...POS }, width: 1828800, height: 914400, rotation: 0, flipH: null, flipV: null, text: 'UNOBSERVABLE:no-public-content-getter', ...over });
const shp = over => ({ kind: 'shape', name: 'SHP', parent: 'D', pos: { fromCol: 7, colOffset: 0, fromRow: 3, rowOffset: 0 }, width: 1828800, height: 914400, rotation: 0, flipH: false, flipV: false, text: 'S2B\r\n', ...over });
const cht = over => ({ kind: 'chart', name: 'CHT', parent: 'D', pos: { fromCol: 2, colOffset: 0, fromRow: 12, rowOffset: 0 }, width: 3600000, height: 2160000, rotation: 0, flipH: null, flipV: null, text: 'UNOBSERVABLE:no-public-content-getter', ...over });
const obs = list => drawings.observe({ D: list }, ['D']);
const why = list => obs(list).gaps.map(g => `${g.field}:${g.why.split(':')[0]}`);

test('required fields per kind: image / chart need no text, a shape does', () => {
  assert.deepEqual(drawings.REQUIRED.shape.includes('text'), true);
  assert.deepEqual(drawings.REQUIRED.image.includes('text'), false);
  assert.deepEqual(drawings.REQUIRED.chart.includes('text'), false);
  assert.equal(obs([img(), shp(), cht()]).complete, true);
});

test('a missing / throwing / partial getter is a gap, never a default', () => {
  for (const f of ['kind', 'name', 'parent', 'pos', 'width', 'height', 'rotation', 'flipH', 'flipV']) {
    assert.equal(obs([img({ [f]: `UNOBSERVABLE:no-Get${f}` })]).complete, false, f);
    const partial = img(); delete partial[f];
    assert.equal(obs([partial]).complete, false, `${f} missing`);
  }
  assert.equal(obs([shp({ text: 'UNOBSERVABLE:paragraph-text' })]).complete, false, 'shape text');
  assert.equal(obs([img({ pos: { fromCol: 2, colOffset: 0, fromRow: 3 } })]).complete, false, 'partial position');
  assert.deepEqual(drawings.observe({ D: 'UNOBSERVABLE:threw' }, ['D']).gaps[0].field, 'drawings');
  assert.deepEqual(drawings.observe({}, ['D']).gaps[0].why, 'not-read');
  assert.equal(obs([img({ kind: 'group' })]).complete, false, 'groups / OLE: no claim');
});

test('the name: missing, empty, "null" or shared by two drawings of a sheet is a gap; the same name on another sheet is not', () => {
  for (const name of ['', 'null', null, 7]) assert.deepEqual(why([img({ name })]), ['name:name-unusable'], JSON.stringify(name));
  assert.deepEqual(why([shp({ name: 'DUP' }), shp({ name: 'DUP', pos: { ...POS, fromRow: 10 } })]), ['name:name-ambiguous']);
  const two = drawings.observe({ D: [shp()], O: [shp({ parent: 'O' })] }, ['D', 'O']);
  assert.equal(two.complete, true);
  assert.deepEqual(why([shp({ parent: 'O' })]), ['parent:parent-mismatch']);
});

test('size: integer EMU (the measured float noise passes, 1 EMU fails); an EXACT 0 x 0 of ANY kind = size unobserved (decisions A / (2)); one zero side is a gap', () => {
  const a = obs([cht()]);
  const b = obs([cht({ height: 2160000.0000000005 })]);
  assert.deepEqual(drawings.drawingDiffs(a, b), []);
  assert.equal(drawings.drawingDiffs(a, obs([cht({ height: 2160001 })])).length > 0, true, '1 EMU differs');
  for (const mk of [img, shp, cht]) {
    assert.equal(obs([mk({ width: 0, height: 0 })]).complete, true);
    assert.equal(obs([mk({ width: 0, height: 0 })]).bySheet.D[0].width, drawings.SIZE_UNOBSERVED);
    assert.deepEqual(why([mk({ width: 0, height: 914400 })]), ['size:size-zero-unexplained'], 'one zero side');
  }
  // the earlier, stricter contract (the run 1 acceptance): an exact 0 x 0 was a gap
  assert.equal(drawings.normalize(img({ width: 0, height: 0 }), 'D', { zeroSize: 'gap' }).gap.why, 'size-zero-unexplained');
  assert.equal(drawings.ZERO_SIZE, 'size-unobserved');
  assert.deepEqual(why([cht({ width: 0, height: 2160000 })]), ['size:size-zero-unexplained'], 'one zero side is not the measured case');
  for (const w of [NaN, Infinity, -5, '10', 'UNOBSERVABLE:GetWidth-threw']) assert.equal(obs([cht({ width: w })]).complete, false, String(w));
});

test('flip: null is a measured value, distinct from false and from a missing getter', () => {
  assert.equal(obs([img({ flipH: null })]).complete, true);
  assert.equal(obs([img({ flipH: 'x' })]).complete, false);
  assert.equal(obs([img({ flipH: 'UNOBSERVABLE:no-GetFlipH' })]).complete, false);
  assert.equal(drawings.drawingDiffs(obs([img({ flipH: null })]), obs([img({ flipH: false })])).length > 0, true, 'null != false');
});

test('L3: same count but swapped / renamed / moved / resized / retexted -> RED; order permutation -> GREEN', () => {
  const base = [img(), shp(), cht()];
  assert.deepEqual(drawings.drawingDiffs(obs(base), obs([cht(), img(), shp()])), [], 'order is not content');
  const cases = {
    swapped: [img(), shp({ kind: 'image', text: undefined }), cht()],
    renamed: [img({ name: 'IMG2' }), shp(), cht()],
    moved: [img({ pos: { ...POS, fromRow: 4 } }), shp(), cht()],
    offset: [img({ pos: { ...POS, rowOffset: 1 } }), shp(), cht()],
    resized: [img({ width: 1828801 }), shp(), cht()],
    retexted: [img(), shp({ text: 'S2B~' }), cht()],
    rotated: [img({ rotation: 90 }), shp(), cht()],
  };
  for (const [k, after] of Object.entries(cases)) assert.equal(drawings.drawingDiffs(obs(base), obs(after)).length > 0, true, k);
  assert.equal(drawings.drawingDiffs(obs(base), obs([img(), shp()])).length > 0, true, 'one gone');
  assert.equal(drawings.drawingDiffs(obs(base), obs([...base, img({ name: 'IMG9' })])).length > 0, true, 'one appeared');
});

test('L3: multiplicity kept (the multiset, not the set)', () => {
  const r = drawings.normalize(img(), 'D').record;
  assert.equal(drawings.multisetDiff([r, r], [r]).length, 1);
  assert.deepEqual(drawings.multisetDiff([r, r], [r, r]), []);
});

test('L3: an observation gap on either side is a difference, never "unchanged"', () => {
  assert.deepEqual(drawings.drawingDiffs(obs([img()]), obs([img({ width: 'UNOBSERVABLE:GetWidth-threw' })]))[0].change, 'unobservable');
  assert.deepEqual(drawings.drawingDiffs(null, obs([img()]))[0].change, 'observation-missing');
});

test('L3, 0 x 0 chart (decision A): unobserved before -> no size claim; positive before, 0 x 0 after -> RED; its name / anchor still compared', () => {
  const z = obs([cht({ width: 0, height: 0 })]);
  assert.deepEqual(drawings.drawingDiffs(z, z), []);
  assert.deepEqual(drawings.drawingDiffs(z, obs([cht()])), [], 'no size claim when unobserved before');
  assert.equal(drawings.drawingDiffs(obs([cht()]), z).length > 0, true, 'became unobservable after the write: never APPLIED');
  assert.equal(drawings.drawingDiffs(z, obs([cht({ width: 0, height: 0, name: 'X' })])).length > 0, true, 'renamed');
  assert.equal(drawings.drawingDiffs(z, obs([cht({ width: 0, height: 0, pos: { fromCol: 3, colOffset: 0, fromRow: 12, rowOffset: 0 } })])).length > 0, true, 'moved');
});

test('driver: the drawing read is generated only when asked, each getter on its own; a bad getter is named, a 0 x 0 is passed through raw', () => {
  assert.doesNotMatch(readQuery(), /GetAllDrawings/);
  const pos = { fromCol: 2, colOffset: 0, fromRow: 3, rowOffset: 0 };
  const ws = { GetName: () => 'D', GetIndex: () => 0, GetUsedRange: () => null };
  const para = t => ({ GetText: () => t });
  // M7 D-B shape text (owner GO 31416), since the DS Bridge cutover (internal tracker point 8) through drawing.text.read/v1 of
  // the REAL served bridge block: the NON-creating Drawing.getDocContent(), wrapped by Api.private_CreateApiDocContent;
  // GetContent itself (it creates a body on a text-less shape) never runs
  const body = { paragraphs: ['a', 'b'] };
  const d = {
    GetClassType: () => 'shape', GetName: () => 'S', GetParentSheet: () => ws, GetPosition: () => pos, GetWidth: () => 10, GetHeight: () => { throw new Error('x'); },
    GetRotation: () => 0, GetFlipH: () => false, GetContent: () => { throw new Error('GetContent must not be called'); }, Drawing: { getDocContent: () => body },
  };
  ws.GetAllDrawings = () => [d];
  const mkApi = (over = {}, bridge = true) => {
    const a = { GetSheets: () => [ws], GetSheet: () => ws, WorksheetFunction: { COUNT: () => 0 }, private_CreateApiDocContent: dc => ({ GetElementsCount: () => dc.paragraphs.length, GetElement: i => para(dc.paragraphs[i]) }), ...over };
    if (bridge) installBridge(a);
    return a;
  };
  const q = () => readQuery([], { drawingSheets: ['*'] });
  const out = new Function('Api', q())(mkApi()).drawingReads;
  assert.deepEqual(out.D[0], { kind: 'shape', name: 'S', parent: 'D', pos, width: 10, height: 'UNOBSERVABLE:GetHeight-threw', rotation: 0, flipH: false, flipV: 'UNOBSERVABLE:no-GetFlipV', text: 'a|b' });
  const textOf = (a = mkApi()) => new Function('Api', q())(a).drawingReads.D[0].text;
  ws.GetAllDrawings = () => [{ ...d, Drawing: { getDocContent: () => null } }];
  assert.equal(textOf(), '', 'no text body: empty, nothing created');
  ws.GetAllDrawings = () => [{ ...d, Drawing: { getDocContent: () => undefined } }];
  assert.equal(textOf(), 'UNOBSERVABLE:BRIDGE:getDocContent-undefined', 'owner 31441: only the measured null is "no text body"');
  ws.GetAllDrawings = () => [{ ...d, Drawing: {} }];
  assert.equal(textOf(), 'UNOBSERVABLE:BRIDGE:no-getDocContent');
  ws.GetAllDrawings = () => [{ ...d, Drawing: { getDocContent: () => { throw new Error('y'); } } }];
  assert.equal(textOf(), 'UNOBSERVABLE:BRIDGE:getDocContent-threw');
  ws.GetAllDrawings = () => [d];
  assert.equal(textOf(mkApi({ private_CreateApiDocContent: () => { throw new Error('z'); } })), 'UNOBSERVABLE:BRIDGE:doc-content-wrap-threw');
  assert.equal(textOf(withPatch(mkApi())), 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', 'an unaccepted served runtime reads no text');
  assert.equal(textOf(mkApi({}, false)), 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', 'no bridge: never a direct read');
  assert.doesNotMatch(q(), /GetContent|getDocContent/, 'the query names no internal text getter');
  ws.GetAllDrawings = undefined;
  assert.equal(new Function('Api', readQuery([], { drawingSheets: ['*'] }))({ GetSheets: () => [ws], GetSheet: () => ws, WorksheetFunction: { COUNT: () => 0 } }).drawingReads.D, 'UNOBSERVABLE:no-GetAllDrawings');
});

// ---- the families, through the fake ----
const cellsAZ = () => { const c = {}; for (let i = 1; i <= 20; i += 1) c[`A${i}`] = { v: `r${i}` }; for (let k = 2; k <= 12; k += 1) c[`${colName(k)}1`] = { v: `c${colName(k)}` }; return c; };
const fakeImg = (over = {}) => ({ kind: 'image', name: 'IMG1', pos: { fromCol: 2, colOffset: 0, fromRow: 3, rowOffset: 0 }, width: 1828800, height: 914400, ...over });
const fakeShp = (over = {}) => ({ kind: 'shape', name: 'SHP1', pos: { fromCol: 7, colOffset: 0, fromRow: 3, rowOffset: 0 }, width: 1828800, height: 914400, text: 'S2B\r\n', ...over });
const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = ({ s = {}, o = {} } = {}) => [{ name: 'S', cells: cellsAZ(), ...s }, { name: 'O', cells: { A1: { v: 'o' } }, ...o }];
async function run(store, steps, script = {}) {
  let n = 0;
  const t = { id: 'm7s2b', steps: Array.isArray(steps) ? steps : [steps] };
  // M11.5c.4: copy_sheet writes only through the /8 Bridge WRITE -- a copy task runs on /8 unless the test names a runtime
  const driver = createFakeXlsxDriver(store, t.steps.some(x => x.intent === 'copy_sheet') && script.runtime === undefined ? { ...script, runtime: '9.3.4|euro-office-bridge-v1/9.3.4.60/8' } : script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}
const op0 = r => r.execution.appliedOperations[0];
const reason = r => { const o = op0(r); return o && o.evidence && o.evidence.evidence ? o.evidence.evidence.reason : null; };
async function refused(b, step, want, script = {}) {
  const { result, m } = await run(createStore(b), step, script);
  assert.equal(result.ok, false);
  assert.equal(reason(result), want, JSON.stringify(op0(result)));
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0);
}
async function applied(b, step, script = {}) {
  const store = createStore(b);
  const first = await run(store, step, script);
  assert.equal(first.result.ok, true, JSON.stringify([op0(first.result), first.result.execution.preservationVerification]));
  assert.equal(op0(first.result).status, OPERATION_STATUS.APPLIED);
  const again = await run(store, step, script);
  assert.equal(again.result.ok, true);
  assert.equal(again.m.commands, 0, 'the reopened retry 0/0');
  assert.equal(op0(again.result).status, OPERATION_STATUS.SATISFIED);
  return store;
}
async function l3Red(b, step, script) {
  const { result } = await run(createStore(b), step, script);
  assert.equal(result.ok, false);
  assert.notEqual(op0(result).status, OPERATION_STATUS.APPLIED);
  assert.match(JSON.stringify(result.execution.preservationVerification), /"axis":"drawings"/);
}
const insRows = (at, sheet = 'S') => ({ intent: 'insert_rows', sheet, at, count: 1, anchor: { column: 'A', value: `r${at}` } });
const insCols = (at, sheet = 'S') => ({ intent: 'insert_columns', sheet, at, count: 1, anchor: { row: 1, value: `c${at}` } });

test('liveScope: structure, sheets and charts read every sheet\'s drawings', () => {
  for (const step of [insRows(15), { intent: 'rename_sheet', sheet: 'S', name: 'Renamed' }, { intent: 'rename_chart', sheet: 'S', name: 'a', newName: 'b' }]) {
    assert.deepEqual(liveScope({ id: 'x', steps: [step] }).drawingSheets, ['*'], step.intent);
  }
});

// M7 D-B (owner GO 2026-10-03): the "before the cut" exception exists only with the build-gated anchor read; on an
// unproven build (anchorBuild false) the S2b refusal stands, now named by its cause (xlsx-m7-db-structure.test.cjs: the rest)
test('structure: without the anchor read (unproven build) a target sheet hosting an image / a shape refuses EVERY insert, even far below / right of it', async () => {
  for (const d of [fakeImg(), fakeShp(), fakeImg({ pos: { fromCol: 0, colOffset: 0, fromRow: 0, rowOffset: 0 } })]) {
    await refused(book({ s: { drawings: [d] } }), insRows(19), 'xlsx-structure-drawing-geometry-unobservable', { anchorBuild: false });
    await refused(book({ s: { drawings: [d] } }), insCols('L'), 'xlsx-structure-drawing-geometry-unobservable', { anchorBuild: false });
    await refused(book({ s: { drawings: [d] } }), { intent: 'delete_rows', sheet: 'S', at: 19, count: 1, anchor: { column: 'A', value: 'r20' } }, 'xlsx-structure-drawing-geometry-unobservable', { anchorBuild: false });
  }
});

test('structure: a 0 x 0 chart on the target sheet still refuses (the guard is independent of the chart size)', async () => {
  const b = book({ s: { objects: { charts: [{ name: 'Z', type: 'bar', title: 'z', series: series('O', ['B']), w: 0, h: 0 }], pivots: [], names: [] } } });
  const { result, m } = await run(createStore(b), insRows(19));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(String(reason(result)), /xlsx-structure-drawing-geometry-unobservable/);
});

test('structure: drawings on ANOTHER sheet -> APPLIED + reopened 0/0, and they are kept exactly', async () => {
  const store = await applied(book({ o: { drawings: [fakeImg(), fakeShp()] } }), insRows(15));
  assert.deepEqual(store.durable.find(x => x.name === 'O').drawings, [fakeImg(), fakeShp()]);
});

test('structure: an unreadable drawing anywhere refuses before the write', async () => {
  await refused(book(), insRows(15), 'xlsx-structure-drawings-unobservable', { drawingsUnobservable: true });
  await refused(book({ o: { drawings: [fakeImg()] } }), insRows(15), 'xlsx-structure-drawings-unobservable', { drawingField: ['width', 'UNOBSERVABLE:GetWidth-threw'] });
  await refused(book({ o: { drawings: [fakeImg({ width: 0, height: 914400 })] } }), insRows(15), 'xlsx-structure-drawings-unobservable');
  await refused(book({ o: { drawings: [fakeShp({ name: 'D' }), fakeShp({ name: 'D', pos: { fromCol: 1, colOffset: 0, fromRow: 10, rowOffset: 0 } })] } }), insRows(15), 'xlsx-structure-drawings-unobservable');
});

for (const fault of ['drawing-moves', 'drawing-renames', 'drawing-resizes', 'drawing-text', 'drawing-swaps']) {
  test(`structure L3: ${fault} on another sheet, same count -> not APPLIED`, async () => {
    await l3Red(book({ o: { drawings: [fakeShp(), fakeImg()] } }), insRows(15), { fault, drawingSheet: 'O' });
  });
}

test('copy_sheet: a source with an image or a shape refuses before the write', async () => {
  for (const d of [fakeImg(), fakeShp()]) await refused(book({ s: { drawings: [d] } }), { intent: 'copy_sheet', sheet: 'S', name: 'S_copy' }, 'xlsx-copy-source-hosts-images-or-shapes');
});

test('copy_sheet: a chart-only source keeps the proven path (APPLIED + reopened 0/0); drawings on another sheet are kept', async () => {
  const b = book({ s: { objects: { charts: [{ name: 'CC', type: 'bar', title: 't', series: series('S', ['B']) }], pivots: [], names: [] } }, o: { drawings: [fakeShp()] } });
  const store = await applied(b, { intent: 'copy_sheet', sheet: 'S', name: 'S_copy' });
  assert.deepEqual(store.durable.map(x => x.name), ['S', 'O', 'S_copy']);
});

test('rename_sheet: the drawings follow the sheet on purpose (parent renamed); every other field kept', async () => {
  const store = await applied(book({ s: { drawings: [fakeImg(), fakeShp()] } }), { intent: 'rename_sheet', sheet: 'S', name: 'Renamed' });
  assert.deepEqual(store.durable.find(x => x.name === 'Renamed').drawings, [fakeImg(), fakeShp()]);
});

test('delete_sheet: the deleted sheet\'s drawings go on purpose; another sheet\'s are kept', async () => {
  await applied(book({ s: { drawings: [fakeImg()] }, o: { drawings: [fakeShp()] } }), { intent: 'delete_sheet', sheet: 'S' });
});

test('sheets L3: a drawing on another sheet changed by the write -> not APPLIED (create / rename / delete)', async () => {
  for (const step of [{ intent: 'create_sheet', name: 'N' }, { intent: 'rename_sheet', sheet: 'S', name: 'Renamed' }, { intent: 'delete_sheet', sheet: 'S' }]) {
    await l3Red(book({ o: { drawings: [fakeShp()] } }), step, { fault: 'drawing-moves', drawingSheet: 'O' });
  }
});

test('sheets: an unreadable drawing anywhere refuses before the write', async () => {
  await refused(book(), { intent: 'create_sheet', name: 'N' }, 'xlsx-sheets-drawings-unobservable', { drawingsUnobservable: true });
});

// ---- charts: decision A ----
const kbook = (charts, otherDrawings = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts, pivots: [], names: [] } },
  { name: 'O', cells: { A1: { v: 'keep' } }, objects: { charts: [{ name: 'Other', type: 'bar', title: 'o', series: series('K', ['B']), w: 0, h: 0 }], pivots: [], names: [] }, drawings: otherDrawings },
];
const zeroC1 = () => ({ name: 'C1', type: 'bar', title: 'Sales', series: series('K', ['B', 'C']), w: 0, h: 0 });
const ren = { intent: 'rename_chart', sheet: 'K', name: 'C1', newName: 'C2' };

test('rename_chart on a 0 x 0 chart (the measured reopen state): APPLIED, reopened retry SATISFIED 0/0 -- the size axis does not block', async () => {
  await applied(kbook([zeroC1()]), ren);
});

test('set_chart with 0 x 0 charts elsewhere: APPLIED + retry 0/0; an image / shape on another sheet kept', async () => {
  await applied(kbook([], [fakeShp()]), { intent: 'set_chart', sheet: 'K', name: 'C1', range: 'A1:C4', type: 'bar', title: 'Sales' });
});

for (const [fault, sheet, note] of [
  ['chart-moves', 'K', 'the renamed 0 x 0 chart moved'],
  ['chart-renames-other', 'O', 'another 0 x 0 chart renamed'],
  ['chart-moves', 'O', 'another 0 x 0 chart moved'],
]) {
  test(`charts L3: ${note} -> not APPLIED`, async () => {
    await l3Red(kbook([zeroC1()]), ren, { fault, drawingSheet: sheet });
  });
}

test('charts L3: another 0 x 0 chart loses a series -> not APPLIED (the chart contract still checked)', async () => {
  const { result } = await run(createStore(kbook([zeroC1()])), ren, { fault: 'chart-series-other', drawingSheet: 'O' });
  assert.equal(result.ok, false);
  assert.notEqual(op0(result).status, OPERATION_STATUS.APPLIED);
});

test('charts L3: a positive chart size that reads 0 x 0 after the write -> not APPLIED', async () => {
  const b = kbook([{ ...zeroC1(), w: 3600000, h: 2160000 }]);
  await l3Red(b, ren, { fault: 'chart-size-lost', drawingSheet: 'K' });
});

test('charts: an image with a zero on ONE side anywhere refuses (only the exact 0 x 0 is a size-observation limit)', async () => {
  await refused(kbook([zeroC1()], [fakeImg({ width: 0, height: 914400 })]), ren, 'xlsx-charts-drawings-unobservable');
});

// ---- decision (2): a 0 x 0 image / shape on ANOTHER sheet does not block; the L3 still protects its observable fields ----
const z = over => ({ width: 0, height: 0, ...over });
const zeroOther = () => book({ o: { drawings: [fakeImg(z()), fakeShp(z())] } });
test('(2) structure / sheets / charts APPLIED + reopened 0/0 next to a 0 x 0 image and shape on another sheet', async () => {
  await applied(zeroOther(), insRows(15));
  await applied(zeroOther(), { intent: 'create_sheet', name: 'New' });
  await applied(kbook([zeroC1()], [fakeImg(z()), fakeShp(z())]), ren);
});
test('(2) the targeted refusals stay: a 0 x 0 drawing on the TARGET sheet still refuses the structural step and the copy', async () => {
  await refused(book({ s: { drawings: [fakeImg(z())] } }), insRows(19), 'xlsx-structure-drawing-geometry-unobservable');
  await refused(book({ s: { drawings: [fakeShp(z())] } }), { intent: 'copy_sheet', sheet: 'S', name: 'S_copy' }, 'xlsx-copy-source-hosts-images-or-shapes');
});
for (const fault of ['drawing-moves', 'drawing-renames', 'drawing-text']) {
  test(`(2) L3 on a 0 x 0 drawing, same count: ${fault} -> not APPLIED`, async () => {
    await l3Red(book({ o: { drawings: [fakeShp(z()), fakeImg(z({ name: 'IMG2' }))] } }), insRows(15), { fault, drawingSheet: 'O' });
  });
}
test('(2) the size observation lost by the write (positive before, 0 x 0 after) -> not APPLIED, for an image and a shape', async () => {
  for (const d of [fakeImg(), fakeShp()]) await l3Red(book({ o: { drawings: [d] } }), insRows(15), { fault: 'drawing-size-lost', drawingSheet: 'O' });
});

// ---- M7 D-B shape text (owner GO 2026-10-04, owner channel): structure / sheets / charts with a shape WITH text
// (positive) and a text-less shape (its text '' -- the non-creating read), APPLIED + 0/0 and kept exactly; an
// unreadable shape text (a missing / throwing getDocContent) refuses each path before the write ----
const bareShp = (over = {}) => fakeShp({ name: 'BARE', pos: { fromCol: 9, colOffset: 0, fromRow: 3, rowOffset: 0 }, text: '', ...over });
const textBook = () => book({ o: { drawings: [fakeShp(), bareShp()] } });
test('shape text: structure / sheets / charts APPLIED + reopened 0/0 next to a shape with text and a text-less one; both kept exactly', async () => {
  const s1 = await applied(textBook(), insRows(15));
  assert.deepEqual(s1.durable.find(x => x.name === 'O').drawings, [fakeShp(), bareShp()]);
  const s2 = await applied(textBook(), { intent: 'create_sheet', name: 'New' });
  assert.deepEqual(s2.durable.find(x => x.name === 'O').drawings, [fakeShp(), bareShp()]);
  await applied(kbook([zeroC1()], [fakeShp(), bareShp()]), ren);
});
test('shape text: an unreadable text (getDocContent missing / throwing) refuses structure / sheets / charts before the write', async () => {
  for (const why of ['UNOBSERVABLE:no-getDocContent', 'UNOBSERVABLE:getDocContent-threw', 'UNOBSERVABLE:getDocContent-undefined']) {
    const field = { drawingField: ['text', why] };
    await refused(textBook(), insRows(15), 'xlsx-structure-drawings-unobservable', field);
    await refused(textBook(), { intent: 'create_sheet', name: 'N' }, 'xlsx-sheets-drawings-unobservable', field);
    await refused(kbook([zeroC1()], [fakeShp(), bareShp()]), ren, 'xlsx-charts-drawings-unobservable', field);
  }
});
test('shape text L3: a text appearing on the text-less shape (the GetContent side effect) -> not APPLIED', async () => {
  await l3Red(book({ o: { drawings: [bareShp(), fakeImg()] } }), insRows(15), { fault: 'drawing-text', drawingSheet: 'O' });
});
