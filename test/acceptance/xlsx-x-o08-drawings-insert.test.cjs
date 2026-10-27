'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const drawingsInsert = require('../../src/adapters/xlsx/families/drawings-insert.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { mapExecutionResult } = require('../../src/server/write-tools.cjs');

/** X-O08 (FORMATTING COMPLETION backlog, docs/X-O08-DRAWING-INSERT-DESIGN-20261007.md v5):
 * insert_shape / insert_image on the fake. Mirrors the already-migrated `charts` family's own
 * test shape (xlsx-w7-charts.test.cjs) -- retry/identity-collision/L1/L3 tables, + direct
 * command-body unit tests, + a direct driver-reader unit test. */

const book = (extra = []) => [
  { name: 'D', cells: { A1: { v: 'keep' } }, drawings: extra },
  { name: 'O', cells: { A1: { v: 'other-keep' } }, drawings: [] },
];
const shapeStep = (over = {}) => ({ intent: 'insert_shape', sheet: 'D', name: 'S1', geometry: 'rect', at: { col: 2, row: 1 }, size: { width: 72, height: 36 }, fill: '#FF0000', ...over });
const imageStep = (over = {}) => ({ intent: 'insert_image', sheet: 'D', name: 'I1', at: { col: 2, row: 1 }, size: { width: 72, height: 36 }, source: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEElEQVR4nGM4o6QERwzEcQDE4xEBYJBBnwAAAABJRU5ErkJggg==', ...over });
const addShape = (over = {}) => ({ id: 'x-o08', steps: [shapeStep(over)] });
const addImage = (over = {}) => ({ id: 'x-o08', steps: [imageStep(over)] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;
const drawingsOf = (store, s) => store.durable.find(x => x.name === s).drawings;

test('insert_shape: APPLIED at the requested cell, retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, addShape());
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.appliedOperations[0].evidence]));
  assert.equal(a.m.commands, 1);
  assert.deepEqual(drawingsOf(store, 'D').map(d => ({ kind: d.kind, name: d.name, fromCol: d.pos.fromCol, fromRow: d.pos.fromRow })), [{ kind: 'shape', name: 'S1', fromCol: 2, fromRow: 1 }]);
  const r = await run(store, addShape());
  assert.equal(r.result.ok, true, JSON.stringify(r.result.execution.appliedOperations[0].evidence));
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('insert_image: APPLIED at the requested cell, retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, addImage());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.deepEqual(drawingsOf(store, 'D').map(d => d.kind), ['image']);
  const r = await run(store, addImage());
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('insert_shape: a DIFFERENT, unrelated drawing on the sheet is untouched', async () => {
  const store = createStore(book([{ kind: 'image', name: 'Other', pos: { fromCol: 9, colOffset: 0, fromRow: 9, rowOffset: 0 }, width: 100, height: 100 }]));
  await run(store, addShape());
  assert.deepEqual(drawingsOf(store, 'D').map(d => d.name).sort(), ['Other', 'S1']);
});

for (const [name, t, why, b] of [
  ['a shape of the name at ANOTHER cell (no silent modify)', addShape(), 'xlsx-drawing-insert-exists-with-another-identity', book([{ kind: 'shape', name: 'S1', pos: { fromCol: 9, colOffset: 0, fromRow: 9, rowOffset: 0 }, width: 100, height: 100, text: '' }])],
  ['the name taken by an IMAGE when inserting a shape', addShape(), 'xlsx-drawing-insert-exists-with-another-identity', book([{ kind: 'image', name: 'S1', pos: { fromCol: 2, colOffset: 0, fromRow: 1, rowOffset: 0 }, width: 100, height: 100 }])],
  ['the name taken by a SHAPE when inserting an image', addImage(), 'xlsx-drawing-insert-exists-with-another-identity', book([{ kind: 'shape', name: 'I1', pos: { fromCol: 2, colOffset: 0, fromRow: 1, rowOffset: 0 }, width: 100, height: 100, text: '' }])],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const { result, driver } = await run(createStore(b), t);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('refused BEFORE dispatch: an unobservable drawing reader', async () => {
  const { result, driver } = await run(createStore(book()), addShape(), { drawingsUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-drawing-insert-drawings-unobservable');
});

test('RED: ignored / wrong position / wrong kind -> L1 FAIL', async () => {
  for (const fault of ['ignore', 'drawing-insert-wrong-position', 'drawing-insert-wrong-kind']) {
    const { result, driver } = await run(createStore(book()), addShape(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

test('RED: a drawing fault on an OTHER drawing after dispatch -> L3 FAIL', async () => {
  const { result } = await run(createStore(book([{ kind: 'image', name: 'Other', pos: { fromCol: 9, colOffset: 0, fromRow: 9, rowOffset: 0 }, width: 100, height: 100 }])), addShape(), { fault: 'drawing-moves', drawingSheet: 'D' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
});

// ---- v2: stroke (X-O08-7, 16fd90b9, LIVE r3) -- width is the READABLE/verified axis. ----
// ---- GH #139 D4/D6 (2026-10-17): colour (fill AND stroke) is now PROVEN through the driver's
// fillLineColor read on an accepted (/14+) runtime -- the fake's default. `script.colourUnobservable`
// simulates a pre-/14 accepted runtime: colour stays unprovable, named in the receipt's `notProven`,
// same as before this capability existed. ----

test('insert_shape + stroke: APPLIED, L1 PASS (width verified), retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, addShape({ stroke: { width: 3, color: '#C81E1E' } }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.equal(drawingsOf(store, 'D')[0].strokeWidth, 38100, 'SetLine wrote 3pt = 38100 EMU, matching X-O08-7 LIVE r3');
  const r = await run(store, addShape({ stroke: { width: 3, color: '#C81E1E' } }));
  assert.equal(r.result.ok, true, JSON.stringify(r.result.execution.appliedOperations[0].evidence));
  assert.equal(r.m.commands, 0, 'same width retried -> SATISFIED, no dispatch');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('insert_shape without stroke: strokeWidth defaults to 0 (the product\'s own hardcoded CreateStroke(0, CreateNoFill())), never a gap', async () => {
  const store = createStore(book());
  const a = await run(store, addShape());
  assert.equal(a.result.ok, true);
  assert.equal(drawingsOf(store, 'D')[0].strokeWidth, 0);
});

test('RED: the SetLine write lands a stroke width 1 EMU off the requested one -> L1 FAIL (the SAME axis the position fault already proves)', async () => {
  const { result, driver } = await run(createStore(book()), addShape({ stroke: { width: 3, color: '#C81E1E' } }), { fault: 'drawing-insert-wrong-stroke-width' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(result.execution.primitiveVerification[0].evidence.evidence.now.strokeWidth, 38101);
  assert.equal(driver.calls.forceSave, 0);
});

test('a stroke-width-off fault is a NO-OP when no `stroke` was requested -- identityOf never compares an axis the operation itself did not write', async () => {
  const { result } = await run(createStore(book()), addShape(), { fault: 'drawing-insert-wrong-stroke-width' });
  assert.equal(result.ok, true);
});

test('wholeTaskVerification: a PROVEN, matching fillLineColor read drops fill/stroke colour from notProven entirely', async () => {
  const withStroke = await run(createStore(book()), addShape({ stroke: { width: 3, color: '#C81E1E' } }));
  assert.equal('notProven' in withStroke.result.execution.wholeTaskVerification.evidence.checks[0], false);
  const withoutStroke = await run(createStore(book()), addShape());
  assert.equal('notProven' in withoutStroke.result.execution.wholeTaskVerification.evidence.checks[0], false);
  // insert_image carries no colour parameter at all -- no notProven entry
  const image = await run(createStore(book()), addImage());
  assert.equal('notProven' in image.result.execution.wholeTaskVerification.evidence.checks[0], false);
});

test('wholeTaskVerification: on a pre-/14 accepted runtime (fillLineColor UNOBSERVABLE), fill-colour stays notProven on every shape, stroke-colour added only when `stroke` was written -- the limit the capability closed', async () => {
  const withStroke = await run(createStore(book()), addShape({ stroke: { width: 3, color: '#C81E1E' } }), { colourUnobservable: true });
  assert.deepEqual(withStroke.result.execution.wholeTaskVerification.evidence.checks[0].notProven, ['shape-fill-colour', 'shape-stroke-colour']);
  const withoutStroke = await run(createStore(book()), addShape(), { colourUnobservable: true });
  assert.deepEqual(withoutStroke.result.execution.wholeTaskVerification.evidence.checks[0].notProven, ['shape-fill-colour']);
});

test('RED: the fillLineColor read reports a DIFFERENT fill colour than written -> L1 FAIL (GH #139 D4, the SAME shape the position/stroke-width faults use)', async () => {
  const { result, driver } = await run(createStore(book()), addShape(), { fault: 'drawing-insert-wrong-colour' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
});

test('RED: the fillLineColor read reports a DIFFERENT stroke colour than written -> L1 FAIL (GH #139 D6), and is a NO-OP when no `stroke` was requested', async () => {
  const withStroke = await run(createStore(book()), addShape({ stroke: { width: 3, color: '#C81E1E' } }), { fault: 'drawing-insert-wrong-stroke-colour' });
  assert.equal(withStroke.result.ok, false);
  assert.equal(withStroke.result.execution.primitiveVerification[0].verdict, 'FAIL');
  const withoutStroke = await run(createStore(book()), addShape(), { fault: 'drawing-insert-wrong-stroke-colour' });
  assert.equal(withoutStroke.result.ok, true, 'identityOf never compares an axis the operation itself did not write');
});

test('planner: stroke field validation', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse({ id: 'x', steps: [shapeStep({ stroke: { color: '#FF0000' } })] }, /stroke\.width must be a positive number/);
  refuse({ id: 'x', steps: [shapeStep({ stroke: { width: 0, color: '#FF0000' } })] }, /stroke\.width must be a positive number/);
  refuse({ id: 'x', steps: [shapeStep({ stroke: { width: 3, color: 'red' } })] }, /stroke\.color must be a 6-hex colour/);
  refuse({ id: 'x', steps: [shapeStep({ stroke: 'thick' })] }, /stroke must be an object/);
  refuse({ id: 'x', steps: [imageStep({ stroke: { width: 3, color: '#FF0000' } })] }, /does not support field/);
});

test('the command body: stroke -> SetLine AFTER SetName, the measured mechanism (not AddShape\'s own stroke argument)', () => {
  const log = [];
  const ws = { GetName: () => 'D', GetAllDrawings: () => [], AddShape: (...x) => { log.push(['AddShape', ...x]); return { GetName: () => null, SetName: x2 => log.push(['SetName', x2]), SetLine: x2 => log.push(['SetLine', x2]) }; } };
  const Api = { GetSheet: () => ws, CreateSolidFill: (...x) => ['fill', ...x], CreateRGBColor: (...x) => ['rgb', ...x], CreateStroke: (...x) => ['stroke', ...x], CreateNoFill: () => 'nofill' };
  const op = planOperations(addShape({ stroke: { width: 3, color: '#C81E1E' } })).operations[0];
  assert.deepEqual(new Function('Api', drawingsInsert.lower(op).body)(Api), { ok: true });
  assert.deepEqual(log[0], ['AddShape', 'rect', 914400, 457200, ['fill', ['rgb', 255, 0, 0]], ['stroke', 0, 'nofill'], 2, 0, 1, 0], 'AddShape keeps the default 0/noFill stroke -- the new stroke is NOT passed here');
  assert.deepEqual(log[1], ['SetName', 'S1']);
  assert.deepEqual(log[2], ['SetLine', ['stroke', 38100, ['fill', ['rgb', 200, 30, 30]]]], 'SetLine called AFTER SetName, with the requested width (3pt=38100 EMU) and colour');
});

test('planner: field validation', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse({ id: 'x', steps: [shapeStep({ sheet: 'My Sheet' })] }, /needs no quoting/);
  refuse({ id: 'x', steps: [shapeStep({ geometry: 'oval' })] }, /geometry must be one of/);
  refuse({ id: 'x', steps: [shapeStep({ fill: 'red' })] }, /fill must be a 6-hex colour/);
  refuse({ id: 'x', steps: [shapeStep({ at: { col: -1, row: 0 } })] }, /col must be a non-negative integer/);
  refuse({ id: 'x', steps: [shapeStep({ at: { col: 1.5, row: 0 } })] }, /col must be a non-negative integer/);
  refuse({ id: 'x', steps: [shapeStep({ size: { width: 0, height: 10 } })] }, /width must be a positive number/);
  refuse({ id: 'x', steps: [shapeStep({ size: { width: -5, height: 10 } })] }, /width must be a positive number/);
  refuse({ id: 'x', steps: [shapeStep({ colOffset: 5 })] }, /does not support field/);
  refuse({ id: 'x', steps: [imageStep({ source: 'http://example.com/x.png' })] }, /only PNG is LIVE-measured/);
  refuse({ id: 'x', steps: [imageStep({ source: 'data:image/jpeg;base64,abcd' })] }, /only PNG is LIVE-measured/);
  refuse({ id: 'x', steps: [imageStep({ fill: '#FF0000' })] }, /does not support field/);
  refuse({ id: 'x', steps: [{ intent: 'delete_shape', sheet: 'D', name: 'S1' }] }, /./);
});

test('the command body: AddShape + SetName; name-taken re-checked at dispatch', () => {
  const log = [];
  const existing = [];
  const mk = n => ({ GetName: () => n, SetName: x => log.push(['SetName', x]) });
  const ws = { GetName: () => 'D', GetAllDrawings: () => existing, AddShape: (...x) => { log.push(['AddShape', ...x]); return { GetName: () => null, SetName: x2 => log.push(['SetName', x2]) }; } };
  const Api = { GetSheet: () => ws, CreateSolidFill: (...x) => ['fill', ...x], CreateRGBColor: (...x) => ['rgb', ...x], CreateStroke: (...x) => ['stroke', ...x], CreateNoFill: () => 'nofill' };
  assert.deepEqual(new Function('Api', drawingsInsert.lower(planOperations(addShape()).operations[0]).body)(Api), { ok: true });
  assert.deepEqual(log[0], ['AddShape', 'rect', 914400, 457200, ['fill', ['rgb', 255, 0, 0]], ['stroke', 0, 'nofill'], 2, 0, 1, 0]);
  assert.deepEqual(log[1], ['SetName', 'S1']);
  log.length = 0;
  existing.push(mk('S1'));
  assert.deepEqual(new Function('Api', drawingsInsert.lower(planOperations(addShape()).operations[0]).body)(Api), { ok: false, stage: 'name-taken' });
});

test('the command body: AddImage + SetName', () => {
  const log = [];
  const ws = { GetName: () => 'D', GetAllDrawings: () => [], AddImage: (...x) => { log.push(['AddImage', ...x]); return { GetName: () => null, SetName: x2 => log.push(['SetName', x2]) }; } };
  const Api = { GetSheet: () => ws };
  const op = planOperations(addImage()).operations[0];
  assert.deepEqual(new Function('Api', drawingsInsert.lower(op).body)(Api), { ok: true });
  assert.deepEqual(log[0], ['AddImage', op.source, 914400, 457200, 2, 0, 1, 0]);
  assert.deepEqual(log[1], ['SetName', 'I1']);
});

test('EMU conversion: pt * 12700, rounded', () => {
  const op = planOperations(addShape({ size: { width: 1, height: 500 } })).operations[0];
  const log = [];
  const ws = { GetName: () => 'D', GetAllDrawings: () => [], AddShape: (...x) => { log.push(x); return { GetName: () => null, SetName: () => {} }; } };
  const Api = { GetSheet: () => ws, CreateSolidFill: () => null, CreateRGBColor: () => null, CreateStroke: () => null, CreateNoFill: () => null };
  new Function('Api', drawingsInsert.lower(op).body)(Api);
  assert.equal(log[0][1], 12700);
  assert.equal(log[0][2], 6350000);
});

// ---- the MCP-facing RECEIPT (mapExecutionResult, write-tools.cjs) -- the gap a maintainer's LIVE r1 found:
// `notProven` reached the executor's own internal wholeTaskVerification, but never this receipt, so
// an MCP caller saw APPLIED + whole-task PASS with no sign the colour axes were unverified. ----

test('receipt: operations[0].semanticsNotProven -- a PROVEN fillLineColor read carries NO colour axis at all (GH #139 D4/D6 closes this limit on an accepted runtime)', async () => {
  const store = createStore(book());
  const t = { id: 'x-o08', steps: [shapeStep({ stroke: { width: 3, color: '#C81E1E' } })] };
  const { result } = await run(store, t);
  const mapped = mapExecutionResult(result, { fileId: 'disposable', task: t });
  assert.equal('semanticsNotProven' in mapped.operations[0], false);
  assert.deepEqual(mapped.preservation, { scope: null }, 'the L3/preservation axis is untouched by the L1/write-semantics one');
});

test('receipt: on a pre-/14 accepted runtime, insert_shape+stroke names BOTH colour axes, additive to (never merged with) preservation.notProven', async () => {
  const store = createStore(book());
  const t = { id: 'x-o08', steps: [shapeStep({ stroke: { width: 3, color: '#C81E1E' } })] };
  const { result } = await run(store, t, { colourUnobservable: true });
  const mapped = mapExecutionResult(result, { fileId: 'disposable', task: t });
  assert.deepEqual(mapped.operations[0].semanticsNotProven, ['shape-fill-colour', 'shape-stroke-colour']);
  assert.deepEqual(mapped.preservation, { scope: null }, 'the L3/preservation axis is untouched by the L1/write-semantics one');
});

test('receipt: insert_shape WITHOUT stroke, pre-/14 accepted runtime -- names only fill-colour', async () => {
  const store = createStore(book());
  const t = { id: 'x-o08', steps: [shapeStep()] };
  const { result } = await run(store, t, { colourUnobservable: true });
  const mapped = mapExecutionResult(result, { fileId: 'disposable', task: t });
  assert.deepEqual(mapped.operations[0].semanticsNotProven, ['shape-fill-colour']);
});

test('receipt: insert_image carries NO semanticsNotProven key at all (no colour parameter, no family check emits one)', async () => {
  const store = createStore(book());
  const t = { id: 'x-o08', steps: [imageStep()] };
  const { result } = await run(store, t);
  const mapped = mapExecutionResult(result, { fileId: 'disposable', task: t });
  assert.equal('semanticsNotProven' in mapped.operations[0], false);
});

test('receipt: an UNRELATED family (write_range) carries NO semanticsNotProven key -- the new field is opt-in per family, not a global default', async () => {
  const store = createStore(book());
  const t = { id: 'x', steps: [{ intent: 'write_range', sheet: 'D', range: 'B1', values: [['x']] }] };
  const { result } = await run(store, t);
  const mapped = mapExecutionResult(result, { fileId: 'disposable', task: t });
  assert.equal('semanticsNotProven' in mapped.operations[0], false);
});

test('receipt: a RETRY (SATISFIED, 0 writes) on a pre-/14 accepted runtime still names the colour limit -- the limit is about the FIELD\'s observability on THIS call, not the write event', async () => {
  const store = createStore(book());
  const t = { id: 'x-o08', steps: [shapeStep({ stroke: { width: 3, color: '#C81E1E' } })] };
  await run(store, t, { colourUnobservable: true });
  const { result } = await run(store, t, { colourUnobservable: true });
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  const mapped = mapExecutionResult(result, { fileId: 'disposable', task: t });
  assert.deepEqual(mapped.operations[0].semanticsNotProven, ['shape-fill-colour', 'shape-stroke-colour']);
});
