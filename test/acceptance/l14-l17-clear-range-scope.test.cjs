'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const clear = require('../../src/adapters/xlsx/families/clear.cjs');
const W = require('../../src/server/write-tools.cjs');
const { req } = require('../fakes/write-wiring.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * L-14 / L-17 (owner decision "A"; LIVE: test/live/xlsx-l14-l17-clear-formats-probe.live.cjs -- ClearFormats() resets the format
 * axes and KEEPS the value, Clear() resets both, ClearContents() equals Clear(), both dissolve a merge): clear_range takes an optional `scope`.
 * 'contents' (the default) is the existing behaviour and stays byte-identical; 'formats' / 'all' refuse any merged area by name.
 */

const rich = { v: 'old', nf: '0.00', fill: 16776960, bold: true, italic: true, wrap: true, rot: 45 };
// GH #141 P1 (clear-formats-l1-teljes-alap): a cell carrying the SEVEN axes BASE_AXES newly
// asserts (border, align, readingOrder, indent, shrink, font.underline, font.strike) -- none of
// these were ever off-base on A1/A2/B1 above, so a regression on any of them would have been
// silently invisible to the pre-existing L-14/L-17 suite.
const g1Axes = { border: { t: { s: 1, c: 16711680 } }, align: { h: 1, v: 4 }, readingOrder: 1, indent: 3, shrink: true, font: { underline: 'xlUnderlineStyleSingle', strike: true } };
const book = () => [
  { name: 'Data', cells: {
    A1: { ...rich }, A2: { ...rich, v: 7 }, B1: { v: 'ring', fill: 255, bold: true },
    C1: { v: 'plain' }, D1: { v: null, f: '=1+1', nf: '0.00' }, G1: { v: 'g', ...g1Axes },
  }, merges: ['E1:F1'] },
];
const task = (scope, range = 'A1:A2', sheet = 'Data') => ({ id: 'l14', steps: [{ intent: 'clear_range', sheet, range, ...(scope ? { scope } : {}) }] });
const cell = (store, a) => store.durable.find(s => s.name === 'Data').cells[a];
const op = result => result.execution.appliedOperations[0];

async function run(store, t, script = {}) {
  let n = 0;
  const bodies = [];
  const driver = createFakeXlsxDriver(store, script);
  const dispatch = driver.dispatch.bind(driver);
  driver.dispatch = async (h, c) => { bodies.push(c && c.body); return dispatch(h, c); };
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, bodies, m: result.metrics };
}

test("L-14: scope 'formats' -- APPLIED, formats back to base, VALUE kept, neighbour untouched; retry SATISFIED 0 commands", async () => {
  const store = createStore(book());
  const a = await run(store, task('formats'));
  assert.equal(a.result.ok, true, JSON.stringify(op(a.result).evidence));
  assert.equal(op(a.result).status, OPERATION_STATUS.APPLIED);
  assert.equal(a.m.commands, 1);
  assert.match(a.bodies[0], /GetRange\("A1:A2"\)\.ClearFormats\(\)/);
  assert.doesNotMatch(a.bodies[0], /SetValue|\.Clear\(\)/);
  assert.equal(cell(store, 'A1').v, 'old');
  assert.equal(cell(store, 'A2').v, 7);
  for (const k of ['nf', 'fill', 'bold', 'italic', 'wrap', 'rot']) assert.equal(cell(store, 'A1')[k], undefined, k);
  assert.equal(cell(store, 'B1').fill, 255, 'the ring cell keeps its own formatting');
  assert.deepEqual(a.driver.reads[0].windows, [{ sheet: 'Data', range: 'A1:A2', style: true }], 'the window is read WITH style');
  const again = await run(store, task('formats'));
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.m.commands, 0);
});

test("L-17: scope 'all' -- APPLIED, value AND formats gone; retry SATISFIED", async () => {
  const store = createStore(book());
  const a = await run(store, task('all'));
  assert.equal(a.result.ok, true, JSON.stringify(op(a.result).evidence));
  assert.match(a.bodies[0], /GetRange\("A1:A2"\)\.Clear\(\)/);
  assert.doesNotMatch(a.bodies[0], /ClearContents|ClearFormats|SetValue/);
  assert.equal(cell(store, 'A1').v, null);
  assert.equal(cell(store, 'A1').fill, undefined);
  assert.equal(op((await run(store, task('all'))).result).status, OPERATION_STATUS.SATISFIED);
});

test("a formula cell keeps its formula under 'formats' and loses it under 'all'", async () => {
  const f = createStore(book());
  assert.equal((await run(f, task('formats', 'D1'))).result.ok, true);
  assert.equal(cell(f, 'D1').f, '=1+1');
  assert.equal(cell(f, 'D1').nf, undefined);
  const al = createStore(book());
  assert.equal((await run(al, task('all', 'D1'))).result.ok, true);
  assert.ok(!cell(al, 'D1').f);
});

test("'contents' (explicit or default) is UNCHANGED: SetValue('') only, formats kept", async () => {
  for (const t of [task('contents'), task(null)]) {
    const store = createStore(book());
    const a = await run(store, t);
    assert.equal(a.result.ok, true);
    assert.match(a.bodies[0], /SetValue\(''\)/);
    assert.doesNotMatch(a.bodies[0], /ClearFormats|\.Clear\(\)/);
    assert.equal(cell(store, 'A1').v, null);
    assert.equal(cell(store, 'A1').fill, 16776960);
  }
  assert.equal(clear.lower(planOperations(task(null)).operations[0]).body, clear.lower(planOperations(task('contents')).operations[0]).body);
});

for (const scope of ['formats', 'all']) {
  for (const range of ['E1', 'F1', 'E1:F1', 'D1:E1']) {
    test(`'${scope}' on a range touching a merged area (${range}) -- refused by name, 0 dispatch, nothing written`, async () => {
      const store = createStore(book());
      const a = await run(store, task(scope, range));
      assert.equal(a.driver.calls.dispatch, 0);
      assert.equal(op(a.result).evidence.evidence.reason, 'xlsx-clear-scope-dissolves-merge');
      assert.deepEqual(store.durable, book());
    });
  }
}

test("'contents' on a merged area keeps its own refusal name", async () => {
  const a = await run(createStore(book()), task('contents', 'E1'));
  assert.equal(op(a.result).evidence.evidence.reason, 'xlsx-clear-in-merged-area-refused');
});

test('RED mutants: a formats clear that erases the value, ignores the call, keeps an axis or leaks elsewhere is NOT APPLIED, nothing persisted', async () => {
  for (const [script, check] of [
    [{ fault: 'clear-formats-erases-value' }, pv => assert.deepEqual(pv.evidence.evidence.valueChanged, ['A1', 'A2'])],
    [{ fault: 'ignore' }, pv => assert.ok(pv.evidence.evidence.offBase.length > 0)],
    [{ fault: 'clear-formats-keeps-axis', keepAxis: 'fill' }, pv => assert.ok(pv.evidence.evidence.offBase.every(x => x.axes.includes('fill')))],
  ]) {
    const store = createStore(book());
    const a = await run(store, task('formats'), script);
    assert.equal(a.result.ok, false, JSON.stringify(script));
    assert.notEqual(op(a.result).status, OPERATION_STATUS.APPLIED);
    check(a.result.execution.primitiveVerification[0]);
    assert.equal(a.driver.calls.forceSave, 0);
    assert.equal(cell(store, 'A1').v, 'old', 'durable untouched');
  }
  const side = await run(createStore(book()), task('formats'), { fault: 'side-effect' });
  assert.equal(side.result.ok, false);
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test("RED mutant for 'all': a clear that leaves the value fails L1", async () => {
  const a = await run(createStore(book()), task('all'), { fault: 'ignore' });
  assert.equal(a.result.ok, false);
  assert.equal(a.result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('fail closed when the style of a target cell is unobservable (no write)', async () => {
  const store = createStore(book());
  const a = await run(store, task('formats'), { styleGetterThrows: false, rotUnobservable: ['A1'] });
  assert.equal(a.driver.calls.dispatch, 0);
  assert.equal(op(a.result).evidence.evidence.reason, 'xlsx-target-style-unobservable');
});

// GH #141 P1 (clear-formats-l1-teljes-alap-20261017): the seven axes added to BASE_AXES.
test("L-14 widened: scope 'formats' on a cell carrying border/align/readingOrder/indent/shrink/font-effects -- ALL seven back to base, value kept", async () => {
  const store = createStore(book());
  const a = await run(store, task('formats', 'G1', 'Data'));
  assert.equal(a.result.ok, true, JSON.stringify(op(a.result).evidence));
  assert.equal(op(a.result).status, OPERATION_STATUS.APPLIED);
  assert.equal(cell(store, 'G1').v, 'g');
  for (const k of ['nf', 'fill', 'bold', 'italic', 'wrap', 'rot', 'border', 'align', 'readingOrder', 'indent', 'shrink', 'font']) assert.equal(cell(store, 'G1')[k], undefined, k);
  const again = await run(store, task('formats', 'G1', 'Data'));
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
});

test('RED mutants, the seven GH #141 axes: a formats clear that keeps border/align/readingOrder/indent/shrink/underline/strike is NOT APPLIED, named on that axis -- proves the new BASE_AXES entries actually discriminate, not pass by construction', async () => {
  for (const axis of ['border', 'align', 'readingOrder', 'indent', 'shrink']) {
    const store = createStore(book());
    const a = await run(store, task('formats', 'G1', 'Data'), { fault: 'clear-formats-keeps-axis', keepAxis: axis });
    assert.equal(a.result.ok, false, axis);
    assert.notEqual(op(a.result).status, OPERATION_STATUS.APPLIED, axis);
    const pv = a.result.execution.primitiveVerification[0];
    assert.ok(pv.evidence.evidence.offBase.every(x => x.axes.includes(axis)), `${axis}: ${JSON.stringify(pv.evidence.evidence.offBase)}`);
  }
  // underline and strike share ONE reader (font) -- keeping `font` itself keeps BOTH off base,
  // the same "one capability covers several axes" shape align.h/v and border already have.
  const store = createStore(book());
  const a = await run(store, task('formats', 'G1', 'Data'), { fault: 'clear-formats-keeps-axis', keepAxis: 'font' });
  assert.equal(a.result.ok, false);
  const pv = a.result.execution.primitiveVerification[0];
  assert.ok(pv.evidence.evidence.offBase.every(x => x.axes.includes('underline') && x.axes.includes('strike')), JSON.stringify(pv.evidence.evidence.offBase));
});

test("RED CONTROL: a cell that was ALREADY at base on the seven new axes (A1, never touched there) is unaffected by the widening -- 'formats' still APPLIES cleanly, proving the new checks are not a false-positive trap on ordinary cells", async () => {
  const store = createStore(book());
  const a = await run(store, task('formats'));
  assert.equal(a.result.ok, true, JSON.stringify(op(a.result).evidence));
  assert.equal(op(a.result).status, OPERATION_STATUS.APPLIED);
});

test('planning: scope must be one of the three; the WRITE tool schema agrees', () => {
  assert.equal(planOperations(task('formats')).ok, true);
  assert.equal(planOperations(task('all')).ok, true);
  for (const bad of ['format', 'ALL', '', 7, null, ['all']]) assert.equal(planOperations({ id: 'x', steps: [{ intent: 'clear_range', sheet: 'Data', range: 'A1', scope: bad }] }).ok, false, JSON.stringify(bad));
  const validate = s => W.validateWriteRequest(req([{ intent: 'clear_range', id: 's', sheet: 'Data', range: 'A1', ...s }]));
  assert.equal(validate({}).ok, true);
  for (const scope of ['contents', 'formats', 'all']) assert.equal(validate({ scope }).ok, true, scope);
  for (const scope of ['format', 'ALL', 7, null]) assert.equal(validate({ scope }).ok, false, JSON.stringify(scope));
  assert.equal(validate({ what: 'all' }).ok, false);
});
