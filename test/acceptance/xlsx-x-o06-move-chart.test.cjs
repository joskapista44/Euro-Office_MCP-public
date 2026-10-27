'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const charts = require('../../src/adapters/xlsx/families/charts.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** X-O06-6 (FORMATTING COMPLETION backlog): move_chart on the fake. Mirrors
 * xlsx-w7-charts.test.cjs's own shape -- retry/refusal tables, RED faults, a command-body unit
 * test, and an explicit check that another chart's SIZE is NOT exact-compared after a move
 * while an unrelated chart's size still is (the `sizeBlind` drawings.cjs extension). */

const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: extra, pivots: [], names: [] } },
];
const mine = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', series: series('K', ['B', 'C']), pos: { fromCol: 2, colOffset: 0, fromRow: 12, rowOffset: 0 }, w: 3600000, h: 2160000, ...over });
const move = (over = {}) => ({ id: 'x-o06', steps: [{ intent: 'move_chart', sheet: 'K', name: 'C1', at: { col: 8, row: 20 }, ...over }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;
const chartOf = (store, s, name) => store.durable.find(x => x.name === s).objects.charts.find(c => c.name === name);

test('move_chart: APPLIED to the requested cell, retry SATISFIED 0/0', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, move());
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.appliedOperations[0].evidence]));
  assert.equal(a.m.commands, 1);
  assert.deepEqual(chartOf(store, 'K', 'C1').pos, { fromCol: 8, colOffset: 0, fromRow: 20, rowOffset: 0 });
  const r = await run(store, move());
  assert.equal(r.result.ok, true, JSON.stringify(r.result.execution.appliedOperations[0].evidence));
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('move_chart: the moved chart\'s own ±1 EMU size drift does NOT fail L3 (sizeBlind)', async () => {
  const store = createStore(book([mine()]));
  const { result } = await run(store, move(), { fault: 'move-chart-size-drifts' });
  assert.equal(result.ok, true, JSON.stringify(result.execution.preservationVerification));
});

for (const [name, t, why, b] of [
  ['the chart is not found', move(), 'xlsx-chart-not-found', book()],
  ['two charts of the name', move(), 'xlsx-chart-name-ambiguous', book([mine(), mine()])],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const { result, driver } = await run(createStore(b), t);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('refused BEFORE dispatch: an unobservable drawing reader', async () => {
  const { result, driver } = await run(createStore(book([mine()])), move(), { drawingsUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-charts-drawings-unobservable');
});

test('RED: ignored / wrong position -> L1 FAIL', async () => {
  for (const fault of ['ignore', 'move-chart-wrong-position']) {
    const { result, driver } = await run(createStore(book([mine()])), move(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

test('RED: moving a chart alters an OTHER chart -> L3 FAIL', async () => {
  const { result } = await run(createStore(book([mine(), mine({ name: 'C2', pos: { fromCol: 20, colOffset: 0, fromRow: 20, rowOffset: 0 } })])), move(), { fault: 'chart-alters-other' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('planner: field validation', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse({ id: 'x', steps: [{ intent: 'move_chart', sheet: "O'Brien", name: 'C1', at: { col: 1, row: 1 } }] }, /sheet-name-apostrophe-unsupported/);
  refuse({ id: 'x', steps: [{ intent: 'move_chart', sheet: 'K', name: 'C1', at: { col: -1, row: 1 } }] }, /col must be a non-negative integer/);
  refuse({ id: 'x', steps: [{ intent: 'move_chart', sheet: 'K', name: 'C1', at: { col: 1.5, row: 1 } }] }, /col must be a non-negative integer/);
  refuse({ id: 'x', steps: [{ intent: 'move_chart', sheet: 'K', name: 'C1', at: { col: 1, row: -1 } }] }, /row must be a non-negative integer/);
  refuse({ id: 'x', steps: [{ intent: 'move_chart', sheet: 'K', name: 'C1', at: { col: 1, row: 1 }, colOffset: 5 }] }, /does not support field/);
  refuse({ id: 'x', steps: [{ intent: 'move_chart', sheet: 'K', name: 'C1' }] }, /at must be an object/);
});

test('the command body: SetPosition(col, 0, row, 0) -- the measured call shape; re-identified at dispatch', () => {
  const log = [];
  const existing = [{ GetName: () => 'C1', SetPosition: (...x) => log.push(['SetPosition', ...x]) }];
  const ws = { GetName: () => 'K', GetAllCharts: () => existing };
  const Api = { GetSheet: () => ws };
  assert.deepEqual(new Function('Api', charts.lower(planOperations(move()).operations[0]).body)(Api), { ok: true });
  assert.deepEqual(log, [['SetPosition', 8, 0, 20, 0]]);
  log.length = 0;
  existing.length = 0;
  assert.deepEqual(new Function('Api', charts.lower(planOperations(move()).operations[0]).body)(Api), { ok: false, stage: 'chart-identity' });
});
