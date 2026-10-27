'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const charts = require('../../src/adapters/xlsx/families/charts.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * X-O07/3 (internal tracker; X-O07/2 LIVE MEASURED and reviewed) -- `delete_chart`, a NEW
 * standalone intent (not a set_chart_style target) through the FULL product stack (plan ->
 * classify -> lower -> dispatch -> verify), proving the mechanism end-to-end against the fake
 * driver's own simulated round-trip. Same loading convention as `xlsx-xo05-3-chart-style-gap-
 * overlap-series-order.test.cjs` (its own B22 runtime-gate pattern, mirrored here for B23) and
 * `xlsx-x-o06-move-chart.test.cjs` (its own standalone-intent / drawing-diff shape). This is the
 * offline counterpart to the LIVE product harness (release-xo07-3-delete-chart-product.live.cjs,
 * RED until the DS /13 install) -- this file measures that OUR OWN wiring is correct,
 * independent of it.
 */

const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: extra, pivots: [], names: [] } },
];
const mine = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', legend: 'bottom', series: series('K', ['B', 'C']), pos: { fromCol: 2, colOffset: 0, fromRow: 12, rowOffset: 0 }, w: 3600000, h: 2160000, ...over });
const del = (over = {}) => ({ id: 'x-o07-3', steps: [{ intent: 'delete_chart', sheet: 'K', name: 'C1', ...over }] });

// X-O07/3: delete_chart is B23-WRITE (classifyDelete's own runtime gate) -- every test in this file
// needs the fake session opened on an ACCEPTED runtime, same convention the X-O05/3 file uses.
// DS9346 (owner D5=B, docs/DS9346-RUNTIME-RETIREMENT.md): the historical /13 (9.3.4.60) is retired --
// the accepted runtime is now the single 9.3.5.0/16 identity.
const RT935 = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT935, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;
const chartsOf = (store, s) => store.durable.find(x => x.name === s).objects.charts;

test('delete_chart: APPLIED, the chart is gone; retry SATISFIED 0/0 (already absent -- the end state already matches the goal)', async () => {
  const store = createStore(book([mine(), mine({ name: 'C2', pos: { fromCol: 20, colOffset: 0, fromRow: 20, rowOffset: 0 } })]));
  const a = await run(store, del());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(chartsOf(store, 'K').find(c => c.name === 'C1'), undefined);
  assert.equal(chartsOf(store, 'K').find(c => c.name === 'C2').name, 'C2', 'the OTHER chart is untouched');
  const r = await run(store, del());
  assert.equal(r.result.ok, true, JSON.stringify(r.result.execution.appliedOperations[0].evidence));
  assert.equal(r.m.commands, 0, 'SATISFIED -- no dispatch needed, the name is already absent');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('delete_chart: a name that never existed -- SATISFIED 0/0 immediately, no dispatch at all', async () => {
  const { result, driver } = await run(createStore(book()), del());
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('refused BEFORE dispatch: two charts of the name -- xlsx-chart-name-ambiguous, never guesses which one, nothing removed', async () => {
  const store = createStore(book([mine(), mine()]));
  const { result, driver } = await run(store, del());
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-name-ambiguous');
  assert.equal(chartsOf(store, 'K').length, 2, 'an ambiguous name must not delete ANY candidate');
});

test('refused BEFORE dispatch: a pre-/13 runtime -- xlsx-chart-delete-bridge-write-runtime-not-proven, named, never a silent attempt', async () => {
  const store = createStore(book([mine()]));
  const driver = createFakeXlsxDriver(store, { runtime: '9.3.4|euro-office-bridge-v1/9.3.4.60/12' });
  const options = { driver, fileId: 'disposable', now: () => 1, ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(del(), { observation: { live: liveScope(del()) } }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-delete-bridge-write-runtime-not-proven');
  assert.equal(chartsOf(store, 'K').length, 1, 'nothing removed on a refused runtime');
});

test('RED: dispatch reports success but the chart is still there (fault "ignore") -- L1 FAIL, never a silent success', async () => {
  const { result, driver } = await run(createStore(book([mine()])), del(), { fault: 'ignore' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
});

test('RED: deleting a chart alters an OTHER chart -- L3 FAIL (the generic "chart-alters-other" fault, same as every other chart intent)', async () => {
  const { result } = await run(createStore(book([mine(), mine({ name: 'C2', pos: { fromCol: 20, colOffset: 0, fromRow: 20, rowOffset: 0 } })])), del(), { fault: 'chart-alters-other' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('L3: the deleted chart\'s OWN drawing record is EXPECTED to disappear; an unrelated chart\'s drawing stays exactly compared', async () => {
  const { result } = await run(createStore(book([mine(), mine({ name: 'C2', pos: { fromCol: 20, colOffset: 0, fromRow: 20, rowOffset: 0 } })])), del());
  assert.equal(result.ok, true, JSON.stringify(result.execution.preservationVerification[0]));
});

test('planner: field validation', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  // L-21 (main): the family planner's sheet-name rule accepts a name that needs quoting; the product
  // constraint (write-tools CONSTRAINTS.delete_chart: bareNameProblem) still refuses it before dispatch.
  assert.equal(planOperations({ id: 'x', steps: [{ intent: 'delete_chart', sheet: 'My Sheet', name: 'C1' }] }).ok, true);
  refuse({ id: 'x', steps: [{ intent: 'delete_chart', sheet: 'K', name: 'C1', extra: 1 }] }, /does not support field/);
  refuse({ id: 'x', steps: [{ intent: 'delete_chart', sheet: 'K' }] }, /./);
  refuse({ id: 'x', steps: [{ intent: 'delete_chart', name: 'C1' }] }, /./);
});

test('plan(): the selector claims the chart (the generic composition-conflict detector picks it up, same shape every other chart intent uses) -- no newName claim, cellFree', () => {
  const p = planOperations(del());
  assert.equal(p.ok, true, JSON.stringify(p));
  const op = p.operations[0];
  assert.deepEqual([...op.selector.claims], ['chart:K:C1']);
  assert.equal(op.selector.cellFree, true);
  assert.equal(op.newName, null);
});

test('the command body: a Bridge WRITE gate + ONE write(), no raw Api chart-identity lookup in-body (the handler resolves sheet+chart server-side, same convention gapOverlap/seriesOrder use)', () => {
  const p = planOperations(del());
  const lowered = charts.lower(p.operations[0]);
  assert.equal(lowered.meta.bridgeChartDeleteWrite, true);
  assert.equal(lowered.meta.family, 'charts');
  assert.equal(lowered.meta.action, 'delete_chart');
  assert.match(lowered.body, /__eoChartDeleteGate\(name, "C1"\)/);
  assert.match(lowered.body, /__eoChartDeleteWrite\(\)/);
  assert.match(lowered.body, /__eoChartDeleteFail\(w\)/);
  assert.doesNotMatch(lowered.body, /GetAllCharts/, 'no raw chart-identity lookup -- the bridge resolves it');
});
