'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const charts = require('../../src/adapters/xlsx/families/charts.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { mapExecutionResult } = require('../../src/server/write-tools.cjs');

/**
 * X-O05/3 (internal tracker; design reviewed and accepted) -- `set_chart_style`'s TWO more
 * targets, `gapOverlap` and `seriesOrder`, through the FULL product stack (plan -> classify ->
 * lower -> dispatch -> verify), proving the mechanism end-to-end against the fake driver's own
 * simulated round-trip (test/fakes/fake-xlsx-driver.cjs's gapOverlap/order fields). This is the
 * offline counterpart to the LIVE harness zero asked for (2-series swap + 3-series cycle, the
 * engine's own intermediate-collision behavior) -- that harness measures the REAL DS /12 runtime;
 * this file measures that OUR OWN plan/classify/lower/verify wiring is correct, independent of it.
 * Same loading convention as `xlsx-xo04-11-chart-style-marker-gridlines.test.cjs`.
 */

const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, D1: { v: 's3' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, D2: { v: 7 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, D3: { v: 8 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 }, D4: { v: 9 } }, objects: { charts: extra, pivots: [], names: [] } },
];
const mine = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', legend: 'bottom', series: series('K', ['B', 'C']), ...over });
const gapOverlap = (over = {}) => ({ id: 'x-o05-3-go', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'gapOverlap', gapWidth: 219, overlap: -26, ...over }] });
const seriesOrder = (order, over = {}) => ({ id: 'x-o05-3-so', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'seriesOrder', order, ...over }] });

// X-O05/3: gapOverlap/seriesOrder are B22-WRITE (classifyStyle's own runtime gate) -- UNLIKE
// marker/gridlines (no Bridge WRITE, no runtime gate), every test in this file needs the fake
// session opened on an ACCEPTED runtime, same convention as the B05/B10/B03 write-cutover tests.
// DS9346 (owner D5=B, docs/DS9346-RUNTIME-RETIREMENT.md): the historical /12 (9.3.4.60) is
// retired -- the accepted runtime is now the single 9.3.5.0/16 identity.
const RT935 = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT935, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics, mapped: mapExecutionResult(result, { fileId: 'disposable', task: t }) };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

// ---- gapOverlap ------------------------------------------------------------------------------

test('gapOverlap: APPLIED, retry SATISFIED -- both fields together', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, gapOverlap());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, gapOverlap());
  assert.equal(r.result.ok, true);
  assert.equal(r.m.commands, 0, 'SATISFIED -- no dispatch needed, the bridge already reads the requested value');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('gapOverlap: gapWidth only / overlap only -- the unrequested field is NOT part of the match', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, gapOverlap({ gapWidth: 100, overlap: undefined }));
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, gapOverlap({ gapWidth: 100, overlap: undefined }));
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED, 'overlap was never requested, so its drifting value does not block SATISFIED');
});

test('gapOverlap: refused on a non-bar chart -- named, never attempted', async () => {
  const { result, driver } = await run(createStore(book([mine({ type: 'lineNormal' })])), gapOverlap());
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-gap-overlap-type-unsupported');
});

test('gapOverlap: bridge UNOBSERVABLE (pre-/12 runtime) -- FAIL_CLOSED, named, never a silent write-only fallback', async () => {
  const { result, driver } = await run(createStore(book([mine()])), gapOverlap(), { chartStyleUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-bridge-unobservable');
});

test('gapOverlap: receipt carries NO semanticsNotProven -- verified, not write-only', async () => {
  const store = createStore(book([mine()]));
  const { mapped } = await run(store, gapOverlap());
  assert.deepEqual(mapped.operations[0].semanticsNotProven, undefined);
});

test('RED (sabotage): gapOverlap dispatch succeeds but the bridge-read value stays stale -- FAIL, never a silent success', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, gapOverlap(), { fault: 'style-bridge-stale' });
  assert.equal(a.result.ok, false);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  const r = await run(store, gapOverlap());
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, 'a clean retry still applies correctly');
});

test('the command body: a direct chart.bar.gapOverlap.write/v1 Bridge call, gated by runtime, only the requested field(s)', () => {
  const op = planOperations(gapOverlap()).operations[0];
  const { body, meta } = charts.lower(op);
  // the capability ID itself is on `meta.bridgeChartStyleWrite`, not in the body text -- the
  // driver's dispatch() resolves it from there and PREPENDS chartStyleBridgeWritePrelude(cap),
  // the SAME split bridgeCopyWrite/bridgeAlignExtWrite/bridgePrintWrite already use.
  assert.equal(meta.bridgeChartStyleWrite, 'chart.bar.gapOverlap.write/v1');
  assert.match(body, /__eoChartStyleGate/);
  assert.match(body, /"gapWidth":219/);
  assert.match(body, /"overlap":-26/);
});

// ---- seriesOrder ------------------------------------------------------------------------------

test('seriesOrder: a 2-series direct swap -- APPLIED, retry SATISFIED', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, seriesOrder([1, 0]));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, seriesOrder([1, 0]));
  assert.equal(r.m.commands, 0, 'SATISFIED -- no dispatch needed');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

// LIVE on 1.0.1-9e03483b93ef the chart.series.order.read/v1 value reached the family as the object { order: n }
// (read_charts' own per-series shape), not a bare number: the applied swap then read back as
// VERIFY_CANNOT_PROVE_TERMINAL_STATE. The fake serves that live shape here (orderReadAsObject).
test('seriesOrder: the LIVE read shape { order: n } -- APPLIED, retry SATISFIED (regression, 1.0.1-9e03483b93ef)', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, seriesOrder([1, 0]), { orderReadAsObject: true });
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, seriesOrder([1, 0]), { orderReadAsObject: true });
  assert.equal(r.m.commands, 0, 'SATISFIED -- no dispatch needed');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('seriesOrder: a 3-series cycle -- APPLIED, retry SATISFIED', async () => {
  const store = createStore(book([mine({ series: series('K', ['B', 'C', 'D']) })]));
  const a = await run(store, seriesOrder([1, 2, 0]));
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  const r = await run(store, seriesOrder([1, 2, 0]));
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('seriesOrder: a control chart (not named by the step) stays untouched by a swap on another chart', async () => {
  const store = createStore(book([mine(), mine({ name: 'C2' })]));
  await run(store, seriesOrder([1, 0]));
  const r = await run(store, seriesOrder([1, 0], { name: 'C2' }));
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, 'C2 had its own, independent order before this call');
});

test('seriesOrder: length mismatch against the chart\'s ACTUAL series count -- named, never attempted', async () => {
  const { result, driver } = await run(createStore(book([mine()])), seriesOrder([1, 2, 0]), {});
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-series-order-length-mismatch');
});

test('seriesOrder: a non-permutation (planner-level, self-contained shape check) refuses BEFORE any live data is needed', () => {
  const p = planOperations(seriesOrder([0, 0]));
  assert.equal(p.ok, false);
  assert.match(p.detail, /order must not repeat a value/);
});

test('seriesOrder: bridge UNOBSERVABLE (pre-/12 runtime) -- FAIL_CLOSED, named', async () => {
  const { result, driver } = await run(createStore(book([mine()])), seriesOrder([1, 0]), { chartStyleUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-bridge-unobservable');
});

test('seriesOrder: receipt carries NO semanticsNotProven -- verified, not write-only', async () => {
  const store = createStore(book([mine()]));
  const { mapped } = await run(store, seriesOrder([1, 0]));
  assert.deepEqual(mapped.operations[0].semanticsNotProven, undefined);
});

test('RED (sabotage): seriesOrder dispatch succeeds but the bridge-read order stays stale -- FAIL, never a silent success', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, seriesOrder([1, 0]), { fault: 'style-bridge-stale' });
  assert.equal(a.result.ok, false);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
});

test('the command body: a single chart.series.order.write/v1 Bridge call carrying the WHOLE permutation, never a per-series seriesIndex arg', () => {
  const op = planOperations(seriesOrder([1, 2, 0])).operations[0];
  const { body, meta } = charts.lower(op);
  assert.equal(meta.bridgeChartStyleWrite, 'chart.series.order.write/v1');
  assert.match(body, /__eoChartStyleGate/);
  assert.match(body, /"order":\[1,2,0\]/);
  assert.doesNotMatch(body, /seriesIndex/);
});

// ---- planner / field validation (charts.cjs's own styleProblem, independent of write-tools.cjs) ---

test('planner: gapOverlap field validation', () => {
  const refuse = (over, re) => { const p = planOperations(gapOverlap(over)); assert.equal(p.ok, false, JSON.stringify(over)); assert.match(p.detail, re, JSON.stringify(over)); };
  refuse({ fill: '#808080' }, /gapOverlap has no fill\/outline/);
  refuse({ axis: 'vertical' }, /axis is only accepted for target gridlines/);
  refuse({ seriesIndex: 0 }, /seriesIndex\/dataPointIndex are not accepted for target gapOverlap/);
  refuse({ gapWidth: undefined, overlap: undefined }, /at least one of gapWidth\/overlap is required/);
  refuse({ gapWidth: 501 }, /gapWidth must be a number between 0 and 500/);
  refuse({ overlap: 101 }, /overlap must be a number between -100 and 100/);
});

test('planner: seriesOrder field validation', () => {
  const refuse = (over, re) => { const p = planOperations(seriesOrder(over.order !== undefined ? over.order : [0, 1], over)); assert.equal(p.ok, false, JSON.stringify(over)); assert.match(p.detail, re, JSON.stringify(over)); };
  refuse({ fill: '#808080' }, /seriesOrder has no fill\/outline/);
  refuse({ axis: 'vertical' }, /axis is only accepted for target gridlines/);
  refuse({ seriesIndex: 0 }, /seriesIndex\/dataPointIndex are not accepted for target seriesOrder/);
  refuse({ order: [] }, /order must be a non-empty array/);
  refuse({ order: [0, 2] }, /order must contain only integers in 0\.\.length-1/);
  refuse({ order: [1, 1] }, /order must not repeat a value/);
});
