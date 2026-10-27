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
 * X-O04/11 (card an internal branch, kioszto: zero) --
 * `set_chart_style`'s TWO new targets, `marker` and `gridlines`, SUPERSEDING the X-O04/1+2 "measured
 * non-functional" finding (round 1, bAllMarkers/bAll*=false). With bAllMarkers/bAll*=true the
 * written colour WORKS_PERSISTED (X-O04/4 round 2 + X-O04/8 LIVE), proven end-to-end through the
 * X-O04/10 DS Bridge /11 READ round-trip (chart.series.markerStyle.read/v1 /
 * chart.axis.majorGridlinesColor.read/v1) -- so UNLIKE the six original targets, these two DO
 * read-and-compare (never WRITE_ONLY, never semanticsNotProven). The fake driver's chartReadsOf
 * (test/fakes/fake-xlsx-driver.cjs) simulates this round-trip: a successful set_chart_style dispatch
 * records the written colour onto the SAME fields a real chartReads carries.
 */

const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: extra, pivots: [], names: [] } },
];
const mine = (over = {}) => ({ name: 'C1', type: 'lineNormal', title: 'Sales', legend: 'bottom', series: series('K', ['B', 'C']), ...over });
const style = (over = {}) => ({ id: 'x-o04-11', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'marker', seriesIndex: 0, fill: '#808080', ...over }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics, mapped: mapExecutionResult(result, { fileId: 'disposable', task: t }) };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

// ---- marker ------------------------------------------------------------------------------------

test('marker: APPLIED, retry SATISFIED (the bridge read now matches -- unlike the six original targets, this one CAN be SATISFIED)', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, style());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, style());
  assert.equal(r.result.ok, true);
  assert.equal(r.m.commands, 0, 'SATISFIED -- no dispatch needed, the bridge already reads the requested colour');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('marker: works on scatter too (the other LIVE-proven type), fill+outline together', async () => {
  const store = createStore(book([mine({ type: 'scatter' })]));
  const a = await run(store, style({ outline: { width: 1, color: '#C81E1E' } }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  const r = await run(store, style({ outline: { width: 1, color: '#C81E1E' } }));
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('marker: refused on an untested chart type (bar) -- named, never attempted', async () => {
  const { result, driver } = await run(createStore(book([mine({ type: 'bar' })])), style());
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-marker-type-unsupported');
});

test('marker: seriesIndex out of range -- the SAME bounds check series/dataPoint already use', async () => {
  const { result, driver } = await run(createStore(book([mine()])), style({ seriesIndex: 2 }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-series-index-out-of-range');
});

test('marker: bridge UNOBSERVABLE (pre-/11 runtime, or any other read gap) -- FAIL_CLOSED, named, never a silent write-only fallback', async () => {
  const { result, driver } = await run(createStore(book([mine()])), style(), { chartStyleUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-bridge-unobservable');
});

test('marker: receipt carries NO semanticsNotProven -- unlike the six original targets, this one is verified, not write-only', async () => {
  const store = createStore(book([mine()]));
  const { mapped } = await run(store, style());
  assert.deepEqual(mapped.operations[0].semanticsNotProven, undefined);
});

test('RED (sabotage): dispatch succeeds but the bridge-read value stays stale -- DISPATCHED_OUTCOME_UNKNOWN / L1 FAIL, never a silent success', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, style(), { fault: 'style-bridge-stale' });
  assert.equal(a.result.ok, false);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  // a clean retry (no fault) still applies correctly
  const r = await run(store, style());
  assert.equal(r.result.ok, true);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
});

test('the command body: SetMarkerFill/SetMarkerOutLine -- extraArgs (seriesIndex, 0, true); the return value is NOT checked (no stage on a false return)', () => {
  const log = [];
  const existing = [{ GetName: () => 'C1', SetMarkerFill: (...x) => { log.push(['SetMarkerFill', ...x]); return false; }, SetMarkerOutLine: (...x) => { log.push(['SetMarkerOutLine', ...x]); return false; } }];
  const ws = { GetName: () => 'K', GetAllCharts: () => existing };
  const Api = { GetSheet: () => ws, CreateSolidFill: v => v, CreateRGBColor: (...x) => x, CreateStroke: (w, f) => [w, f] };
  const op = planOperations({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'marker', seriesIndex: 1, fill: '#FF0000', outline: { width: 1, color: '#000000' } }] }).operations[0];
  // unlike plotArea/series/etc., a `false` return here does NOT fail the call -- the trusted signal is the bridge read
  assert.deepEqual(new Function('Api', charts.lower(op).body)(Api), { ok: true });
  assert.deepEqual(log[0], ['SetMarkerFill', [255, 0, 0], 1, 0, true]);
  assert.deepEqual(log[1], ['SetMarkerOutLine', [12700, [0, 0, 0]], 1, 0, true]);
});

// ---- gridlines -----------------------------------------------------------------------------------

const gridlines = (over = {}) => ({ id: 'x-o04-11-g', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'gridlines', axis: 'vertical', outline: { width: 1, color: '#3CC88C' }, ...over }] });

test('gridlines: APPLIED, retry SATISFIED -- vertical axis', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, gridlines());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, gridlines());
  assert.equal(r.m.commands, 0, 'SATISFIED -- no dispatch needed');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('gridlines: the horizontal axis is addressed independently -- writing vertical does not satisfy a horizontal request, and vice versa', async () => {
  const store = createStore(book([mine()]));
  await run(store, gridlines({ axis: 'vertical' }));
  const h = await run(store, gridlines({ axis: 'horizontal', outline: { width: 1, color: '#DC5A0F' } }));
  assert.equal(h.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, 'the vertical write does not satisfy a horizontal request');
  const hAgain = await run(store, gridlines({ axis: 'horizontal', outline: { width: 1, color: '#DC5A0F' } }));
  assert.equal(hAgain.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  const vAgain = await run(store, gridlines({ axis: 'vertical' }));
  assert.equal(vAgain.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED, 'the earlier vertical write is still in effect');
});

test('gridlines: bridge UNOBSERVABLE -- FAIL_CLOSED, named', async () => {
  const { result, driver } = await run(createStore(book([mine()])), gridlines(), { chartStyleUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-bridge-unobservable');
});

test('gridlines: receipt carries NO semanticsNotProven', async () => {
  const store = createStore(book([mine()]));
  const { mapped } = await run(store, gridlines());
  assert.deepEqual(mapped.operations[0].semanticsNotProven, undefined);
});

test('RED (sabotage): gridlines dispatch succeeds but the bridge-read value stays stale -- FAIL, never a silent success', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, gridlines(), { fault: 'style-bridge-stale' });
  assert.equal(a.result.ok, false);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
});

test('the command body: SetMajorVerticalGridlines / SetMajorHorizontalGridlines -- no index, setter name keyed by axis, return value NOT checked', () => {
  const log = [];
  const existing = [{ GetName: () => 'C1', SetMajorVerticalGridlines: v => { log.push(['SetMajorVerticalGridlines', v]); return false; }, SetMajorHorizontalGridlines: v => { log.push(['SetMajorHorizontalGridlines', v]); return false; } }];
  const ws = { GetName: () => 'K', GetAllCharts: () => existing };
  const Api = { GetSheet: () => ws, CreateSolidFill: v => v, CreateRGBColor: (...x) => x, CreateStroke: (w, f) => [w, f] };
  const opV = planOperations(gridlines({ axis: 'vertical' })).operations[0];
  assert.deepEqual(new Function('Api', charts.lower(opV).body)(Api), { ok: true });
  assert.deepEqual(log[0], ['SetMajorVerticalGridlines', [12700, [60, 200, 140]]]);
  log.length = 0;
  const opH = planOperations(gridlines({ axis: 'horizontal', outline: { width: 1, color: '#DC5A0F' } })).operations[0];
  assert.deepEqual(new Function('Api', charts.lower(opH).body)(Api), { ok: true });
  assert.deepEqual(log[0], ['SetMajorHorizontalGridlines', [12700, [220, 90, 15]]]);
});

// ---- planner / field validation (charts.cjs's own styleProblem, independent of write-tools.cjs) ---

test('planner: gridlines field validation', () => {
  const refuse = (over, re) => { const p = planOperations(gridlines(over)); assert.equal(p.ok, false, JSON.stringify(over)); assert.match(p.detail, re, JSON.stringify(over)); };
  refuse({ fill: '#808080' }, /gridlines has no fill/);
  refuse({ outline: undefined }, /outline is required for target gridlines/);
  refuse({ axis: 'diagonal' }, /axis must be one of vertical, horizontal/);
  refuse({ seriesIndex: 0 }, /seriesIndex\/dataPointIndex are not accepted for target gridlines/);
  refuse({ outline: { width: 0, color: '#000000' } }, /outline\.width must be a positive number/);
  refuse({ outline: { width: 1, color: 'oops' } }, /outline\.color must be a 6-hex colour/);
});

test('planner: marker field validation', () => {
  const refuse = (over, re) => { const p = planOperations(style(over)); assert.equal(p.ok, false, JSON.stringify(over)); assert.match(p.detail, re, JSON.stringify(over)); };
  refuse({ fill: undefined }, /at least one of fill\/outline is required/);
  refuse({ seriesIndex: -1 }, /seriesIndex must be a non-negative integer/);
  refuse({ axis: 'vertical' }, /axis is only accepted for target gridlines/);
  refuse({ fill: 'red' }, /fill must be a 6-hex colour/);
});
