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

/** X-O04/1+2 chart STYLE (1bafc2ca, a LIVE measurement round): set_chart_style on the fake.
 * Mirrors xlsx-x-o06-move-chart.test.cjs's own shape -- retry/refusal tables, RED faults, command-
 * body unit tests -- PLUS the receipt-level semanticsNotProven checks 16fd90b9 already established
 * for drawings-insert.cjs, since this intent is WRITE-ONLY by design (no getter exists at all) and
 * must never be read as "verified". */

const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: extra, pivots: [], names: [] } },
];
const mine = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', legend: 'bottom', series: series('K', ['B', 'C']), ...over });
// a default `fill` so a bare style() is a valid step on its own -- an outline-only override
// MUST explicitly pass `fill: undefined` to clear it (NOT just omit it), or the default silently
// leaks in and calls a fill setter the test never meant to exercise.
const style = (over = {}) => ({ id: 'x-o04', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'plotArea', fill: '#808080', ...over }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics, mapped: mapExecutionResult(result, { fileId: 'disposable', task: t }) };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('set_chart_style: APPLIED, retry ALSO APPLIED (never SATISFIED -- no getter to confirm it already matches)', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, style());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, style());
  assert.equal(r.result.ok, true);
  assert.equal(r.m.commands, 1, 'a retry re-dispatches the SAME idempotent setter -- it is never 0');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, 'never SATISFIED for this intent');
});

test('receipt: semanticsNotProven names both fill AND outline, capability-specific per target', async () => {
  const store = createStore(book([mine()]));
  const { mapped } = await run(store, style({ target: 'title', fill: '#808080', outline: { width: 2, color: '#C81E1E' } }));
  assert.deepEqual(mapped.operations[0].semanticsNotProven, ['chart-style-title-fill', 'chart-style-title-outline']);
});

test('receipt: fill-only / outline-only each name just their own field', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, style({ target: 'legend', fill: '#808080' }));
  assert.deepEqual(a.mapped.operations[0].semanticsNotProven, ['chart-style-legend-fill']);
  const store2 = createStore(book([mine()]));
  const b = await run(store2, style({ target: 'legend', fill: undefined, outline: { width: 1, color: '#000000' } }));
  assert.deepEqual(b.mapped.operations[0].semanticsNotProven, ['chart-style-legend-outline']);
});

test('receipt: the "chart" target (bare Fill/SetOutLine) is named chart-style-chart-fill/-outline', async () => {
  const store = createStore(book([mine()]));
  const { mapped } = await run(store, style({ target: 'chart', fill: '#808080', outline: { width: 1, color: '#000000' } }));
  assert.deepEqual(mapped.operations[0].semanticsNotProven, ['chart-style-chart-fill', 'chart-style-chart-outline']);
});

test('receipt: series/dataPoint targets are named the same way, index-qualified operation unaffected', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, style({ target: 'series', fill: '#808080', seriesIndex: 0 }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.deepEqual(a.mapped.operations[0].semanticsNotProven, ['chart-style-series-fill']);
  const store2 = createStore(book([mine()]));
  const b = await run(store2, style({ target: 'dataPoint', fill: undefined, outline: { width: 1, color: '#000000' }, seriesIndex: 1, dataPointIndex: 2 }));
  assert.equal(b.result.ok, true, JSON.stringify(b.result.execution.appliedOperations[0].evidence));
  assert.deepEqual(b.mapped.operations[0].semanticsNotProven, ['chart-style-datapoint-outline']);
});

for (const [name, t, why, b] of [
  ['the chart is not found', style(), 'xlsx-chart-not-found', book()],
  ['two charts of the name', style(), 'xlsx-chart-name-ambiguous', book([mine(), mine()])],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const { result, driver } = await run(createStore(b), t);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('refused BEFORE dispatch: seriesIndex out of range (the chart has 2 series -- indices 0/1)', async () => {
  const { result, driver } = await run(createStore(book([mine()])), style({ target: 'series', fill: '#FF0000', seriesIndex: 2 }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-series-index-out-of-range');
});

test('RED: the setter reports rejection -- DISPATCHED_OUTCOME_UNKNOWN, ok=false, never a silent success; a clean retry still APPLIES', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, style(), { fault: 'style-setter-rejected' });
  assert.equal(a.result.ok, false);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  const r = await run(store, style());
  assert.equal(r.result.ok, true);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
});

test('RED (sabotage): a style write that also changes the chart\'s legend as a side effect -> L1 FAIL (the "kept" check fires)', async () => {
  const { result } = await run(createStore(book([mine()])), style(), { fault: 'style-alters-legend' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

// X-O04/11: SUPERSEDED -- `marker` was REFUSED here (measured non-functional, X-O04/1+2 round 1).
// With bAllMarkers=true it WORKS_PERSISTED (X-O04/4 round 2 + X-O04/8 LIVE) and is now a released,
// verifiable target (see xlsx-xo04-11-chart-style-marker-gridlines.test.cjs). `gridline` (singular)
// was never a real target name -- the real one, `gridlines` (plural), is ALSO now released; the
// singular typo stays refused by the enum, unchanged.
test('refused at plan time: an unknown target name (not in the enum, including the "gridline" singular typo) -- 0 operations, never attempted', () => {
  for (const target of ['gridline', 'bogus']) {
    const p = planOperations({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target, fill: '#FF0000' }] });
    assert.equal(p.ok, false, target);
    assert.match(p.detail, /target must be one of plotArea, title, legend, chart, series, dataPoint, marker, gridlines/, target);
  }
});

test('planner: field validation', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'bogus', fill: '#FF0000' }] }, /target must be one of/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'title' }] }, /at least one of fill\/outline is required/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'title', fill: 'red' }] }, /fill must be a 6-hex colour/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'title', outline: { width: 0, color: '#FF0000' } }] }, /outline\.width must be a positive number/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'title', outline: { width: 1, color: 'oops' } }] }, /outline\.color must be a 6-hex colour/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'title', fill: '#FF0000', seriesIndex: 0 }] }, /seriesIndex is only accepted for target series\/dataPoint/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'series', fill: '#FF0000' }] }, /seriesIndex must be a non-negative integer/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'series', fill: '#FF0000', seriesIndex: -1 }] }, /seriesIndex must be a non-negative integer/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'series', fill: '#FF0000', seriesIndex: 0, dataPointIndex: 0 }] }, /dataPointIndex is only accepted for target dataPoint/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'dataPoint', fill: '#FF0000', seriesIndex: 0 }] }, /dataPointIndex must be a non-negative integer/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'title', fill: '#FF0000', colOffset: 5 }] }, /does not support field/);
});

test('the command body: plotArea fill+outline -- SetPlotAreaFill/SetPlotAreaOutLine, fail-closed on a false return', () => {
  const log = [];
  const existing = [{ GetName: () => 'C1', SetPlotAreaFill: v => { log.push(['SetPlotAreaFill', v]); return true; }, SetPlotAreaOutLine: v => { log.push(['SetPlotAreaOutLine', v]); return true; } }];
  const ws = { GetName: () => 'K', GetAllCharts: () => existing };
  const Api = { GetSheet: () => ws, CreateSolidFill: (...x) => ['fill', ...x], CreateRGBColor: (...x) => ['rgb', ...x], CreateStroke: (...x) => ['stroke', ...x] };
  const op = planOperations(style({ target: 'plotArea', fill: '#808080', outline: { width: 2, color: '#C81E1E' } })).operations[0];
  assert.deepEqual(new Function('Api', charts.lower(op).body)(Api), { ok: true });
  assert.deepEqual(log[0], ['SetPlotAreaFill', ['fill', ['rgb', 128, 128, 128]]]);
  assert.deepEqual(log[1], ['SetPlotAreaOutLine', ['stroke', 25400, ['fill', ['rgb', 200, 30, 30]]]]);
});

test('the command body: a false return from the setter fails closed, named by WHICH call rejected', () => {
  const ws1 = { GetName: () => 'K', GetAllCharts: () => [{ GetName: () => 'C1', SetPlotAreaFill: () => false }] };
  const Api1 = { GetSheet: () => ws1, CreateSolidFill: () => null, CreateRGBColor: () => null };
  const op1 = planOperations(style({ target: 'plotArea', fill: '#808080' })).operations[0];
  assert.deepEqual(new Function('Api', charts.lower(op1).body)(Api1), { ok: false, stage: 'style-fill-rejected' });

  const ws2 = { GetName: () => 'K', GetAllCharts: () => [{ GetName: () => 'C1', SetPlotAreaOutLine: () => false }] };
  const Api2 = { GetSheet: () => ws2, CreateSolidFill: () => null, CreateRGBColor: () => null, CreateStroke: () => null };
  const op2 = planOperations(style({ target: 'plotArea', fill: undefined, outline: { width: 1, color: '#000000' } })).operations[0];
  assert.deepEqual(new Function('Api', charts.lower(op2).body)(Api2), { ok: false, stage: 'style-outline-rejected' });
});

test('the command body: series (indexed) -- extra args (seriesIndex, false); dataPoint -- (seriesIndex, dataPointIndex, false)', () => {
  const log = [];
  const existing = [{ GetName: () => 'C1', SetSeriesFill: (...x) => { log.push(['SetSeriesFill', ...x]); return true; } }];
  const ws = { GetName: () => 'K', GetAllCharts: () => existing };
  const Api = { GetSheet: () => ws, CreateSolidFill: v => v, CreateRGBColor: (...x) => x };
  const op = planOperations(style({ target: 'series', fill: '#FF0000', seriesIndex: 1 })).operations[0];
  new Function('Api', charts.lower(op).body)(Api);
  assert.deepEqual(log[0], ['SetSeriesFill', [255, 0, 0], 1, false]);

  const log2 = [];
  const existing2 = [{ GetName: () => 'C1', SetDataPointOutLine: (...x) => { log2.push(['SetDataPointOutLine', ...x]); return true; } }];
  const ws2 = { GetName: () => 'K', GetAllCharts: () => existing2 };
  const Api2 = { GetSheet: () => ws2, CreateStroke: (w, f) => [w, f], CreateSolidFill: v => v, CreateRGBColor: (...x) => x };
  const op2 = planOperations(style({ target: 'dataPoint', fill: undefined, outline: { width: 1, color: '#000000' }, seriesIndex: 2, dataPointIndex: 3 })).operations[0];
  new Function('Api', charts.lower(op2).body)(Api2);
  assert.deepEqual(log2[0], ['SetDataPointOutLine', [12700, [0, 0, 0]], 2, 3, false]);
});

test('the command body: the "chart" target calls the BARE Fill/SetOutLine names', () => {
  const log = [];
  const existing = [{ GetName: () => 'C1', Fill: v => { log.push(['Fill', v]); return true; }, SetOutLine: v => { log.push(['SetOutLine', v]); return true; } }];
  const ws = { GetName: () => 'K', GetAllCharts: () => existing };
  const Api = { GetSheet: () => ws, CreateSolidFill: v => v, CreateRGBColor: (...x) => x, CreateStroke: (w, f) => [w, f] };
  const op = planOperations(style({ target: 'chart', fill: '#808080', outline: { width: 1, color: '#000000' } })).operations[0];
  new Function('Api', charts.lower(op).body)(Api);
  assert.deepEqual(log[0], ['Fill', [128, 128, 128]]);
  assert.deepEqual(log[1], ['SetOutLine', [12700, [0, 0, 0]]]);
});

test('EMU conversion for outline.width: pt * 12700, rounded', () => {
  const log = [];
  const existing = [{ GetName: () => 'C1', SetTitleOutLine: (...x) => { log.push(x); return true; } }];
  const ws = { GetName: () => 'K', GetAllCharts: () => existing };
  const Api = { GetSheet: () => ws, CreateStroke: (w, f) => [w, f], CreateSolidFill: v => v, CreateRGBColor: (...x) => x };
  const op = planOperations(style({ target: 'title', fill: undefined, outline: { width: 0.5, color: '#000000' } })).operations[0];
  new Function('Api', charts.lower(op).body)(Api);
  assert.equal(log[0][0][0], 6350);
});
