'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { mapExecutionResult } = require('../../src/server/write-tools.cjs');

/**
 * X-O07/4b: diagram copy AS RECREATION, through the FULL product stack (plan -> classify ->
 * lower -> dispatch -> verify), against the fake driver's own simulated round-trip. v1 scope
 * (type/series/title/legend/axis-titles/position -- NEVER the style targets, see charts.cjs's
 * own COPY_CHART_ALWAYS_OMITTED / classifyCopy header comment). This is the offline counterpart
 * to the LIVE product harness (release-xo07-4b-copy-chart-product.live.cjs) -- this file measures
 * that OUR OWN wiring is correct, independent of it.
 */

const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extraSrc = [], extraDest = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: extraSrc, pivots: [], names: [] } },
  { name: 'Dest', cells: {}, objects: { charts: extraDest, pivots: [], names: [] } },
];
const srcChart = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', legend: 'top', series: series('K', ['B', 'C']), horAxisTitle: 'Cat', verAxisTitle: 'Val', ...over });
const copyStep = (over = {}) => ({ intent: 'copy_chart', sheet: 'K', name: 'Copy1', sourceSheet: 'K', sourceName: 'C1', at: { col: 10, row: 0 }, ...over });
const copyTask = (over = {}) => ({ id: 'x-o07-4b', steps: [copyStep(over)] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics, mapped: mapExecutionResult(result, { fileId: 'disposable', task: t }) };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;
const chartsOf = (store, s) => store.durable.find(x => x.name === s).objects.charts;

test('copy_chart: APPLIED, the copy matches the source (type/series/title/legend/axis-titles/position); retry SATISFIED 0/0', async () => {
  const store = createStore(book([srcChart()]));
  const a = await run(store, copyTask());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const copy = chartsOf(store, 'K').find(c => c.name === 'Copy1');
  assert.equal(copy.type, 'bar');
  assert.equal(copy.title, 'Sales');
  assert.equal(copy.legend, 'top');
  assert.deepEqual(copy.series, chartsOf(store, 'K').find(c => c.name === 'C1').series);
  assert.deepEqual(copy.pos, { fromCol: 10, colOffset: 0, fromRow: 0, rowOffset: 0 });
  // the source is UNTOUCHED
  const src = chartsOf(store, 'K').find(c => c.name === 'C1');
  assert.equal(src.title, 'Sales');
  const r = await run(store, copyTask());
  assert.equal(r.result.ok, true, JSON.stringify(r.result.execution.appliedOperations[0].evidence));
  assert.equal(r.m.commands, 0, 'SATISFIED -- no dispatch needed, the copy already matches the source');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('copy_chart: ROWS orientation source reconstructs correctly too', async () => {
  // a hand-built ROWS-orientation chart: name cells share a COLUMN (A), increasing rows.
  const rowsChart = { name: 'R1', type: 'bar', title: 'Rows', legend: 'none', series: [
    { name: '=K!$A$2', values: '=K!$B$2:$C$2', cat: '=K!$B$1:$C$1', x: '' },
    { name: '=K!$A$3', values: '=K!$B$3:$C$3', cat: '=K!$B$1:$C$1', x: '' },
  ] };
  const store = createStore(book([rowsChart]));
  const a = await run(store, copyTask({ sourceName: 'R1' }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  const copy = chartsOf(store, 'K').find(c => c.name === 'Copy1');
  assert.deepEqual(copy.series, rowsChart.series);
});

test('copy_chart: REGRESSION -- a category-less source (every series own cat reads "") reconstructs and copies correctly (zero LIVE r1, 2026-10-15: a 3-series bar chart with single-cell values and cat:"" failed xlsx-chart-copy-series-ref-not-parseable before this)', async () => {
  const noCatChart = { name: 'NoCat1', type: 'bar', title: null, legend: 'bottom', series: [
    { name: '=K!$F$1', values: '=K!$F$2', cat: '', x: '' },
    { name: '=K!$G$1', values: '=K!$G$2', cat: '', x: '' },
    { name: '=K!$H$1', values: '=K!$H$2', cat: '', x: '' },
  ] };
  const store = createStore(book([noCatChart]));
  const a = await run(store, copyTask({ sourceName: 'NoCat1' }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  const copy = chartsOf(store, 'K').find(c => c.name === 'Copy1');
  assert.deepEqual(copy.series, noCatChart.series);
  // retry: SATISFIED, 0 dispatch -- the identity check must recognize the no-cat shape as matching
  const r = await run(store, copyTask({ sourceName: 'NoCat1' }));
  assert.equal(r.result.ok, true, JSON.stringify(r.result.execution.appliedOperations[0].evidence));
  assert.equal(r.m.commands, 0, 'SATISFIED on retry -- the no-cat copy already matches the no-cat source');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('refused BEFORE dispatch: source series disagree on having a cat range (some have one, some don\'t) -- never guesses, xlsx-chart-copy-series-cat-mismatch', async () => {
  const mixedChart = { name: 'Mixed1', type: 'bar', title: null, legend: 'bottom', series: [
    { name: '=K!$F$1', values: '=K!$F$2:$F$4', cat: '=K!$A$2:$A$4', x: '' },
    { name: '=K!$G$1', values: '=K!$G$2:$G$4', cat: '', x: '' },
  ] };
  const store = createStore(book([mixedChart]));
  const { result, driver } = await run(store, copyTask({ sourceName: 'Mixed1' }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-copy-series-cat-mismatch');
});

test('copy_chart: no axis titles on the source -> none on the copy (not "undefined", not a throw)', async () => {
  const store = createStore(book([srcChart({ horAxisTitle: undefined, verAxisTitle: undefined })]));
  const a = await run(store, copyTask());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  const copy = chartsOf(store, 'K').find(c => c.name === 'Copy1');
  assert.equal(copy.horAxisTitle, undefined);
  assert.equal(copy.verAxisTitle, undefined);
});

test('refused BEFORE dispatch: source not found -- xlsx-chart-copy-source-not-found, nothing created', async () => {
  const store = createStore(book([]));
  const { result, driver } = await run(store, copyTask());
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-copy-source-not-found');
  assert.equal(chartsOf(store, 'K').length, 0);
});

test('refused BEFORE dispatch: source ambiguous (two charts of the source name) -- never guesses which one', async () => {
  const store = createStore(book([srcChart(), srcChart()]));
  const { result, driver } = await run(store, copyTask());
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-copy-source-ambiguous');
});

test('refused BEFORE dispatch: dest name already taken with a DIFFERENT identity -- never silently overwrites', async () => {
  // the conflicting chart must be on the SAME sheet as the destination ('K', the default copyStep sheet).
  const store = createStore(book([srcChart(), srcChart({ name: 'Copy1', title: 'Something Else' })]));
  const { result, driver } = await run(store, copyTask());
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-copy-dest-exists-with-another-identity');
});

test('destination on a DIFFERENT sheet (K2) -- attempted, same mechanism, UNMEASURED live but offline-consistent', async () => {
  const store = createStore(book([srcChart()]));
  const a = await run(store, copyTask({ sheet: 'Dest' }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  const copy = chartsOf(store, 'Dest').find(c => c.name === 'Copy1');
  assert.equal(copy.title, 'Sales');
  assert.deepEqual(copy.series, chartsOf(store, 'K').find(c => c.name === 'C1').series);
  // the source sheet's OWN chart list is untouched (no accidental duplicate there)
  assert.equal(chartsOf(store, 'K').filter(c => c.name === 'Copy1').length, 0);
});

test('RED: dispatch drops a series (fault "copy-wrong-series") -- L1 FAIL, never a silent partial copy', async () => {
  const store = createStore(book([srcChart()]));
  const { result } = await run(store, copyTask(), { fault: 'copy-wrong-series' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('RED: copying alters the SOURCE as a side effect (fault "copy-alters-source") -- L3 FAIL', async () => {
  const store = createStore(book([srcChart()]));
  const { result } = await run(store, copyTask(), { fault: 'copy-alters-source' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('receipt: the style targets this v1 never copies are ALWAYS named in semanticsNotProven, never silent', async () => {
  const store = createStore(book([srcChart()]));
  const { result, mapped } = await run(store, copyTask());
  assert.equal(result.ok, true);
  const notProven = mapped.operations[0].semanticsNotProven;
  assert.ok(Array.isArray(notProven) && notProven.length > 0, JSON.stringify(mapped.operations[0]));
  for (const f of ['chart-style-plotarea', 'chart-style-marker', 'chart-style-gridlines', 'chart-style-gapoverlap', 'chart-style-seriesorder']) {
    assert.ok(notProven.includes(f), `${f} missing from semanticsNotProven: ${JSON.stringify(notProven)}`);
  }
  // GH #302 G6: the data labels are copied now; this source's per-point labels are not read by the fake -> named as such
  assert.equal(notProven.includes('chart-data-labels'), false, JSON.stringify(notProven));
  assert.ok(notProven.includes('chart-point-data-labels-unobservable'), JSON.stringify(notProven));
});

test('planner: field validation', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  // copy_chart is cellFree (like rename_chart/move_chart), so it accepts the WIDER sheetNameProblem
  // class (spaces ARE allowed there, unlike set_chart's own narrower "needs no quoting" class) --
  // an apostrophe is the one thing that class refuses (L-21 LIVE: mishandled by the DS engine).
  refuse({ id: 'x', steps: [copyStep({ sheet: "O'Brien" })] }, /apostrophe/);
  refuse({ id: 'x', steps: [copyStep({ extra: 1 })] }, /does not support field/);
  refuse({ id: 'x', steps: [copyStep({ at: undefined })] }, /./);
  refuse({ id: 'x', steps: [copyStep({ sourceSheet: undefined })] }, /./);
  refuse({ id: 'x', steps: [copyStep({ sourceName: undefined })] }, /./);
  refuse({ id: 'x', steps: [copyStep({ sheet: 'K', name: 'C1', sourceSheet: 'K', sourceName: 'C1' })] }, /must differ from the source/);
});

test('plan(): the selector claims only the DESTINATION (the generic composition-conflict detector picks it up) -- cellFree', () => {
  const p = planOperations(copyTask());
  assert.equal(p.ok, true, JSON.stringify(p));
  const op = p.operations[0];
  assert.deepEqual([...op.selector.claims], ['chart:K:Copy1']);
  assert.equal(op.selector.cellFree, true);
  assert.equal(op.sourceSheet, 'K');
  assert.equal(op.sourceName, 'C1');
});

test('the dispatched command actually carries the source-derived fields (type/title/legend/axis-titles/position) through meta -- proven end-to-end via the fake driver, not by inspecting the body string directly', async () => {
  const store = createStore(book([srcChart()]));
  const { result, driver } = await run(store, copyTask());
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.equal(driver.calls.dispatch, 1);
  const copy = chartsOf(store, 'K').find(c => c.name === 'Copy1');
  assert.equal(copy.type, 'bar');
  assert.equal(copy.title, 'Sales');
  assert.equal(copy.legend, 'top');
  assert.equal(copy.horAxisTitle, 'Cat\r\n');
  assert.equal(copy.verAxisTitle, 'Val\r\n');
  assert.deepEqual(copy.pos, { fromCol: 10, colOffset: 0, fromRow: 0, rowOffset: 0 });
});
