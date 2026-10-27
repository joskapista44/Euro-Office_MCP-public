'use strict';

/**
 * GH #302 G6: copy_chart carries the source's data labels -- the chart-level flags (SetShowDataLabels) and every
 * POINT-level effective label (SetShowPointDataLabel), verified on the copy by read_charts' dataLabels + effectiveLabels
 * (the LIVE design probe docs/evidence/gh302-g6-copy-chart-labels-probe-20261026.json). What it cannot carry is named per
 * copy in semanticsNotProven, never claimed.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const charts = require('../../src/adapters/xlsx/families/charts.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { mapExecutionResult } = require('../../src/server/write-tools.cjs');

const OFF = Object.freeze({ showSerName: false, showCatName: false, showVal: false, showPercent: false });
const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const KEYS = ['0:0', '0:1', '0:2', '1:0', '1:1', '1:2'];
const effective = (over = {}) => Object.fromEntries(KEYS.map(k => [k, over[k] || { level: 'SERIES', ...OFF }]));
const LABELED = () => effective({ '0:1': { level: 'POINT', ...OFF, showCatName: true, showVal: true }, '1:0': { level: 'POINT', ...OFF } });
const srcChart = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', legend: 'top', series: series('K', ['B', 'C']), dataLabels: { ...OFF, showVal: true }, effectiveLabels: LABELED(), ...over });
const book = src => [{ name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: [src], pivots: [], names: [] } }];
const copyTask = () => ({ id: 'gh302-g6', steps: [{ intent: 'copy_chart', sheet: 'K', name: 'Copy1', sourceSheet: 'K', sourceName: 'C1', at: { col: 10, row: 0 } }] });
const copyOf = store => store.durable.find(x => x.name === 'K').objects.charts.find(c => c.name === 'Copy1');

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics, mapped: mapExecutionResult(result, { fileId: 'disposable', task: t }) };
}

test('G6 copy_chart: the copy carries the chart-level flags and every POINT label; retry SATISFIED 0 commands', async () => {
  const store = createStore(book(srcChart()));
  const a = await run(store, copyTask());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.deepEqual(copyOf(store).dataLabels, { ...OFF, showVal: true });
  assert.deepEqual(copyOf(store).effectiveLabels, LABELED());
  const notProven = a.mapped.operations[0].semanticsNotProven || [];
  assert.equal(notProven.some(f => /data-labels/.test(f)), false, JSON.stringify(notProven));
  const r = await run(store, copyTask());
  assert.equal(r.result.ok, true);
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('G6 RED: the point labels not carried (fault copy-drops-point-labels) -> the copy L1 fails', async () => {
  const store = createStore(book(srcChart()));
  const a = await run(store, copyTask(), { fault: 'copy-drops-point-labels' });
  assert.equal(a.result.ok, false);
  assert.match(JSON.stringify(a.result.execution), /"pointLabelsOk":false/, 'it fails on the point labels, not on something else');
});

test('G6: a series-level label that SHOWS a flag has no setter -> APPLIED for the rest, named chart-series-level-data-labels', async () => {
  const src = srcChart({ effectiveLabels: effective({ '0:1': { level: 'POINT', ...OFF, showVal: true }, '1:2': { level: 'SERIES', ...OFF, showVal: true } }) });
  const store = createStore(book(src));
  const a = await run(store, copyTask());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.deepEqual(copyOf(store).effectiveLabels['0:1'], { level: 'POINT', ...OFF, showVal: true });
  assert.ok((a.mapped.operations[0].semanticsNotProven || []).includes('chart-series-level-data-labels'));
});

test('G6: per-point labels unobservable -> the chart level still carried, the points named chart-point-data-labels-unobservable', async () => {
  const store = createStore(book(srcChart({ effectiveLabels: 'UNOBSERVABLE:CAPABILITY_MISSING' })));
  const a = await run(store, copyTask());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.deepEqual(copyOf(store).dataLabels, { ...OFF, showVal: true });
  assert.ok((a.mapped.operations[0].semanticsNotProven || []).includes('chart-point-data-labels-unobservable'));
});

test('G6: the source chart-level labels unreadable -> refused before the write (they would be lost)', async () => {
  const store = createStore(book(srcChart()));
  const a = await run(store, copyTask(), { chartLabelsUnobservable: true });
  assert.equal(a.result.ok, false);
  assert.equal(a.m.commands, 0);
  assert.equal(copyOf(store), undefined);
});

test('G6 lower: chart level only when a flag is on, then one SetShowPointDataLabel per POINT label', () => {
  const op = charts.plan(copyTask().steps[0], 0).operation;
  const obs = { format: 'xlsx', complete: true, sheets: [{ name: 'K', index: 0, used: 'A1:C4' }], cells: {}, chartReads: { K: { charts: [srcChart()], pivots: [] } }, drawingReads: { K: [] } };
  charts.classify(obs, op);
  const body = charts.lower(op).body;
  assert.match(body, /c\.SetShowDataLabels\(false, false, true, false\);/);
  assert.match(body, /SetShowPointDataLabel\(0, 1, false, true, true, false\) === false/);
  assert.match(body, /SetShowPointDataLabel\(1, 0, false, false, false, false\) === false/);
  assert.equal((body.match(/SetShowPointDataLabel/g) || []).length, 2);
  const plain = charts.plan(copyTask().steps[0], 0).operation;
  charts.classify({ ...obs, chartReads: { K: { charts: [srcChart({ dataLabels: { ...OFF }, effectiveLabels: effective() })], pivots: [] } } }, plain);
  assert.doesNotMatch(charts.lower(plain).body, /SetShowDataLabels|SetShowPointDataLabel/, 'nothing to carry, no label call (the v1 body)');
});
