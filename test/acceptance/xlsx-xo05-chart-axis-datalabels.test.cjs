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
 * X-O05 (LIVE MEASURED, dispatched #ref): set_chart_axis_title (SetHorAxisTitle /
 * SetVerAxisTitle, one text argument, CRLF read-back like set_chart_title) and set_chart_data_labels
 * (SetShowDataLabels(showSerName, showCatName, showVal, showPercent), four positional booleans;
 * GetDataLabels() -> { showSerName, showCatName, showVal, showPercent }). An untouched chart reads
 * both axis titles null and every label flag false (the control chart in #ref).
 */

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const op = result => result.execution.appliedOperations[0];
const reason = result => op(result).evidence.evidence.reason;

const chartBook = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, A2: { v: 'a' }, B2: { v: 1 }, A3: { v: 'b' }, B3: { v: 2 } }, objects: { charts: extra, pivots: [], names: [] } },
  { name: 'O', cells: { A1: { v: 'keep' } }, objects: { charts: [], pivots: [], names: [] } },
];
const addChart = (extra = {}) => ({ id: 'xo05', steps: [{ intent: 'set_chart', sheet: 'K', name: 'R1', range: 'A1:B3', type: 'bar', title: 'Sales', ...extra }] });
const chartOf = store => store.durable.find(s => s.name === 'K').objects.charts.find(c => c.name === 'R1');

const axisTitle = (axis, title) => ({ id: 'xo05-axis', steps: [{ intent: 'set_chart_axis_title', sheet: 'K', name: 'R1', axis, title }] });
const dataLabels = labels => ({ id: 'xo05-labels', steps: [{ intent: 'set_chart_data_labels', sheet: 'K', name: 'R1', showSerName: false, showCatName: false, showVal: false, showPercent: false, ...labels }] });

// ---- axis titles ----------------------------------------------------------------------------------------------

test('set_chart_axis_title horizontal: APPLIED, read back with CRLF trimmed, retry SATISFIED 0/0; other axis untouched', async () => {
  const store = createStore(chartBook());
  await run(store, addChart());
  const { result, m } = await run(store, axisTitle('horizontal', 'Months'));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, op(result).evidence]));
  assert.equal(m.commands, 1);
  assert.equal(chartOf(store).horAxisTitle, 'Months\r\n');
  assert.equal(chartOf(store).verAxisTitle, undefined, 'the vertical axis was never touched');
  const again = await run(store, axisTitle('horizontal', 'Months'));
  assert.equal(again.m.commands, 0);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
});

test('set_chart_axis_title vertical: APPLIED and coexists with an already-set horizontal title', async () => {
  const store = createStore(chartBook());
  await run(store, addChart());
  await run(store, axisTitle('horizontal', 'Months'));
  const { result } = await run(store, axisTitle('vertical', 'Revenue'));
  assert.equal(result.ok, true, JSON.stringify(op(result).evidence));
  assert.equal(chartOf(store).horAxisTitle, 'Months\r\n', 'the horizontal title set earlier stays');
  assert.equal(chartOf(store).verAxisTitle, 'Revenue\r\n');
});

test('set_chart_axis_title refused: no chart of the name; axis title of two lines; axis must be horizontal/vertical', () => {
  assert.match(planOperations(axisTitle('sideways', 'X')).detail, /axis must be/);
  assert.equal(planOperations(axisTitle('horizontal', 'a\nb')).ok, false);
});

test('set_chart_axis_title RED: the wrong text lands -> L1 FAIL', async () => {
  const store = createStore(chartBook());
  await run(store, addChart());
  const { result } = await run(store, axisTitle('horizontal', 'Months'), { fault: 'axis-title-wrong' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('set_chart_axis_title RED: setting one axis silently clears the OTHER axis -> L1 FAIL (the other-axis-kept check)', async () => {
  const store = createStore(chartBook());
  await run(store, addChart());
  await run(store, axisTitle('vertical', 'Revenue'));
  const { result } = await run(store, axisTitle('horizontal', 'Months'), { fault: 'axis-title-drops-other-axis' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('no chart of the name: refused before any write', async () => {
  const a = await run(createStore(chartBook()), axisTitle('horizontal', 'X'));
  assert.equal(a.driver.calls.dispatch, 0);
  assert.equal(reason(a.result), 'xlsx-chart-not-found');
});

test('the command calls SetHorAxisTitle / SetVerAxisTitle on the one named chart only', () => {
  const horBody = charts.lower(planOperations(axisTitle('horizontal', 'Months')).operations[0]).body;
  assert.match(horBody, /hit\[0\]\.SetHorAxisTitle\("Months"\);/);
  assert.doesNotMatch(horBody, /SetVerAxisTitle|SetTitle\(/);
  const verBody = charts.lower(planOperations(axisTitle('vertical', 'Revenue')).operations[0]).body;
  assert.match(verBody, /hit\[0\]\.SetVerAxisTitle\("Revenue"\);/);
});

// ---- data labels: NOT a write intent (OWNER-GATE / LIMIT / #ref) --------------------------------

test('set_chart_data_labels is refused at plan time (no public reader of the effective series-level state)', () => {
  const p = planOperations(dataLabels({ showVal: true }));
  assert.equal(p.ok, false);
});

// ---- set_chart's own identity still holds for a freshly created chart (the wanted() baseline widened) -----------

test('set_chart (regression): a freshly created chart still identifies correctly now that wanted() carries axis/label defaults', async () => {
  const store = createStore(chartBook());
  const { result, m } = await run(store, addChart());
  assert.equal(result.ok, true, JSON.stringify(op(result).evidence));
  assert.equal(m.commands, 1);
  const again = await run(store, addChart());
  assert.equal(again.m.commands, 0, 'retry SATISFIED 0/0');
});

// ---- a maintainer review: an unobservable read is never defaulted -----------------------------------------------------

test('UNOBSERVABLE data labels also stop a title change: the kept-check cannot prove the labels stayed put', async () => {
  const store = createStore(chartBook());
  await run(store, addChart());
  const { result } = await run(store, { id: 'xo05-t', steps: [{ intent: 'set_chart_title', sheet: 'K', name: 'R1', title: 'Other' }] }, { chartLabelsUnobservable: true });
  assert.notEqual(op(result).status, OPERATION_STATUS.APPLIED);
});

