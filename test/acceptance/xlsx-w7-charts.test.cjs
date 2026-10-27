'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const charts = require('../../src/adapters/xlsx/families/charts.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M6 wave 7: set_chart (create) / rename_chart on the fake; the identity is derived from the data range (audit CH2). */

const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: extra, pivots: [], names: [] } },
  { name: 'O', cells: { A1: { v: 'keep' } }, objects: { charts: [{ name: 'Other', type: 'bar', title: 'o', legend: 'bottom', series: series('K', ['B']) }], pivots: [{ name: 'PV', source: 'K!$A$1:$C$4' }], names: [] } },
];
const mine = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', series: series('K', ['B', 'C']), ...over });
const add = (extra = {}) => ({ id: 'm6-w7c', steps: [{ intent: 'set_chart', sheet: 'K', name: 'C1', range: 'A1:C4', type: 'bar', title: 'Sales', ...extra }] });
const ren = (name = 'C1', newName = 'C2') => ({ id: 'm6-w7c', steps: [{ intent: 'rename_chart', sheet: 'K', name, newName }] });
const leg = (legend, name = 'C1') => ({ id: 'm6-w7c', steps: [{ intent: 'set_chart_legend', sheet: 'K', name, legend }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;
const chartsOf = (store, s) => store.durable.find(x => x.name === s).objects.charts;

test('set_chart: APPLIED with the derived series, retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, add());
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.appliedOperations[0].evidence]));
  assert.equal(a.m.commands, 1);
  assert.deepEqual(chartsOf(store, 'K').map(c => c.name), ['C1']);
  // the W3h-measured column orientation, written out -- not derived by the code under test
  assert.deepEqual(chartsOf(store, 'K')[0].series, [{ name: '=K!$B$1', values: '=K!$B$2:$B$4', cat: '=K!$A$2:$A$4', x: '' }, { name: '=K!$C$1', values: '=K!$C$2:$C$4', cat: '=K!$A$2:$A$4', x: '' }]);
  const r = await run(store, add());
  assert.equal(r.result.ok, true);
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

// X-O03: lineNormal/area/pie/doughnut/radar/stock, LIVE-measured to hold the same series contract
// as bar. radar's OWN read-side type is the measured "unknown" quirk (fake-xlsx-driver.cjs
// simulates it, charts.cjs's READBACK_TYPE/readbackType() holds it fixed) -- `chartsOf()` here
// reads the fake's INTERNAL store object directly (never through chartReadsOf), so it still shows
// the real requested type, not the quirk.
for (const type of ['lineNormal', 'area', 'pie', 'doughnut', 'radar', 'stock']) {
  test(`set_chart type=${type}: APPLIED with the derived series, retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    const a = await run(store, add({ type }));
    assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
    assert.equal(chartsOf(store, 'K')[0].type, type);
    const r = await run(store, add({ type }));
    assert.equal(r.m.commands, 0, JSON.stringify(r.result.execution.appliedOperations[0].evidence));
    assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  });
}

// X-O04: the 5 LIVE-measured legend positions, each APPLIED + retry SATISFIED, identity (type/title/series) kept.
for (const legend of ['bottom', 'top', 'left', 'right', 'none']) {
  test(`set_chart_legend legend=${legend}: APPLIED keeping the chart's identity, retry SATISFIED 0/0`, async () => {
    const store = createStore(book([mine({ legend: 'bottom' })]));
    const a = await run(store, leg(legend));
    assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
    assert.equal(chartsOf(store, 'K')[0].legend, legend);
    assert.equal(chartsOf(store, 'K')[0].title, 'Sales', 'title kept');
    const r = await run(store, leg(legend));
    assert.equal(r.m.commands, 0);
    assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  });
}

test('set_chart_legend: changing one chart leaves another chart\'s legend untouched', async () => {
  const store = createStore(book([mine({ legend: 'bottom' })]));
  await run(store, leg('top'));
  assert.equal(chartsOf(store, 'O')[0].legend, 'bottom', 'the Other chart on sheet O is unaffected');
});

test('rename_chart: APPLIED keeping the identity, retry SATISFIED 0/0', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, ren());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.primitiveVerification));
  assert.deepEqual(chartsOf(store, 'K').map(c => c.name), ['C2']);
  const r = await run(store, ren());
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

for (const [name, t, why, b, script] of [
  ['a chart of the name over ANOTHER range (audit CH2: no false SATISFIED)', add(), 'xlsx-chart-exists-with-another-identity', book([mine({ series: series('K', ['B']) })])],
  ['a chart of the name with another title (no silent modify)', add(), 'xlsx-chart-exists-with-another-identity', book([mine({ title: 'Old' })])],
  ['a chart of the name with another type (X-O03: type is part of the identity)', add({ type: 'area' }), 'xlsx-chart-exists-with-another-identity', book([mine({ type: 'bar' })])],
  ['two charts of the name', add(), 'xlsx-chart-name-ambiguous', book([mine(), mine()])],
  ['a formula in the data', add(), 'xlsx-chart-data-shape-not-measured', (() => { const x = book(); x[0].cells.B3 = { v: 2, f: '=1+1' }; return x; })()],
  ['a text value in the body', add(), 'xlsx-chart-data-shape-not-measured', (() => { const x = book(); x[0].cells.C3 = { v: 'five' }; return x; })()],
  ['a numeric header', add(), 'xlsx-chart-data-shape-not-measured', (() => { const x = book(); x[0].cells.B1 = { v: 2024 }; return x; })()],
  ['an unobservable chart reader', add(), 'xlsx-chart-unobservable', book(), { chartsUnobservable: true }],
  ['rename: the new name already taken', ren(), 'xlsx-chart-rename-state-conflict', book([mine(), mine({ name: 'C2' })])],
  ['rename: neither name present', ren(), 'xlsx-chart-rename-state-conflict', book()],
  ['legend: the chart is not found', leg('top'), 'xlsx-chart-not-found', book()],
  ['legend: two charts of the name', leg('top'), 'xlsx-chart-name-ambiguous', book([mine(), mine()])],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const { result, driver } = await run(createStore(b), t, script || {});
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('RED: ignored / the wrong title / the wrong range / added twice / an extra chart / a duplicate masking the count -> L1 FAIL; another chart altered / a cell elsewhere -> L3 FAIL', async () => {
  for (const fault of ['ignore', 'chart-wrong-title', 'chart-wrong-range', 'chart-duplicate', 'chart-extra-unnamed', 'chart-duplicate-drops-neighbour']) {
    const { result, driver } = await run(createStore(book([{ name: 'N', type: 'bar', title: 'n', series: series('K', ['C']) }])), add(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
  for (const fault of ['chart-alters-other', 'side-effect']) {
    const { result } = await run(createStore(book()), add(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', fault);
  }
});

test('RED: a rename that drops a series / was ignored / copied instead of moved -> L1 FAIL', async () => {
  for (const fault of ['rename-drops-series', 'ignore', 'rename-copies']) {
    const { result } = await run(createStore(book([mine()])), ren(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
  }
});

// X-O04: 'legend-rejected' simulates the MEASURED SetLegendPos behaviour (rejects an unsupported
// position by returning false, without throwing and without changing it) -- never APPLIED.
test('RED: a legend write that is rejected / ignored / drops a series -> L1 FAIL', async () => {
  for (const fault of ['legend-rejected', 'ignore', 'legend-drops-series']) {
    const { result } = await run(createStore(book([mine({ legend: 'bottom' })])), leg('top'), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
  }
});

test('planner: one sheet name without quoting, bar only, one title line, a header + category + data range', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse(add({ sheet: 'My Sheet' }), /needs no quoting/);
  refuse(add({ type: 'scatter' }), /type must be one of/);
  refuse(add({ type: 'surface' }), /type must be one of/);
  refuse(add({ type: 'notachart' }), /type must be one of/);
  refuse(add({ title: 'a\nb' }), /one non-empty line/);
  refuse(add({ range: 'A1:A4' }), /header row, a category column and data/);
  refuse(add({ inRows: true }), /does not support field/);
  refuse(add({ width: 100 }), /does not support field/);
  refuse(ren('C1', 'C1'), /different from name/);
  refuse(leg('topRight'), /legend must be one of/);
  refuse(leg('notalegend'), /legend must be one of/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_legend', sheet: 'K', name: 'C1', legend: 'top', extra: 1 }] }, /does not support field/);
  // X-O07/3: `delete_chart` is now a REAL intent too (chart.drawing.delete.write/v1) -- this
  // "an unsupported intent name refuses" probe moves to a name that stays fictional.
  refuse({ id: 'x', steps: [{ intent: 'erase_chart', sheet: 'K', name: 'C1' }] }, /./);
});

test('the command bodies: AddChart (columns, bar) + SetName + SetTitle; rename re-identified at dispatch', () => {
  const log = [];
  const existing = [];
  const mk = n => ({ GetName: () => n, SetName: x => log.push(['SetName', x]), SetTitle: (...x) => log.push(['SetTitle', ...x]) });
  const ws = { GetName: () => 'K', GetAllCharts: () => existing, AddChart: (...x) => { log.push(['AddChart', ...x]); return mk(null); } };
  const Api = { GetSheet: () => ws };
  assert.deepEqual(new Function('Api', charts.lower(planOperations(add()).operations[0]).body)(Api), { ok: true });
  assert.deepEqual(log, [['AddChart', 'K!$A$1:$C$4', false, 'bar', 2, 3600000, 2160000, 3, 0, 0, 0], ['SetName', 'C1'], ['SetTitle', 'Sales', 12, false]]);
  // X-O03: a non-bar type substitutes into the SAME AddChart call, nothing else changes.
  log.length = 0;
  assert.deepEqual(new Function('Api', charts.lower(planOperations(add({ type: 'pie' })).operations[0]).body)(Api), { ok: true });
  assert.deepEqual(log, [['AddChart', 'K!$A$1:$C$4', false, 'pie', 2, 3600000, 2160000, 3, 0, 0, 0], ['SetName', 'C1'], ['SetTitle', 'Sales', 12, false]]);
  log.length = 0;
  existing.push(mk('C1'));
  assert.deepEqual(new Function('Api', charts.lower(planOperations(add()).operations[0]).body)(Api), { ok: false, stage: 'name-taken' });
  log.length = 0;
  assert.deepEqual(new Function('Api', charts.lower(planOperations(ren()).operations[0]).body)(Api), { ok: true });
  assert.deepEqual(log, [['SetName', 'C2']]);
  existing.push(mk('C2'));
  assert.deepEqual(new Function('Api', charts.lower(planOperations(ren()).operations[0]).body)(Api), { ok: false, stage: 'rename-identity' });
});

test('the command body: SetLegendPos; its measured false return (an unsupported position) is checked, not assumed from the absence of a throw', () => {
  const log = [];
  let legendReturn = true;
  const existing = [{ GetName: () => 'C1', SetLegendPos: x => { log.push(['SetLegendPos', x]); return legendReturn; } }];
  const ws = { GetName: () => 'K', GetAllCharts: () => existing };
  const Api = { GetSheet: () => ws };
  assert.deepEqual(new Function('Api', charts.lower(planOperations(leg('top')).operations[0]).body)(Api), { ok: true });
  assert.deepEqual(log, [['SetLegendPos', 'top']]);
  legendReturn = false;
  log.length = 0;
  assert.deepEqual(new Function('Api', charts.lower(planOperations(leg('right')).operations[0]).body)(Api), { ok: false, stage: 'legend-rejected' });
  assert.deepEqual(log, [['SetLegendPos', 'right']], 'the call still happened -- the rejection is read from its return value');
});

test('the driver reader: every chart with name / type / title / legend / series, pivot NAMES only; a throwing sheet is unobservable', () => {
  const sx = { GetName: () => '=K!$B$1', GetValues: () => '=K!$B$2:$B$4', GetCatFormula: () => '=K!$A$2:$A$4' };
  // X-O05: GetHorAxisTitle/GetVerAxisTitle/GetDataLabels -- the measured untouched-chart baseline (null / null /
  // every flag false).
  const ch = { GetName: () => 'C1', GetChartType: () => 'bar', GetTitle: () => 'Sales', GetLegendPos: () => 'bottom', GetAllSeries: () => [sx], GetHorAxisTitle: () => null, GetVerAxisTitle: () => null, GetDataLabels: () => ({ showSerName: false, showCatName: false, showVal: false, showPercent: false }) };
  const ws = { GetName: () => 'K', GetIndex: () => 0, GetUsedRange: () => null, GetAllCharts: () => [ch], GetAllPivotTables: () => [{ GetName: () => 'PV', GetSource: () => { throw new Error('ws'); } }] };
  const bad = { ...ws, GetName: () => 'B', GetAllCharts: () => { throw new Error('boom'); } };
  const Api = { GetSheets: () => [ws, bad], GetSheet: n => (n === 'K' ? ws : bad), WorksheetFunction: { COUNT: () => 0 } };
  const r = new Function('Api', readQuery([], { chartSheets: ['*'] }))(Api).chartReads;
  // X-O05/4 round 2 (a maintainer's LIVE r1): GetValues() here is a formula STRING -- the REAL shape -- a single-column
  // range $B$2:$B$4 (3 rows, same column) -> 3 points. This Api has no GetEuroOfficeBridge, so each point's
  // bridge call fails BRIDGE_UNAVAILABLE -- verbatim per point, never a guessed {} or a thrown error.
  // X-O04/10: gridlinesColor is unconditional per chart (this fake has no GetEuroOfficeBridge, so
  // it fails BRIDGE_UNAVAILABLE the same way effectiveLabels does above); markerStyle is scoped to
  // scatter+lineNormal series only -- absent here, this chart's type is 'bar'.
  // X-O05/2: order is unconditional per series (unlike markerStyle), same BRIDGE_UNAVAILABLE path;
  // gapOverlap is unconditional per chart (same as gridlinesColor), same path too.
  // GH139 B /16: lineStyle.read is scoped to bar+lineNormal series (this chart's type is 'bar', so
  // it applies here too, same BRIDGE_UNAVAILABLE path); markerSymbol.read is scoped to
  // scatter+lineNormal only -- 'bar' is OUT of scope, so it is a NAMED UNOBSERVABLE (a maintainer's
  // mechanism-probe #ref: an out-of-scope type must not simply omit the field), not absent.
  assert.deepEqual(r.K, { charts: [{ name: 'C1', type: 'bar', title: 'Sales', legend: 'bottom', series: [{ name: '=K!$B$1', values: '=K!$B$2:$B$4', cat: '=K!$A$2:$A$4', order: 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', markerSymbol: 'UNOBSERVABLE:chart-type-not-in-scope:bar', lineStyle: 'UNOBSERVABLE:BRIDGE_UNAVAILABLE' }], horAxisTitle: null, verAxisTitle: null, dataLabels: { showSerName: false, showCatName: false, showVal: false, showPercent: false },
    effectiveLabels: { '0:0': 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', '0:1': 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', '0:2': 'UNOBSERVABLE:BRIDGE_UNAVAILABLE' }, gridlinesColor: 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', gapOverlap: 'UNOBSERVABLE:BRIDGE_UNAVAILABLE' }], pivots: ['PV'] });
  assert.match(r.B, /^UNOBSERVABLE:/);
});
