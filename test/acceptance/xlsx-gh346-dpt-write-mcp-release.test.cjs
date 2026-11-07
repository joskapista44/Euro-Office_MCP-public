'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const charts = require('../../src/adapters/xlsx/families/charts.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { mapExecutionResult, validateWriteRequest } = require('../../src/server/write-tools.cjs');
const { XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS } = require('../../src/adapters/xlsx/runtime.cjs');

/**
 * GH #346 AUTO-REPAIR Egyseg 3 (design docs/GH346-DPT-AUTO-REPAIR-DESIGN-20261104.md Dontesi pont 5)
 * -- `set_chart_style`'s `dataPoint`
 * target, through the FULL product stack (schema -> CONSTRAINTS -> plan -> classify -> lower ->
 * dispatch -> verify), proving the mechanism end-to-end against the fake driver's own simulated
 * round-trip (test/fakes/fake-xlsx-driver.cjs's `pointStyles` field) AND the write-tools.cjs release
 * gate itself. Offline counterpart to the LIVE harness (test/live/xlsx-gh346-dpt-normalize-write-27.live.cjs)
 * -- that harness measures the REAL DS /27 runtime; this file measures that OUR OWN
 * schema/plan/classify/lower/verify wiring is correct, independent of it. Same loading convention as
 * xlsx-xo05-3-chart-style-gap-overlap-series-order.test.cjs.
 */

const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: extra, pivots: [], names: [] } },
];
const mine = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', legend: 'bottom', series: series('K', ['B', 'C']), ...over });
const dataPoint = (over = {}) => ({ id: 'gh346-u3', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'dataPoint', seriesIndex: 0, dataPointIndex: 1, fill: '#334455', ...over }] });

// GH #346 unit 3: dataPoint is B29-WRITE (/27-ONLY) -- UNLIKE gapOverlap/seriesOrder (B22, /12../27) or
// markerSymbol/lineStyle (B24, /15../27), every test here needs the fake session opened on the /27
// runtime, never the generic default (XLSX_RUNTIME_ID), same convention the B22/B24 cutover tests use.
const RT27 = [...XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS][0];
const RT_PRE27 = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
// an explicit, OBSERVED "nothing stored at this idx yet" value -- the fake driver's `pointStyles` is
// RAW (unfiltered, same as the real driver's chartReads), so a point that was genuinely READ but
// carries no dPt is this shape, never a missing key (a missing key means "never read", a DIFFERENT,
// unobservable case -- see styleBridgeNow's own 'UNOBSERVABLE:point-not-read' branch).
const ABSENT = { element: 'absent', fill: null, outline: null, outlineWidthEmu: null };

async function run(store, t, script = {}, runtime = RT27) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics, mapped: mapExecutionResult(result, { fileId: 'disposable', task: t }) };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;
const point = (series0, abs0) => `${series0}:${abs0}`;

/** a chart fixture carrying an explicit `pointStyles` dict (the fake driver's own convention -- see
 * fake-xlsx-driver.cjs's chartReadsOf, mirroring effectiveLabels) */
const chartWithPoints = (pointStyles, over = {}) => mine({ pointStyles, ...over });

// ---- plan() / styleProblem field validation (charts.cjs's own layer, independent of write-tools.cjs) ----

test('planner: dataPoint field validation', () => {
  const refuse = (over, re) => { const p = planOperations(dataPoint(over)); assert.equal(p.ok, false, JSON.stringify(over)); assert.match(p.detail, re, JSON.stringify(over)); };
  refuse({ seriesIndex: undefined }, /seriesIndex must be a non-negative integer/);
  refuse({ dataPointIndex: undefined }, /dataPointIndex must be a non-negative integer/);
  refuse({ fill: undefined, outline: undefined }, /at least one of fill\/outline is required/);
  refuse({ axis: 'vertical' }, /axis is only accepted for target gridlines/);
});

// ---- write-tools.cjs release gate: RED (before this unit) -> GREEN (after) ----------------------

test('RED->GREEN: write-tools.cjs schema -- dataPoint is a RECOGNISED target enum value and dataPointIndex is a SCHEMA field (both NEW in this unit; before it the enum excluded dataPoint and the schema had no dataPointIndex property at all)', () => {
  const v = validateWriteRequest({ fileId: '123', task: dataPoint() });
  assert.equal(v.ok, true, JSON.stringify(v));
});

test('RED CONTROL, reproduced on a NEIGHBOURING target: an unlisted target is still refused by the schema enum (proves the enum actually discriminates, not just accepts everything)', () => {
  const v = validateWriteRequest({ fileId: '123', task: { id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'plotArea', fill: '#334455' }] } });
  assert.equal(v.ok, false);
  assert.equal(v.code, 'REFUSED_SHAPE');
  assert.match(v.reason, /target must be one of/);
});

test('write-tools.cjs CONSTRAINTS.set_chart_style: dataPointIndex required for target dataPoint, and NOT accepted for marker', () => {
  const miss = validateWriteRequest({ fileId: '123', task: { id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'dataPoint', seriesIndex: 0, fill: '#334455' }] } });
  assert.equal(miss.ok, false);
  assert.match(miss.reason, /dataPointIndex is required for target dataPoint/);
  const leaks = validateWriteRequest({ fileId: '123', task: { id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'marker', seriesIndex: 0, dataPointIndex: 1, fill: '#334455' }] } });
  assert.equal(leaks.ok, false);
  assert.match(leaks.reason, /dataPointIndex is only accepted for target dataPoint/);
});

// ---- full stack: gap (no entry yet) -> APPLY -> write -> SATISFIED on retry ----------------------

test('dataPoint: gap (no stored entry) -- APPLIED, retry SATISFIED', async () => {
  const store = createStore(book([chartWithPoints({ [point(0, 1)]: ABSENT })]));
  const a = await run(store, dataPoint());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, dataPoint());
  assert.equal(r.result.ok, true);
  assert.equal(r.m.commands, 0, 'SATISFIED -- no dispatch needed, the bridge already reads the requested value');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('dataPoint: an ALREADY-present, matching, non-duplicate entry is SATISFIED without any dispatch', async () => {
  const store = createStore(book([chartWithPoints({ [point(0, 1)]: { element: 'present', fill: { form: 'rgb', hex: '334455' }, outline: null, outlineWidthEmu: null, effective: { brush: { form: 'rgb', hex: '334455' }, pen: null } } })]));
  const { result, m } = await run(store, dataPoint());
  assert.equal(m.commands, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('dataPoint: fill-only / outline-only -- the unrequested field is NOT part of the match', async () => {
  const already = { element: 'present', fill: { form: 'rgb', hex: '334455' }, outline: { form: 'rgb', hex: 'AABBCC' }, outlineWidthEmu: 12700, effective: { brush: { form: 'rgb', hex: '334455' }, pen: { form: 'rgb', hex: 'AABBCC' } } };
  const store = createStore(book([chartWithPoints({ [point(0, 1)]: already })]));
  const r = await run(store, dataPoint({ outline: undefined }));
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED, 'fill already matches, outline was never requested');
});

// ---- a DUPLICATE (count>1) is NEVER satisfied, even when the effective (first-hit) fill already matches ----

test('dataPoint: a stored DUPLICATE is APPLY even when the effective fill already matches the request -- normalisation itself is the pending work', async () => {
  const alreadyDup = { element: 'present', fill: { form: 'rgb', hex: '334455' }, outline: null, outlineWidthEmu: null, effective: { brush: { form: 'rgb', hex: '334455' }, pen: null }, duplicated: { count: 2, order: [{ fill: { form: 'rgb', hex: '334455' }, line: null }, { fill: { form: 'rgb', hex: '999999' }, line: null }] } };
  const store = createStore(book([chartWithPoints({ [point(0, 1)]: alreadyDup })]));
  const { result, m } = await run(store, dataPoint());
  assert.equal(m.commands, 1, 'must dispatch -- a duplicate is never SATISFIED, regardless of the effective value');
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
});

// ---- runtime gate: B29 is /27-ONLY, narrower than B22 (gapOverlap/seriesOrder) or B24 (markerSymbol/lineStyle) ----

test('dataPoint: refused BEFORE dispatch on a pre-/27 runtime -- named, never attempted (B29 is narrower than B22/B24)', async () => {
  const { result, driver } = await run(createStore(book([chartWithPoints({})])), dataPoint(), {}, RT_PRE27);
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-bridge-write-runtime-not-proven');
});

test('dataPoint: bridge UNOBSERVABLE (pre-/21 runtime simulated) -- FAIL_CLOSED, named, never a silent write-only fallback to the old public setter', async () => {
  const { result, driver } = await run(createStore(book([chartWithPoints({})])), dataPoint(), { chartStyleUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-bridge-unobservable');
});

test('dataPoint: refused BEFORE dispatch: seriesIndex out of range (the chart has 2 series -- indices 0/1)', async () => {
  const { result, driver } = await run(createStore(book([chartWithPoints({})])), dataPoint({ seriesIndex: 2 }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-series-index-out-of-range');
});

test('dataPoint: receipt carries NO semanticsNotProven -- verified, not write-only (unlike the five remaining original targets)', async () => {
  const store = createStore(book([chartWithPoints({})]));
  const { mapped } = await run(store, dataPoint());
  assert.deepEqual(mapped.operations[0].semanticsNotProven, undefined);
});

// ---- RED (sabotage): the Bridge write dispatches ok, but the read-back value stays stale ----------

test('RED (sabotage): dataPoint dispatch succeeds but the bridge-read value stays stale -- FAIL, never a silent success; a clean retry still APPLIES', async () => {
  const store = createStore(book([chartWithPoints({ [point(0, 1)]: ABSENT })]));
  const a = await run(store, dataPoint(), { fault: 'style-bridge-stale' });
  assert.equal(a.result.ok, false);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  const r = await run(store, dataPoint());
  assert.equal(r.result.ok, true);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
});

// ---- the command body: the Bridge-write args shape, and the MUTATION CONTROL that the public,
// position-based setter is NEVER called on this path -------------------------------------------

test('the command body: a direct chart.series.dataPoint.style.write/v1 Bridge call, gated by B29, carrying series/point/fill -- outline/outlineWidthEmu omitted when outline is not requested', () => {
  const op = planOperations(dataPoint({ outline: undefined })).operations[0];
  const { body, meta } = charts.lower(op);
  assert.equal(meta.bridgeDataPointStyleWrite, true);
  assert.match(body, /__eoDataPointStyleGate/);
  assert.match(body, /__eoDataPointStyleWrite/);
  assert.match(body, /"series":0/);
  assert.match(body, /"point":1/);
  assert.match(body, /"fill":\{"form":"rgb","hex":"334455"\}/);
  assert.doesNotMatch(body, /outline/);
  assert.doesNotMatch(body, /outlineWidthEmu/);
});

test('the command body: outline carries BOTH outline AND outlineWidthEmu (the EMU conversion, pt * 12700 rounded) together, never one without the other', () => {
  const op = planOperations(dataPoint({ fill: undefined, outline: { width: 0.5, color: '#000000' } })).operations[0];
  const { body } = charts.lower(op);
  assert.match(body, /"outline":\{"form":"rgb","hex":"000000"\}/);
  assert.match(body, /"outlineWidthEmu":6350/);
  assert.doesNotMatch(body, /"fill":/);
});

test('MUTATION CONTROL: the dataPoint command body NEVER calls SetDataPointFill/SetDataPointOutLine (the public, position-based setter -- GH #346\'s own root cause), even against a fake editor that supplies both', () => {
  const calls = [];
  const existing = [{ GetName: () => 'C1', SetDataPointFill: () => { calls.push('SetDataPointFill'); return true; }, SetDataPointOutLine: () => { calls.push('SetDataPointOutLine'); return true; } }];
  const ws = { GetName: () => 'K', GetAllCharts: () => existing };
  const Api = { GetSheet: () => ws, CreateSolidFill: v => v, CreateRGBColor: (...x) => x, CreateStroke: (w, f) => [w, f] };
  const op = planOperations(dataPoint({ outline: { width: 1, color: '#000000' } })).operations[0];
  // the body references the Bridge gate/write helpers, which this bare Api fake does not implement --
  // it MUST throw 'ReferenceError: __eoDataPointStyleGate is not defined' rather than fall back to the
  // public setter in any way; either outcome proves the setter path is unreachable, so both are accepted.
  try { new Function('Api', charts.lower(op).body)(Api); } catch (e) { /* expected: no bridge helpers defined in this bare fake */ }
  assert.equal(calls.length, 0, 'neither public setter was ever invoked');
});

// ---- whole-task preservation (L1): every OTHER point/series on the same chart unaffected ----------

test('whole-task: PRIMITIVE.set_chart_style for dataPoint reads-and-compares (SATISFIED/APPLY), and a sabotaged OTHER chart field still fails the "kept" check', async () => {
  const store = createStore(book([chartWithPoints({ [point(0, 1)]: ABSENT })]));
  const { result } = await run(store, dataPoint(), { fault: 'style-alters-legend' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});
