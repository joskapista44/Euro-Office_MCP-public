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
const { seriesStyleBridgeWritePrelude, B24_BRIDGE_WRITE_RUNTIMES } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { BRIDGE_WRITES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');

/**
 * X-O05/5-C (card an internal branch, kioszto zero;
 * owner GO the owner 36662/36663/36664 via X-O05/5-B) -- set_chart_style's TWO more targets,
 * `markerSymbol` (chart.series.marker.write/v1: symbol/size) and `lineStyle`
 * (chart.series.lineStyle.write/v1: dash/smooth), through the FULL product stack (plan ->
 * classify -> lower -> dispatch -> verify) against the fake driver's generic dispatch success
 * path. UNLIKE gapOverlap/seriesOrder (xlsx-xo05-3-...test.cjs) -- which ARE read-and-compare
 * verifiable, SATISFIED on retry -- these two are WRITE-ONLY, like the SIX ORIGINAL style
 * targets (xlsx-xo04-chart-style.test.cjs): never SATISFIED, always re-dispatched on retry,
 * named in the receipt's semanticsNotProven. Served ONLY on DS /15 (B24), never /12../14 (B22) --
 * the runtime-gate tests below are the SAME discipline xlsx-xo05-3's own RT12-vs-unscoped split
 * already establishes, just on the NARROWER /15-only list.
 */

// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): RT15/RT16 (historical 9.3.4.60) are RETIRED --
// the accepted runtime is now the single 9.3.5.0/16 identity. RT14 stays as a NEGATIVE example (never carried
// this /15-only capability in the first place, and is now doubly unaccepted).
const RT935 = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
const RT14 = '9.3.4|euro-office-bridge-v1/9.3.4.60/14';

const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extra = [], type = 'bar') => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: extra, pivots: [], names: [] } },
];
const mine = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', legend: 'bottom', series: series('K', ['B', 'C']), ...over });
const markerSymbol = (over = {}) => ({ id: 'x-o05-5c-ms', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'markerSymbol', seriesIndex: 0, symbol: 'star', ...over }] });
const lineStyle = (over = {}) => ({ id: 'x-o05-5c-ls', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'lineStyle', seriesIndex: 0, dash: 'dashDot', ...over }] });

// markerSymbol is scoped to the SAME MARKER_CHART_TYPES as the fill/outline `marker` target
// (scatter/lineNormal); lineStyle is scoped to its OWN R8q-measured scope (bar/lineNormal) --
// `mine()`'s default `type: 'bar'` works for lineStyle but NOT for markerSymbol, so these tests
// name their own chart type explicitly where it matters.

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT935, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics, mapped: mapExecutionResult(result, { fileId: 'disposable', task: t }) };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

// ---- markerSymbol ------------------------------------------------------------------------------

test('markerSymbol: APPLIED, retry ALSO APPLIED (WRITE-ONLY, like the six original targets -- never SATISFIED)', async () => {
  const store = createStore(book([mine({ type: 'lineNormal' })]));
  const a = await run(store, markerSymbol());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, markerSymbol());
  assert.equal(r.result.ok, true);
  assert.equal(r.m.commands, 1, 'a retry re-dispatches the SAME idempotent write -- it is never 0');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, 'never SATISFIED -- no getter for symbol/size exists');
});

test('markerSymbol: receipt names symbol/size in semanticsNotProven, size-only names just size', async () => {
  const store = createStore(book([mine({ type: 'lineNormal' })]));
  const a = await run(store, markerSymbol({ symbol: 'star', size: 9 }));
  assert.deepEqual(a.mapped.operations[0].semanticsNotProven, ['chart-style-markersymbol-symbol', 'chart-style-markersymbol-size']);
  const store2 = createStore(book([mine({ type: 'lineNormal' })]));
  const b = await run(store2, markerSymbol({ symbol: undefined, size: 9 }));
  assert.deepEqual(b.mapped.operations[0].semanticsNotProven, ['chart-style-markersymbol-size']);
});

test('markerSymbol: refused BEFORE dispatch on an unscoped chart type (bar is not scatter/lineNormal)', async () => {
  const { result, driver } = await run(createStore(book([mine({ type: 'bar' })])), markerSymbol());
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-marker-symbol-type-unsupported');
});

test('markerSymbol: refused BEFORE dispatch -- seriesIndex out of range', async () => {
  const { result, driver } = await run(createStore(book([mine({ type: 'lineNormal' })])), markerSymbol({ seriesIndex: 5 }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-series-index-out-of-range');
});

test('markerSymbol: refused BEFORE dispatch -- the DS runtime does not carry this /15-only capability', async () => {
  const store = createStore(book([mine({ type: 'lineNormal' })]));
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT14 });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(markerSymbol(), { observation: { live: liveScope(markerSymbol()) } }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0, 'refused BEFORE touching the editor, same discipline as gapOverlap/seriesOrder on an unscoped runtime');
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-chart-style-bridge-write-runtime-not-proven');
});

// ---- lineStyle ----------------------------------------------------------------------------------

test('lineStyle: APPLIED, retry ALSO APPLIED (WRITE-ONLY -- never SATISFIED)', async () => {
  const store = createStore(book([mine()]));
  const a = await run(store, lineStyle());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const r = await run(store, lineStyle());
  assert.equal(r.m.commands, 1, 'a retry re-dispatches the SAME idempotent write -- it is never 0');
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
});

test('lineStyle: receipt names dash/smooth in semanticsNotProven, dash-only names just dash', async () => {
  // smooth is ONLY measured on lineNormal (round 2) -- see LINE_STYLE_SMOOTH_CHART_TYPES.
  const store = createStore(book([mine({ type: 'lineNormal' })]));
  const a = await run(store, lineStyle({ dash: 'dashDot', smooth: true }));
  assert.deepEqual(a.mapped.operations[0].semanticsNotProven, ['chart-style-linestyle-dash', 'chart-style-linestyle-smooth']);
  const store2 = createStore(book([mine()]));
  const b = await run(store2, lineStyle({ dash: 'dashDot', smooth: undefined }));
  assert.deepEqual(b.mapped.operations[0].semanticsNotProven, ['chart-style-linestyle-dash']);
});

test('lineStyle: smooth=false is a REQUESTED value, not "unrequested" -- named in semanticsNotProven too', async () => {
  const store = createStore(book([mine({ type: 'lineNormal' })]));
  const { mapped } = await run(store, lineStyle({ dash: undefined, smooth: false }));
  assert.deepEqual(mapped.operations[0].semanticsNotProven, ['chart-style-linestyle-smooth'], 'styleSmooth !== null must stay true for smooth:false');
});

test('lineStyle: smooth refused BEFORE dispatch on bar -- dash IS proven there (R8q B1), smooth is NOT (round 2 LIVE)', async () => {
  const { result, driver } = await run(createStore(book([mine()])), lineStyle({ dash: undefined, smooth: true }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-line-style-smooth-type-unsupported');
});

test('lineStyle: smooth-on-bar is refused even when dash is ALSO requested (smooth is the thing that is unscoped, not the whole op)', async () => {
  const { result, driver } = await run(createStore(book([mine()])), lineStyle({ dash: 'dashDot', smooth: true }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-line-style-smooth-type-unsupported');
});

test('lineStyle: refused BEFORE dispatch on an unscoped chart type (pie is not bar/lineNormal, R8q\'s own measured scope)', async () => {
  const { result, driver } = await run(createStore(book([mine({ type: 'pie' })])), lineStyle());
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-line-style-type-unsupported');
});

test('lineStyle: refused BEFORE dispatch -- seriesIndex out of range', async () => {
  const { result, driver } = await run(createStore(book([mine()])), lineStyle({ seriesIndex: 9 }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-chart-style-series-index-out-of-range');
});

// ---- cross-target refusals (shared enum/shape, both new targets) ------------------------------

test('refused at plan time: the enum now includes markerSymbol/lineStyle', () => {
  const p = planOperations({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'bogus', fill: '#FF0000' }] });
  assert.equal(p.ok, false);
  assert.match(p.detail, /target must be one of plotArea, title, legend, chart, series, dataPoint, marker, gridlines, gapOverlap, seriesOrder, markerSymbol, lineStyle/);
});

test('planner: markerSymbol/lineStyle field validation', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  const step = over => ({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'markerSymbol', seriesIndex: 0, ...over }] });
  refuse(step({ symbol: undefined, size: undefined }), /at least one of symbol\/size is required/);
  refuse(step({ symbol: 'picture' }), /symbol must be one of/);
  refuse(step({ size: 1 }), /size must be an integer between 2 and 72/);
  refuse(step({ seriesIndex: undefined, symbol: 'dot' }), /seriesIndex must be a non-negative integer/);
  refuse(step({ fill: '#FFFFFF' }), /markerSymbol has no fill\/outline/);
  refuse(step({ dash: 'dot' }), /dash\/smooth are not accepted for target markerSymbol/);
  const stepLs = over => ({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'lineStyle', seriesIndex: 0, ...over }] });
  refuse(stepLs({ dash: undefined, smooth: undefined }), /at least one of dash\/smooth is required/);
  refuse(stepLs({ dash: 'invalid' }), /dash must be one of/);
  refuse(stepLs({ smooth: 'yes' }), /smooth must be a boolean/);
  refuse(stepLs({ symbol: 'dot' }), /symbol\/size are not accepted for target lineStyle/);
  // cross-target: the six-original/gapOverlap/seriesOrder/gridlines branches must reject the two new fields too
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'title', fill: '#808080', dash: 'solid' }] }, /symbol\/size\/dash\/smooth are only accepted for target markerSymbol\/lineStyle/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'gapOverlap', gapWidth: 100, symbol: 'dot' }] }, /symbol\/size\/dash\/smooth are not accepted for target gapOverlap/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'seriesOrder', order: [0, 1], smooth: true }] }, /symbol\/size\/dash\/smooth are not accepted for target seriesOrder/);
  refuse({ id: 'x', steps: [{ intent: 'set_chart_style', sheet: 'K', name: 'C1', target: 'gridlines', axis: 'vertical', outline: { width: 1, color: '#000000' }, dash: 'dot' }] }, /symbol\/size\/dash\/smooth are not accepted for target gridlines/);
});

test('the command body: a direct chart.series.marker.write/v1 / chart.series.lineStyle.write/v1 Bridge call, gated by runtime, carrying the series index', () => {
  const opMarker = planOperations(markerSymbol()).operations[0];
  const { body: bodyMarker, meta: metaMarker } = charts.lower(opMarker);
  assert.equal(metaMarker.bridgeSeriesStyleWrite, 'chart.series.marker.write/v1');
  assert.match(bodyMarker, /__eoSeriesStyleGate/);
  assert.match(bodyMarker, /"series":0/);
  assert.match(bodyMarker, /"symbol":"star"/);

  const opLine = planOperations(lineStyle({ dash: 'dashDot', smooth: true })).operations[0];
  const { body: bodyLine, meta: metaLine } = charts.lower(opLine);
  assert.equal(metaLine.bridgeSeriesStyleWrite, 'chart.series.lineStyle.write/v1');
  assert.match(bodyLine, /__eoSeriesStyleGate/);
  // smooth MUST precede dash in the args/requested order -- it matches the handler's own
  // apply-order (build-rev15.py's LINE_STYLE_WRITE_IMPL: smooth applied before dash).
  assert.ok(bodyLine.indexOf('"smooth"') < bodyLine.indexOf('"dash"'), 'smooth requested before dash');
});

test('STATIC: the in-editor series-style write client is syntactically valid JS and tolerates an optional widthSource/fillSource receipt that marker.write/v1 never sends', () => {
  const cap = BRIDGE_WRITES.LINE_STYLE_WRITE;
  const prelude = seriesStyleBridgeWritePrelude(cap);
  // a minimal fake Api/bridge -- proves the TEMPLATE STRING parses and the gate/write wiring
  // runs end-to-end, not that the real DS handler's exact behaviour is reproduced (that is the
  // round8q offline test's job, on the DS-side handler itself).
  const fakeManifest = { bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', patchVersion: 'euro-office-bridge-v1/9.3.5.0/16', capabilities: [{ id: cap.id, version: cap.version, kind: 'WRITE', mutation: 'TARGETED' }] };
  const fakeBridge = { manifest: () => fakeManifest, write: (id, args) => ({ ok: true, value: { applied: ['dash'], widthSource: 'preserved', fillSource: 'preserved' } }) };
  const fakeApi = { GetEuroOfficeBridge: () => fakeBridge };
  const body = `${prelude}
  const g = __eoSeriesStyleGate('S', 'C1');
  if (!g.ok) return g;
  const w = __eoSeriesStyleWrite({ series: 0, dash: 'dashDot' }, ['dash']);
  return w;`;
  const fn = new Function('Api', body);
  const r = fn(fakeApi);
  assert.deepEqual(r, { ok: true, widthSource: 'preserved', fillSource: 'preserved' });
});

test('STATIC: the same write client REFUSES a receipt with an unknown extra key, never silently accepting it', () => {
  const cap = BRIDGE_WRITES.MARKER_WRITE;
  const prelude = seriesStyleBridgeWritePrelude(cap);
  const fakeManifest = { bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', patchVersion: 'euro-office-bridge-v1/9.3.5.0/16', capabilities: [{ id: cap.id, version: cap.version, kind: 'WRITE', mutation: 'TARGETED' }] };
  const fakeBridge = { manifest: () => fakeManifest, write: () => ({ ok: true, value: { applied: ['symbol'], somethingElse: 'x' } }) };
  const fakeApi = { GetEuroOfficeBridge: () => fakeBridge };
  const body = `${prelude}
  const g = __eoSeriesStyleGate('S', 'C1');
  if (!g.ok) return g;
  return __eoSeriesStyleWrite({ series: 0, symbol: 'dot' }, ['symbol']);`;
  const fn = new Function('Api', body);
  const r = fn(fakeApi);
  assert.equal(r.ok, false);
  assert.equal(r.reason, 'UNEXPECTED_SHAPE');
});

test('B24_BRIDGE_WRITE_RUNTIMES is /16 + /17 (DS9346: the historical /15, /16 9.3.4.60 chain is retired; GH #326/#327 widens again, byte-identical handler proven by b24-write-handler-identity.test.cjs)', () => {
  assert.deepEqual(B24_BRIDGE_WRITE_RUNTIMES, [RT935, '9.3.5|euro-office-bridge-v1/9.3.5.0/17']);
});
