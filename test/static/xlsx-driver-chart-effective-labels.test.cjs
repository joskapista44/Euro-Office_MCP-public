'use strict';

/**
 * X-O05/4 (design docs/X-O05-3-DATALABEL-READ-CONSUMER-DESIGN-20261013.md): the per-point
 * effectiveLabels loop in xlsx-driver.cjs's chartReads, run against the REAL generated script
 * body (new Function('Api', readQuery(...))) and a fake Api -- the SAME "evaluate the real
 * generated body" technique test/static/xlsx-driver-read-query.test.cjs and
 * test/static/ds-bridge-client.test.cjs already use. The bridge itself is a hand-rolled stub
 * (manifest + read), not the real documentserver-patches bridge-block -- that block's own
 * internal chart-model shape is exercised LIVE only (the per-point loop here is driver code,
 * independent of what the bridge's OWN implementation does with a given (series,point)).
 *
 * X-O05/4 round 2 (a maintainer's LIVE r1, internal tracker): round 1's fixture used a plain JS array for
 * GetValues(), which is NOT the real shape -- a real chart's GetValues() is a FORMULA STRING
 * ('$B$2:$C$2', or a bare single-cell ref with no colon), which round 1's Array.isArray guard
 * silently read as 0 points for every series. These fixtures now use formula strings (the single
 * remaining array-based case is kept as its own, explicitly-labelled defensive-input test, not
 * the default shape).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { readQuery, MAX_CHART_LABEL_POINTS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');

const CAP = BRIDGE_CAPABILITIES.CHART_LABELS_EFFECTIVE_READ;
// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the default manifest now carries the
// ACCEPTED 9.3.5.0/16 identity -- the historical 9.3.4.60/5 baseline is retired.
const manifest = (over = {}) => ({
  bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', documentServerBuild: '9.3.5-x',
  patchVersion: 'euro-office-bridge-v1/9.3.5.0/16',
  capabilities: [{ id: CAP.id, version: 1, kind: 'READ', format: 'XLSX', mutation: 'NO' }],
  ...over,
});

/** A single-row, multi-column range formula -- the real shape AddChart(seriesInRows=true) gives a
 * row-series' values (a maintainer's r1: '=X5O4_70bf66!$B$2:$C$2'). `n` points -> columns B..(B+n-1),
 * using full multi-letter column names (AA, AB, ...) so a cap-sized (300+) range is well-formed. */
const colName = num => { let s = ''; let n = num; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const colLetter = i => colName(2 + i); // 0 -> B (col 2), 1 -> C (col 3), ...
const rowRange = (row, n) => (n === 1 ? `=S!$${colLetter(0)}$${row}` : `=S!$${colLetter(0)}$${row}:$${colLetter(n - 1)}$${row}`);

/** A fake chart: `seriesValues[i]` is series i's GetValues() return -- a formula string (the real
 * shape) or an array (a defensive-input case, see the dedicated test below). */
function fakeChart(name, seriesValues) {
  return {
    GetName: () => name, GetChartType: () => 'bar', GetTitle: () => 'T', GetLegendPos: () => 'right',
    GetAllSeries: () => seriesValues.map((v, i) => ({ GetName: () => `=S!$A$${i + 2}`, GetValues: () => v, GetCatFormula: () => 'x' })),
    GetHorAxisTitle: () => null, GetVerAxisTitle: () => null,
    GetDataLabels: () => ({ showSerName: false, showCatName: false, showVal: false, showPercent: false }),
  };
}

function fakeApi({ charts = [], read, m = manifest(), noBridge = false } = {}) {
  // GetUsedRange() -> null: no used range, so the unrelated readRange(ws, ws.GetRange(s.row.used), ...) default-sweep
  // (usedRanges, true by default) never fires and never needs a GetRange stub this fixture does not model.
  const ws = { GetName: () => 'S1', GetIndex: () => 0, GetUsedRange: () => null, GetAllCharts: () => charts, GetAllPivotTables: () => [] };
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'S1' ? ws : null), WorksheetFunction: { COUNT: () => 1 } };
  if (!noBridge) api.GetEuroOfficeBridge = () => ({ manifest: () => m, read });
  return api;
}

function run(api) {
  // eslint-disable-next-line no-new-func
  return new Function('Api', readQuery([], { chartSheets: ['*'] }))(api);
}

const PASS_VALUE = { level: 'SERIES', showSerName: false, showCatName: false, showVal: false, showPercent: false };

test('a chart under the cap gets effectiveLabels from REAL formula-string values, one bridge call per (series,point), old fields unchanged', () => {
  const seen = [];
  const read = (id, args) => { seen.push([id, args]); return { ok: true, value: PASS_VALUE }; };
  // series 0: row 2, columns B-C (2 points) -- a maintainer's r1 shape exactly; series 1: a bare single-cell ref (1 point)
  const chart = fakeChart('C1', [rowRange(2, 2), rowRange(3, 1)]);
  const out = run(fakeApi({ charts: [chart], read }));
  assert.equal(out.ok, true);
  const c = out.chartReads.S1.charts[0];
  assert.equal(c.name, 'C1'); assert.equal(c.type, 'bar'); assert.equal(c.title, 'T'); assert.equal(c.legend, 'right');
  assert.equal(c.series.length, 2);
  assert.deepEqual(c.dataLabels, { showSerName: false, showCatName: false, showVal: false, showPercent: false });
  assert.deepEqual(Object.keys(c.effectiveLabels).sort(), ['0:0', '0:1', '1:0']);
  assert.deepEqual(c.effectiveLabels['0:0'], PASS_VALUE);
  assert.equal(seen.length, 3);
  assert.deepEqual(seen[0], [CAP.id, { sheet: 'S1', chart: 'C1', series: 0, point: 0 }]);
  assert.deepEqual(seen[2], [CAP.id, { sheet: 'S1', chart: 'C1', series: 1, point: 0 }]);
});

test('a single-column range (same column, rows differ) counts rows, not columns', () => {
  const read = () => ({ ok: true, value: PASS_VALUE });
  const chart = fakeChart('COLRANGE', ['=S!$C$2:$C$6']); // 5 rows, same column -> 5 points
  const out = run(fakeApi({ charts: [chart], read }));
  const c = out.chartReads.S1.charts[0];
  assert.deepEqual(Object.keys(c.effectiveLabels).sort(), ['0:0', '0:1', '0:2', '0:3', '0:4']);
});

test('a chart over MAX_CHART_LABEL_POINTS gets ONE named reason string, never a partial per-point list; the chart\'s other fields are unaffected', () => {
  let calls = 0;
  const read = () => { calls += 1; return { ok: true, value: PASS_VALUE }; };
  const chart = fakeChart('BIG', [rowRange(2, MAX_CHART_LABEL_POINTS + 1)]);
  const out = run(fakeApi({ charts: [chart], read }));
  const c = out.chartReads.S1.charts[0];
  assert.equal(c.effectiveLabels, 'UNOBSERVABLE:chart-label-points-truncated');
  assert.equal(c.name, 'BIG'); assert.equal(c.series.length, 1);
  assert.equal(calls, 0, 'no bridge call is made once the chart is over the cap');
});

test('a chart at exactly the cap still reads every point -- the boundary is inclusive, not exclusive', () => {
  let calls = 0;
  const read = () => { calls += 1; return { ok: true, value: PASS_VALUE }; };
  const chart = fakeChart('EXACT', ['=S!$A$1:$A$' + MAX_CHART_LABEL_POINTS]); // a single-column range, MAX points
  const out = run(fakeApi({ charts: [chart], read }));
  const c = out.chartReads.S1.charts[0];
  assert.notEqual(c.effectiveLabels, 'UNOBSERVABLE:chart-label-points-truncated');
  assert.equal(calls, MAX_CHART_LABEL_POINTS);
});

test('an unparseable values formula (a 2D range) makes the WHOLE chart UNOBSERVABLE:point-count-unknown, never a guessed {} -- round 2\'s fix for the round-1 bug', () => {
  const read = () => { throw new Error('must not be called -- the point count is unknown before any bridge call'); };
  const chart = fakeChart('TWO_D', ['=S!$B$2:$C$5']); // both rows AND columns differ -- not a single-axis range
  const out = run(fakeApi({ charts: [chart], read }));
  const c = out.chartReads.S1.charts[0];
  assert.equal(c.effectiveLabels, 'UNOBSERVABLE:point-count-unknown');
  assert.equal(c.name, 'TWO_D'); // the chart's own non-label fields still read
});

test('ONE unparseable series among several poisons the WHOLE chart (point-count-unknown), not just that series -- a per-chart decision, made before any bridge call', () => {
  let calls = 0;
  const read = () => { calls += 1; return { ok: true, value: PASS_VALUE }; };
  const chart = fakeChart('MIXED', [rowRange(2, 2), '=S!$B$2:$C$5' /* 2D, unparseable */, rowRange(4, 1)]);
  const out = run(fakeApi({ charts: [chart], read }));
  const c = out.chartReads.S1.charts[0];
  assert.equal(c.effectiveLabels, 'UNOBSERVABLE:point-count-unknown');
  assert.equal(calls, 0, 'no partial per-series bridge calls happen once any series is unparseable');
});

test('fail-closed: BRIDGE_UNAVAILABLE surfaces verbatim per point; other points, and the chart\'s own non-label fields, are unaffected', () => {
  const chart = fakeChart('C2', [rowRange(2, 2)]);
  const out = run(fakeApi({ charts: [chart], noBridge: true }));
  const c = out.chartReads.S1.charts[0];
  assert.deepEqual(c.effectiveLabels, { '0:0': 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', '0:1': 'UNOBSERVABLE:BRIDGE_UNAVAILABLE' });
  assert.equal(c.name, 'C2'); assert.equal(c.type, 'bar');
});

test('fail-closed: RUNTIME_VERSION_MISMATCH surfaces verbatim per point, independent of the per-point loop', () => {
  const chart = fakeChart('C3', [rowRange(2, 2)]);
  const out = run(fakeApi({ charts: [chart], read: () => ({ ok: true, value: PASS_VALUE }), m: manifest({ documentServerVersion: '9.9.9' }) }));
  const c = out.chartReads.S1.charts[0];
  assert.deepEqual(c.effectiveLabels, { '0:0': 'UNOBSERVABLE:RUNTIME_VERSION_MISMATCH', '0:1': 'UNOBSERVABLE:RUNTIME_VERSION_MISMATCH' });
});

test('DEFENSIVE INPUT (not the real shape, see the header note): GetValues() as a plain array is still supported -- length-0 asks the bridge nothing, effectiveLabels is {} not UNOBSERVABLE', () => {
  const read = () => { throw new Error('must not be called'); };
  const chart = fakeChart('EMPTY', [[]]);
  const out = run(fakeApi({ charts: [chart], read }));
  const c = out.chartReads.S1.charts[0];
  assert.deepEqual(c.effectiveLabels, {});
});

test('a sheet with no charts at all stays an empty charts array -- effectiveLabels never appears on nothing', () => {
  const out = run(fakeApi({ charts: [] }));
  assert.deepEqual(out.chartReads.S1.charts, []);
});
