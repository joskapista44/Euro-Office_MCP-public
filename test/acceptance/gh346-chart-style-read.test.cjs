'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const rt = require('../../src/server/read-tools.cjs');

/** GH #346: read_charts carries the fill / outline of the six style targets through the Bridge /21 chart.series.style.read/v1. */

const CAP = 'chart.series.style.read/v1';
const RGB = hex => ({ form: 'rgb', hex });
const STYLE = {
  chart: { element: 'present', fill: RGB('112233'), outline: RGB('445566'), outlineWidthEmu: 19050 },
  plotArea: { element: 'present', fill: RGB('778899'), outline: null, outlineWidthEmu: null },
  title: { element: 'present', fill: { form: 'scheme', id: 4 }, outline: null, outlineWidthEmu: null },
  legend: { element: 'absent', fill: null, outline: null, outlineWidthEmu: null },
  series: { element: 'present', fill: RGB('AA0000'), outline: RGB('00AA00'), outlineWidthEmu: 12700 },
  'dataPoint:1': { element: 'present', fill: RGB('0000AA'), outline: null, outlineWidthEmu: null },
};

// GH #346 AUTO-REPAIR (/27, documentserver-patches/9.3.5.0/euro-office-bridge-v1/rev27-candidate): the Bridge value for a
// duplicated dataPoint on /27+ -- effective:{brush,pen} (first document-order match) plus duplicated:{count,order}. The
// bridge-block.js handler logic itself (hits.length > 1, the first-match-wins rule) is covered in
// ds935-bridge27-datapoint-effective.test.cjs; the tests below are the PRODUCT side (xlsx-driver.cjs + read-tools.cjs
// pointStylesOf), i.e. that this value rides through unchanged.
const DUP_VALUE = {
  element: 'present', fill: RGB('0000AA'), outline: null, outlineWidthEmu: null,
  effective: { brush: RGB('0000AA'), pen: null },
  duplicated: { count: 2, order: [{ fill: RGB('0000AA'), line: null }, { fill: RGB('112233'), line: RGB('445566') }] },
};

function api({ patchVersion = 'euro-office-bridge-v1/9.3.5.0/21', ambiguousPoint = null, dupPoint = null, dupValue = null } = {}) {
  const calls = [];
  const bridge = {
    manifest: () => ({ bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', patchVersion,
      capabilities: /\/(21|26|27)$/.test(patchVersion) ? [{ id: CAP, version: 1, kind: 'READ', mutation: 'NO' }] : [] }),
    read: (id, a) => {
      calls.push([id, a]);
      if (id !== CAP) return { ok: false, reason: 'unexpected' };
      if (a.target === 'dataPoint') {
        if (a.point === ambiguousPoint) return { ok: false, reason: 'datapoint-ambiguous' };
        if (a.point === dupPoint) return { ok: true, value: dupValue };
        return { ok: true, value: STYLE[`dataPoint:${a.point}`] || { element: 'absent', fill: null, outline: null, outlineWidthEmu: null } };
      }
      return { ok: true, value: STYLE[a.target] };
    },
  };
  const sx = { GetName: () => '=K!$B$1', GetValues: () => '=K!$B$2:$B$4', GetCatFormula: () => '=K!$A$2:$A$4' };
  const ch = { GetName: () => 'C1', GetChartType: () => 'bar', GetTitle: () => 'T', GetLegendPos: () => 'bottom', GetAllSeries: () => [sx], GetHorAxisTitle: () => null, GetVerAxisTitle: () => null, GetDataLabels: () => ({}) };
  const ws = { GetName: () => 'K', GetIndex: () => 0, GetUsedRange: () => null, GetAllCharts: () => [ch], GetAllPivotTables: () => [] };
  return { Api: { GetSheets: () => [ws], GetSheet: n => (n === 'K' ? ws : null), WorksheetFunction: { COUNT: () => 0 }, GetEuroOfficeBridge: () => bridge }, calls };
}
const chartOf = (o, opts) => new Function('Api', readQuery([], { chartSheets: ['K'], ...opts }))(o.Api).chartReads.K.charts[0];

test('the six targets: chart / plotArea / title / legend per chart, the series and every point (the effectiveLabels bound) per series', () => {
  const o = api();
  const c = chartOf(o, { chartStyle: true });
  assert.deepEqual(c.styles, { chart: STYLE.chart, plotArea: STYLE.plotArea, title: STYLE.title, legend: STYLE.legend });
  assert.deepEqual(c.series[0].style, STYLE.series);
  assert.deepEqual(Object.keys(c.pointStyles), ['0:0', '0:1', '0:2']);
  assert.deepEqual(c.pointStyles['0:1'], STYLE['dataPoint:1']);
  const targets = o.calls.filter(([id]) => id === CAP).map(([, a]) => `${a.target}${a.series !== undefined ? ':' + a.series : ''}${a.point !== undefined ? ':' + a.point : ''}`);
  assert.deepEqual(targets, ['chart', 'plotArea', 'title', 'legend', 'series:0', 'dataPoint:0:0', 'dataPoint:0:1', 'dataPoint:0:2']);
});

test('not asked -> no style read at all (every other chart read stays as it was)', () => {
  const o = api();
  const c = chartOf(o, {});
  assert.equal('styles' in c, false);
  assert.equal('pointStyles' in c, false);
  assert.equal('style' in c.series[0], false);
  assert.equal(o.calls.filter(([id]) => id === CAP).length, 0);
});

test('NEG: a /19 runtime names every target, never a guessed null; an ambiguous point is ITS gap only', () => {
  const c = chartOf(api({ patchVersion: 'euro-office-bridge-v1/9.3.5.0/19' }), { chartStyle: true });
  for (const t of ['chart', 'plotArea', 'title', 'legend']) assert.match(c.styles[t], /^UNOBSERVABLE:CAPABILITY_MISSING$/, t);
  assert.match(c.series[0].style, /^UNOBSERVABLE:/);
  const a = chartOf(api({ ambiguousPoint: 2 }), { chartStyle: true });
  assert.equal(a.pointStyles['0:2'], 'UNOBSERVABLE:BRIDGE:datapoint-ambiguous');
  assert.deepEqual(a.pointStyles['0:1'], STYLE['dataPoint:1']);
});

test('read_charts: asks for the styles; the projection lists only the points with their own dPt, and every named gap', () => {
  assert.deepEqual(rt.xlsxRequest('read_charts', { sheet: 'K' }), { windows: [], chartSheets: ['K'], chartStyle: true });
  const c = chartOf(api({ ambiguousPoint: 2 }), { chartStyle: true });
  const out = rt.XLSX_PROJECTIONS.read_charts({ sheets: [{ name: 'K' }], chartReads: { K: { charts: [c], pivots: [] } } }, { sheet: 'K' });
  const p = out.sheets[0].charts[0];
  assert.deepEqual(Object.keys(p.pointStyles), ['0:1', '0:2'], 'the dPt-less point 0:0 inherits the series style and is not listed');
  assert.deepEqual(p.styles.title.fill, { form: 'scheme', id: 4 }, 'a scheme colour is the engine id, never resolved to rgb');
});

// ---- GH #346 AUTO-REPAIR, v2 (review, point 1) ----------------------------------------------------------

test('GH #346 AUTO-REPAIR (/27): a duplicated dataPoint Bridge value carries effective + duplicated through xlsx-driver.cjs and read_charts/pointStylesOf unchanged', () => {
  const o = api({ patchVersion: 'euro-office-bridge-v1/9.3.5.0/27', dupPoint: 1, dupValue: DUP_VALUE });
  const c = chartOf(o, { chartStyle: true });
  assert.deepEqual(c.pointStyles['0:1'], DUP_VALUE, 'bridgeAxis (xlsx-driver.cjs) passes an ok:true Bridge value through as-is');
  const out = rt.XLSX_PROJECTIONS.read_charts({ sheets: [{ name: 'K' }], chartReads: { K: { charts: [c], pivots: [] } } }, { sheet: 'K' });
  assert.deepEqual(out.sheets[0].charts[0].pointStyles['0:1'], DUP_VALUE,
    'pointStylesOf filters entries by element !== "absent", never by field name -- effective/duplicated are not stripped');
});

test('RED (regression, /26): the SAME duplicate shape on a pre-/27 Bridge still refuses datapoint-ambiguous through read_charts/pointStylesOf, unchanged', () => {
  const o = api({ patchVersion: 'euro-office-bridge-v1/9.3.5.0/26', ambiguousPoint: 1 });
  const c = chartOf(o, { chartStyle: true });
  assert.equal(c.pointStyles['0:1'], 'UNOBSERVABLE:BRIDGE:datapoint-ambiguous');
  const out = rt.XLSX_PROJECTIONS.read_charts({ sheets: [{ name: 'K' }], chartReads: { K: { charts: [c], pivots: [] } } }, { sheet: 'K' });
  assert.equal(out.sheets[0].charts[0].pointStyles['0:1'], 'UNOBSERVABLE:BRIDGE:datapoint-ambiguous');
});
