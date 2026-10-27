'use strict';

/**
 * LIVE finding (a maintainer, X-O03/7 r2 on production 1.0.1-cb5d558f0be2 with the rev9 DS): a sheet holding a SCATTER chart
 * made read_charts fail for the WHOLE sheet -- `charts: null, reason: "UNOBSERVABLE:TypeError: GetValues: this series
 * has no values reference"`: the public series GetValues() THROWS on scatter (X-O03/1), and it was called bare in
 * the chartReads series map, so one throwing getter took down every chart and pivot of the sheet. Fixed: the
 * per-series PUBLIC getters (GetName / GetValues / GetCatFormula) are read field by field -- a throw becomes THAT
 * field's 'UNOBSERVABLE:<why>', every other field / series / chart of the sheet is read normally.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

function api(charts) {
  const ws = { GetName: () => 'S1', GetIndex: () => 0, GetUsedRange: () => null, GetAllCharts: () => charts, GetAllPivotTables: () => [{ GetName: () => 'P1' }] };
  return { GetSheets: () => [ws], GetSheet: n => (n === 'S1' ? ws : null), WorksheetFunction: { COUNT: () => 1 } };
}
const chart = (name, type, series) => ({
  GetName: () => name, GetChartType: () => type, GetTitle: () => 'T', GetLegendPos: () => 'right',
  GetAllSeries: () => series, GetHorAxisTitle: () => null, GetVerAxisTitle: () => null,
  GetDataLabels: () => ({ showSerName: false, showCatName: false, showVal: false, showPercent: false }),
});
const throwing = () => { throw new TypeError('GetValues: this series has no values reference'); };
const run = charts => new Function('Api', readQuery([], { chartSheets: ['*'] }))(api(charts));

test('a scatter series whose public GetValues() THROWS does not take down the sheet: the other chart, the pivots and the series name survive', () => {
  const out = run([
    chart('Scatter', 'scatter', [{ GetName: () => '=S1!$B$1', GetValues: throwing, GetCatFormula: () => '=S1!$A$2:$A$4' }]),
    chart('Bar', 'bar', [{ GetName: () => '=S1!$E$1', GetValues: () => '=S1!$E$2:$E$4', GetCatFormula: () => '=S1!$D$2:$D$4' }]),
  ]);
  const r = out.chartReads.S1;
  assert.equal(typeof r, 'object', `the sheet read is not a whole-sheet UNOBSERVABLE string: ${JSON.stringify(r)}`);
  assert.deepEqual(r.pivots, ['P1']);
  assert.deepEqual(r.charts.map(c => [c.name, c.type]), [['Scatter', 'scatter'], ['Bar', 'bar']]);
  const s = r.charts[0].series[0];
  assert.equal(s.name, '=S1!$B$1');
  assert.match(String(s.values), /^UNOBSERVABLE:/);
  assert.equal(s.cat, '=S1!$A$2:$A$4');
  assert.equal(r.charts[1].series[0].values, '=S1!$E$2:$E$4', 'the other chart is read normally');
});

test('GetCatFormula / GetName throwing on a series degrade the same way, field by field', () => {
  const out = run([chart('C', 'bar', [{ GetName: throwing, GetValues: () => '=S1!$B$2:$B$4', GetCatFormula: throwing }])]);
  const s = out.chartReads.S1.charts[0].series[0];
  assert.match(String(s.name), /^UNOBSERVABLE:/);
  assert.match(String(s.cat), /^UNOBSERVABLE:/);
  assert.equal(s.values, '=S1!$B$2:$B$4');
});
