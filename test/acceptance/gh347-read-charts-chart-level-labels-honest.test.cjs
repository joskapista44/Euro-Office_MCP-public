'use strict';

/**
 * GH #347 (CH02) unit 2: read_charts must not state an EFFECTIVE data-label state it has not read. The chart-level `dataLabels` is
 * the chart-type element's own c:dLbls (GetDataLabels); every series keeps its own all-false c:dLbls that overrides it, so after a
 * chart-level SetShowDataLabels the flags read true while no point shows anything (LIVE r2, docs/evidence/gh347-chart-labels-probe-live-r2-20261031.json).
 * The projection therefore NAMES the old field's scope and adds `dataLabelsEffective`, summarised ONLY from effectiveLabels.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { FORMATS } = require('../../src/server/read-tools.cjs');

const F = (o = {}) => ({ showSerName: false, showCatName: false, showVal: false, showPercent: false, ...o });
const pt = (level, o) => ({ level, ...F(o) });
const read = charts => FORMATS.xlsx.project({ chartReads: { S: { charts, pivots: [] } } }, { operation: 'read_charts', args: {} }).sheets[0].charts;
const chart = (extra = {}) => ({ name: 'c', type: 'bar', title: null, legend: null, series: [], horAxisTitle: null, verAxisTitle: null, dataLabels: F(), ...extra });
/** the LIVE r2 bar_b shape: the chart-level flags read showVal true, every one of the 2 x 3 points is SERIES, all false */
const BAR_B = () => chart({ dataLabels: F({ showVal: true }), effectiveLabels: Object.fromEntries([0, 1].flatMap(s => [0, 1, 2].map(p => [`${s}:${p}`, pt('SERIES')]))) });

test('RED on the old projection: the chart-level showVal:true is NOT presented as an effective state -- the scope is named and the effective summary says every point shows nothing', () => {
  const [c] = read([BAR_B()]);
  assert.deepEqual(c.dataLabels, F({ showVal: true }), 'the old field stays, backward compatible');
  assert.equal(c.dataLabelsScope, 'chart-group-setting');
  assert.deepEqual(c.dataLabelsEffective, F(), 'NOT showVal:true -- no point shows a value');
});

test('all points identical -> that one four-flag state (the level is dropped: it is per point); differing points -> "mixed"', () => {
  const same = Object.fromEntries([0, 1, 2].map(p => [`0:${p}`, pt('POINT', { showVal: true })]));
  assert.deepEqual(read([chart({ effectiveLabels: same })])[0].dataLabelsEffective, F({ showVal: true }));
  assert.deepEqual(read([chart({ effectiveLabels: { ...same, '0:1': pt('POINT', { showCatName: true }) } })])[0].dataLabelsEffective, 'mixed', 'LIVE r2 bar_c: one point written, the rest SERIES');
  assert.deepEqual(read([chart({ effectiveLabels: { '0:0': pt('SERIES'), '0:1': pt('POINT', { showSerName: true }) } })])[0].dataLabelsEffective, 'mixed');
});

test('NEG: unreadable effective state is a named UNOBSERVABLE, never derived from the chart-level GetDataLabels', () => {
  const lab = F({ showVal: true });
  assert.equal(read([chart({ dataLabels: lab, effectiveLabels: 'UNOBSERVABLE:point-count-unknown' })])[0].dataLabelsEffective, 'UNOBSERVABLE:point-count-unknown');
  assert.equal(read([chart({ dataLabels: lab, effectiveLabels: 'UNOBSERVABLE:chart-label-points-truncated' })])[0].dataLabelsEffective, 'UNOBSERVABLE:chart-label-points-truncated');
  assert.equal(read([chart({ dataLabels: lab, effectiveLabels: { '0:0': pt('SERIES'), '0:1': 'UNOBSERVABLE:CAPABILITY_MISSING' } })])[0].dataLabelsEffective, 'UNOBSERVABLE:CAPABILITY_MISSING', 'one unreadable point makes the whole summary that reason, never a partial list');
  assert.equal(read([chart({ dataLabels: lab })])[0].dataLabelsEffective, 'UNOBSERVABLE:not-read', 'effectiveLabels absent');
  assert.equal(read([chart({ dataLabels: lab, effectiveLabels: {} })])[0].dataLabelsEffective, 'UNOBSERVABLE:no-points');
  assert.equal(read([chart({ dataLabels: lab, effectiveLabels: { '0:0': { level: 'SERIES', showVal: true } } })])[0].dataLabelsEffective, 'UNOBSERVABLE:point-flags-incomplete', 'a point without all four booleans');
});

test('a sheet whose charts are not a list (a named gap) and every other chart field pass through untouched', () => {
  const r = FORMATS.xlsx.project({ chartReads: { S: 'UNOBSERVABLE:sheet-missing' } }, { operation: 'read_charts', args: {} }).sheets[0];
  assert.deepEqual(r, { sheet: 'S', charts: null, pivots: null, reason: 'UNOBSERVABLE:sheet-missing' });
  const src = BAR_B();
  const [c] = read([src]);
  for (const k of Object.keys(src)) assert.deepEqual(c[k], src[k], k);
});
