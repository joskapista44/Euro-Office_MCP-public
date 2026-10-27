'use strict';

/**
 * GH #302 G5: set_chart_point_data_label -- one point's four data-label flags through the public SetShowPointDataLabel,
 * verified by read_charts' effectiveLabels (chart.dataLabels.effective.read/v1) in the measured shape
 * (docs/evidence/gh302-g5-point-data-label-probe-20261026.json).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/chart-point-labels.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const OFF = Object.freeze({ showSerName: false, showCatName: false, showVal: false, showPercent: false });
const lbl = (level, extra = {}) => ({ level, ...OFF, ...extra });
const chart = (name, type, labels, extra = {}) => ({ name, type, title: null, legend: 'right', series: [{ name: 'S1', values: '$B$2:$D$2', cat: '$B$1:$D$1' }], dataLabels: { ...OFF }, effectiveLabels: labels, ...extra });
const BAR = () => ({ '0:0': lbl('SERIES'), '0:1': lbl('SERIES'), '0:2': lbl('SERIES') });
const obs = charts => ({ format: 'xlsx', complete: true, sheets: [{ name: 'S', index: 0, used: null }], cells: {}, chartReads: { S: { charts, pivots: [] } }, drawingReads: { S: [] } });
const stepOf = (extra = {}) => ({ intent: 'set_chart_point_data_label', sheet: 'S', name: 'C1', series: 0, point: 1, labels: { ...OFF, showVal: true }, ...extra });
const opOf = s => { const p = fam.plan(s, 0); assert.equal(p.ok, true, p.detail); return p.operation; };

test('G5 schema + plan: all four flags required; an extra flag, a missing one, a non-boolean, a negative index, an extra field refused', () => {
  assert.equal(W.validateWriteRequest(req([stepOf()])).ok, true);
  for (const labels of [{ showVal: true }, { ...OFF, showLegendKey: true }, { ...OFF, showVal: 1 }]) {
    assert.equal(fam.plan(stepOf({ labels }), 0).ok, false, JSON.stringify(labels));
    assert.equal(W.validateWriteRequest(req([stepOf({ labels })])).ok, false, JSON.stringify(labels));
  }
  assert.equal(fam.plan(stepOf({ point: -1 }), 0).ok, false);
  assert.equal(fam.plan(stepOf({ series: 1.5 }), 0).ok, false);
  assert.equal(W.validateWriteRequest(req([stepOf({ color: 'red' })])).ok, false);
});

test('G5 classify: other flags -> APPLY; these flags already effective (whatever level) -> SATISFIED; named FAIL_CLOSED reasons', () => {
  const op = opOf(stepOf());
  assert.equal(fam.classify(obs([chart('C1', 'bar', BAR())]), op).state, 'APPLY');
  assert.equal(fam.classify(obs([chart('C1', 'bar', { ...BAR(), '0:1': lbl('SERIES', { showVal: true }) })]), op).state, 'SATISFIED', 'a series-level value label already shows exactly this');
  const off = opOf(stepOf({ labels: { ...OFF } }));
  assert.equal(fam.classify(obs([chart('C1', 'pie', { '0:0': lbl('POINT'), '0:1': lbl('POINT'), '0:2': lbl('POINT') })]), off).state, 'SATISFIED', 'the untouched pie: every point POINT, all off (measured)');
  const reason = o => fam.classify(o, op).evidence.reason;
  assert.equal(reason(obs([chart('Other', 'bar', BAR())])), 'xlsx-chart-point-label-no-such-chart');
  assert.equal(reason(obs([chart('C1', 'bar', BAR()), chart('C1', 'bar', BAR())])), 'xlsx-chart-point-label-no-such-chart');
  assert.equal(reason(obs([chart('C1', 'bar', 'UNOBSERVABLE:chart-label-points-truncated')])), 'xlsx-chart-point-label-unobservable');
  assert.equal(reason(obs([chart('C1', 'bar', { '0:0': lbl('SERIES') })])), 'xlsx-chart-point-label-no-such-point');
  assert.equal(reason(obs([chart('C1', 'bar', { ...BAR(), '0:1': 'UNOBSERVABLE:CAPABILITY_MISSING' })])), 'xlsx-chart-point-label-unobservable');
  assert.equal(reason({ ...obs([]), chartReads: { S: 'UNOBSERVABLE:sheet-missing' } }), 'xlsx-chart-unobservable');
});

test('G5 classify: showPercent only on a pie (the measured type); a pie percent -> APPLY', () => {
  const pct = opOf(stepOf({ labels: { ...OFF, showPercent: true } }));
  const bar = fam.classify(obs([chart('C1', 'bar', BAR())]), pct);
  assert.equal(bar.state, 'FAIL_CLOSED');
  assert.equal(bar.evidence.reason, 'xlsx-chart-point-label-percent-not-measured-on-type');
  assert.equal(fam.classify(obs([chart('C1', 'pie', { '0:0': lbl('POINT'), '0:1': lbl('POINT'), '0:2': lbl('POINT') })]), pct).state, 'APPLY');
});

test('G5 lower: one chart of the name, the series re-checked, the setter with the four flags positionally', () => {
  const body = fam.lower(opOf(stepOf({ series: 1, point: 2, labels: { showSerName: false, showCatName: true, showVal: true, showPercent: false } }))).body;
  assert.match(body, /chart-identity/);
  assert.match(body, /GetAllSeries\(\) \|\| \[\]\)\.length <= 1\)/);
  assert.match(body, /SetShowPointDataLabel\(1, 2, false, true, true, false\) === false/);
});

test('G5 verify: L1 the point\'s effective flags; L3 every other point, chart field, chart and drawing kept', () => {
  const op = opOf(stepOf());
  const before = { primary: obs([chart('C1', 'bar', BAR()), chart('C2', 'bar', BAR())]) };
  const good = obs([chart('C1', 'bar', { ...BAR(), '0:1': lbl('POINT', { showVal: true }) }), chart('C2', 'bar', BAR())]);
  assert.equal(fam.verifyPrimitive(op, before, { primary: good }).ok, true);
  assert.equal(fam.verifyPrimitive(op, before, { primary: obs([chart('C1', 'bar', { ...BAR(), '0:1': lbl('POINT', { showVal: true, showCatName: true }) })]) }).ok, false, 'one flag too many');
  assert.equal(fam.verifyPreservation(op, before, { primary: good }).ok, true);
  const otherPoint = obs([chart('C1', 'bar', { ...BAR(), '0:1': lbl('POINT', { showVal: true }), '0:2': lbl('POINT') }), chart('C2', 'bar', BAR())]);
  assert.equal(fam.verifyPreservation(op, before, { primary: otherPoint }).ok, false, 'another point of the chart changed level');
  const otherChart = obs([chart('C1', 'bar', { ...BAR(), '0:1': lbl('POINT', { showVal: true }) }), chart('C2', 'bar', { ...BAR(), '0:0': lbl('POINT', { showVal: true }) })]);
  assert.equal(fam.verifyPreservation(op, before, { primary: otherChart }).ok, false, 'another chart\'s label changed');
  const chartLevel = obs([chart('C1', 'bar', { ...BAR(), '0:1': lbl('POINT', { showVal: true }) }, { dataLabels: { ...OFF, showVal: true } }), chart('C2', 'bar', BAR())]);
  assert.equal(fam.verifyPreservation(op, before, { primary: chartLevel }).ok, false, 'the chart-level labels changed (the probe: they must not)');
});

test('G5 routing: set_chart_point_data_label plans through the chart-point-labels family, its chart claimed', () => {
  const p = planOperations({ id: 't', steps: [stepOf()] });
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operations[0].family, 'chart-point-labels');
  assert.equal(planOperations({ id: 't', steps: [stepOf(), { intent: 'set_chart_title', sheet: 'S', name: 'C1', title: 'x' }] }).ok, false, 'two steps on one chart clash');
});
