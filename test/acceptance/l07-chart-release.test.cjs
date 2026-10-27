'use strict';

/**
 * L-07 (owner decision L-07): set_chart (create) and set_chart_legend on the PRODUCT surface. The LIVE measurement on the prod DS (DS /12,
 * docs/evidence/l07-charts-fail-closed-20261015) found a chart WRITABLE for all seven types and both orientations, the legend WRITABLE for four positions (the fifth,
 * bottom, is the default: NOOP), and REFUSED, named, 0 writes / 0 commands: a chart of that name with another identity, a closed type, a one-column range, an
 * unsupported legend position, a missing chart. Offline, the product WRITE tool with the fake XLSX editor; the LIVE proof is test/live/release-l07-charts-product.live.cjs
 * (RED on the pre-release build, GREEN on the rehearsal build). The style targets (set_chart_style) are not part of this release.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

const J = x => JSON.stringify(x);
const book = () => [{ name: 'Data', cells: {
  A1: { v: 'k' }, B1: { v: 'S1' }, C1: { v: 'S2' }, D1: { v: 'S3' },
  A2: { v: 'Q1' }, B2: { v: 1 }, C2: { v: 2 }, D2: { v: 3 },
  A3: { v: 'Q2' }, B3: { v: 4 }, C3: { v: 5 }, D3: { v: 6 },
  A4: { v: 'Q3' }, B4: { v: 7 }, C4: { v: 8 }, D4: { v: 9 } } }];
const mk = (type, extra = {}) => ({ intent: 'set_chart', id: 's', sheet: 'Data', name: 'Ch', range: 'A1:D4', type, title: 'T1', ...extra });
const leg = pos => ({ intent: 'set_chart_legend', id: 's', sheet: 'Data', name: 'Ch', legend: pos });
const statuses = r => (r.operations || []).map(o => o.status);
const validate = step => W.validateWriteRequest(req([step]));
const charts = h => h.store.durable.find(x => x.name === 'Data').objects.charts;

test('L-07: set_chart / set_chart_legend are released intents, schema + constraints wired (not sole steps); the style targets stay out', () => {
  for (const i of ['set_chart', 'set_chart_legend']) assert.ok(W.XLSX_WRITE_ALLOWLIST.includes(i), i);
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(W.SOLE_STEP.includes('set_chart'), false);
  for (const t of ['bar', 'lineNormal', 'area', 'pie', 'doughnut', 'radar', 'stock']) assert.equal(validate(mk(t)).ok, true, t);
  assert.equal(validate(mk('bar', { orientation: 'rows' })).ok, true);
  for (const p of ['bottom', 'top', 'left', 'right', 'none']) assert.equal(validate(leg(p)).ok, true, p);
});

test('L-07: the schema + constraints refuse a malformed step before any dispatch', () => {
  const bad = [
    [mk('scatter'), 'a closed type'],
    [mk('surface'), 'a closed type (AddChart throws)'],
    [mk('bar', { range: 'A1:A4' }), 'a one-column range (no category column / data)'],
    [mk('bar', { range: 'A1:D1' }), 'a one-row range (no data)'],
    [mk('bar', { range: 'not a range' }), 'a range that is not A1'],
    [mk('bar', { title: '' }), 'an empty title'],
    [mk('bar', { title: 'a\nb' }), 'a two-line title'],
    [mk('bar', { name: '1bad' }), 'a chart name that starts with a digit'],
    [mk('bar', { sheet: 'My Sheet' }), 'a sheet name that needs quoting in the series reference'],
    [mk('bar', { orientation: 'diagonal' }), 'an unknown orientation'],
    [mk('bar', { legend: 'top' }), 'an unknown field (legend belongs to set_chart_legend)'],
    [{ ...mk('bar'), title: undefined }, 'title missing'],
    [leg('topRight'), 'an unsupported legend position'],
    [leg(''), 'an empty legend position'],
    [{ ...leg('top'), title: 'x' }, 'an unknown field'],
    [{ ...leg('top'), name: '' }, 'an empty chart name'],
  ];
  for (const [s, why] of bad) assert.equal(validate(s).ok, false, why);
});

test('L-07: create a chart -- APPLIED, the chart reads back with its type / title / three series, a retry is SATISFIED with 0 dispatch; each of the seven types', async () => {
  for (const [type, want] of [['bar', 'bar'], ['lineNormal', 'lineNormal'], ['area', 'area'], ['pie', 'pie'], ['doughnut', 'doughnut'], ['radar', 'unknown'], ['stock', 'stock']]) {
    const h = harness({ store: createStore(book()) });
    const r = await h.tool.handler(req([mk(type)]));
    assert.equal(r.ok, true, `${type}: ${J(r).slice(0, 500)}`);
    assert.deepEqual(statuses(r), ['APPLIED'], type);
    assert.equal(charts(h).length, 1, type);
    assert.equal(charts(h)[0].name, 'Ch');
    assert.equal(charts(h)[0].series.length, 3, `${type}: the series are in columns`);
    assert.equal(charts(h)[0].type === want || charts(h)[0].type === type, true, `${type}: ${charts(h)[0].type}`);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([mk(type)]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `${type}: retry`);
  }
});

test('L-07: the rows orientation puts the series in ROWS (one per data row)', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([mk('bar', { orientation: 'rows' })]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.equal(charts(h)[0].series.length, 3, 'A1:D4 has three data rows');
});

test('L-07: set_chart_legend -- a position other than the current one is APPLIED, a retry SATISFIED / 0 dispatch; the current position is a NOOP', async () => {
  for (const pos of ['top', 'left', 'right', 'none']) {
    const h = harness({ store: createStore(book()) });
    assert.equal((await h.tool.handler(req([mk('bar')]))).ok, true);
    const r = await h.tool.handler(req([leg(pos)]));
    assert.equal(r.ok, true, `${pos}: ${J(r).slice(0, 400)}`);
    assert.deepEqual(statuses(r), ['APPLIED'], pos);
    assert.equal(charts(h)[0].legend, pos);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([leg(pos)]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `${pos}: retry`);
  }
  // bottom is where a created chart starts: the target equals the current state
  const h = harness({ store: createStore(book()) });
  await h.tool.handler(req([mk('bar')]));
  const d0 = h.calls().dispatch;
  const same = await h.tool.handler(req([leg('bottom')]));
  assert.deepEqual([same.ok, statuses(same), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], J(same).slice(0, 300));
  // and back to bottom from another position is a real write
  await h.tool.handler(req([leg('top')]));
  const back = await h.tool.handler(req([leg('bottom')]));
  assert.deepEqual([back.ok, statuses(back)], [true, ['APPLIED']]);
  assert.equal(charts(h)[0].legend, 'bottom');
});

test('L-07 NAMED limits: a chart of that name with ANOTHER identity, and a legend on a chart that is not there, are refused -- 0 dispatch, the document untouched, the family reason passes through', async () => {
  const h = harness({ store: createStore(book()) });
  assert.equal((await h.tool.handler(req([mk('bar')]))).ok, true);
  let before = J(h.store.durable); let d0 = h.calls().dispatch;
  const taken = await h.tool.handler(req([mk('lineNormal')]));
  assert.equal(taken.ok, false);
  assert.ok(J(taken).includes('xlsx-chart-exists-with-another-identity'), J(taken).slice(0, 500));
  assert.equal(h.calls().dispatch - d0, 0);
  assert.equal(J(h.store.durable), before);
  const h2 = harness({ store: createStore(book()) });
  before = J(h2.store.durable);
  const none = await h2.tool.handler(req([leg('top')]));
  assert.equal(none.ok, false);
  assert.ok(J(none).includes('xlsx-chart-not-found'), J(none).slice(0, 500));
  assert.equal(h2.calls().dispatch, 0);
  assert.equal(J(h2.store.durable), before);
});

test('L-07 live harness verdict: a synthetic GREEN run passes; the pre-release shape (set_chart not allowlisted) and each broken gate fail', { skip: needsPrivate('test/live/release-l07-charts-product.live.cjs') }, () => {
  const h = require('../live/release-l07-charts-product.live.cjs');
  const ap = { ok: true, persistence: 'FORCESAVE_CONFIRMED', ops: [{ status: 'APPLIED' }] };
  const sat = { ok: true, persistence: null, ops: [{ status: 'SATISFIED' }] };
  const chart = want => ({ name: 'Ch', type: want.type, title: `${want.title}\r\n`, legend: want.legend === undefined ? 'bottom' : want.legend, series: Array.from({ length: want.series }, () => ({})) });
  const green = () => {
    const o = { SEEDS: { a: ap }, PAIRS: {}, NEG_NAMED: {}, NEG_SHAPE: {}, READ: {}, CLEANUP: { gone: true }, FIXTURE_MUTATED: 'NO' };
    for (const p of h.POS) { o.PAIRS[p.key] = { first: ap, second: sat }; o.READ[p.key] = { ok: true, charts: [chart(p.want)] }; }
    for (const n of h.NEG_NAMED) { o.NEG_NAMED[n.key] = { ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: n.reason } }] }; o.READ[n.key] = { ok: true, charts: n.keep ? [chart(n.keep)] : [] }; }
    for (const k of Object.keys(h.NEG_SHAPE)) o.NEG_SHAPE[k] = { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] };
    o.READ.SHAPE = { ok: true, charts: [] };
    return o;
  };
  assert.equal(h.verdict(green()).PASS, true, J(h.verdict(green()).failed));
  assert.equal(h.POS.filter(p => p.key.startsWith('CT_')).length, 8);
  assert.equal(h.POS.filter(p => p.key.startsWith('LG_')).length, 5);
  assert.deepEqual(h.POS.find(p => p.key === 'LG_BOTTOM').setup.map(s => s.intent === 'set_chart_legend' ? s.legend : s.intent), ['set_chart', 'top'], 'bottom is measured from a NON-bottom start (from the default it is a NOOP)');
  const red = green(); red.PAIRS.CT_BAR = { first: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] }, second: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] } };
  assert.ok(h.verdict(red).failed.includes('CT_BAR_applied'), 'the pre-release build is RED');
  const wrongType = green(); wrongType.READ.CT_PIE.charts[0].type = 'bar';
  assert.ok(h.verdict(wrongType).failed.includes('CT_PIE_saved_chart'));
  const noLegend = green(); noLegend.READ.LG_TOP.charts[0].legend = 'bottom';
  assert.ok(h.verdict(noLegend).failed.includes('LG_TOP_saved_chart'), 'a legend that did not reach the saved file');
  const changed = green(); changed.READ.TAKEN.charts[0].type = 'lineNormal';
  assert.ok(h.verdict(changed).failed.includes('neg_TAKEN_unchanged'), 'the refused step must not modify the chart');
  const unnamed = green(); unnamed.NEG_NAMED.NOCHART.ops[0].failure = { reason: 'other' };
  assert.ok(h.verdict(unnamed).failed.includes('neg_NOCHART'));
  const dispatched = green(); dispatched.NEG_SHAPE.scatter = { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [{ status: 'FAILED' }] };
  assert.ok(h.verdict(dispatched).failed.includes('neg_shape_scatter'));
  const strayChart = green(); strayChart.READ.SHAPE.charts = [chart({ type: 'bar', title: 'T1', series: 3 })];
  assert.ok(h.verdict(strayChart).failed.includes('shape_sheet_has_no_chart'));
});
