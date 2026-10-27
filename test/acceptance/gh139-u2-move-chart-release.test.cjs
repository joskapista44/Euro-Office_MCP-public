'use strict';

/**
 * GH #139 U2: move_chart on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The family-level
 * behaviour is in test/acceptance/xlsx-x-o06-move-chart.test.cjs (X-O06-6, already adapter-complete and
 * LIVE-measured); this file is the product-tool level, mirroring gh138-u2-print-titles-release.test.cjs. The
 * READ mirror is read_drawings's own pos fields on the chart kind (GH #139 U1 matrix row C16,
 * docs/GH139-CHART-DRAWING-SYMMETRY-MATRIX.md); the LIVE proof is test/live/release-gh139-u2-move-chart-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const series = (s, cols) => cols.map(c => ({ name: `=${s}!$${c}$1`, values: `=${s}!$${c}$2:$${c}$4`, cat: `=${s}!$A$2:$A$4`, x: '' }));
const book = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, A2: { v: 'a' }, B2: { v: 1 }, A3: { v: 'b' }, B3: { v: 2 }, A4: { v: 'c' }, B4: { v: 3 } }, objects: { charts: extra, pivots: [], names: [] } },
];
const mine = (over = {}) => ({ name: 'C1', type: 'bar', title: 'Sales', series: series('K', ['B']), pos: { fromCol: 2, colOffset: 0, fromRow: 12, rowOffset: 0 }, w: 3600000, h: 2160000, ...over });
const move = (over = {}) => ({ intent: 'move_chart', id: 's', sheet: 'K', name: 'C1', at: { col: 8, row: 20 }, ...over });
const statuses = r => (r.operations || []).map(o => o.status);
const posOf = (store, sheet, name) => store.durable.find(x => x.name === sheet).objects.charts.find(c => c.name === name).pos;

test('GH #139 U2: move_chart is a released intent (allowlist + accepted set + schema + constraints), not a sole step, the release gate passes', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('move_chart'));
  assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes('move_chart'));
  assert.equal(W.SOLE_STEP.includes('move_chart'), false);
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].sort());
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].filter(i => i !== 'move_chart'), docxAllowlist: [] }).ok, false, 'move_chart missing');
});

test('GH #139 U2: the schema + CONSTRAINTS refuse a malformed step before any dispatch (the "before the release" shape, REFUSED_NOT_ALLOWLISTED at the family level -- here checked at the schema/plan level)', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(move()).ok, true);
  const bad = [
    [move({ at: undefined }), 'at missing'],
    [{ intent: 'move_chart', id: 's', sheet: 'K', name: 'C1' }, 'at entirely absent'],
    [move({ at: { col: -1, row: 1 } }), 'negative col'],
    [move({ at: { col: 1.5, row: 1 } }), 'non-integer col'],
    [move({ at: { col: 1, row: -1 } }), 'negative row'],
    [move({ at: { col: 1, row: 1.5 } }), 'non-integer row'],
    [move({ at: { col: 1, row: 1, colOffset: 0 } }), 'at carries an extra key'],
    [move({ colOffset: 5 }), 'move_chart does not support colOffset (offset is always 0 at this boundary)'],
    [move({ sheet: '' }), 'an empty sheet name'],
    [move({ name: '' }), 'an empty chart name'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('GH #139 U2: APPLIED to the requested cell, retry SATISFIED 0 dispatch', async () => {
  const h = harness({ store: createStore(book([mine()])) });
  const first = await h.tool.handler(req([move()]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  assert.deepEqual(posOf(h.store, 'K', 'C1'), { fromCol: 8, colOffset: 0, fromRow: 20, rowOffset: 0 });
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([move()]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('GH #139 U2: a nonexistent chart name is refused named xlsx-chart-not-found, nothing dispatched -- the FAIL_CLOSED evidence this release was missing on the product surface (census.json, move_chart, M11.0 checkpoint)', async () => {
  const h = harness({ store: createStore(book()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([move()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-chart-not-found/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('GH #139 U2: two charts of the same name are refused named xlsx-chart-name-ambiguous, nothing dispatched', async () => {
  const h = harness({ store: createStore(book([mine(), mine()])) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([move()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-chart-name-ambiguous/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('GH #139 U2: the large-book path is the NAMED refusal, NOT the tiered observation (same as every other chart intent) -- move_chart is NOT in TIERED_INTENTS', () => {
  const sc = s => liveScope({ id: 't', steps: [s] });
  assert.equal(sc(move()).tiered, undefined, 'not in TIERED_INTENTS -- the key is absent, not false');
  assert.equal(sc({ intent: 'copy_chart', id: 's', sheet: 'K', name: 'C2', sourceSheet: 'K', sourceName: 'C1', at: { col: 0, row: 0 } }).tiered, undefined, 'the sibling chart intent stays not-tiered too, unaffected');
});
