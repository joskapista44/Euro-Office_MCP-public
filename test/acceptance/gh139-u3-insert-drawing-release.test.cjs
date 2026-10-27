'use strict';

/**
 * GH #139 U3: insert_shape / insert_image on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The
 * family-level behaviour is in test/acceptance/xlsx-x-o08-drawings-insert.test.cjs (X-O08, already adapter-complete
 * and LIVE-measured); this file is the product-tool level, mirroring gh139-u2-move-chart-release.test.cjs. The READ
 * mirror is read_drawings's own parent/kind/name/pos/width/height/stroke-width fields (GH #139 U1 matrix rows D1-D6,
 * docs/GH139-CHART-DRAWING-SYMMETRY-MATRIX.md); the LIVE proof is
 * test/live/release-gh139-u3-insert-drawing-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEElEQVR4nGM4o6QERwzEcQDE4xEBYJBBnwAAAABJRU5ErkJggg==';
const book = (extra = []) => [
  { name: 'D', cells: { A1: { v: 'keep' } }, drawings: extra },
  { name: 'O', cells: { A1: { v: 'other-keep' } }, drawings: [] },
];
const shape = (over = {}) => ({ intent: 'insert_shape', id: 's', sheet: 'D', name: 'S1', geometry: 'rect', at: { col: 2, row: 1 }, size: { width: 72, height: 36 }, fill: '#FF0000', ...over });
const image = (over = {}) => ({ intent: 'insert_image', id: 's', sheet: 'D', name: 'I1', at: { col: 2, row: 1 }, size: { width: 72, height: 36 }, source: PNG, ...over });
const statuses = r => (r.operations || []).map(o => o.status);
const drawingsOf = (store, s) => store.durable.find(x => x.name === s).drawings;

test('GH #139 U3: insert_shape / insert_image are released intents (allowlist + accepted set + schema + constraints), not a sole step, the release gate passes', () => {
  for (const intent of ['insert_shape', 'insert_image']) {
    assert.ok(W.XLSX_WRITE_ALLOWLIST.includes(intent), intent);
    assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes(intent), intent);
    assert.equal(W.SOLE_STEP.includes(intent), false, intent);
  }
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].sort());
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].filter(i => i !== 'insert_shape' && i !== 'insert_image'), docxAllowlist: [] }).ok, false, 'insert_shape/insert_image missing');
});

test('GH #139 U3: the schema + CONSTRAINTS refuse a malformed step before any dispatch', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(shape()).ok, true);
  assert.equal(v(image()).ok, true);
  const bad = [
    [shape({ geometry: 'oval' }), 'geometry outside rect'],
    [shape({ fill: undefined }), 'fill missing'],
    [shape({ fill: 'red' }), 'fill not a hex colour'],
    [shape({ at: { col: -1, row: 1 } }), 'negative col'],
    [shape({ at: { col: 1.5, row: 1 } }), 'non-integer col'],
    [shape({ size: { width: 0, height: 36 } }), 'width shape-level 0 is schema-valid, refused by the family plan()'],
    [shape({ stroke: { width: 1 } }), 'stroke missing color'],
    [shape({ stroke: { width: 1, color: 'blue' } }), 'stroke.color not a hex colour'],
    [{ intent: 'insert_shape', id: 's', sheet: 'D', name: 'S1', geometry: 'rect', at: { col: 1, row: 1 }, size: { width: 1, height: 1 }, fill: '#FF0000', source: PNG }, 'insert_shape does not support source'],
    [image({ source: 'http://example.com/x.png' }), 'an http(s) URL is refused by the schema pattern'],
    [image({ source: 'data:image/jpeg;base64,AAAA' }), 'a non-PNG mime type is refused'],
    [image({ source: undefined }), 'source missing'],
    [shape({ sheet: '' }), 'an empty sheet name'],
    [shape({ name: '' }), 'an empty drawing name'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('GH #139 U3: insert_shape APPLIED at the requested cell, retry SATISFIED 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const first = await h.tool.handler(req([shape()]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  assert.equal(drawingsOf(h.store, 'D').length, 1);
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([shape()]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('GH #139 U3: insert_image APPLIED at the requested cell, retry SATISFIED 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const first = await h.tool.handler(req([image()]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  assert.equal(drawingsOf(h.store, 'D').length, 1);
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([image()]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('GH #139 U3: insert_shape with stroke APPLIED, width verified, retry SATISFIED 0 dispatch; a PROVEN fillLineColor read carries NO semanticsNotProven key (GH #139 D4/D6)', async () => {
  const h = harness({ store: createStore(book()) });
  const first = await h.tool.handler(req([shape({ stroke: { width: 2, color: '#000000' } })]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  assert.equal('semanticsNotProven' in first.operations[0], false);
  const retry = await h.tool.handler(req([shape({ stroke: { width: 2, color: '#000000' } })]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
});

test('GH #139 U3: on a pre-/14 accepted runtime, insert_shape with stroke names both colour axes, without stroke only fill-colour', async () => {
  const h = harness({ store: createStore(book()), script: { colourUnobservable: true } });
  const first = await h.tool.handler(req([shape({ stroke: { width: 2, color: '#000000' } })]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  assert.deepEqual(first.operations[0].semanticsNotProven, ['shape-fill-colour', 'shape-stroke-colour']);
  const h2 = harness({ store: createStore(book()), script: { colourUnobservable: true } });
  const r = await h2.tool.handler(req([shape()]));
  assert.deepEqual(statuses(r), ['APPLIED']);
  assert.deepEqual(r.operations[0].semanticsNotProven, ['shape-fill-colour']);
});

test('GH #139 U3: a name collision with another identity is refused named xlsx-drawing-insert-exists-with-another-identity, nothing dispatched', async () => {
  const h = harness({ store: createStore(book([{ kind: 'image', name: 'S1', pos: { fromCol: 1, colOffset: 0, fromRow: 1, rowOffset: 0 }, width: 914400, height: 457200 }])) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([shape()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-drawing-insert-exists-with-another-identity/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('GH #139 U3: an unrelated drawing on the sheet is untouched (L3 preservation)', async () => {
  const other = { kind: 'shape', name: 'CONTROL', pos: { fromCol: 0, colOffset: 0, fromRow: 0, rowOffset: 0 }, width: 914400, height: 457200, text: '', strokeWidth: 0 };
  const h = harness({ store: createStore(book([other])) });
  const r = await h.tool.handler(req([shape()]));
  assert.deepEqual(statuses(r), ['APPLIED'], J(r).slice(0, 400));
  const control = drawingsOf(h.store, 'D').find(d => d.name === 'CONTROL');
  assert.deepEqual(control, other);
});

test('GH #139 U3: the large-book path is the NAMED refusal, NOT the tiered observation -- insert_shape/insert_image are NOT in TIERED_INTENTS', () => {
  const sc = s => liveScope({ id: 't', steps: [s] });
  assert.equal(sc(shape()).tiered, undefined, 'not in TIERED_INTENTS -- the key is absent, not false');
  assert.equal(sc(image()).tiered, undefined, 'the sibling drawing intent stays not-tiered too, unaffected');
});

test('GH #139 U3: the source size ceiling -- 1,400,000 characters pass the schema, one more is REFUSED_SHAPE before any dispatch', () => {
  const prefix = 'data:image/png;base64,';
  const at = n => image({ source: prefix + 'A'.repeat(n - prefix.length) });
  assert.equal(W.validateWriteRequest(req([at(1400000)])).ok, true);
  const over = W.validateWriteRequest(req([at(1400001)]));
  assert.equal(over.ok, false);
  assert.equal(over.code, 'REFUSED_SHAPE', JSON.stringify(over).slice(0, 300));
});
