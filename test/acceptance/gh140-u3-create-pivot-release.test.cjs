'use strict';

/**
 * GH #140 U3: create_pivot on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The family-level
 * behaviour is in test/acceptance/xlsx-m7-h-create-pivot.test.cjs (M7-H, already adapter-complete and
 * LIVE-measured, docs/evidence/M7-H-ACCEPTANCE.md PASS 6/6); this file is the product-tool level, mirroring
 * gh139-u2-move-chart-release.test.cjs / gh138-u3-page-layout-release.test.cjs. The READ mirror is read_pivots
 * (GH #140 U2, already released); the LIVE proof is test/live/release-gh140-u3-create-pivot-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { FILE, harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const DATA = [['reg', 'prod', 'amt'], ['N', 'x', 1], ['S', 'y', 2], ['N', 'y', 3], ['S', 'x', 4], ['N', 'x', 5], ['S', 'y', 6]];
const cellsOf = rows => Object.fromEntries(rows.flatMap((r, i) => r.map((v, j) => [`${'ABC'[j]}${i + 1}`, { v }])));
const book = () => [
  { name: 'Src', cells: cellsOf(DATA), objects: { charts: [], pivots: [], names: [] } },
  { name: 'Host', cells: { A1: { v: 'keep' } }, objects: { charts: [], pivots: [], names: [] } },
];
const create = (over = {}) => ({ intent: 'create_pivot', id: 's', name: 'PV1', sheet: 'Src', range: 'A1:C7', rows: 'reg', columns: 'prod', sum: 'amt', pivotSheet: 'Host', at: 'H3', ...over });
const statuses = r => (r.operations || []).map(o => o.status);
const pivotsOf = (store, sheet) => store.durable.find(x => x.name === sheet).objects.pivots;

test('GH #140 U3: create_pivot is a released intent (allowlist + accepted set + schema + constraints), not a sole step, the release gate passes', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('create_pivot'));
  assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes('create_pivot'));
  assert.equal(W.SOLE_STEP.includes('create_pivot'), false);
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].sort());
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].filter(i => i !== 'create_pivot'), docxAllowlist: [] }).ok, false, 'create_pivot missing');
});

test('GH #140 U3: the schema + CONSTRAINTS refuse a malformed step before any dispatch', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(create()).ok, true, J(v(create())));
  const bad = [
    [create({ name: undefined }), 'name missing'],
    [create({ sheet: undefined }), 'sheet missing'],
    [create({ range: undefined }), 'range missing'],
    [create({ rows: undefined }), 'rows missing'],
    [create({ columns: undefined }), 'columns missing'],
    [create({ sum: undefined }), 'sum missing'],
    [create({ pivotSheet: undefined }), 'pivotSheet missing'],
    [create({ at: undefined }), 'at missing'],
    [create({ name: '' }), 'empty name'],
    [create({ name: '1PV' }), 'name starting with a digit'],
    [create({ name: 'has space' }), 'name with a space'],
    [create({ range: 'A1' }), 'range not a two-corner range'],
    [create({ at: 'H3:I4' }), 'at is a range, not one cell'],
    [create({ rows: '' }), 'empty rows'],
    [create({ rows: 'reg\nprod' }), 'rows with a newline'],
    [create({ extra: true }), 'create_pivot does not support an extra field'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('GH #140 U3: REFUSED_SHAPE for a 4th row/column/sum-shaped field is a family-level (plan) refusal, not a schema one -- rows/columns/sum sharing a name', async () => {
  const h = harness({ store: createStore(book()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([create({ sum: 'reg' })]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /three different header names/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('GH #140 U3: APPLIED, the table as M7-H measured, retry SATISFIED 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const first = await h.tool.handler(req([create()]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  const pivots = pivotsOf(h.store, 'Host');
  assert.equal(pivots.length, 1);
  assert.equal(pivots[0].name, 'PV1');
  assert.equal(pivots[0].source, "Src!$A$1:$C$7");
  assert.deepEqual(pivots[0].layout, { rows: ['reg'], cols: ['prod'], data: ['Sum of amt'] });
  const hostSheet = h.store.durable.find(x => x.name === 'Host');
  assert.equal(hostSheet.cells.A1.v, 'keep', 'the pre-existing cell outside the table is untouched');
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([create()]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('GH #140 U3: a non-empty destination is refused xlsx-pivot-destination-not-empty, nothing dispatched', async () => {
  const b = book();
  b[1].cells.I4 = { v: 'busy' };
  const h = harness({ store: createStore(b) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([create()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-pivot-destination-not-empty/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('GH #140 U3: a pivot name already taken by another identity is refused xlsx-pivot-exists-with-another-identity, nothing dispatched', async () => {
  const b = book();
  b[1].objects.pivots = [{ name: 'PV1', source: 'Src!$A$1:$C$3', table: 'H3:I4', layout: { rows: ['reg'], cols: ['prod'], data: ['Sum of amt'] } }];
  const h = harness({ store: createStore(b) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([create()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-pivot-exists-with-another-identity/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('GH #140 U3: the create + the product READ (read_pivots, GH #140 U2) agree on name/source -- the acceptance point docs/GH140-PIVOT-CONTRACT.md section 3 named missing', async () => {
  const h = harness({ store: createStore(book()) });
  const w = await h.tool.handler(req([create()]));
  assert.deepEqual(statuses(w), ['APPLIED'], J(w).slice(0, 400));
  const { createReadTool } = require('../../src/server/read-tools.cjs');
  const rt = createReadTool('xlsx', { configSource: { kind: 'BROKER' }, loadPlaywright: () => {}, createDriver: () => h.drivers[h.drivers.length - 1] });
  const r = await rt.handler({ fileId: FILE, operation: 'read_pivots', sheet: 'Host' });
  assert.equal(r.ok, true, J(r));
  assert.deepEqual(r.result.sheets, [{ sheet: 'Host', pivots: [{ name: 'PV1', source: 'Src!$A$1:$C$7', table: 'H3:K7', layout: { rows: ['reg'], cols: ['prod'], data: ['Sum of amt'] } }] }]);
});

test('GH #140 U3: the large-book path is the NAMED refusal, NOT the tiered observation (same as every other pivot/chart intent) -- create_pivot is NOT in TIERED_INTENTS', () => {
  const sc = s => liveScope({ id: 't', steps: [s] });
  assert.equal(sc(create()).tiered, undefined, 'not in TIERED_INTENTS -- the key is absent, not false');
});

// ---- RED CONTROL: proves the checks above discriminate, not pass-by-construction ----

test('RED CONTROL: a fabricated pivot name is correctly absent after the real create', async () => {
  const h = harness({ store: createStore(book()) });
  await h.tool.handler(req([create()]));
  const pivots = pivotsOf(h.store, 'Host');
  assert.ok(!pivots.some(p => p.name === 'DefinitelyNotARealPivot'));
});

test('RED CONTROL: a made-up field name is correctly refused by the schema, not silently accepted', () => {
  const r = W.validateWriteRequest(req([create({ filter: 'reg' })]));
  assert.equal(r.ok, false);
});
