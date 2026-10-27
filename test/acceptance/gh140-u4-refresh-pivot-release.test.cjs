'use strict';

/**
 * GH #140 U4: refresh_pivot on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The family-level behaviour is in
 * xlsx-m7-p3-refresh-pivot.test.cjs (plan / classify / lower / verify, 30+ tests); this file is the product-tool level, mirroring
 * gh138-u1-defined-name-release.test.cjs: the release wiring, the strict schema, the family's named refusals reached THROUGH the product tool,
 * and the READ mirror (read_pivots, GH #140 U2) showing the NEW source after the refresh. The LIVE proof is
 * test/live/release-gh140-u4-refresh-pivot-product.live.cjs.
 *
 * The pivot to refresh is SEEDED through the canonical internal create_pivot (test setup only -- create_pivot is not part of this release).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const rt = require('../../src/server/read-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { fakeBroker } = require('../fakes/read-editors.cjs');

const J = x => JSON.stringify(x);
const DATA = [['reg', 'prod', 'amt'], ['N', 'x', 1], ['S', 'y', 2], ['N', 'y', 3], ['S', 'x', 4], ['N', 'x', 5], ['S', 'y', 6]];
const MORE = [['reg', 'prod', 'amt'], ['N', 'x', 10], ['S', 'y', 20], ['W', 'z', 30], ['N', 'z', 40], ['S', 'x', 50], ['W', 'y', 60]];
const OTHERHEAD = [['area', 'prod', 'amt'], ['N', 'x', 1], ['S', 'y', 2], ['N', 'y', 3]];
const cellsOf = rows => Object.fromEntries(rows.flatMap((r, i) => r.map((v, j) => [`${'ABC'[j]}${i + 1}`, { v }])));
const book = () => [
  { name: 'Src', cells: cellsOf(DATA) },
  { name: 'Src2', cells: cellsOf(MORE) },
  { name: 'Bad', cells: cellsOf(OTHERHEAD) },
  { name: 'Host', cells: { A1: { v: 'keep' } } },
];
const CREATE = { intent: 'create_pivot', name: 'PE', sheet: 'Src', range: 'A1:C7', rows: 'reg', columns: 'prod', sum: 'amt', pivotSheet: 'Host', at: 'H3' };

/** test SETUP only: the canonical internal create_pivot builds the pivot the fake's reader serves */
async function seed(store, over = {}) {
  const t = { id: 'seed', steps: [{ ...CREATE, ...over }] };
  let n = 0;
  const r = await withSession({ driver: createFakeXlsxDriver(store, {}), fileId: 'seed', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } },
    xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  if (!r.ok) throw new Error(`pivot seed failed: ${r.outcome}`);
  return store;
}
const step = (over = {}) => ({ intent: 'refresh_pivot', id: 'r', pivot: 'PE', sourceSheet: 'Src2', sourceRange: 'A1:C7', ...over });
const statuses = r => (r.operations || []).map(o => o.status);
const host = store => store.durable.find(s => s.name === 'Host');
const pivotOf = (store, name = 'PE') => host(store).objects.pivots.find(p => p.name === name);

test('GH #140 U4: refresh_pivot is released (allowlist + accepted set + schema + constraints), a normal (not sole) step, and the release gate passes', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('refresh_pivot'));
  assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes('refresh_pivot'));
  assert.ok(!W.SOLE_STEP.includes('refresh_pivot'));
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].sort());
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].filter(i => i !== 'refresh_pivot'), docxAllowlist: [] }).ok, false, 'refresh_pivot missing');
});

test('GH #140 U4: the step schema is EXACTLY the family FIELDS (pivot, sourceSheet, sourceRange) -- no layout / function key exists', () => {
  const fields = require('../../src/adapters/xlsx/families/pivot-refresh.cjs');
  const props = Object.keys(W.STEP_SCHEMAS.refresh_pivot.properties).filter(k => k !== 'intent' && k !== 'id').sort();
  assert.deepEqual(props, ['pivot', 'sourceRange', 'sourceSheet']);
  assert.deepEqual([...W.STEP_SCHEMAS.refresh_pivot.required].sort(), ['intent', 'pivot', 'sourceRange', 'sourceSheet']);
  assert.equal(W.STEP_SCHEMAS.refresh_pivot.additionalProperties, false);
  assert.deepEqual(fields.INTENTS, ['refresh_pivot']);
});

test('GH #140 U4: the schema + the constraints refuse a malformed step BEFORE any dispatch, named', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(step()).ok, true, J(v(step())));
  const bad = [
    [step({ rows: 'reg' }), 'a layout key (the family never writes a layout)'],
    [step({ sum: 'amt' }), 'an aggregation key'],
    [step({ colour: 'red' }), 'an unknown key'],
    [step({ pivot: undefined }), 'pivot missing'],
    [step({ sourceSheet: undefined }), 'sourceSheet missing'],
    [step({ sourceRange: undefined }), 'sourceRange missing'],
    [step({ pivot: '1bad' }), 'a pivot name starting with a digit'],
    [step({ pivot: 'has space' }), 'a pivot name with a space'],
    [step({ sourceSheet: "O'Brien" }), 'an apostrophe sheet name (the public setter builds an UNQUOTED reference)'],
    [step({ sourceSheet: 'My Sheet' }), 'a sheet name that needs quoting'],
    [step({ sourceRange: 'A1' }), 'a single cell'],
    [step({ sourceRange: 'A1:C1' }), 'a header row and no data row'],
    [step({ sourceRange: 'A1:B7' }), 'fewer than three columns'],
    [step({ sourceRange: 'a1:c7' }), 'a lower-case range'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('GH #140 U4: APPLIED on the product tool -- the pivot is re-bound to the NAMED source and the table follows; the retry is SATISFIED with 0 dispatch', async () => {
  const store = await seed(createStore(book()));
  const before = J(pivotOf(store));
  const h = harness({ store });
  const r = await h.tool.handler(req([step()]));
  assert.equal(r.ok, true, J(r).slice(0, 600));
  assert.deepEqual(statuses(r), ['APPLIED']);
  assert.notEqual(J(pivotOf(store)), before, 'the pivot changed');
  assert.match(J(pivotOf(store)), /Src2/, 'the data reference names the new source');
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([step()]));
  assert.deepEqual(statuses(again), ['SATISFIED'], J(again).slice(0, 500));
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('GH #140 U4: the family refusals reach the product caller NAMED, and nothing is written (missing pivot / wrong header / quoted-name sheet is refused by the schema above)', async () => {
  const store = await seed(createStore(book()));
  const h = harness({ store });
  const snap = () => J(store.durable);
  const before = snap();
  const d0 = h.calls().dispatch;
  const missing = await h.tool.handler(req([step({ pivot: 'NoSuchPivot' })]));
  assert.equal(missing.ok, false);
  assert.match(J(missing), /xlsx-pivot-not-found/, J(missing).slice(0, 500));
  const header = await h.tool.handler(req([step({ sourceSheet: 'Bad' })]));   // a header row that is not exactly the pivot's field list
  assert.equal(header.ok, false);
  assert.match(J(header), /header/i, J(header).slice(0, 500));
  const nosheet = await h.tool.handler(req([step({ sourceSheet: 'Nope' })]));
  assert.equal(nosheet.ok, false);
  assert.match(J(nosheet), /xlsx-sheet-not-found/, J(nosheet).slice(0, 500));
  assert.equal(h.calls().dispatch, d0, 'every refusal is before the write');
  assert.equal(snap(), before, 'nothing changed');
});

test('GH #140 U4: a pivot that is ambiguous (the same name on two sheets) is refused, named; a SECOND pivot is kept unchanged by a refresh of the first', async () => {
  const store = await seed(createStore(book()));
  await seed(store, { name: 'PF', pivotSheet: 'Host', at: 'N3' });
  const other0 = J(pivotOf(store, 'PF'));
  const h = harness({ store });
  const r = await h.tool.handler(req([step()]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.equal(J(pivotOf(store, 'PF')), other0, 'the other pivot is untouched');
  // the same name on a second sheet -> ambiguous
  store.durable.find(s => s.name === 'Src').objects = { charts: [], names: [], pivots: [JSON.parse(J(pivotOf(store, 'PE')))] };
  const amb = await h.tool.handler(req([step({ sourceSheet: 'Src2' })]));
  assert.equal(amb.ok, false);
  assert.match(J(amb), /xlsx-pivot-name-ambiguous/, J(amb).slice(0, 500));
});

test('GH #140 U4: the READ mirror -- read_pivots (GH #140 U2) shows the NEW source after the refresh, and the OLD source before it', async () => {
  const store = await seed(createStore(book()));
  const b = fakeBroker('xlsx');
  const read = rt.createReadTool('xlsx', { configSource: b.source, loadPlaywright: () => {}, createDriver: () => createFakeXlsxDriver(store) });
  const sourceOf = async () => {
    const r = await read.handler({ fileId: '1', operation: 'read_pivots', sheet: 'Host' });
    const text = J(r);
    return { r, text };
  };
  const before = await sourceOf();
  assert.match(before.text, /Src/, before.text.slice(0, 400));
  assert.doesNotMatch(before.text, /Src2/, 'before the refresh the source is the old one');
  const h = harness({ store });
  assert.deepEqual(statuses(await h.tool.handler(req([step()]))), ['APPLIED']);
  const after = await sourceOf();
  assert.match(after.text, /Src2/, after.text.slice(0, 400));
});

test('GH #140 U4: the step is scoped for the live observation like every pivot step; two refreshes in ONE task are composed by the canonical planner (same source sheet = overlapping writes, refused named)', () => {
  const sc = liveScope({ id: 't', steps: [step()] });
  assert.deepEqual(sc.windows, [{ sheet: 'Src2', range: 'A1:C7' }]);
  assert.deepEqual(sc.objectSheets, ['*'], 'the pivot identity of EVERY sheet is observed (the family needs it to find the pivot by name)');
  const two = W.validateWriteRequest(req([step(), step({ id: 'r2', pivot: 'PF' })]));
  assert.equal(two.ok, false);
  assert.equal(two.code, 'REFUSED_COMPOSITION');
  assert.match(two.reason, /overlapping-writes-refused/);
  const apart = W.validateWriteRequest(req([step(), step({ id: 'r2', pivot: 'PF', sourceSheet: 'Src3' })]));
  assert.equal(apart.ok, true, J(apart));
});
