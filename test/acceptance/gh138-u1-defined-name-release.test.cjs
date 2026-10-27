'use strict';

/**
 * GH #138 U1: set_defined_name / delete_defined_name on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The family-level behaviour is in
 * xlsx-w6-names.test.cjs (+ the M7 R16 scope tests); this file is the product-tool level, mirroring l11-filter-range-release.test.cjs. The READ mirror is
 * read_defined_names (GH #134 P2); the LIVE proof is test/live/release-gh138-u1-defined-name-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const store0 = () => {
  const s = createStore([{ name: 'Data', cells: { A1: { v: 7 }, B1: { v: 14, f: '=Used*2' } } }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
  s.names = [{ name: 'Used', ref: 'Data!$A$1' }, { name: 'Free', ref: 'Other!$A$1' }];
  return s;
};
const set = (name, refersTo, extra = {}) => ({ intent: 'set_defined_name', id: 's', name, refersTo, ...extra });
const del = (name, extra = {}) => ({ intent: 'delete_defined_name', id: 's', name, ...extra });
const statuses = r => (r.operations || []).map(o => o.status);

test('GH #138 U1: both intents are released (allowlist + accepted set + schema + constraints), normal (not sole) steps, and the release gate passes', () => {
  for (const i of ['set_defined_name', 'delete_defined_name']) {
    assert.ok(W.XLSX_WRITE_ALLOWLIST.includes(i), i);
    assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes(i), i);
    assert.ok(!W.SOLE_STEP.includes(i), i);
  }
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].sort());
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].filter(i => i !== 'set_defined_name'), docxAllowlist: [] }).ok, false, 'set_defined_name missing');
});

test('GH #138 U1: the schema + the family plan refuse a malformed step before any dispatch', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(set('Rate', '=Data!$A$1')).ok, true);
  assert.equal(v(set('Rate', "='My Sheet'!$A$1:$B$2", { scope: 'Data' })).ok, true);
  assert.equal(v(del('Rate')).ok, true);
  assert.equal(v(del('Rate', { scope: 'Data' })).ok, true);
  const bad = [
    [set('Rate', '=A1'), 'a relative reference'],
    [set('Rate', '=0.27'), 'a constant'],
    [set('Rate', 'Data!$A$1'), 'no ='],
    [set('A1', '=Data!$A$1'), 'a cell-reference-shaped name'],
    [set('Print_Area', '=Data!$A$1', { scope: 'Data' }), 'a reserved name'],
    [set('Print_Titles', '=Data!$1:$1', { scope: 'Data' }), 'a reserved name'],
    [set('1abc', '=Data!$A$1'), 'a name starting with a digit'],
    [set('Rate', '=Data!$A$1', { scope: '' }), 'an empty scope'],
    [set('Rate', '=Data!$A$1', { colour: 'red' }), 'an unknown key'],
    [{ intent: 'set_defined_name', id: 's', name: 'Rate' }, 'refersTo missing'],
    [del('Rate', { refersTo: '=Data!$A$1' }), 'a delete with a refersTo'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('GH #138 U1: set APPLIED (workbook level), the retry SATISFIED 0 dispatch; a re-point of an existing name refused, nothing written', async () => {
  const h = harness({ store: store0() });
  const r = await h.tool.handler(req([set('Total', '=Data!$A$1:$A$3')]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.deepEqual(statuses(r), ['APPLIED']);
  assert.ok(h.store.names.some(x => x.name === 'Total' && x.ref === 'Data!$A$1:$A$3'), J(h.store.names));
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([set('Total', '=Data!$A$1:$A$3')]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
  const rep = await h.tool.handler(req([set('Free', '=Data!$A$1')]));
  assert.equal(rep.ok, false, J(rep).slice(0, 300));
  assert.equal(h.calls().dispatch, d0, 'a re-point is refused before the write');
  assert.ok(h.store.names.some(x => x.name === 'Free' && x.ref === 'Other!$A$1'), 'the existing name kept its reference');
});

test('GH #138 U1: delete -- an unused name APPLIED then SATISFIED; a name a formula uses is refused, kept', async () => {
  const h = harness({ store: store0() });
  const r = await h.tool.handler(req([del('Free')]));
  assert.deepEqual(statuses(r), ['APPLIED'], J(r).slice(0, 400));
  assert.ok(!h.store.names.some(x => x.name === 'Free'));
  assert.deepEqual(statuses(await h.tool.handler(req([del('Free')]))), ['SATISFIED']);
  const used = await h.tool.handler(req([del('Used')]));
  assert.equal(used.ok, false, J(used).slice(0, 300));
  assert.ok(h.store.names.some(x => x.name === 'Used'), 'the name in use is kept');
});

test('GH #138 U1: a sheet-level name (scope) APPLIED on its sheet; the large-book path is the tiered observation, not a refusal', async () => {
  const h = harness({ store: store0() });
  const r = await h.tool.handler(req([set('Local', '=Other!$A$1', { scope: 'Other' })]));
  assert.deepEqual(statuses(r), ['APPLIED'], J(r).slice(0, 400));
  const other = h.store.durable.find(x => x.name === 'Other');
  assert.ok(((other.objects || {}).names || []).some(n => n.name === 'Local'), J(other.objects));
  const sc = s => liveScope({ id: 't', steps: [s] });
  assert.equal(sc(set('X', '=Data!$A$1')).tiered, true);
  assert.equal(sc(del('X')).tiered, true);
});

// set_print_titles / clear_print_titles were the U1-era neighbours here (not yet released); GH #138 U2 released
// them separately (gh138-u2-print-titles-release.test.cjs) -- create_pivot (still not released) is the neighbour now.
test('GH #138 U1: the release does not widen other neighbours', () => {
  assert.equal(W.validateWriteRequest(req([{ intent: 'create_pivot', id: 's', sheet: 'Data' }])).ok, false);
});
