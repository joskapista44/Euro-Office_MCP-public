'use strict';

/**
 * GH #138 U2: set_print_titles / clear_print_titles on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The family-level behaviour is in
 * xlsx-m7-f-printnames.test.cjs (+ xlsx-m7-p57); this file is the product-tool level, mirroring gh138-u1-defined-name-release.test.cjs. The READ mirror is
 * read_defined_names's own `titles` projection (GH #134 P2, defined-names-read.cjs); the LIVE proof is test/live/release-gh138-u2-print-titles-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const store0 = () => createStore([{ name: 'Data', cells: { A1: { v: 7 } } }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
const set = (sheet, extra = {}) => ({ intent: 'set_print_titles', id: 's', sheet, ...extra });
const clear = (sheet, extra = {}) => ({ intent: 'clear_print_titles', id: 's', sheet, ...extra });
const statuses = r => (r.operations || []).map(o => o.status);
const titlesOf = (store, sheet) => ((store.durable.find(s => s.name === sheet).objects || {}).names || []).find(n => n.name === 'Print_Titles');

test('GH #138 U2: both intents are released (allowlist + accepted set + schema + constraints), normal (not sole) steps, and the release gate passes', () => {
  for (const i of ['set_print_titles', 'clear_print_titles']) {
    assert.ok(W.XLSX_WRITE_ALLOWLIST.includes(i), i);
    assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes(i), i);
    assert.ok(!W.SOLE_STEP.includes(i), i);
  }
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].sort());
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].filter(i => i !== 'set_print_titles'), docxAllowlist: [] }).ok, false, 'set_print_titles missing');
});

test('GH #138 U2: the schema + the family plan refuse a malformed step before any dispatch', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(set('Data', { rows: '1:1' })).ok, true);
  assert.equal(v(set('Data', { columns: 'A:B' })).ok, true);
  assert.equal(v(set('Data', { rows: '1:1', columns: 'A:B' })).ok, true);
  assert.equal(v(clear('Data')).ok, true);
  const bad = [
    [set('Data'), 'neither rows nor columns'],
    [set('Data', { rows: '1' }), 'rows missing its second part'],
    [set('Data', { columns: 'A' }), 'columns missing its second part'],
    [set('Data', { rows: '2:1' }), 'rows reversed'],
    [set('Data', { columns: 'B:A' }), 'columns reversed'],
    [set('Data', { rows: '1:1', range: 'A1:B2' }), 'range does not belong to set_print_titles'],
    [clear('Data', { rows: '1:1' }), 'clear does not take rows'],
    [set('', { rows: '1:1' }), 'an empty sheet name'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('GH #138 U2: set (rows, columns, both) APPLIED on its sheet, the retry SATISFIED 0 dispatch', async () => {
  const h = harness({ store: store0() });
  const rowsOnly = await h.tool.handler(req([set('Data', { rows: '1:1' })]));
  assert.deepEqual(statuses(rowsOnly), ['APPLIED'], J(rowsOnly).slice(0, 400));
  assert.deepEqual(titlesOf(h.store, 'Data'), { name: 'Print_Titles', refersTo: '=Data!$1:$1' });
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([set('Data', { rows: '1:1' })]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
  // columns-first storage order when both are given (M7 X-P03, names.cjs planPrint)
  const both = await h.tool.handler(req([set('Other', { rows: '2:3', columns: 'A:B' })]));
  assert.deepEqual(statuses(both), ['APPLIED'], J(both).slice(0, 400));
  assert.deepEqual(titlesOf(h.store, 'Other'), { name: 'Print_Titles', refersTo: '=Other!$A:$B,Other!$2:$3' });
});

test('GH #138 U2: a re-set with a DIFFERENT value on the SAME sheet is refused named xlsx-name-exists-with-another-reference, 0 writes -- zero\'s 1st LIVE round: set_print_titles has NO `replaces` field (unlike set_print_area), so there is no silent re-point; a titles CHANGE is clear_print_titles + set_print_titles', async () => {
  const h = harness({ store: store0() });
  await h.tool.handler(req([set('Data', { rows: '1:1' })]));
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([set('Data', { columns: 'A:B' })]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-name-exists-with-another-reference/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
  assert.deepEqual(titlesOf(h.store, 'Data'), { name: 'Print_Titles', refersTo: '=Data!$1:$1' }, 'the ORIGINAL titles survive untouched');
  // the SAME value is a genuine retry (SATISFIED), not a re-point -- canon(ref) equality, unaffected by this change
  assert.deepEqual(statuses(await h.tool.handler(req([set('Data', { rows: '1:1' })]))), ['SATISFIED']);
});

test('GH #138 U2: clear -- the Print_Titles name is removed then APPLIED/SATISFIED; a sheet with no titles clears cleanly to SATISFIED', async () => {
  const h = harness({ store: store0() });
  await h.tool.handler(req([set('Data', { rows: '1:1' })]));
  const r = await h.tool.handler(req([clear('Data')]));
  assert.deepEqual(statuses(r), ['APPLIED'], J(r).slice(0, 400));
  assert.equal(titlesOf(h.store, 'Data'), undefined);
  assert.deepEqual(statuses(await h.tool.handler(req([clear('Data')]))), ['SATISFIED']);
});

test('GH #138 U2: a nonexistent sheet is refused named xlsx-sheet-not-found, nothing dispatched -- the FAIL_CLOSED evidence this release was missing (census.json, clear_print_titles, M11.0 checkpoint)', async () => {
  const h = harness({ store: store0() });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([clear('NoSuchSheet')]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-sheet-not-found/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
  const r2 = await h.tool.handler(req([set('NoSuchSheet', { rows: '1:1' })]));
  assert.equal(r2.ok, false, J(r2).slice(0, 300));
  assert.match(J(r2), /xlsx-sheet-not-found/);
});

test('GH #138 U2: the release does not widen the reserved print names or the other scope rule (names.cjs, unchanged)', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v({ intent: 'set_defined_name', id: 's', name: 'Print_Titles', refersTo: '=Data!$1:$1', scope: 'Data' }).ok, false, 'Print_Titles stays reserved for set_defined_name');
});

test('GH #138 U2: the large-book path is the NAMED refusal, NOT the tiered observation (card\'s own boundary) -- UNLIKE set_defined_name/delete_defined_name', () => {
  const sc = s => liveScope({ id: 't', steps: [s] });
  assert.equal(sc(set('Data', { rows: '1:1' })).tiered, undefined, 'not in TIERED_INTENTS -- the key is absent, not false');
  assert.equal(sc(clear('Data')).tiered, undefined);
  assert.equal(sc({ intent: 'set_defined_name', id: 's', name: 'X', refersTo: '=Data!$A$1' }).tiered, true, 'the sibling intent stays tiered, unaffected');
});
