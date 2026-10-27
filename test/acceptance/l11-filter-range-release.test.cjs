'use strict';

/**
 * L-11/2 (owner decision "A", the owner, owner channel, the answer to 36637; assigned by a maintainer, card
 * an internal branch): filter_range on the PRODUCT surface (the WRITE tool with the fake
 * XLSX editor) for SMALL workbooks. The LARGE-book refusal is UNCHANGED (the owner rule, owner channel): filter_range
 * is NOT in TIERED_INTENTS and has no bounded variant, so above the read bound it stays refused
 * (IMPLEMENTATION_LIMIT:family-not-migrated-to-the-tiered-observation, pinned in m7-s32-tiered.test.cjs) -- this file
 * re-pins the SCOPE that decides it (liveScope), it does not loosen it.
 * The family-level behaviour (criteria, the X-L05 manual-hidden-row gate) is in xlsx-w6-filter.test.cjs and
 * xlsx-m7-d-filter.test.cjs; this file is the product-tool level, mirroring l06b-move-range-release.test.cjs.
 * The LIVE proof is test/live/release-l11-filter-range-product.live.cjs (zero runs it, per the kiosztas).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const book = () => [
  { name: 'F', cells: { A1: { v: 'id' }, B1: { v: 'tag' }, A2: { v: 1 }, B2: { v: 'x' }, A3: { v: 2 }, B3: { v: 'y' }, A4: { v: 3 }, B4: { v: 'x' } } },
  { name: 'O', cells: { A1: { v: 'keep' } } },
];
// a key given as `undefined` is REMOVED (the wire JSON has no undefined), so `value: undefined` means "no value key"
const flt = (extra = {}) => Object.fromEntries(Object.entries({ intent: 'filter_range', id: 's', sheet: 'F', range: 'A1:B4', field: 2, value: 'x', ...extra }).filter(([, v]) => v !== undefined));
const validate = step => W.validateWriteRequest(req([step]));
const statuses = r => (r.operations || []).map(o => o.status);

test('L-11/2: filter_range is a released intent (allowlist + accepted set + schema + constraints), a normal (not sole) step', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('filter_range'));
  assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes('filter_range'));
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].sort());
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.ok(!W.SOLE_STEP.includes('filter_range'));
  assert.equal(validate(flt()).ok, true);
  assert.equal(validate(flt({ value: undefined, values: ['x', 'y'] })).ok, true);
  assert.equal(validate(flt({ value: undefined, and: ['>1', '<9'] })).ok, true);
  assert.equal(validate(flt({ value: 5 })).ok, true);
});

test('L-11/2: the writeReleaseGate passes with the owner-named filter_range and fails when an UNNAMED intent joins / filter_range goes missing / it is duplicated', () => {
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
  const accepted = [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST];
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...accepted, 'clear_print_titles'], docxAllowlist: [] }).ok, false, 'an intent that was not named by an owner decision');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: accepted.filter(i => i !== 'filter_range'), docxAllowlist: [] }).ok, false, 'filter_range missing');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...accepted, 'filter_range'], docxAllowlist: [] }).ok, false, 'a duplicate');
});

test('L-11/2: the schema + constraints refuse a malformed step before any dispatch', () => {
  const bad = [
    [flt({ field: 0 }), 'field below 1'],
    [flt({ field: 3 }), 'field outside the range columns (A:B has 2)'],
    [flt({ field: 1.5 }), 'a fractional field'],
    [flt({ field: undefined }), 'field missing'],
    [flt({ value: undefined }), 'no criteria at all'],
    [flt({ values: ['x', 'y'] }), 'value AND values together'],
    [flt({ value: '>5' }), 'an operator inside value'],
    [flt({ value: '=x' }), 'a formula-like value'],
    [flt({ value: '' }), 'an empty value'],
    [flt({ value: undefined, values: ['x'] }), 'values with one element'],
    [flt({ value: undefined, values: ['x', 'x'] }), 'two identical values'],
    [flt({ value: undefined, values: ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k'] }), 'values with eleven elements (GH #301 D2 caps the list at 10)'],
    [flt({ value: undefined, and: ['>=5', '<9'] }), 'an unmeasured operator'],
    [flt({ value: undefined, and: ['>5'] }), 'and with one comparison'],
    [flt({ range: 'A1:B1' }), 'a header-only range'],
    [flt({ range: 'A1' }), 'a single cell'],
    [flt({ range: 'a1:b4' }), 'a lower-case range'],
    [flt({ sheet: '' }), 'an empty sheet'],
    [flt({ colour: 'red' }), 'an unknown key'],
  ];
  for (const [s, why] of bad) assert.equal(validate(s).ok, false, why);
});

test('L-11/2: small book -- APPLIED, the filter reads back on that field, the other sheet untouched; the retry is SATISFIED 0/0', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([flt()]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.deepEqual(statuses(r), ['APPLIED']);
  const F = h.store.durable.find(x => x.name === 'F');
  assert.equal(F.autoFilter.range, 'A1:B4');
  assert.ok(F.autoFilter.filters.some(f => f.field === 2 && f.on !== false), J(F.autoFilter));
  assert.deepEqual(h.store.durable.find(x => x.name === 'O').cells, { A1: { v: 'keep' } });
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([flt()]));
  assert.equal(again.ok, true, J(again).slice(0, 300));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('L-11/2: NEG -- the X-L05 gate holds on the product surface: a manually hidden data row, no active filter -> refused, 0 dispatch', async () => {
  const b = book();
  b[0].layout = { rows: { 3: { hidden: true } } };
  const h = harness({ store: createStore(b) });
  const r = await h.tool.handler(req([flt()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.equal(h.calls().dispatch, 0);
  assert.match(J(r), /xlsx-filter-would-unhide-manual-rows/);
});

test('L-11/2: NEG -- the LARGE-book rule is unchanged: filter_range is NOT tiered, has no bounded variant (clear_filter, the POZ control, is tiered)', () => {
  const sc = s => liveScope({ id: 't', steps: [s] });
  assert.notEqual(sc({ intent: 'filter_range', sheet: 'S', range: 'A1:B2', field: 1, value: 'x' }).tiered, true, 'filter_range must stay outside the tiered observation (owner rule, owner channel)');
  assert.equal(sc({ intent: 'clear_filter', sheet: 'S', range: 'A1:B2' }).tiered, true, 'the control: a tiered intent reads as tiered');
});

test('L-11/2: the release does not widen anything else -- the sole-step / neighbour intents keep their classes', () => {
  // GH #138 U2 released clear_print_titles -- create_pivot (not released) is the neighbour example now.
  assert.equal(W.validateWriteRequest(req([{ intent: 'create_pivot', id: 's', sheet: 'F' }])).ok, false, 'a not-released neighbour stays refused');
  assert.equal(validate({ intent: 'clear_filter', id: 's', sheet: 'F', range: 'A1:B4' }).ok, true);
});
