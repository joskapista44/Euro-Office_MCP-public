'use strict';

/**
 * GH #344 colour filter WRITE (internal tracker, docs/DS-9.3.5-COLOR-FILTER-WRITE-PROPOSAL.md): filter_range's sixth,
 * mutually-exclusive criteria alternative, `colour: { kind: 'cellColor' | 'fontColor', rgb: [r, g, b] }`.
 * Verification reuses filter-read.cjs's own `colourOf()` (Bridge /21) -- the public GetCriteria1/2 getters read a
 * colour filter back null (unchanged), so this is the only in-family resolution path, never duplicated.
 * The call shape (`Api.CreateColorFromRGB(r, g, b), 'xlFilter<Cell|Font>Color'`) is the one already LIVE-measured in
 * xlsx-gh344-gh345-filter-color-empty-public-probe.live.cjs:118-119 -- this unit only wires it into filter_range.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const filter = require('../../src/adapters/xlsx/families/filter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const book = () => [{ name: 'F', cells: {
  A1: { v: 'id' }, B1: { v: 'tag' },
  A2: { v: 1 }, B2: { v: 'w' }, A3: { v: 2 }, B3: { v: 'x' }, A4: { v: 3 }, B4: { v: 'y' }, A5: { v: 4 }, B5: { v: 'z' }, A6: { v: 5 }, B6: { v: 'w' },
} }];
const flt = extra => ({ intent: 'filter_range', id: 's', sheet: 'F', range: 'A1:B6', ...extra });
const statuses = r => (r.operations || []).map(o => o.status);
const v = s => W.validateWriteRequest(req([s]));
const RED = [255, 0, 0];
const BLUE = [0, 0, 255];

test('schema: colour alone is accepted (both kinds); combined with any other criteria, or malformed, is refused', () => {
  assert.equal(v(flt({ field: 1, colour: { kind: 'cellColor', rgb: RED } })).ok, true);
  assert.equal(v(flt({ field: 1, colour: { kind: 'fontColor', rgb: BLUE } })).ok, true);
  for (const [step, why] of [
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: RED }, value: 'x' }), 'colour with value'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: RED }, values: ['x', 'y'] }), 'colour with values'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: RED }, top: 3 }), 'colour with top'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: RED }, dynamic: 'aboveAverage' }), 'colour with dynamic'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: RED }, and: ['>1', '<9'] }), 'colour with and'],
    [flt({ field: 1, colour: { kind: 'themeColor', rgb: RED } }), 'invalid kind'],
    [flt({ field: 1, colour: { rgb: RED } }), 'missing kind'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: [256, 0, 0] } }), 'rgb component over 255'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: [-1, 0, 0] } }), 'rgb component negative'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: [1.5, 0, 0] } }), 'rgb component not an integer'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: [0, 0] } }), 'rgb only two components'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: [0, 0, 0, 0] } }), 'rgb four components'],
    [flt({ field: 1, colour: { kind: 'cellColor', rgb: 'red' } }), 'rgb not an array'],
    [flt({ field: 1, colour: {} }), 'empty colour object'],
  ]) assert.equal(v(step).ok, false, why);
});

test('the pre-write refusal names are exactly colour-and-other-criteria / colour-kind-invalid / colour-rgb-out-of-range', () => {
  const problem = step => filter.plan(step, 0).detail;
  assert.match(problem(flt({ field: 1, colour: { kind: 'cellColor', rgb: RED }, value: 'x' })), /colour-and-other-criteria/);
  assert.match(problem(flt({ field: 1, colour: { kind: 'bad', rgb: RED } })), /colour-kind-invalid/);
  assert.match(problem(flt({ field: 1, colour: { kind: 'cellColor', rgb: [999, 0, 0] } })), /colour-rgb-out-of-range/);
});

test('lower(): the SetAutoFilter call is exactly the measured colour shape, for both kinds', () => {
  const body = step => filter.lower(filter.plan(flt(step), 0).operation).body;
  assert.match(body({ field: 1, colour: { kind: 'cellColor', rgb: RED } }), /SetAutoFilter\(1, Api\.CreateColorFromRGB\(255, 0, 0\), 'xlFilterCellColor'\)/);
  assert.match(body({ field: 1, colour: { kind: 'fontColor', rgb: BLUE } }), /SetAutoFilter\(1, Api\.CreateColorFromRGB\(0, 0, 255\), 'xlFilterFontColor'\)/);
});

test('APPLIED, then retry SATISFIED, 0 extra dispatch -- both kinds', async () => {
  for (const colour of [{ kind: 'cellColor', rgb: RED }, { kind: 'fontColor', rgb: BLUE }]) {
    const h = harness({ store: createStore(book()) });
    const first = await h.tool.handler(req([flt({ field: 1, colour })]));
    assert.equal(first.ok, true, J(first).slice(0, 600));
    assert.deepEqual(statuses(first), ['APPLIED']);
    const before = h.calls().dispatch;
    const again = await h.tool.handler(req([flt({ field: 1, colour })]));
    assert.deepEqual(statuses(again), ['SATISFIED']);
    assert.equal(h.calls().dispatch, before, 'the retry dispatches nothing');
  }
});

test('colour-unverifiable: the Bridge capability never lists the column -- dispatched, but NOT reported APPLIED', async () => {
  const h = harness({ store: createStore(book()), script: { fault: 'filter-colour-unlisted' } });
  const r = await h.tool.handler(req([flt({ field: 1, colour: { kind: 'cellColor', rgb: RED } })]));
  assert.ok(h.calls().dispatch >= 1, 'the write was dispatched -- this is the VERIFY failing, not a pre-write refusal');
  assert.equal(r.ok, false);
});

test('a fake that resolves a DIFFERENT rgb / the wrong kind fails the verify (the Bridge read-back is the proof)', async () => {
  for (const fault of ['filter-wrong-value', 'filter-colour-wrong-kind']) {
    const h = harness({ store: createStore(book()), script: { fault } });
    const r = await h.tool.handler(req([flt({ field: 1, colour: { kind: 'cellColor', rgb: RED } })]));
    assert.ok(h.calls().dispatch >= 1, `${fault}: the write was dispatched`);
    assert.equal(r.ok, false, `${fault} must not pass`);
  }
});

test('regression: a colour filter on one field does not disturb a later value filter on another field, and the colour still holds', async () => {
  const h = harness({ store: createStore(book()) });
  const r1 = await h.tool.handler(req([flt({ field: 1, colour: { kind: 'cellColor', rgb: RED } })], 't1'));
  assert.equal(r1.ok, true, J(r1).slice(0, 600));
  const r2 = await h.tool.handler(req([flt({ field: 2, value: 'x' })], 't2'));
  assert.equal(r2.ok, true, J(r2).slice(0, 600));
  // retry the colour filter alone -- SATISFIED proves field 2's write left it exactly as it was (others() kept)
  const r3 = await h.tool.handler(req([flt({ field: 1, colour: { kind: 'cellColor', rgb: RED } })], 't3'));
  assert.deepEqual(statuses(r3), ['SATISFIED'], J(r3).slice(0, 600));
});

// classify()/verifyPrimitive() on a hand-built observation in the /21-colours shape (no harness/driver, no dispatch)
const ALL_VISIBLE = Object.fromEntries([2, 3, 4, 5, 6].map(r => [String(r), { hidden: false }]));
const obs = (filters, colours) => ({
  format: 'xlsx', complete: true, sheets: [{ name: 'F' }],
  filterReads: { F: { range: 'A1:B6', filters, ...(colours !== undefined ? { colours } : {}) } },
  layout: { F: { rows: ALL_VISIBLE } },
});
const op = colour => filter.plan({ intent: 'filter_range', id: 's', sheet: 'F', range: 'A1:B6', field: 1, colour }, 0).operation;
const cellColourFilter = { field: 1, on: true, operator: 'xlFilterCellColor', criteria1: null, criteria2: null };

test('classify(): no active filter yet -> APPLY; the matching colour already there -> SATISFIED', () => {
  const o1 = op({ kind: 'cellColor', rgb: RED });
  assert.equal(filter.classify(obs([]), o1).state, 'APPLY');
  const matchObs = obs([cellColourFilter], { columns: [{ field: 1, kind: 'cellColor', colour: { form: 'rgb', hex: 'FF0000' } }] });
  assert.equal(filter.classify(matchObs, o1).state, 'SATISFIED');
});

test('verifyPrimitive(): rgb form compares exact hex; indexed/theme form accepts a matching kind only (G1 not run)', () => {
  const o1 = op({ kind: 'cellColor', rgb: RED });
  const p1 = { primary: obs([cellColourFilter], { columns: [{ field: 1, kind: 'cellColor', colour: { form: 'rgb', hex: 'FF0000' } }] }) };
  assert.equal(filter.verifyPrimitive(o1, p1, p1).ok, true, 'exact rgb hex match');
  const pWrong = { primary: obs([cellColourFilter], { columns: [{ field: 1, kind: 'cellColor', colour: { form: 'rgb', hex: '00FF00' } }] }) };
  assert.equal(filter.verifyPrimitive(o1, pWrong, pWrong).ok, false, 'a different rgb hex must not verify');
  const pIndexed = { primary: obs([cellColourFilter], { columns: [{ field: 1, kind: 'cellColor', colour: { form: 'indexed', indexed: 2 } }] }) };
  assert.equal(filter.verifyPrimitive(o1, pIndexed, pIndexed).ok, true, 'an indexed-form resolution of the matching kind is accepted (no palette table to compare exactly, G1 open)');
  const pKindMismatch = { primary: obs([cellColourFilter], { columns: [{ field: 1, kind: 'fontColor', colour: { form: 'indexed', indexed: 2 } }] }) };
  assert.equal(filter.verifyPrimitive(o1, pKindMismatch, pKindMismatch).ok, false, 'a resolved kind that disagrees with the request must not verify, regardless of form');
});

test('verifyPrimitive(): colour-unverifiable -- the capability never lists the field, or the dxf holds no colour', () => {
  const o1 = op({ kind: 'cellColor', rgb: RED });
  const pUnlisted = { primary: obs([cellColourFilter], { columns: [] }) };
  assert.equal(filter.verifyPrimitive(o1, pUnlisted, pUnlisted).ok, false, 'column not listed -> UNOBSERVABLE, not APPLIED');
  const pNoCapability = { primary: obs([cellColourFilter], undefined) };
  assert.equal(filter.verifyPrimitive(o1, pNoCapability, pNoCapability).ok, false, 'no colours capability at all (a runtime without Bridge /21) -> not APPLIED');
  const pNoColour = { primary: obs([cellColourFilter], { columns: [{ field: 1, kind: 'cellColor', colour: null }] }) };
  assert.equal(filter.verifyPrimitive(o1, pNoColour, pNoColour).ok, false, 'COLOR_NONE (dxf has no colour) -> not APPLIED');
});

// review: the real driver reads a colour filter's colour ONLY when the request asks for the /21 axis
// (xlsx-driver readQuery `filterColor`); the fake driver serves `colours` regardless, so this is asserted on the
// observation request itself -- RED before the review fix (liveScope carried no filterColor at all)
test('GH #344 review: a colour filter_range asks the live observation for the /21 colour axis; a value filter does not', () => {
  const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
  const colourTask = { id: 't', steps: [{ intent: 'filter_range', sheet: 'S', range: 'A1:B5', field: 1, colour: { kind: 'cellColor', rgb: [255, 0, 0] } }] };
  const valueTask = { id: 't', steps: [{ intent: 'filter_range', sheet: 'S', range: 'A1:B5', field: 1, value: 'x' }] };
  assert.equal(liveScope(colourTask).filterColor, true);
  assert.deepEqual(liveScope(colourTask).filterSheets, ['S']);
  assert.equal(liveScope(valueTask).filterColor, undefined);
});
