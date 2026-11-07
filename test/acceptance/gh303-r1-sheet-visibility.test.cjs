'use strict';

/**
 * GH #303 R1: set_sheet_visibility -- one sheet visible / hidden through the public SetVisible, verified by
 * read_sheet_state's visibility. The engine hides the LAST visible sheet too (docs/evidence/gh303-r1-sheet-visibility-probe-20261026.json),
 * so that refusal is ours, before the write and in the dispatch.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/sheet-visibility.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const obs = vis => ({ format: 'xlsx', complete: true, sheets: Object.keys(vis).map((name, index) => ({ name, index, used: null })), cells: {}, visibility: vis });
const stepOf = (extra = {}) => ({ intent: 'set_sheet_visibility', sheet: 'B', visibility: 'hidden', ...extra });
const opOf = s => { const p = fam.plan(s, 0); assert.equal(p.ok, true, p.detail); return p.operation; };

test('R1 schema + plan: visible | hidden only; veryHidden / a boolean / an extra field refused; the only step of its task', () => {
  assert.equal(W.validateWriteRequest(req([stepOf()])).ok, true);
  assert.equal(W.validateWriteRequest(req([stepOf({ visibility: 'visible' })])).ok, true);
  for (const visibility of ['veryHidden', true, 'shown']) {
    assert.equal(fam.plan(stepOf({ visibility }), 0).ok, false, JSON.stringify(visibility));
    assert.equal(W.validateWriteRequest(req([stepOf({ visibility })])).ok, false, JSON.stringify(visibility));
  }
  assert.equal(W.validateWriteRequest(req([stepOf({ active: true })])).ok, false);
  assert.equal(W.validateWriteRequest(req([stepOf(), stepOf({ sheet: 'C' })])).ok, false, 'two hides in one task could leave no visible sheet');
});

test('R1 classify: another state -> APPLY; this state -> SATISFIED; hiding the last visible sheet -> named FAIL_CLOSED', () => {
  const hide = opOf(stepOf());
  assert.equal(fam.classify(obs({ A: true, B: true }), hide).state, 'APPLY');
  assert.equal(fam.classify(obs({ A: true, B: false }), hide).state, 'SATISFIED');
  const last = fam.classify(obs({ A: false, B: true }), hide);
  assert.equal(last.state, 'FAIL_CLOSED');
  assert.equal(last.evidence.reason, 'xlsx-sheet-visibility-last-visible-sheet');
  const show = opOf(stepOf({ visibility: 'visible' }));
  assert.equal(fam.classify(obs({ A: true, B: false }), show).state, 'APPLY');
  assert.equal(fam.classify(obs({ A: true }), hide).evidence.reason, 'xlsx-sheet-not-found');
  assert.equal(fam.classify(obs({ A: true, B: 'UNOBSERVABLE:visibility-not-boolean' }), hide).evidence.reason, 'xlsx-sheet-visibility-unobservable');
  assert.equal(fam.classify({ ...obs({ A: true, B: true }), visibility: null }, hide).evidence.reason, 'xlsx-sheet-visibility-unobservable');
});

test('R1 lower: a hide re-checks another visible sheet in the dispatch; a show does not need to', () => {
  const hide = fam.lower(opOf(stepOf())).body;
  assert.match(hide, /last-visible-sheet/);
  assert.match(hide, /ws\.SetVisible\(false\)/);
  const show = fam.lower(opOf(stepOf({ visibility: 'visible' }))).body;
  assert.doesNotMatch(show, /last-visible-sheet/);
  assert.match(show, /ws\.SetVisible\(true\)/);
});

// GH #365 option C U2 (card an internal branch): this family's own `gate()` reads ONLY the public
// GetVisible() boolean (`p.visibility`), which never distinguishes hidden from veryHidden (both
// read `false`) -- so a veryHidden sheet is, BY CONSTRUCTION, already APPLY-eligible for 'visible'
// through the exact same code path the 'A: true, B: false' case above already proves. NOTHING in
// this family changed for U2: this test NAMES that pre-existing behaviour explicitly, so the card's
// own stated requirement ("veryHidden lapra 'visible' iras engedett") has a test that fails loudly
// if a FUTURE change (e.g. gating on veryHiddenReads here too) ever narrows it. Writing 'veryHidden'
// itself stays refused by the SAME schema check the first test above already covers -- this
// family's STATES never gained a third entry.
test('GH #365 C U2: revealing a veryHidden sheet (write visibility:"visible") is NOT blocked -- gate()/classify() cannot tell hidden and veryHidden apart, so both reach APPLY the same way; writing "veryHidden" itself remains refused', () => {
  const show = opOf(stepOf({ visibility: 'visible' }));
  // the obs() fixture's `visibility` map is exactly what GetVisible() would report for either a
  // plain-hidden OR a veryHidden sheet (both `false`) -- there is no third input shape to give it.
  const veryHiddenLookingObs = obs({ A: true, B: false });
  assert.equal(fam.classify(veryHiddenLookingObs, show).state, 'APPLY', 'revealing a (possibly veryHidden) hidden sheet must reach APPLY, never a named refusal');
  const lowered = fam.lower(show).body;
  assert.match(lowered, /ws\.SetVisible\(true\)/, 'the dispatch still calls the SAME public SetVisible(true) -- no veryHidden-specific engine path exists or is needed here');
  assert.equal(fam.plan(stepOf({ visibility: 'veryHidden' }), 0).ok, false, 'STATES has no veryHidden entry -- writing it directly stays refused, unchanged by U2');
});

test('R1 verify: L1 the read state; L3 every other sheet\'s visibility and the sheet list kept', () => {
  const op = opOf(stepOf());
  const before = { primary: obs({ A: true, B: true, C: false }) };
  assert.equal(fam.verifyPrimitive(op, before, { primary: obs({ A: true, B: false, C: false }) }).ok, true);
  assert.equal(fam.verifyPrimitive(op, before, { primary: obs({ A: true, B: true, C: false }) }).ok, false);
  assert.equal(fam.verifyPreservation(op, before, { primary: obs({ A: true, B: false, C: false }) }).ok, true);
  assert.equal(fam.verifyPreservation(op, before, { primary: obs({ A: true, B: false, C: true }) }).ok, false, 'another sheet shown');
  assert.equal(fam.verifyPreservation(op, before, { primary: obs({ A: true, B: false }) }).ok, false, 'a sheet gone');
});

test('R1 routing: set_sheet_visibility plans through the sheet-visibility family', () => {
  const p = planOperations({ id: 't', steps: [stepOf()] });
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operations[0].family, 'sheet-visibility');
});
