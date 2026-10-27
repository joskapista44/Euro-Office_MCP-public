'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { xlsxAdapter, planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const DC = require('../../src/adapters/xlsx/families/delete-cells.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

/**
 * L-18/1 (card an internal branch, kiosztás) --
 * the FIRST implementation unit: delete_cells wired into the product. docs/L18-DELETE-CELLS-SHIFT-PLAN.md
 * section 5 records the design: retry safety reuses `structure.cjs`'s OWN LIVE-proven anchor mechanism
 * (option (a) of a maintainer's two choices), scoped to the block's own band; everything the 5-round probe did NOT
 * measure (any name/rule/chart/pivot/AutoFilter reaching the sheet, any formula shape other than a bare
 * same-sheet cell reference, a merge the block only partially covers) is refused named, PRE_DISPATCH.
 *
 * This file is UNIT-level (plan/gate/classify/lower against hand-built observation objects, no fake driver,
 * no live connection) -- the family's own `p` shape mirrors `structure.cjs`'s and `range.cjs`'s conventions
 * exactly (p.cells keyed 'SHEET\\u0000ADDR', p.dependencies.refs, p.rules.bySheet, p.objects, p.filterReads).
 * A fake-driver END-TO-END test (the xlsx-w5-structure.test.cjs convention) is STILL NOT in this round --
 * extending the shared fake-xlsx-driver.cjs to simulate a NEW family's dispatch is real, separate work this
 * round's budget did not reach; named honestly, not silently dropped. The LIVE release harness
 * (test/live/release-l18-1-delete-cells-product.live.cjs, a maintainer's own ask #ref) IS now included -- section 9
 * below exercises its `verdict()` function against synthetic data, the l06b-move-range-release.test.cjs
 * convention (zero runs the harness itself; this file only proves the verdict function's own logic).
 */

const key = (sheet, addr) => `${sheet}\u0000${addr}`;

/** A minimal, complete `p` observation: one sheet (or more), no dependents/rules/objects/filter by default --
 * callers add what their case needs. */
function obs(sheets, { dependents = [], rulesBySheet = {}, objects = {}, filterReads = {}, broken = [] } = {}) {
  const cells = {};
  for (const [name, c] of Object.entries(sheets)) for (const [addr, v] of Object.entries(c)) cells[key(name, addr)] = v;
  // every sheet defaults to "AutoFilter read, none present" ({ range: null }) -- a sheet named in
  // `filterReads` by the caller (any axis under test) keeps its own value.
  const fr = { ...Object.fromEntries(Object.keys(sheets).map(name => [name, { range: null }])), ...filterReads };
  return {
    format: 'xlsx', complete: true, truncated: false,
    sheets: Object.keys(sheets).map(name => ({ name })),
    cells,
    dependencies: { complete: true, unobservable: [], refs: dependents, broken },
    rules: { complete: true, unobservable: [], bySheet: rulesBySheet },
    objects,
    filterReads: fr,
  };
}

// ---- 1. RED-first: today, before this patch, delete_cells is an unknown intent -----------------------------

test('RED/GREEN: delete_cells is now a KNOWN intent -- validateWriteRequest accepts a well-formed step (the "before this patch" shape was REFUSED_UNKNOWN_INTENT)', () => {
  const task = { id: 't', steps: [{ intent: 'delete_cells', id: 's', sheet: 'P', range: 'C2:D3', shift: 'up', anchor: { column: 'C', value: 'now-below' } }] };
  // RED, reproduced structurally (not re-run against the old code): an intent absent from the allowlist/schema
  // refuses UNKNOWN_INTENT/REFUSED_NOT_ALLOWLISTED before this patch -- the GREEN is that it no longer does.
  const v = W.validateWriteRequest({ fileId: '123', task });
  assert.notEqual(v.code, 'REFUSED_UNKNOWN_INTENT');
  assert.notEqual(v.code, 'REFUSED_NOT_ALLOWLISTED');
  // this specific step still refuses at PLAN (the anchor is never found in the fixture below) -- that is a
  // DIFFERENT, later check (section 3); this test only proves the intent itself is now recognized.
});

test('delete_cells is released: allowlist + accepted set + schema + constraints + SOLE_STEP + routed candidate', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('delete_cells'));
  assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes('delete_cells'));
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].sort());
  assert.ok(Object.prototype.hasOwnProperty.call(W.STEP_SCHEMAS, 'delete_cells'));
  assert.ok(Object.prototype.hasOwnProperty.call(W.CONSTRAINTS, 'delete_cells'));
  assert.ok(W.SOLE_STEP.includes('delete_cells'), 'delete_cells shifts its own band, same reasoning as structure -- must be the only step of its task');
  assert.ok(T.xlsxRoutedIntents().includes('delete_cells'));
});

// ---- 2. plan(): field/shape validation --------------------------------------------------------------------

const baseStep = (over = {}) => ({ intent: 'delete_cells', id: 's', sheet: 'P', range: 'C2:D3', shift: 'up', anchor: { column: 'C', value: 'x' }, ...over });

test('plan: an unknown field is refused, named', () => {
  const p = DC.plan({ ...baseStep(), extra: 1 }, 0);
  assert.equal(p.ok, false);
  assert.match(p.detail, /does not support field\(s\): extra/);
});

test('plan: shift must be up or left', () => {
  assert.equal(DC.plan({ ...baseStep(), shift: 'down' }, 0).ok, false);
});

test('plan: anchor.row is rejected for shift up (must use anchor.column)', () => {
  const p = DC.plan({ ...baseStep(), anchor: { row: 2, value: 'x' } }, 0);
  assert.equal(p.ok, false);
  assert.match(p.detail, /anchor\.row is not used for shift 'up'/);
});

test('plan: anchor.column is rejected for shift left (must use anchor.row)', () => {
  const p = DC.plan({ ...baseStep(), shift: 'left', anchor: { column: 'C', value: 'x' } }, 0);
  assert.equal(p.ok, false);
  assert.match(p.detail, /anchor\.column is not used for shift 'left'/);
});

test('plan: anchor.column must be WITHIN the block\'s own column band for shift up', () => {
  const p = DC.plan({ ...baseStep(), anchor: { column: 'A', value: 'x' } }, 0); // block is C2:D3, A is outside
  assert.equal(p.ok, false);
  assert.match(p.detail, /within the block's own column band/);
});

test('plan: anchor.row must be WITHIN the block\'s own row band for shift left', () => {
  const p = DC.plan({ ...baseStep(), shift: 'left', anchor: { row: 9, value: 'x' } }, 0); // block rows 2-3, 9 is outside
  assert.equal(p.ok, false);
  assert.match(p.detail, /within the block's own row band/);
});

test('plan: a well-formed UP step builds the expected operation shape (shift/band/anchor now+terminal)', () => {
  const p = DC.plan(baseStep(), 0);
  assert.equal(p.ok, true);
  assert.deepEqual(p.operation.block, { sheet: 'P', r1: 2, r2: 3, c1: 3, c2: 4 });
  assert.deepEqual(p.operation.shift, { sheet: 'P', axis: 'rows', at: 2, count: -2 });
  assert.deepEqual(p.operation.band, { lo: 3, hi: 4 });
  assert.deepEqual(p.operation.anchor, { lane: 3, value: 'x', now: 4, terminal: 2 });
});

test('plan: a well-formed LEFT step builds the expected operation shape', () => {
  const p = DC.plan({ ...baseStep(), shift: 'left', anchor: { row: 2, value: 'x' } }, 0);
  assert.equal(p.ok, true);
  assert.deepEqual(p.operation.shift, { sheet: 'P', axis: 'cols', at: 3, count: -2 });
  assert.deepEqual(p.operation.band, { lo: 2, hi: 3 });
  assert.deepEqual(p.operation.anchor, { lane: 2, value: 'x', now: 5, terminal: 3 });
});

test('plan: must be the only step of its task (SOLE_STEP via the adapter\'s own composition rule)', () => {
  const r = planOperations({ id: 't', steps: [baseStep(), { intent: 'write_range', id: 'w', sheet: 'P', range: 'A1', values: [[1]] }] });
  assert.equal(r.ok, false);
  assert.match(r.detail, /delete_cells must be the only step of its task/);
});

// ---- 3. classify(): retry-safety anchor mechanism (the card's point (ii)) ----------------------------------

const upOp = () => DC.plan(baseStep(), 0).operation;

test('classify: SATISFIED (retry-safe, zero writes) when the anchor is ALREADY at its terminal position', () => {
  const p = obs({ P: { C2: { v: 'x', f: '' }, D2: { v: 'd', f: '' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'SATISFIED');
  assert.equal(c.evidence.proof, 'fresh-live-anchor-at-its-terminal-position');
});

test('classify: APPLY when the anchor is at its pre-shift ("now") position and nothing refuses', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' }, D4: { v: 'd', f: '' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'APPLY');
});

test('classify: FAIL_CLOSED, named, when the anchor is at neither position (a real retry-unsafe state is never guessed)', () => {
  const p = obs({ P: { C7: { v: 'x', f: '' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-anchor-state-conflict');
});

test('classify: FAIL_CLOSED when the anchor value is not found at all, or found more than once', () => {
  const none = DC.classify(obs({ P: { C4: { v: 'other', f: '' } } }), upOp());
  assert.equal(none.evidence.reason, 'xlsx-delete-cells-anchor-not-found');
  const two = DC.classify(obs({ P: { C4: { v: 'x', f: '' }, C9: { v: 'x', f: '' } } }), upOp());
  assert.equal(two.evidence.reason, 'xlsx-delete-cells-anchor-not-unique');
});

// ---- 4. classify(): the merge guard -- the card's own key finding (full-cover vs partial) -------------------

test('classify: a merge ENTIRELY inside the block does NOT refuse (measured DELETED_WITH_BLOCK -- it goes with the block)', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' }, C2: { v: 'm', f: '', merge: 'C2:D2' }, D2: { v: '', f: '', merge: 'C2:D2' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'APPLY');
});

test('classify: a merge that only PARTIALLY overlaps the block is refused, named (the measured TRUNCATED finding -- the DS silently shrinks it otherwise)', () => {
  // C3:D4 straddles the block's bottom edge (block is C2:D3) -- round 5's exact measured case.
  const p = obs({ P: { C4: { v: 'x', f: '' }, C3: { v: 'm', f: '', merge: 'C3:D4' }, D3: { v: '', f: '', merge: 'C3:D4' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-merge-partial-overlap');
});

// RED-first (a maintainer's review, #ref): BEFORE this fix, a merge that never touches the block's own rows/cols
// at all -- it sits further down the SHIFT PATH, where the shift still moves part of it -- passed mergeGuard
// completely unseen (overlapsBlock is false: the block's rows 2-3 never reach row 6). The shift then tears it
// apart exactly like a block-straddling merge does (B6 stays, C6 shifts up into the block's own band).
test('classify: a merge BEYOND the block, in the shift path, crossing the band edge is refused, named -- zero\'s #ref finding (B6:C6, block C2:D3 up)', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' }, B6: { v: 'm', f: '', merge: 'B6:C6' }, C6: { v: '', f: '', merge: 'B6:C6' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-merge-crosses-band');
});

test('classify: a merge ENTIRELY WITHIN the band, in the shift path, is ALSO refused -- unmeasured whether it shifts intact (C6:D6, block C2:D3 up)', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' }, C6: { v: 'm', f: '', merge: 'C6:D6' }, D6: { v: '', f: '', merge: 'C6:D6' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-merge-crosses-band');
});

test('classify: a merge in the shift path but ENTIRELY OUTSIDE the band is inert -- never refused (A6:A7, block C2:D3 up)', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' }, A6: { v: 'm', f: '', merge: 'A6:A7' }, A7: { v: '', f: '', merge: 'A6:A7' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'APPLY');
});

const leftOp = () => DC.plan({ ...baseStep(), shift: 'left', anchor: { row: 2, value: 'x' } }, 0).operation;

test('classify: the LEFT mirror of the shift-path merge finding -- a merge beyond the block\'s columns, crossing the band\'s row edge, is refused, named (F1:F2, block C2:D3 left)', () => {
  const p = obs({ P: { E2: { v: 'x', f: '' }, F1: { v: 'm', f: '', merge: 'F1:F2' }, F2: { v: '', f: '', merge: 'F1:F2' } } });
  const c = DC.classify(p, leftOp());
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-merge-crosses-band');
});

test('classify: LEFT -- a shift-path merge entirely outside the band (rows 2-3) is inert (F4:F5, block C2:D3 left)', () => {
  const p = obs({ P: { E2: { v: 'x', f: '' }, F4: { v: 'm', f: '', merge: 'F4:F5' }, F5: { v: '', f: '', merge: 'F4:F5' } } });
  const c = DC.classify(p, leftOp());
  assert.equal(c.state, 'APPLY');
});

// ---- 5. classify(): the blanket dependent guard (names/rules/charts/pivots/AutoFilter -- NEM-MERT) ----------

test('classify: a defined name reaching the sheet refuses, named, blanket (not measured for a block-shift)', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' } } }, { dependents: [{ kind: 'name', sheets: ['P'], host: null, name: 'N' }] });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-dependent-not-measured');
});

test('classify: a validation/CF rule hosted on the sheet refuses, named, blanket', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' } } }, { rulesBySheet: { P: [{ kind: 'validation', range: 'A1' }] } });
  const c = DC.classify(p, upOp());
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-dependent-not-measured');
});

test('classify: a chart or pivot hosted on the sheet refuses, named, blanket', () => {
  const chart = DC.classify(obs({ P: { C4: { v: 'x', f: '' } } }, { objects: { P: { charts: [{}] } } }), upOp());
  assert.equal(chart.evidence.reason, 'xlsx-delete-cells-dependent-not-measured');
  const pivot = DC.classify(obs({ P: { C4: { v: 'x', f: '' } } }, { objects: { P: { pivots: [{}] } } }), upOp());
  assert.equal(pivot.evidence.reason, 'xlsx-delete-cells-dependent-not-measured');
});

test('classify: an AutoFilter on the sheet refuses, named, blanket', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' } } }, { filterReads: { P: { range: 'A1:B5' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-dependent-not-measured');
});

// ---- 6. classify(): the formula-shape guard (bare same-sheet cell reference ONLY) ----------------------------

test('classify: a bare same-sheet cell formula in the shift path is ACCEPTED (the measured C5=\'=C4+1\' shape)', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' }, C9: { v: 7, f: '=C4+1' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'APPLY');
});

test('classify: a RANGE-shaped formula on the same sheet refuses, named (not the measured bare-cell shape)', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' }, C9: { v: 7, f: '=SUM(C1:C3)' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-formula-shape-not-measured');
});

test('classify: a formula on ANOTHER sheet reaching INTO the shifting sheet refuses, named -- the earlier draft of this guard, scoped to one sheet only, would have missed this', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' } }, Q: { A1: { v: 9, f: '=P!C4' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.evidence.reason, 'xlsx-delete-cells-formula-shape-not-measured');
  assert.ok(c.evidence.cells.includes('Q!A1'), JSON.stringify(c.evidence));
});

test('classify: a formula on another sheet NOT reaching the shifting sheet at all does not refuse', () => {
  const p = obs({ P: { C4: { v: 'x', f: '' } }, Q: { A1: { v: 9, f: '=Q!B1' }, B1: { v: 1, f: '' } } });
  const c = DC.classify(p, upOp());
  assert.equal(c.state, 'APPLY');
});

test('classify: #REF! (C) -- a broken formula on ANOTHER sheet does not refuse; ON the shifting sheet it refuses, named xlsx-ref-on-target-unmeasured', () => {
  const elsewhere = obs({ P: { C4: { v: 'x', f: '' } }, Q: { A1: { v: '#REF!', f: '= #REF!*2' } } }, { broken: [{ kind: 'formula', host: 'Q', name: 'A1', text: '= #REF!*2' }] });
  assert.equal(DC.classify(elsewhere, upOp()).state, 'APPLY');
  const here = obs({ P: { C4: { v: 'x', f: '' }, H1: { v: '#REF!', f: '=#REF!A1' } } }, { broken: [{ kind: 'formula', host: 'P', name: 'H1', text: '=#REF!A1' }] });
  const c = DC.classify(here, upOp());
  assert.notEqual(c.state, 'APPLY');
  assert.match(JSON.stringify(c), /xlsx-ref-on-target-unmeasured/);
});

// ---- 6b. verifyPrimitive(): the G1-style cross-reference case (round 4/5's exact LIVE finding) -------------

const wrap = p => ({ primary: p });

test('verifyPrimitive: an out-of-band formula pointing AT a cell inside the deleted block correctly becomes #REF! (measured: G1 "=C3" -> "=#REF!")', () => {
  const before = obs({ P: { C4: { v: 'x', f: '' }, G1: { v: 0, f: '=C3' } } });
  const after = obs({ P: { C2: { v: 'x', f: '' }, G1: { v: '#REF!', f: '=#REF!' } } });
  const r = DC.verifyPrimitive(upOp(), wrap(before), wrap(after));
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});

test('verifyPrimitive: an out-of-band formula pointing AT a cell BEHIND the block correctly re-points with the shift (measured: a shift-path formula outside the band, e.g. a row-1 cell referencing into the band from outside it, is NOT this shape -- but a same-sheet cell reference whose target is in-band and survives DOES move)', () => {
  // G1 = "=C4" (C4 is in-band, behind the block -- survives the shift, lands at C2): the reference must follow.
  const before = obs({ P: { C4: { v: 'x', f: '' }, G1: { v: 'x', f: '=C4' } } });
  const after = obs({ P: { C2: { v: 'x', f: '' }, G1: { v: 'x', f: '=C2' } } });
  const r = DC.verifyPrimitive(upOp(), wrap(before), wrap(after));
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});

test('verifyPrimitive: an out-of-band formula pointing at a cell OUTSIDE the band entirely never changes', () => {
  const before = obs({ P: { C4: { v: 'x', f: '' }, G1: { v: 5, f: '=A1' }, A1: { v: 5, f: '' } } });
  const after = obs({ P: { C2: { v: 'x', f: '' }, G1: { v: 5, f: '=A1' }, A1: { v: 5, f: '' } } });
  const r = DC.verifyPrimitive(upOp(), wrap(before), wrap(after));
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});

test('verifyPrimitive: FAILS when a measured cross-reference did NOT update as the shift implies', () => {
  const before = obs({ P: { C4: { v: 'x', f: '' }, G1: { v: 0, f: '=C3' } } });
  const after = obs({ P: { C2: { v: 'x', f: '' }, G1: { v: 0, f: '=C3' } } }); // wrong: should be #REF! now
  const r = DC.verifyPrimitive(upOp(), wrap(before), wrap(after));
  assert.equal(r.ok, false);
  assert.ok(r.evidence.wrong.some(w => w.from === 'G1'));
});

test('verifyPrimitive: the in-band shift-path formula (measured C5="=C4+1" -> lands at C3 as "=C2+1") re-points correctly', () => {
  // C4 doubles as the operation's own anchor cell (upOp()'s anchor.value is 'x') -- its landing at C2
  // is what wholeTaskChecks/classify's retry-safety reads, same as every other case in this file.
  const before = obs({ P: { C4: { v: 'x', f: '' }, C5: { v: 34, f: '=C4+1' } } });
  const after = obs({ P: { C2: { v: 'x', f: '' }, C3: { v: 34, f: '=C2+1' } } });
  const r = DC.verifyPrimitive(upOp(), wrap(before), wrap(after));
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});

// ---- 7. lower(): the raw callCommand body -------------------------------------------------------------------

test('lower: UP produces GetRange(block).Delete(\'up\'), LEFT produces Delete(\'left\')', () => {
  const up = DC.lower(upOp());
  assert.equal(up.kind, 'callCommand');
  assert.match(up.body, /GetRange\("C2:D3"\)\.Delete\('up'\)/);
  require('node:assert').doesNotThrow(() => new Function(up.body)); // the generated body itself is valid JS
  const left = DC.lower(DC.plan({ ...baseStep(), shift: 'left', anchor: { row: 2, value: 'x' } }, 0).operation);
  assert.match(left.body, /GetRange\("C2:D3"\)\.Delete\('left'\)/);
});

// ---- 8. wholeTaskChecks() ------------------------------------------------------------------------------------

test('wholeTaskChecks: ok when the anchor ends at its terminal position', () => {
  const checks = DC.wholeTaskChecks(upOp(), { primary: obs({ P: { C2: { v: 'x', f: '' } } }) });
  assert.equal(checks.length, 1);
  assert.equal(checks[0].ok, true);
});

test('wholeTaskChecks: not ok when the anchor is missing from the final observation', () => {
  const checks = DC.wholeTaskChecks(upOp(), { primary: obs({ P: { C9: { v: 'other', f: '' } } }) });
  assert.equal(checks[0].ok, false);
});

// ---- 9. LIVE harness verdict() -- the l06b-move-range-release.test.cjs convention: a synthetic GREEN run
// passes, and each broken gate fails on its OWN named check, not a different one. ----------------------------

test('L-18/1 LIVE harness verdict: a synthetic GREEN run passes; the pre-release shape (delete_cells not allowlisted) and each broken gate fail', { skip: needsPrivate('test/live/release-l18-1-delete-cells-product.live.cjs') }, () => {
  const h = require('../live/release-l18-1-delete-cells-product.live.cjs');
  const UP_SHEET = 'FAKEUP';
  const ap = () => ({ ok: true, persistence: 'FORCESAVE_CONFIRMED', ops: [{ status: 'APPLIED' }] });
  // a seed task with MULTIPLE ops (e.g. UP_v's 9 steps) -- a maintainer's 2nd LIVE round: `checks.seeds`
  // must accept this shape, not just a single-op one.
  const seedAp = (n = 1) => ({ ok: true, persistence: 'FORCESAVE_CONFIRMED', ops: Array.from({ length: n }, () => ({ status: 'APPLIED' })) });
  const satisfied = () => ({ ok: true, persistence: null, ops: [{ status: 'SATISFIED' }] });
  const namedFail = reason => ({ ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason } }] });
  // the pre-fix answer on the shared copy: LEFT refused workbook-wide because of UP's G1 "=#REF!"
  const blockedFail = axis => ({ ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.LEFT.preFixReason, unobservable: [{ axis, why: 'UNOBSERVABLE:broken-reference', host: UP_SHEET }] } }] });
  const readOf = want => { const o = {}; for (const [a, w] of Object.entries(want || {})) o[a] = { v: w.v, f: w.f || '' }; return o; };
  const green = () => {
    const o = { SHEET_NAMES: { UP: UP_SHEET }, SEEDS: { a: seedAp(9) }, MERGE_SEEDS: { a: ap() }, PAIRS: {}, PAIRS2: {}, MERGE_IN: {}, MERGE_SPAN: {}, MERGE_BAND: {}, NG_SHEET: {}, READ: {}, READ2: {}, CLEANUP: { gone: true }, CLEANUP2: { gone: true }, FIXTURE_MUTATED: 'NO' };
    for (const p of h.POS) {
      o.PAIRS[p.key] = { first: ap(), second: satisfied() }; o.READ[p.key] = readOf(p.want || p.want2);
    }
    // LEFT's own happy path, on the second, clean copy (+ #ref).
    o.PAIRS2.LEFT = { first: ap(), second: satisfied() };
    o.READ2.LEFT = readOf(h.LEFT.want2);
    o.MERGE_IN = { first: ap(), second: satisfied() };
    o.READ.MERGE_IN = readOf(h.MERGE_IN.want);
    o.MERGE_SPAN = namedFail(h.MERGE_SPAN.reason);
    o.READ.MERGE_SPAN = readOf(h.MERGE_SPAN.keep);
    o.MERGE_BAND = namedFail(h.MERGE_BAND.reason);
    o.READ.MERGE_BAND = readOf(h.MERGE_BAND.keep);
    o.NG_SHEET = namedFail('xlsx-sheet-not-found');
    o.UP_REF = namedFail(h.UP_REF.reason);
    o.READ.UT = { C4: { v: 33, f: '' }, C5: { v: 34, f: '=C4+1' }, G1: { v: 'h7', f: '' } };
    return o;
  };
  assert.equal(h.verdict(green()).PASS, true, JSON.stringify(h.verdict(green()).failed));
  const red = green(); red.PAIRS.UP = { first: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] }, second: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] } };
  assert.equal(h.verdict(red).PASS, false, 'the pre-release build is RED');
  assert.equal(h.verdict(red).outcome, 'DS_NOT_RELEASED');
  const stillApplying = green(); stillApplying.PAIRS.UP.second = ap();
  assert.ok(h.verdict(stillApplying).failed.includes('UP_retry_satisfied'), 'a second APPLIED is now WRONG, not right (retry must be SATISFIED)');
  // #REF! (C): the pre-fix workbook-wide refusal of LEFT (UP's G1 on ANOTHER sheet) is now RED, on LEFT's own check
  const preFix = green(); preFix.PAIRS.LEFT = { first: blockedFail(`formula:${UP_SHEET}:G1`), second: blockedFail(`formula:${UP_SHEET}:G1`) };
  assert.ok(h.verdict(preFix).failed.includes('LEFT_applied'), 'a #REF! on ANOTHER sheet must no longer block LEFT');
  const leftMutated = green(); leftMutated.READ.LEFT.C2 = { v: 'WRONG', f: '' };
  assert.ok(h.verdict(leftMutated).failed.includes('LEFT_saved'), 'LEFT on the shared copy must land the same values as on the clean copy');
  const upRefApplied = green(); upRefApplied.UP_REF = ap();
  assert.ok(h.verdict(upRefApplied).failed.includes('UP_REF_refused'), 'a delete_cells on the sheet hosting the #REF! must be refused');
  const upRefOtherReason = green(); upRefOtherReason.UP_REF = namedFail('xlsx-delete-cells-dependency-observation-incomplete');
  assert.ok(h.verdict(upRefOtherReason).failed.includes('UP_REF_refused'), 'refused by the named target guard, not the old workbook-wide refusal');
  const seedPartial = green(); seedPartial.SEEDS.a = { ok: true, persistence: 'FORCESAVE_CONFIRMED', ops: [{ status: 'APPLIED' }, { status: 'PENDING' }] };
  assert.ok(h.verdict(seedPartial).failed.includes('seeds'), 'a seed with ANY non-APPLIED op is not fully seeded');
  const mergeSurvived = green(); mergeSurvived.READ.MERGE_IN.D2 = { v: 34, f: '', merge: 'C2:D2' };
  assert.ok(h.verdict(mergeSurvived).failed.includes('MERGE_IN_merge_gone'), 'a dissolved merge must not survive the read-back');
  const notRefused = green(); notRefused.MERGE_SPAN = ap();
  assert.ok(h.verdict(notRefused).failed.includes('MERGE_SPAN_refused'), 'an APPLIED merge-partial-overlap is now WRONG, not right');
  const spanMutated = green(); spanMutated.READ.MERGE_SPAN.C2 = { v: 'CHANGED', f: '' };
  assert.ok(h.verdict(spanMutated).failed.includes('MERGE_SPAN_untouched'), 'a refused dispatch must leave the sheet byte-for-byte as seeded');
  const bandNotRefused = green(); bandNotRefused.MERGE_BAND = ap();
  assert.ok(h.verdict(bandNotRefused).failed.includes('MERGE_BAND_refused'), 'zero\'s #ref finding: an APPLIED shift-path merge crossing the band is now WRONG, not right');
  const bandMutated = green(); bandMutated.READ.MERGE_BAND.C4 = { v: 'CHANGED', f: '' };
  assert.ok(h.verdict(bandMutated).failed.includes('MERGE_BAND_untouched'));
  const wrongReason = green(); wrongReason.NG_SHEET = namedFail('other');
  assert.ok(h.verdict(wrongReason).failed.includes('NG_SHEET_refused'));
  const utChanged = green(); utChanged.READ.UT.C4 = { v: 'CHANGED', f: '' };
  assert.ok(h.verdict(utChanged).failed.includes('UT_untouched'));
  const notCleanedUp = green(); notCleanedUp.CLEANUP = { gone: false };
  assert.ok(h.verdict(notCleanedUp).failed.includes('cleanup'));
  const left2StillApplying = green(); left2StillApplying.PAIRS2.LEFT.second = ap();
  assert.ok(h.verdict(left2StillApplying).failed.includes('LEFT_retry_satisfied2'), 'LEFT\'s OWN happy path (2nd copy): a second APPLIED is wrong, retry must be SATISFIED');
  const left2WrongShift = green(); left2WrongShift.READ2.LEFT.C2 = { v: 'WRONG', f: '' };
  assert.ok(h.verdict(left2WrongShift).failed.includes('LEFT_saved2'));
  const notCleanedUp2 = green(); notCleanedUp2.CLEANUP2 = { gone: false };
  assert.ok(h.verdict(notCleanedUp2).failed.includes('cleanup2'));
});
