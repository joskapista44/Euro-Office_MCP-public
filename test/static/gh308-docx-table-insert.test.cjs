'use strict';

/**
 * GH #308 (write side, offline): the `insert_table` family (src/adapters/docx/families/table-insert.cjs).
 * plan/classify/resolve/verify against fabricated `normalizeInventory` observations -- no editor, no
 * VM. The commandBody itself (the actual callCommand string) is covered separately, against a real
 * VM, in gh308-docx-table-insert-vm.test.cjs.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const ti = require('../../src/adapters/docx/families/table-insert.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');

const ANCHOR = 'GH308 anchor';
const SIBLING = 'GH308 sibling';
const OTHER = 'GH308 other';

function obsOf(html) {
  return { primary: normalizeInventory({ html }) };
}

const noTable = () => `<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`;
const withTable = cells => `<p>${OTHER}</p><p>${ANCHOR}</p><table><tr>${cells[0].map(c => `<td>${c}</td>`).join('')}</tr><tr>${cells[1].map(c => `<td>${c}</td>`).join('')}</tr></table><p>${SIBLING}</p>`;

// ---- plan ---------------------------------------------------------------------------------------

test('plan: accepts { afterParagraph }, rows, cols, and an optional matching cells matrix', () => {
  const p = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 2, cols: 2 }, 0);
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operation.preservationMode, 'PRIMITIVE_INTEGRAL');
  assert.equal(p.operation.cells, null);
  const withCells = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 2, cols: 2, cells: [['a', 'b'], ['c', 'd']] }, 0);
  assert.equal(withCells.ok, true, withCells.detail);
  assert.deepEqual(withCells.operation.cells, [['a', 'b'], ['c', 'd']]);
});

test('plan: refuses a wrong-shaped selector, out-of-bound rows/cols, and a cells matrix that does not match rows x cols', () => {
  assert.equal(ti.plan({ intent: 'insert_table', selector: {}, rows: 1, cols: 1 }, 0).ok, false);
  assert.equal(ti.plan({ intent: 'insert_table', selector: { afterParagraph: '' }, rows: 1, cols: 1 }, 0).ok, false);
  assert.equal(ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR, extra: 1 }, rows: 1, cols: 1 }, 0).ok, false);
  for (const bad of [0, -1, 1.5, 51, 'x']) {
    assert.equal(ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: bad, cols: 1 }, 0).ok, false, `rows=${bad}`);
    assert.equal(ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: bad }, 0).ok, false, `cols=${bad}`);
  }
  assert.equal(ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 2, cols: 2, cells: [['a', 'b']] }, 0).ok, false, 'wrong row count');
  assert.equal(ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 2, cells: [['a']] }, 0).ok, false, 'wrong col count');
  assert.equal(ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1, cells: [['a\tb']] }, 0).ok, false, 'tab in a cell text');
});

// ---- classify / resolve --------------------------------------------------------------------------

test('classify: APPLY when the anchor is unique and no matching table follows it yet', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 2, cols: 2 }, 0).operation;
  const c = ti.classify(obsOf(noTable()).primary, op);
  assert.equal(c.state, 'APPLY');
});

test('classify: SATISFIED when the table right after the anchor already has this operation\'s own shape and (blank) cell texts', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 2, cols: 2 }, 0).operation;
  const c = ti.classify(obsOf(withTable([['', ''], ['', '']])).primary, op);
  assert.equal(c.state, 'SATISFIED');
});

test('classify: SATISFIED only when the cell texts also match -- a same-shape table with different text is APPLY, not a guessed match', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 2, cols: 2, cells: [['a', 'b'], ['c', 'd']] }, 0).operation;
  assert.equal(ti.classify(obsOf(withTable([['a', 'b'], ['c', 'd']])).primary, op).state, 'SATISFIED');
  assert.equal(ti.classify(obsOf(withTable([['x', 'y'], ['z', 'w']])).primary, op).state, 'APPLY');
  assert.equal(ti.classify(obsOf(noTable()).primary, op).state, 'APPLY');
});

test('classify: FAIL_CLOSED when the anchor is missing or ambiguous, never guessed', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: 'nowhere' }, rows: 1, cols: 1 }, 0).operation;
  assert.equal(ti.classify(obsOf(noTable()).primary, op).state, 'FAIL_CLOSED');
  const dup = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1 }, 0).operation;
  assert.equal(ti.classify(obsOf(`<p>${ANCHOR}</p><p>${ANCHOR}</p>`).primary, dup).state, 'FAIL_CLOSED');
});

test('classify: an anchor found only inside a table\'s own joined cell text is refused -- "csak body-szint", no cell-level anchor', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: 'cell one cell two' }, rows: 1, cols: 1 }, 0).operation;
  const html = '<table><tr><td>cell one</td><td>cell two</td></tr></table>';
  assert.equal(ti.classify(obsOf(html).primary, op).state, 'FAIL_CLOSED');
});

test('resolve / preDispatchCheck: round-trip through the real planOperations pipeline', () => {
  const planned = planOperations({ steps: [{ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 2, cols: 2 }] });
  assert.equal(planned.ok, true, planned.detail);
  const op = planned.operations[0];
  const r = ti.resolve(obsOf(noTable()).primary, op.selector);
  assert.equal(r.ok, true);
  assert.equal(r.target.anchorText, ANCHOR);
  assert.equal(ti.preDispatchCheck(op, r.target).ok, true);
  assert.equal(ti.preDispatchCheck(op, { kind: 'insert_table_anchor', anchorText: 'wrong' }).ok, false);
});

// ---- verify ---------------------------------------------------------------------------------------

test('verifyPrimitive: ok only when the after-state carries a matching table immediately after the (still-unique) anchor', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 2, cols: 2, cells: [['a', 'b'], ['c', 'd']] }, 0).operation;
  const before = obsOf(noTable());
  const good = ti.verifyPrimitive(op, before, obsOf(withTable([['a', 'b'], ['c', 'd']])));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  const bad = ti.verifyPrimitive(op, before, obsOf(noTable()));
  assert.equal(bad.ok, false);
});

// ---- verifyPreservation (L3) -----------------------------------------------------------------------

test('verifyPreservation: anchor NOT last -- OK when the rest is unchanged and nothing trails the new table', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1, cells: [['x']] }, 0).operation;
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const good = ti.verifyPreservation(op, before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><table><tr><td>x</td></tr></table><p>${SIBLING}</p>`));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  assert.equal(good.evidence.wasLast, false);
});

test('verifyPreservation: anchor NOT last -- FAILS if a trailing empty paragraph shows up anyway (that artifact is only measured when the anchor was last)', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1, cells: [['x']] }, 0).operation;
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const bad = ti.verifyPreservation(op, before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><table><tr><td>x</td></tr></table><p>${SIBLING}</p><p></p>`));
  assert.equal(bad.ok, false);
});

test('verifyPreservation: anchor NOT last -- FAILS if anything else outside the table changed', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1, cells: [['x']] }, 0).operation;
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const bad = ti.verifyPreservation(op, before, obsOf(`<p>${OTHER} CHANGED</p><p>${ANCHOR}</p><table><tr><td>x</td></tr></table><p>${SIBLING}</p>`));
  assert.equal(bad.ok, false);
});

test('verifyPreservation: anchor WAS last -- REQUIRES exactly one trailing empty paragraph after the table (the measured engine behaviour)', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1, cells: [['x']] }, 0).operation;
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p>`);
  const good = ti.verifyPreservation(op, before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><table><tr><td>x</td></tr></table><p></p>`));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  assert.equal(good.evidence.wasLast, true);
});

test('verifyPreservation: anchor WAS last -- FAILS if the trailing empty paragraph is missing or doubled', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1, cells: [['x']] }, 0).operation;
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p>`);
  const missing = ti.verifyPreservation(op, before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><table><tr><td>x</td></tr></table>`));
  assert.equal(missing.ok, false);
  const doubled = ti.verifyPreservation(op, before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><table><tr><td>x</td></tr></table><p></p><p></p>`));
  assert.equal(doubled.ok, false);
});

// ---- verifyPreservation: the anchor-only border-alt equivalence (the LIVE review round that found this) --------
// LIVE-measured: a re-save can flip the ANCHOR paragraph's own style between
// "mso-border-<side>-alt:none" and "border-<side>:none" -- same meaning, different spelling. The
// equivalence is scoped to that ONE block; it must not hide a real border change anywhere.

const BEFORE_STYLE = 'margin-top:0pt;margin-bottom:10pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none';
const AFTER_STYLE = 'margin-top:0pt;margin-bottom:10pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none';

test('normalizeAnchorBorderNoneAlt: rewrites mso-border-<side>-alt:none to border-<side>:none, nothing else', () => {
  assert.equal(ti.normalizeAnchorBorderNoneAlt(BEFORE_STYLE), AFTER_STYLE);
  assert.equal(ti.normalizeAnchorBorderNoneAlt('border-left:1px solid'), 'border-left:1px solid');
  assert.equal(ti.normalizeAnchorBorderNoneAlt(null), null);
});

test('verifyPreservation: the exact measured before/after anchor style pair is the SAME anchor (border-alt equivalence) -- PASS', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1, cells: [['x']] }, 0).operation;
  const before = obsOf(`<p>${OTHER}</p><p style="${BEFORE_STYLE}">${ANCHOR}</p><p>${SIBLING}</p>`);
  const good = ti.verifyPreservation(op, before, obsOf(`<p>${OTHER}</p><p style="${AFTER_STYLE}">${ANCHOR}</p><table><tr><td>x</td></tr></table><p>${SIBLING}</p>`));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

test('verifyPreservation: a REAL border change on the anchor (not the measured alt-spelling swap) still FAILS', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1, cells: [['x']] }, 0).operation;
  const before = obsOf(`<p>${OTHER}</p><p style="${BEFORE_STYLE}">${ANCHOR}</p><p>${SIBLING}</p>`);
  const bad = ti.verifyPreservation(op, before, obsOf(`<p>${OTHER}</p><p style="border-left:1px solid">${ANCHOR}</p><table><tr><td>x</td></tr></table><p>${SIBLING}</p>`));
  assert.equal(bad.ok, false);
});

test('verifyPreservation: the SAME before/after swap on a NON-anchor block (not the anchor) still FAILS', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 1, cols: 1, cells: [['x']] }, 0).operation;
  const before = obsOf(`<p style="${BEFORE_STYLE}">${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const bad = ti.verifyPreservation(op, before, obsOf(`<p style="${AFTER_STYLE}">${OTHER}</p><p>${ANCHOR}</p><table><tr><td>x</td></tr></table><p>${SIBLING}</p>`));
  assert.equal(bad.ok, false);
});

test('wholeTaskChecks: one named check, true only on a matching terminal table', () => {
  const op = ti.plan({ intent: 'insert_table', selector: { afterParagraph: ANCHOR }, rows: 2, cols: 2 }, 0).operation;
  const pass = ti.wholeTaskChecks(op, obsOf(withTable([['', ''], ['', '']])));
  assert.equal(pass.length, 1);
  assert.equal(pass[0].ok, true);
  const fail = ti.wholeTaskChecks(op, obsOf(noTable()));
  assert.equal(fail[0].ok, false);
});

test('producedText: a table is not a text unit a later step can ref', () => {
  assert.equal(ti.producedText({}), null);
});
