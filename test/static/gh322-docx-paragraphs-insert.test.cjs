'use strict';

/**
 * GH #322 (write side, offline): the `insert_paragraphs` family
 * (src/adapters/docx/families/paragraphs-insert.cjs). plan/classify/resolve/verify against
 * fabricated `normalizeInventory` observations -- no editor, no VM. The commandBody itself (the
 * actual callCommand string) is covered by the LIVE probe (test/live/docx-gh322-bulk-insert-probe.
 * live.cjs, an internal branch @48afe5074) and by the dedicated LIVE product harness, not here.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const pi = require('../../src/adapters/docx/families/paragraphs-insert.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');

const ANCHOR = 'GH322 anchor';
const SIBLING = 'GH322 sibling';
const OTHER = 'GH322 other';

function obsOf(html, paragraphs = []) {
  return { primary: normalizeInventory({ html, paragraphs }) };
}

const items = texts => texts.map(text => ({ text }));
const op = (overrides = {}) => pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: items(['a', 'b']), ...overrides }, 0).operation;

// ---- plan -----------------------------------------------------------------------------------------

test('plan: accepts { afterParagraph } and 1..200 paragraph items, each with an optional style/format', () => {
  const p = pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: [{ text: 'a' }, { text: 'b', style: 'Normal', format: { bold: true, color: '#FF0000', alignment: 'center', spaceBefore: 6 } }] }, 0);
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operation.preservationMode, 'PRIMITIVE_INTEGRAL');
  assert.equal(p.operation.paragraphs.length, 2);
  assert.equal(p.operation.paragraphs[1].style, 'Normal');
  assert.deepEqual(p.operation.paragraphs[1].format, { bold: true, spaceBefore: 6, alignment: 'center', color: '#FF0000' });
});

test('plan: refuses a wrong-shaped selector, an out-of-range paragraph count, a tab/newline in text, an empty text, a non-string style, and an unknown format key', () => {
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: {}, paragraphs: items(['a']) }, 0).ok, false);
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: '' }, paragraphs: items(['a']) }, 0).ok, false);
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR, extra: 1 }, paragraphs: items(['a']) }, 0).ok, false);
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: [] }, 0).ok, false, 'zero items');
  const many = items(Array.from({ length: 201 }, (_, i) => `p${i}`));
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: many }, 0).ok, false, '201 items');
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: many.slice(0, 200) }, 0).ok, true, '200 items is the ceiling, not over it');
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: [{ text: 'a\tb' }] }, 0).ok, false, 'tab in text');
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: [{ text: '' }] }, 0).ok, false, 'empty text');
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: [{ text: 'a', style: '' }] }, 0).ok, false, 'empty style');
  assert.equal(pi.plan({ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: [{ text: 'a', format: { highlight: 'yellow' } }] }, 0).ok, false, 'highlight is not in the accepted subset');
});

// ---- classify / resolve -----------------------------------------------------------------------------

test('classify: APPLY when the sequence after the anchor does not match (shorter, or different text)', () => {
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`;
  assert.equal(pi.classify(obsOf(html).primary, op()).state, 'APPLY', 'nothing there yet');
  assert.equal(pi.classify(obsOf(`<p>${ANCHOR}</p><p>a</p>`).primary, op()).state, 'APPLY', 'only one of the two landed');
  assert.equal(pi.classify(obsOf(`<p>${ANCHOR}</p><p>x</p><p>y</p>`).primary, op()).state, 'APPLY', 'wrong text, never a guessed match');
});

test('classify: SATISFIED when the N blocks right after the anchor already match text, in order', () => {
  const c = pi.classify(obsOf(`<p>${ANCHOR}</p><p>a</p><p>b</p><p>${SIBLING}</p>`).primary, op());
  assert.equal(c.state, 'SATISFIED');
});

test('classify: SATISFIED only when a requested style/format ALSO already matches -- a same-text run with the wrong format is APPLY', () => {
  const withOp = op({ paragraphs: [{ text: 'a', format: { bold: true } }] });
  const matched = obsOf(`<p>${ANCHOR}</p><p>a</p>`, [{ text: 'a', format: { bold: true } }]);
  assert.equal(pi.classify(matched.primary, withOp).state, 'SATISFIED');
  const wrongFormat = obsOf(`<p>${ANCHOR}</p><p>a</p>`, [{ text: 'a', format: { bold: false } }]);
  assert.equal(pi.classify(wrongFormat.primary, withOp).state, 'APPLY');
  const unknownFormat = obsOf(`<p>${ANCHOR}</p><p>a</p>`); // no paragraphs read -> block.format is null
  assert.equal(pi.classify(unknownFormat.primary, withOp).state, 'APPLY', 'unobservable format is never a guessed match');
});

test('classify: FAIL_CLOSED when the anchor is missing or ambiguous, never guessed', () => {
  assert.equal(pi.classify(obsOf(`<p>${OTHER}</p>`).primary, op()).state, 'FAIL_CLOSED');
  assert.equal(pi.classify(obsOf(`<p>${ANCHOR}</p><p>${ANCHOR}</p>`).primary, op()).state, 'FAIL_CLOSED');
});

test('resolve / preDispatchCheck: round-trip through the real planOperations pipeline', () => {
  const planned = planOperations({ steps: [{ intent: 'insert_paragraphs', selector: { afterParagraph: ANCHOR }, paragraphs: items(['a']) }] });
  assert.equal(planned.ok, true, planned.detail);
  const o = planned.operations[0];
  const r = pi.resolve(obsOf(`<p>${ANCHOR}</p>`).primary, o.selector);
  assert.equal(r.ok, true);
  assert.equal(r.target.anchorText, ANCHOR);
  assert.equal(pi.preDispatchCheck(o, r.target).ok, true);
  assert.equal(pi.preDispatchCheck(o, { kind: 'insert_paragraphs_anchor', anchorText: 'wrong' }).ok, false);
});

// ---- blockMatches / sequenceMatches --------------------------------------------------------------

test('blockMatches: text must match exactly; a table block never matches; an empty wanted (no style/format) matches on text alone', () => {
  const block = { kind: 'paragraph', text: 'a', format: null };
  assert.equal(pi.blockMatches(block, { text: 'a' }), true);
  assert.equal(pi.blockMatches(block, { text: 'b' }), false);
  assert.equal(pi.blockMatches({ kind: 'table', text: 'a' }, { text: 'a' }), false);
  assert.equal(pi.blockMatches(null, { text: 'a' }), false);
});

test('wantedOf: folds style into the SAME key compare() reads from a block\'s own .format.style', () => {
  assert.deepEqual(pi.wantedOf({ text: 'a', style: 'Heading 1', format: { bold: true } }), { bold: true, style: 'Heading 1' });
  assert.deepEqual(pi.wantedOf({ text: 'a' }), {});
});

// ---- verify ---------------------------------------------------------------------------------------

test('verifyPrimitive: ok only when the after-state carries the exact sequence right after the (still-unique) anchor', () => {
  const before = obsOf(`<p>${ANCHOR}</p>`);
  const good = pi.verifyPrimitive(op(), before, obsOf(`<p>${ANCHOR}</p><p>a</p><p>b</p>`));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  const bad = pi.verifyPrimitive(op(), before, obsOf(`<p>${ANCHOR}</p><p>a</p>`));
  assert.equal(bad.ok, false);
});

// ---- verifyPreservation (L3) -----------------------------------------------------------------------

test('verifyPreservation: OK when the N new blocks are exactly the requested sequence and the prefix + rest are unchanged', () => {
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const good = pi.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>a</p><p>b</p><p>${SIBLING}</p>`));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

test('verifyPreservation: FAILS if a new block is wrong/missing, if the count is not exactly +N, or if anything outside the N new blocks changed', () => {
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const shortRun = pi.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>a</p><p>${SIBLING}</p>`));
  assert.equal(shortRun.ok, false, 'only one of the two landed');
  const wrongText = pi.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>x</p><p>y</p><p>${SIBLING}</p>`));
  assert.equal(wrongText.ok, false);
  const prefixChanged = pi.verifyPreservation(op(), before, obsOf(`<p>${OTHER} CHANGED</p><p>${ANCHOR}</p><p>a</p><p>b</p><p>${SIBLING}</p>`));
  assert.equal(prefixChanged.ok, false, 'a block before the anchor changed');
  const restChanged = pi.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>a</p><p>b</p><p>${SIBLING} CHANGED</p>`));
  assert.equal(restChanged.ok, false, 'a block after the new run changed');
  const extra = pi.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>a</p><p>b</p><p>EXTRA</p><p>${SIBLING}</p>`));
  assert.equal(extra.ok, false, 'an extra, unrequested block appeared');
});

test('verifyPreservation: style/format on a new paragraph is part of the "exactly the requested sequence" check', () => {
  const withOp = op({ paragraphs: [{ text: 'a', format: { bold: true } }] });
  const before = obsOf(`<p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const good = pi.verifyPreservation(withOp, before, obsOf(`<p>${ANCHOR}</p><p>a</p><p>${SIBLING}</p>`, [{ text: 'a', format: { bold: true } }]));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  const missingFormat = pi.verifyPreservation(withOp, before, obsOf(`<p>${ANCHOR}</p><p>a</p><p>${SIBLING}</p>`));
  assert.equal(missingFormat.ok, false, 'the format was never observed as applied');
});

// ---- verifyPreservation: the anchor-only border-alt equivalence (LIVE, gh271 representative document) -----
// The engine's re-save flips the ANCHOR paragraph's "no border" spelling between
// "border-<side>:none" and "mso-border-<side>-alt:none" -- the SAME pair GH #308 measured for
// insert_table. Scoped to that ONE block; it must not hide a real border change anywhere.

const BEFORE_STYLE = 'margin-top:0pt;margin-bottom:10pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none';
const AFTER_STYLE = 'margin-top:0pt;margin-bottom:10pt;border:none;mso-border-left-alt:none;mso-border-top-alt:none;mso-border-right-alt:none;mso-border-bottom-alt:none;mso-border-between:none';

test('verifyPreservation: the exact measured before/after anchor style pair is the SAME anchor (border-alt equivalence) -- PASS', () => {
  const before = obsOf(`<p>${OTHER}</p><p style="${BEFORE_STYLE}">${ANCHOR}</p><p>${SIBLING}</p>`);
  const good = pi.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p style="${AFTER_STYLE}">${ANCHOR}</p><p>a</p><p>b</p><p>${SIBLING}</p>`));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

test('verifyPreservation: a REAL border change on the anchor (not the measured alt-spelling swap) still FAILS', () => {
  const before = obsOf(`<p>${OTHER}</p><p style="${BEFORE_STYLE}">${ANCHOR}</p><p>${SIBLING}</p>`);
  const bad = pi.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p style="border-left:1px solid">${ANCHOR}</p><p>a</p><p>b</p><p>${SIBLING}</p>`));
  assert.equal(bad.ok, false);
});

test('verifyPreservation: the SAME before/after swap on a NON-anchor block still FAILS', () => {
  const before = obsOf(`<p style="${BEFORE_STYLE}">${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const bad = pi.verifyPreservation(op(), before, obsOf(`<p style="${AFTER_STYLE}">${OTHER}</p><p>${ANCHOR}</p><p>a</p><p>b</p><p>${SIBLING}</p>`));
  assert.equal(bad.ok, false);
});

test('wholeTaskChecks: one named check, true only when the exact sequence is in place', () => {
  const pass = pi.wholeTaskChecks(op(), obsOf(`<p>${ANCHOR}</p><p>a</p><p>b</p>`));
  assert.equal(pass.length, 1);
  assert.equal(pass[0].ok, true);
  const fail = pi.wholeTaskChecks(op(), obsOf(`<p>${ANCHOR}</p>`));
  assert.equal(fail[0].ok, false);
});

test('producedText: a batch of N paragraphs has no single semantic result a later step can ref', () => {
  assert.equal(pi.producedText({}), null);
});
