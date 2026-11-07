'use strict';

/**
 * GH #313 (write side, offline): the `insert_section_break` family
 * (src/adapters/docx/families/section-break-insert.cjs). plan/classify/resolve/verify against
 * fabricated `normalizeInventory` observations -- no editor, no VM. The commandBody string itself
 * is covered by its own LIVE probe (test/live/docx-gh313-section-break-probe.live.cjs), not here.
 *
 * R2 review correction (internal tracker/#ref, mini-probe PASS): classify/L1/wholeTaskChecks are
 * now anchor-specific, using the per-paragraph `sectionIndex` field (value-based GetPrevious()-
 * chain length, docx-driver.cjs's sectionIndexOf) that inventory.cjs attaches to every
 * paragraph/heading block. `obsOf`'s third argument simulates it: a { text: sectionIndex } map,
 * turned into the `payload.paragraphs` entries attachProperties() text-matches against -- the SAME
 * mechanism the real driver read uses, just fabricated here instead of VM-executed.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const sbi = require('../../src/adapters/docx/families/section-break-insert.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');

const ANCHOR = 'GH313 anchor';
const SIBLING = 'GH313 sibling';
const OTHER = 'GH313 other';

/** One page-setup-shaped section fixture, every field defaulted so a test only states what it cares about. */
function sec(fields = {}) {
  return {
    pageWidth: 11906, pageHeight: 16838,
    margins: { top: 1134, right: 850, bottom: 1134, left: 1701 },
    columnsCount: 1, columnsSpaces: [], startPageNumber: -1,
    type: 'nextPage', titlePage: false, evenAndOdd: false,
    ...fields,
  };
}
/** `idx`, when given, is a { paragraphText: sectionIndex } map -- simulates the driver's per-
 * paragraph `sectionIndex` read via the SAME text-match attach path normalizeInventory uses for
 * every other per-paragraph field. Omitted entirely: every block's sectionIndex stays null
 * (simulates an unreadable/never-measured signal). */
function obsOf(html, sections, idx) {
  const paragraphs = idx ? Object.entries(idx).map(([text, sectionIndex]) => ({ text, sectionIndex })) : undefined;
  return { primary: normalizeInventory({ html, document: { sections: sections.map((s, i) => ({ ...s, index: i })) }, ...(paragraphs ? { paragraphs } : {}) }) };
}

const noBreak = () => `<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`;

// ---- plan -----------------------------------------------------------------------------------------

test('plan: accepts { afterParagraph } + a measured type', () => {
  for (const t of ['nextPage', 'continuous']) {
    const p = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: t }, 0);
    assert.equal(p.ok, true, p.detail);
    assert.equal(p.operation.preservationMode, 'PRIMITIVE_INTEGRAL');
    assert.equal(p.operation.sectionType, t);
  }
});

test('plan: refuses a wrong-shaped selector and an unmeasured/missing type', () => {
  assert.equal(sbi.plan({ intent: 'insert_section_break', selector: {}, type: 'nextPage' }, 0).ok, false);
  assert.equal(sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: '' }, type: 'nextPage' }, 0).ok, false);
  assert.equal(sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR, extra: 1 }, type: 'nextPage' }, 0).ok, false);
  for (const bad of [undefined, '', 'evenPage', 'oddPage', 'nextColumn', 'NEXTPAGE']) {
    assert.equal(sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: bad }, 0).ok, false, `type=${bad}`);
  }
});

// ---- classify / resolve ----------------------------------------------------------------------------

test('classify: APPLY when the anchor does not yet close a section boundary (its sectionIndex matches what follows)', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'nextPage' }, 0).operation;
  const idx = { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 0 };
  const c = sbi.classify(obsOf(noBreak(), [sec()], idx).primary, op);
  assert.equal(c.state, 'APPLY');
});

test('classify: FAIL_CLOSED (neither SATISFIED nor APPLY) when the anchor sectionIndex is unreadable', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'nextPage' }, 0).operation;
  const c = sbi.classify(obsOf(noBreak(), [sec()]).primary, op); // no idx given -> sectionIndex stays null everywhere
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'insert-section-break-section-index-unreadable');
});

test('classify: SATISFIED -- the anchor already closes a section boundary, and the next section already carries the requested type', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'continuous' }, 0).operation;
  const idx = { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 1 };
  const c = sbi.classify(obsOf(noBreak(), [sec(), sec({ type: 'continuous' })], idx).primary, op);
  assert.equal(c.state, 'SATISFIED', JSON.stringify(c.evidence));
  assert.equal(c.evidence.sectionIndex, 1);
});

test('classify: FAIL_CLOSED, named -- the anchor already closes a section boundary, but with a DIFFERENT type (never silently re-applied or accepted)', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'continuous' }, 0).operation;
  const idx = { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 1 };
  const c = sbi.classify(obsOf(noBreak(), [sec(), sec({ type: 'nextPage' })], idx).primary, op);
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'insert-section-break-anchor-already-closed-with-different-type');
  assert.equal(c.evidence.nextSectionType, 'nextPage');
});

test('classify: FAIL_CLOSED when the anchor is missing or ambiguous, never guessed', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: 'nowhere' }, type: 'nextPage' }, 0).operation;
  assert.equal(sbi.classify(obsOf(noBreak(), [sec()]).primary, op).state, 'FAIL_CLOSED');
  const dup = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'nextPage' }, 0).operation;
  assert.equal(sbi.classify(obsOf(`<p>${ANCHOR}</p><p>${ANCHOR}</p>`, [sec()]).primary, dup).state, 'FAIL_CLOSED');
});

test('resolve / preDispatchCheck: round-trip through the real planOperations pipeline', () => {
  const planned = planOperations({ steps: [{ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'nextPage' }] });
  assert.equal(planned.ok, true, planned.detail);
  const op = planned.operations[0];
  const r = sbi.resolve(obsOf(noBreak(), [sec()]).primary, op.selector);
  assert.equal(r.ok, true);
  assert.equal(r.target.anchorText, ANCHOR);
  assert.equal(sbi.preDispatchCheck(op, r.target).ok, true);
  assert.equal(sbi.preDispatchCheck(op, { kind: 'insert_section_break_anchor', anchorText: 'wrong' }).ok, false);
});

// ---- verifyPrimitive (L1) ---------------------------------------------------------------------------

test('verifyPrimitive: ok when the anchor closes a boundary exactly one past its own index, AND that section carries the requested type', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'continuous' }, 0).operation;
  const before = obsOf(noBreak(), [sec()], { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 0 });
  const after = obsOf(noBreak(), [sec(), sec({ type: 'continuous' })], { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 1 });
  const good = sbi.verifyPrimitive(op, before, after);
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

test('verifyPrimitive: the semantic-trap regression -- type set on an EXISTING (non-final) section, not the final one -- must FAIL', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'continuous' }, 0).operation;
  const before = obsOf(noBreak(), [sec()], { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 0 });
  // the type landed on sections[0] (the anchor's OWN, pre-existing section) instead of sections[1]
  // (the new one, which stays at its inherited default 'nextPage')
  const after = obsOf(noBreak(), [sec({ type: 'continuous' }), sec({ type: 'nextPage' })], { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 1 });
  const trapped = sbi.verifyPrimitive(op, before, after);
  assert.equal(trapped.ok, false, 'type set on the wrong (non-final) section must not verify as ok');
});

test('verifyPrimitive: no new section at all -- FAILS', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'continuous' }, 0).operation;
  const before = obsOf(noBreak(), [sec()], { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 0 });
  assert.equal(sbi.verifyPrimitive(op, before, before).ok, false);
});

test('verifyPrimitive: anchor NOT in the last pre-existing section -- verified via the NEXT paragraph, not "the last section" (v2 known limitation, now fixed)', () => {
  const MID = 'GH313 mid';
  const TAIL = 'GH313 tail';
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p><p>${MID}</p><p>${SIBLING}</p><p>${TAIL}</p>`;
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'continuous' }, 0).operation;
  // before: 3 pre-existing sections (0,1,2); the anchor sits in section 0, which is followed by
  // sections 1 and 2 -- i.e. NOT the last section.
  const before = obsOf(html, [sec({ columnsCount: 1 }), sec({ columnsCount: 2 }), sec({ columnsCount: 4 })],
    { [OTHER]: 0, [ANCHOR]: 0, [MID]: 0, [SIBLING]: 1, [TAIL]: 2 });
  // after: a NEW section (continuous) inserted right after the anchor; MID now belongs to it;
  // everything that used to be section 1/2 shifts to 2/3.
  const after = obsOf(html, [sec({ columnsCount: 1 }), sec({ columnsCount: 1, type: 'continuous' }), sec({ columnsCount: 2 }), sec({ columnsCount: 4 })],
    { [OTHER]: 0, [ANCHOR]: 0, [MID]: 1, [SIBLING]: 2, [TAIL]: 3 });
  const good = sbi.verifyPrimitive(op, before, after);
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

// ---- verifyPreservation (L3) ------------------------------------------------------------------------
// UNCHANGED since the R1 review (zero confirmed, #ref): position-agnostic, no sectionIndex needed.

test('verifyPreservation: single pre-existing section -- body unchanged, exactly one new section appended, reproduces before when removed', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'nextPage' }, 0).operation;
  const before = obsOf(noBreak(), [sec({ columnsCount: 1, type: 'continuous' })]);
  const after = obsOf(noBreak(), [sec({ columnsCount: 1, type: 'continuous' }), sec({ columnsCount: 1, type: 'nextPage' })]);
  const good = sbi.verifyPreservation(op, before, after);
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  assert.equal(good.evidence.newSectionIndex, 1);
});

test('verifyPreservation: FAILS if the body text/order changed (a section break must add no visible block)', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'nextPage' }, 0).operation;
  const before = obsOf(noBreak(), [sec()]);
  const after = obsOf(`<p>${OTHER} CHANGED</p><p>${ANCHOR}</p><p>${SIBLING}</p>`, [sec(), sec({ type: 'nextPage' })]);
  assert.equal(sbi.verifyPreservation(op, before, after).ok, false);
});

test('verifyPreservation: FAILS if a PRE-EXISTING section\'s own fields changed as a side effect', () => {
  // the pre-existing section uses a non-default type/columns combo so it cannot be confused
  // with the newly-inserted one (which legitimately carries the op's requested type).
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'nextPage' }, 0).operation;
  const before = obsOf(noBreak(), [sec({ columnsCount: 1, type: 'continuous' })]);
  // index0 = the pre-existing section, corrupted (columns silently 1 -> 3); index1 = the real
  // new section (requested type, but SAME columns as before -- so it can't masquerade as "before").
  const after = obsOf(noBreak(), [sec({ columnsCount: 3, type: 'continuous' }), sec({ columnsCount: 1, type: 'nextPage' })]);
  assert.equal(sbi.verifyPreservation(op, before, after).ok, false);
});

test('verifyPreservation: position-agnostic -- the new section can land anywhere in the after-list, not only at the end', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'continuous' }, 0).operation;
  const before = obsOf(noBreak(), [sec({ columnsCount: 1 }), sec({ columnsCount: 4, startPageNumber: 9 })]);
  // the new section lands BETWEEN the two pre-existing ones, not at the end
  const after = obsOf(noBreak(), [sec({ columnsCount: 1 }), sec({ columnsCount: 2, type: 'continuous' }), sec({ columnsCount: 4, startPageNumber: 9 })]);
  const good = sbi.verifyPreservation(op, before, after);
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  assert.equal(good.evidence.newSectionIndex, 1);
});

test('verifyPreservation: FAILS if a later, untouched section\'s own fields changed too (not just the one new insertion)', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'continuous' }, 0).operation;
  const before = obsOf(noBreak(), [sec({ columnsCount: 1 }), sec({ columnsCount: 4 })]);
  const after = obsOf(noBreak(), [sec({ columnsCount: 1 }), sec({ columnsCount: 2, type: 'continuous' }), sec({ columnsCount: 2 })]); // last section's columns silently changed 4 -> 2
  assert.equal(sbi.verifyPreservation(op, before, after).ok, false);
});

test('verifyPreservation: FAILS if the section count did not increase by exactly one', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'nextPage' }, 0).operation;
  const before = obsOf(noBreak(), [sec()]);
  assert.equal(sbi.verifyPreservation(op, before, before).ok, false, 'no new section');
  const twoNew = obsOf(noBreak(), [sec(), sec({ type: 'nextPage' }), sec({ type: 'continuous' })]);
  assert.equal(sbi.verifyPreservation(op, before, twoNew).ok, false, 'two new sections');
});

// ---- wholeTaskChecks (L2) ---------------------------------------------------------------------------

test('wholeTaskChecks: one named check, true only when the section right after the anchor carries the requested type', () => {
  const op = sbi.plan({ intent: 'insert_section_break', selector: { afterParagraph: ANCHOR }, type: 'continuous' }, 0).operation;
  const idx = { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 1 };
  const pass = sbi.wholeTaskChecks(op, obsOf(noBreak(), [sec(), sec({ type: 'continuous' })], idx));
  assert.equal(pass.length, 1);
  assert.equal(pass[0].ok, true);
  // the anchor does not close any boundary here (sectionIndex matches what follows) -- no section
  // "right after" the anchor can be identified, so the check must be false
  const fail = sbi.wholeTaskChecks(op, obsOf(noBreak(), [sec()], { [OTHER]: 0, [ANCHOR]: 0, [SIBLING]: 0 }));
  assert.equal(fail[0].ok, false);
});

test('producedText: a section break is not a text unit a later step can ref', () => {
  assert.equal(sbi.producedText({}), null);
});
