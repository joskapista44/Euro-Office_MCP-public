'use strict';

const { norm } = require('../inventory.cjs');
const { resolveStableTarget } = require('../resolver.cjs');

/**
 * DOCX `section-break-insert` family (GH #313, NEW_CAPABILITY): insert_section_break.
 *
 * No source intent: set_page_setup (page-setup.cjs) addresses an EXISTING section by {index, of};
 * it never had to CREATE one, because every migrated source document had a fixed section count.
 * Without a way to insert a break, a 2-3 column block confined to PART of a page cannot be built --
 * columns can only ever be set for a whole, pre-existing section.
 *
 * LIVE-measured (test/live/docx-gh313-section-break-probe.live.cjs, run
 * gh313-sectionprobe-1793165686897-6de018):
 *   - the anchor paragraph itself has no InsertSection method (SURFACE: GetSection/SetSection only);
 *     `doc.CreateSection(anchorParagraph)` is the call that exists and works: it returns a NEW
 *     section object, the section count goes from N to N+1, and the stored document.xml carries
 *     the new <w:sectPr> embedded in the ANCHOR paragraph's own <w:pPr> -- i.e. the RETURNED section
 *     is the one the anchor now CLOSES (call it section k), not the one that starts right after it.
 *   - SEMANTIC TRAP (named by the LIVE reviewer): OOXML's <w:type> inside a section's own sectPr
 *     describes how THAT section BEGINS. The card's own "szakasztores a horgony UTAN, next page /
 *     continuous" request is therefore a property of the NEXT section (k+1, the one starting right
 *     after the anchor), not of the section CreateSection returned (k). Setting SetType on the
 *     returned object -- the probe's own first attempt -- writes the type on the WRONG section.
 *   - the fix does not need `doc.GetSections()` indexing at all (and deliberately avoids it: this
 *     driver's own read code already documents that a section object is a NEW WRAPPER PER CALL,
 *     so matching CreateSection's return against a separately-obtained GetSections() array entry by
 *     identity is not safe). Instead: after CreateSection(anchor), the paragraph immediately
 *     following the anchor in body order is re-located (by text, not identity -- the SAME
 *     anchor-lookup convention every other family uses) and ITS OWN `GetSection()` is called
 *     directly -- one getter call, its result used immediately, never compared against anything
 *     obtained from a different call. That section is k+1; SetType is called on it.
 *   - columns are NOT part of this intent: the card's own point 4 -- set_page_setup already
 *     addresses a section by {index, of}; once the break exists (and the section count is one
 *     higher), the SAME intent sets the new section's columns. This keeps the two concerns
 *     (creating a boundary vs. configuring what is inside it) in the families that already own them.
 *
 * Preservation placement: PRIMITIVE_INTEGRAL. A section break adds NO new body element (it lives in
 * the anchor paragraph's own metadata); every body block (text, order, kind) is unchanged. Exactly
 * one BRAND NEW section appears; every PRE-EXISTING section's own fields (including its own type)
 * are unchanged -- verified position-agnostically: removing exactly one entry from the after-list
 * must reproduce the before-list exactly, in order (no section index bookkeeping needed).
 *
 * R1 REVIEW CORRECTION (LIVE-measured, zero): v1 originally tried to
 * correlate a body position to a section index via a NEW docx-driver.cjs inventory field
 * (closingBlockIndex, built from paragraph.GetSection() comparisons). Its own self-validating
 * sanity check worked exactly as designed: on the real DS, paragraph.GetSection() returns a NEW,
 * never-equal wrapper even between adjacent paragraphs in the SAME section, so the check degraded
 * to UNOBSERVABLE on every single read -- the field was REMOVED. v2 (this file, interim) rebuilt
 * classify/L1 on section COUNT and FIELD CONTENT only, with no idempotency and only a
 * last-section-shaped L1 -- documented as known limitations pending a follow-up probe.
 *
 * R2 REVIEW RESOLUTION (LIVE-measured, zero, mini-probe PASS:
 * test/live/docx-gh313-section-index-probe.live.cjs): the follow-up identity/surface probe found no
 * stable per-section VALUE identifier (GetId/GetIndex missing; ToJSON is just the section's own
 * page-setup fields, which collide across identically-configured sections) and confirmed `===` on
 * separately-obtained `paragraph.GetSection()` results is unsafe even for the SAME paragraph called
 * twice -- but the prototype carries `GetPrevious()`/`GetNext()`. `sectionIndexOf(sec) = length of
 * the GetPrevious() chain to null` is a plain INTEGER, computed with ONE call per paragraph and no
 * cross-call comparison, confirmed monotonic and in agreement with `doc.GetSections()` array order
 * and the stored sectPr census (3/2/1 columns across 3 sections, index 0/1/2). docx-driver.cjs now
 * reads this per paragraph as `sectionIndex` (see its own `sectionIndexOf` comment); inventory.cjs
 * attaches it to every paragraph/heading block by the same text-match convention as every other
 * per-paragraph field. classify/L1/wholeTaskChecks below are rebuilt on it:
 *   - classify regains idempotency: if the anchor ALREADY closes a section boundary (its own
 *     sectionIndex differs from the next indexed block's) AND that next section already carries the
 *     requested type -> SATISFIED. If it already closes a boundary with a DIFFERENT type, that is a
 *     named FAIL_CLOSED (not silently re-applied, not silently accepted) -- per instruction: flag
 *     it rather than pick a resolution.
 *   - L1 (verifyPrimitive) is now anchor-specific, not last-section-shaped: the anchor's own
 *     sectionIndex in the AFTER state, plus the next indexed block's sectionIndex being EXACTLY one
 *     more, plus that section's type matching the request. This also correctly handles an anchor
 *     that is NOT in the (then-)last section -- the v2 known limitation.
 *   - L3 (verifyPreservation) is UNCHANGED -- zero confirmed the position-agnostic remove-one-and-
 *     compare design needs no section-index bookkeeping and should stay exactly as it is.
 *   - wholeTaskChecks is re-aimed at the section right after the anchor (consistent with L1), not
 *     "the last section" (which a non-last-section anchor would make meaningless).
 * If a block's `sectionIndex` is unreadable (no unique live match, or the driver read reported
 * UNOBSERVABLE), classify and L1 both fail toward the SAFE side: classify refuses with
 * insert-section-break-section-index-unreadable (neither a guessed SATISFIED nor an APPLY that could
 * insert a second break on a retry), L1 fails `ok: false` (never guesses a pass).
 *
 * Scope (MI NEM TARTOZIK BELE, per the card): no column write (set_page_setup's job, after this
 * op), no LIVE run by this agent, no main-merge.
 */

const TARGET_KINDS = Object.freeze(['paragraph', 'heading']);
const TYPES = Object.freeze(['nextPage', 'continuous']);

function plan(step, index) {
  if (step.intent !== 'insert_section_break') return { ok: false, detail: `step ${index}: ${step.intent} is not a section-break-insert intent` };
  const sel = step.selector;
  if (!sel || typeof sel !== 'object' || Object.keys(sel).sort().join(',') !== 'afterParagraph') {
    return { ok: false, detail: `step ${index}: insert_section_break requires a selector of exactly { afterParagraph }` };
  }
  if (typeof sel.afterParagraph !== 'string' || !norm(sel.afterParagraph)) {
    return { ok: false, detail: `step ${index}: insert_section_break selector.afterParagraph must be a non-empty string` };
  }
  if (!TYPES.includes(step.type)) {
    return { ok: false, detail: `step ${index}: insert_section_break requires type to be one of ${TYPES.join(', ')}` };
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      selector: Object.freeze({ afterParagraph: sel.afterParagraph }),
      preservationMode: 'PRIMITIVE_INTEGRAL',
      sectionType: step.type,
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

// ---- identity -------------------------------------------------------------------------------------

const blocksOf = primary => (primary && primary.blocks) || [];
const sectionsOf = primary => {
  const s = primary && primary.document && primary.document.sections;
  return Array.isArray(s) ? s : null;
};

/** The unique body-level paragraph/heading named by `text`, never a table or list. */
function resolveAnchor(primary, text) {
  const r = resolveStableTarget(primary, { text });
  if (!r.ok) return r;
  if (!TARGET_KINDS.includes(r.target.kind)) {
    return { ok: false, target: null, selector: r.selector, ambiguity: 'UNSUPPORTED_TARGET_KIND', kind: r.target.kind };
  }
  return r;
}

/**
 * The next block AFTER `anchorBlock` (by reference, since resolveAnchor's target IS the live block
 * object in `primary.blocks`) that carries a readable, integer `sectionIndex` -- i.e. the next
 * paragraph/heading in body order whose section membership we can actually name. Returns null if
 * the anchor is the last such block, or cannot be located in `primary.blocks` at all.
 */
function nextIndexedBlockAfter(primary, anchorBlock) {
  const blocks = blocksOf(primary);
  const i = blocks.indexOf(anchorBlock);
  if (i < 0) return null;
  for (let j = i + 1; j < blocks.length; j += 1) {
    if (Number.isInteger(blocks[j].sectionIndex)) return blocks[j];
  }
  return null;
}

// ---- classify / resolve --------------------------------------------------------------------------

function classify(primary, op) {
  const anchor = resolveAnchor(primary, op.selector.afterParagraph);
  if (!anchor.ok) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `insert-section-break-anchor-${String(anchor.ambiguity).toLowerCase().replace(/_/g, '-')}` } };
  }
  const anchorIdx = anchor.target.sectionIndex;
  const nextBlock = nextIndexedBlockAfter(primary, anchor.target);
  const nextIdx = nextBlock ? nextBlock.sectionIndex : null;
  // An unreadable anchor index cannot tell "no break yet" from "break already here": APPLY on it
  // could insert a SECOND break on a retry, so the safe side is a named refusal, not APPLY.
  if (!Number.isInteger(anchorIdx)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'insert-section-break-section-index-unreadable', anchor: anchor.target.text } };
  }
  if (!Number.isInteger(nextIdx) || nextIdx === anchorIdx) {
    // No later block with a readable index (the anchor is the last one), or the anchor does not
    // close a section boundary yet: no break exists here.
    return { state: 'APPLY', evidence: { anchor: anchor.target.text, type: op.sectionType } };
  }
  // The anchor already closes a section boundary (nextIdx > anchorIdx). Is the section that starts
  // right after it already the requested type?
  const sections = sectionsOf(primary);
  const nextSection = Array.isArray(sections) ? sections[nextIdx] : null;
  if (nextSection && nextSection.type === op.sectionType) {
    return { state: 'SATISFIED', evidence: { anchor: anchor.target.text, type: op.sectionType, sectionIndex: nextIdx } };
  }
  // Closes a boundary, but with a DIFFERENT type -- flagged, never silently re-applied or accepted.
  return {
    state: 'FAIL_CLOSED',
    evidence: {
      reason: 'insert-section-break-anchor-already-closed-with-different-type',
      anchorSectionIndex: anchorIdx, nextSectionIndex: nextIdx, nextSectionType: nextSection && nextSection.type,
    },
  };
}

function resolve(primary, selector) {
  const sel = selector || {};
  const anchor = resolveAnchor(primary, sel.afterParagraph);
  if (!anchor.ok) return { ok: false, target: null, selector, ambiguity: anchor.ambiguity };
  return { ok: true, target: { kind: 'insert_section_break_anchor', anchorText: norm(anchor.target.text) }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && target.kind === 'insert_section_break_anchor' && target.anchorText === norm(op.selector.afterParagraph)
    ? { ok: true }
    : { ok: false, dispatched: false, detail: 'insert-section-break-anchor-target-unresolved' };
}

// ---- lower ----------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({ family: 'section-break-insert', action: 'insert', anchorText: target.anchorText, sectionType: op.sectionType });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (GH #313 LIVE PASS). Re-locates the anchor by the same text-matching convention
 * every other DOCX write family uses (re-checked for uniqueness inside the sandbox), calls
 * doc.CreateSection(anchor) to create the break, then re-locates the paragraph immediately
 * following the anchor and calls ITS OWN GetSection() directly -- never doc.GetSections() indexing
 * -- to reach the NEW section and set its type (the fix for the measured semantic trap above). Any
 * guard miss returns unchanged; the break may still have been created with its inherited (default)
 * type, which L1 will correctly report as not matching the request, never silently accepted.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  if (typeof doc.CreateSection !== 'function') return;
  const textOpts = { Numbering: false };
  const lookupTextOf = p => String(p.GetText(textOpts) || '').replace(/[\\r\\n]+$/, '');
  const n = doc.GetElementsCount();
  let anchorIndex = -1;
  for (let i = 0; i < n; i += 1) {
    const e = doc.GetElement(i);
    if (e && typeof e.GetClassType === 'function' && e.GetClassType() === 'paragraph' && lookupTextOf(e) === m.anchorText) {
      if (anchorIndex >= 0) return; // ambiguous at dispatch time: refuse rather than guess
      anchorIndex = i;
    }
  }
  if (anchorIndex < 0) return;
  const anchor = doc.GetElement(anchorIndex);
  doc.CreateSection(anchor);
  const nextEl = doc.GetElement(anchorIndex + 1);
  if (!nextEl || typeof nextEl.GetSection !== 'function') return;
  const newSection = nextEl.GetSection();
  if (newSection && typeof newSection.SetType === 'function') newSection.SetType(m.sectionType);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/**
 * L1 (R2, anchor-specific -- see the file header): in the AFTER state, the section count rose by
 * exactly one, the anchor's own sectionIndex and the next indexed block's sectionIndex differ by
 * EXACTLY one (the anchor now closes a brand-new boundary, not merely some pre-existing one further
 * along), and that next section carries the requested type. Unlike v2's "last section" shape, this
 * is correct regardless of whether the anchor sits in the document's final section.
 */
function verifyPrimitive(op, before, after) {
  const anchorAfter = resolveAnchor(after && after.primary, op.selector.afterParagraph);
  if (!anchorAfter.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the after-state' } };
  const beforeSections = sectionsOf(before && before.primary);
  const afterSections = sectionsOf(after && after.primary);
  if (!beforeSections || !afterSections) return { ok: false, evidence: { reason: 'sections-unreadable' } };
  const sectionCountIncreased = afterSections.length === beforeSections.length + 1;

  const anchorIdx = anchorAfter.target.sectionIndex;
  const nextBlock = nextIndexedBlockAfter(after.primary, anchorAfter.target);
  const nextIdx = nextBlock ? nextBlock.sectionIndex : null;
  const anchorClosesNewBoundary = Number.isInteger(anchorIdx) && Number.isInteger(nextIdx) && nextIdx === anchorIdx + 1;
  const nextSection = anchorClosesNewBoundary ? afterSections[nextIdx] : null;
  const nextHasRequestedType = !!nextSection && nextSection.type === op.sectionType;

  const ok = sectionCountIncreased && anchorClosesNewBoundary && nextHasRequestedType;
  return {
    ok,
    evidence: {
      sectionCountIncreased, anchorClosesNewBoundary, nextHasRequestedType,
      anchorSectionIndex: anchorIdx, nextSectionIndex: nextIdx, nextSectionType: nextSection && nextSection.type,
      sectionCountBefore: beforeSections.length, sectionCountAfter: afterSections.length,
    },
  };
}

/** Canonical, order-sensitive key for a body block (table-insert.cjs's own shape, reused). */
const canonicalBlock = b => (b.kind === 'table'
  ? ['table', (b.rows || []).map(row => row.map(c => [norm(c.text), c.span, c.rowspan]))]
  : [b.kind, b.level || null, norm(b.text), b.style || null]);

/** Every page-setup field, INCLUDING type -- a pre-existing section's type must not silently
 * change any more than its size/margins/columns/startPageNumber. */
const sectionFields = s => JSON.stringify({ pageWidth: s.pageWidth, pageHeight: s.pageHeight, margins: s.margins, columnsCount: s.columnsCount, columnsSpaces: s.columnsSpaces, startPageNumber: s.startPageNumber, titlePage: s.titlePage, evenAndOdd: s.evenAndOdd, type: s.type });

/**
 * L3, PRIMITIVE_INTEGRAL (UNCHANGED since R1 -- zero confirmed this design needs no section-index
 * bookkeeping and should stay as-is): every body block unchanged (a break adds no visible element);
 * EXACTLY one brand-new section appears, and removing it (at WHATEVER its index turns out to be)
 * from the after-list reproduces the before-list exactly, in order -- position-agnostic by
 * construction. (This also catches the R1-review semantic-trap shape: if the type lands on a
 * PRE-EXISTING section instead of the new one, that section's own `type` field now differs from its
 * own before-value, and no single removal from after can reproduce before.)
 */
function verifyPreservation(op, before, after) {
  const beforeBlocks = blocksOf(before.primary);
  const afterBlocks = blocksOf(after.primary);
  const bodySame = JSON.stringify(beforeBlocks.map(canonicalBlock)) === JSON.stringify(afterBlocks.map(canonicalBlock));

  const beforeSections = sectionsOf(before.primary);
  const afterSections = sectionsOf(after.primary);
  if (!beforeSections || !afterSections) return { ok: false, evidence: { reason: 'sections-unreadable' } };

  const beforeKeys = beforeSections.map(sectionFields);
  const afterKeys = afterSections.map(sectionFields);
  let removedOneMatches = false;
  let newSectionIndex = -1;
  if (afterKeys.length === beforeKeys.length + 1) {
    for (let r = 0; r < afterKeys.length; r += 1) {
      const candidate = afterKeys.filter((_, i) => i !== r);
      if (JSON.stringify(candidate) === JSON.stringify(beforeKeys)) { removedOneMatches = true; newSectionIndex = r; break; }
    }
  }

  const ok = bodySame && removedOneMatches;
  return { ok, evidence: { bodySame, removedOneMatches, newSectionIndex, sectionCountBefore: beforeSections.length, sectionCountAfter: afterSections.length } };
}

/** R2: re-aimed at the section right after the anchor (consistent with L1), not "the last
 * section" -- a non-last-section anchor would make that check meaningless. */
function wholeTaskChecks(op, finalObservation) {
  const anchor = resolveAnchor(finalObservation && finalObservation.primary, op.selector.afterParagraph);
  const sections = sectionsOf(finalObservation && finalObservation.primary);
  let ok = false;
  if (anchor.ok && Array.isArray(sections)) {
    const nextBlock = nextIndexedBlockAfter(finalObservation.primary, anchor.target);
    const nextIdx = nextBlock ? nextBlock.sectionIndex : null;
    const nextSection = Number.isInteger(nextIdx) ? sections[nextIdx] : null;
    ok = !!nextSection && nextSection.type === op.sectionType;
  }
  return [{ name: `${op.id}: the section right after the anchor carries the requested ${op.sectionType} type`, ok }];
}

// refs: none -- a section break is not a text unit a later step can name
function producedText() {
  return null;
}

module.exports = {
  plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  resolveAnchor, nextIndexedBlockAfter, canonicalBlock, sectionFields, TYPES,
};
