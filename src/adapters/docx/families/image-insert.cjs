'use strict';

const { norm, textUnits, tablesContent } = require('../inventory.cjs');
const { resolveStableTarget } = require('../resolver.cjs');
const { imageRows } = require('./image.cjs');
const { canonicalBlock } = require('./table-insert.cjs');
const { otherParts, paragraphDescriptor } = require('./header-footer.cjs');

/**
 * DOCX `image-insert` family (GH #306, NEW_CAPABILITY): insert_image.
 *
 * `families/image.cjs` only resizes/re-wraps an EXISTING image (GH #268c) -- this family creates
 * a NEW one. Same split as `table-insert.cjs` vs the table-cell/table-row/... families.
 *
 * LIVE-measured (test/live/docx-gh306-insert-image-probe.live.cjs, g268v1's own base mechanism +
 * this probe's own measured rounds, reviewed before landing -- commit 749a4f607):
 *   - `Api.CreateImage(dataUri, widthEmu, heightEmu)` + `paragraph.AddDrawing(img)` lands an INLINE
 *     image at the end of the TARGET paragraph's own run flow. A PNG and a JPEG data: URI both
 *     land; the saved `word/media/*` file's bytes sha256-match the ORIGINAL input exactly.
 *   - `doc.AddElement(anchorIndex + 1, newParagraph)` (the SAME mechanism GH #308/`table-insert.cjs`
 *     already measured for a table) lands a NEW, otherwise-empty paragraph carrying the image
 *     immediately after the anchor -- the anchor's own paragraph is untouched.
 *   - the anchor's body-level INDEX is found by the SAME text-matching walk every other DOCX write
 *     family uses (`paragraph-lookup.cjs`'s own convention) -- an IDENTITY walk
 *     (`doc.GetElement(i) === anchor`) never matches: the wrapper object is NEW per call (measured
 *     THREE times now across #308/#313/#306; see the shared memory-lap
 *     `euro-office-builder-wrapper-identity-never-matches.md`).
 *   - the DS can fold two drawings that reference BYTE-IDENTICAL source images onto ONE
 *     `word/media/*` file (two distinct `<w:drawing>` rows, the SAME `r:embed`) -- L1 therefore
 *     never counts media FILES, only image ROWS (`read_images`) and their OWN embed's bytes.
 *   - a requested EMU size can round-trip ±1 EMU off (measured: height 2514600 requested, read
 *     back 2514599) -- the SAME tolerance the XLSX `insert_image` precedent (`drawings-insert.cjs`)
 *     already states as an open, unresolved drift, never exact-matched there either.
 *   - insertion time, in-session, for a PNG/JPEG a few hundred bytes to ~12.6 KB: 1-12 ms -- no
 *     measured ceiling was hit; a payload cap is a SCHEMA choice (below), not a measured limit.
 *
 * ADDRESSING: `{ afterParagraph }` -- the SAME unique body-level paragraph/heading contract
 * `insert_table`'s own selector already uses (a table/list, or text only inside a table cell, is
 * refused the same way `resolveStableTarget` refuses it for every other family).
 *
 * POSITION (the card's own open question, left to this probe to decide -- both routes measured
 * and both kept, since neither is a strict subset of the other):
 *   'after' (DEFAULT) -- a NEW, otherwise-empty paragraph immediately after the anchor (ROUTE B).
 *   'end'             -- appended to the END of the anchor's OWN paragraph, as a new run (ROUTE A).
 *
 * RETRY IDENTITY -- a NAMED, ACCEPTED LIMIT, not hidden (the SAME class of limit `image.cjs`'s own
 * header comment already states for SOURCE identity: "cannot tell two different-image
 * same-position-same-name images apart, and does not claim to"):
 *   position 'end'   -- SATISFIED when the anchor's OWN paragraph already carries an inline image
 *                       of the requested size (±1 EMU). Precise: the anchor's text is the row's
 *                       own `paragraph` field.
 *   position 'after' -- the paragraph THIS write creates carries NO text of its own (the body has
 *                       no name for it), so a retry falls back to POSITION: is the block right
 *                       after the anchor an EMPTY paragraph, and does the document have an inline
 *                       image of the requested size whose OWN paragraph is ALSO empty? This cannot
 *                       distinguish two DIFFERENT 'after' inserts of the SAME size at the SAME
 *                       anchor from one call that already landed, nor two 'after' inserts of
 *                       COINCIDENTALLY the same size at TWO DIFFERENT anchors in the same document.
 *                       Accepted for v1; a caller needing a hard retry guarantee for 'after' should
 *                       target a DIFFERENT anchor per call, or use 'end'.
 *
 * Preservation placement: PRIMITIVE_INTEGRAL. Every OTHER image (by its own placement/wrap/size/
 * paragraph key) is preserved as a multiset; every OTHER body block is unchanged, in sequence and
 * content; 'after' additionally requires the ONE new block to be exactly the empty paragraph this
 * write creates, immediately after the anchor, nothing else shifted.
 *
 * GH #306 part 4 -- FLOATING placement (LIVE: docx-gh306-cell-floating-image-probe.live.cjs, 10/10 PASS): an optional step-level
 * `placement { wrap, horizontal, vertical }` on the body target with position 'after' -- the EXACT shape and measured value set
 * insert_shape's own placement uses (shape-insert.cjs planPlacement, reused): square / topAndBottom / inFront / behind, horizontal
 * margin-align or page-offset, vertical paragraph- or page-offset. The ApiDrawing setters (SetWrappingStyle, SetHorAlign |
 * SetHorPosition, SetVerPosition) are called on the image BEFORE AddDrawing -- each returned true for an image too; the stored
 * drawing is a wp:anchor with the wrap element, behindDoc and positionH / positionV byte-exact, and read_images carries the same
 * ToJSON projection (placement 'floating', wrap, behindDoc, hRel / hAlign / hOffset, vRel / vAlign / vOffset). The retry identity
 * and L1 compare that placement too: an inline request is not SATISFIED by a floating image, nor the other way round. Not measured,
 * so refused: placement with position 'end' (a floating image inside an existing text paragraph), and on a header/footer target.
 *
 * GH #306 part 4 -- TABLE-CELL target (same probe): selector { tableCell: { text, tableAnchor? } } -- replace_table_cell's own cell identity
 * (table.cjs read + scopeOf: the cell's FULL text unique among the TOP-LEVEL table cells, optionally inside the ONE table a tableAnchor cell
 * names). Only a cell of plain paragraphs of runs; the image is appended inline to the cell's LAST paragraph (GetCell().GetContent()
 * .GetElement(n - 1).AddDrawing, measured: stored inside that w:tc, read_images row part 'body', paragraphIndex null, paragraph = the
 * cell paragraph's text). RETRY IDENTITY, a named limit of the same class as position 'end': SATISFIED when an inline image of the
 * requested size sits in a table-cell paragraph with the cell's last-paragraph text AND the cell reads an image marker -- a same-size
 * image already in that cell is not told apart from one this call landed. No floating placement and no nested-table cell here.
 *
 * Scope (the card's own MI NEM TARTOZIK BELE, as first landed -- GH #306 part 4 since adds floating and the table cell above): no #317 two-column
 * layout (a later round), no NC-file source (a separate jogosultsagi-ut decision -- this v1 is
 * `data:` URI only, the SAME restriction the XLSX precedent already applies), no alt-text/title
 * (no public setter, GH #268 section 3).
 *
 * GH #306b 2. resz: a HEADER/FOOTER target, `selector: { section: { index, of }, part: 'header' |
 * 'footer', variant }` -- the SAME `section: {index,of}` + `part` convention edit_header/edit_footer
 * already use (header-footer.cjs); mutually exclusive with `afterParagraph` (the body target above).
 * `variant` was restricted to 'default' ONLY when this target first landed (docx-gh306b-header-
 * image-probe.live.cjs, docx-gh306b-image-docpr-id-probe.live.cjs used the default variant only).
 * GH #296 part 3: 'title' and 'even' are now accepted -- LIVE-measured
 * (docx-gh296-hf-title-even-image-insert-probe.live.cjs, 10/10 PASS: the SAME GetHeader/GetFooter
 * (variant, false) + AddDrawing reach lands the image in the stored first-page header / even footer
 * part, read_images labels it header:<i>:title / footer:<i>:even, read_headers_footers' hasDrawing
 * sees it, the default part stays drawing-free, both flags stay on) -- but ONLY on a part that
 * ALREADY EXISTS and is ALREADY ACTIVE (title: the section's own titlePage === true; even: every
 * section's evenAndOdd === true, the flag is document-wide). An inactive variant is FAIL_CLOSED, never
 * activated: unlike edit_header/edit_footer, this family has no activation path (an image written to
 * an inactive part would be invisible, and activating it would switch which slot other pages show --
 * header-footer.cjs's own activationGuard territory, out of scope here). Every section's variant
 * state must be KNOWN (the SAME B21 discipline header-footer.cjs's variantStateGuard uses), and the
 * command re-checks the target section's flag in-sandbox through the SAME __eoB21Variant bridge READ
 * (meta.bridgeSectionVariantRead, gated on the word /3 runtime by the driver) before writing.
 *
 * MECHANISM (LIVE-measured, docx-gh306b-header-image-probe.live.cjs's own SETUP: `sec.GetHeader
 * ('default', true)` + `hdr.RemoveAllElements()` + `hdr.GetElement(0)` then `.AddText(...)` AND
 * `.AddDrawing(...)` on that SAME existing paragraph object): appending a drawing to an EXISTING
 * paragraph of a header/footer part is measured; creating a NEW paragraph inside a header/footer
 * part (an `AddElement` call on the part itself, as opposed to on the top-level `doc`) is NOT --
 * this capability therefore ONLY appends the new image to the part's OWN LAST paragraph (`AddDrawing`
 * on that paragraph, re-located via `GetElement(n-1)`, the exact object shape the probe already
 * exercised), never creating a paragraph of its own. Consequently there is no `position` choice here
 * (unlike the body target): the target part must ALREADY EXIST (`create: false` -- a missing part is
 * FAIL_CLOSED, never created implicitly, a narrower v1 than header-footer.cjs's own create-on-write;
 * the inheritance/activation machinery that create would otherwise need is out of scope here), and
 * must hold only plain paragraphs of runs (`part.simple`, the SAME shape guard header-footer.cjs's
 * own full-replacement path uses, re-checked in-sandbox) -- a part holding a table is refused rather
 * than risking a mis-addressed append.
 *
 * RETRY IDENTITY (header/footer target): SATISFIED when the part's own LAST paragraph (whatever text
 * it already holds) carries a `read_images` row of the requested size (±1 EMU), inline, whose OWN
 * `part` label matches this exact section+kind+variant (GH #306b's own part-join, LIVE PASS via
 * docPr.id) -- the SAME class of position-based limit `matchingImageAtEnd` already accepts for the
 * body target's 'end' position, scoped additionally by `part` so a same-size image that happens to
 * sit in a DIFFERENT part (or the body) is never mistaken for a match. GH #317 follow-up: the target
 * paragraph is matched by the row's own `paragraphIndex` (the same part-walk), NOT its text -- the
 * row text is GetText(), which drops a real tab, so '\tA' and 'A' collided and 'A\tB' never matched.
 *
 * Preservation (header/footer target): every image anywhere in the document, by a key that now
 * includes its OWN `part` label (so a header image is never confused with a same-shape body image),
 * is preserved except the one new image appended; the body (text units, tables, notes, bookmarks)
 * is unchanged; every OTHER section/kind/variant part is unchanged (`header-footer.cjs`'s own
 * `otherParts`, REUSED as-is -- the SAME comparison dict edit_header/edit_footer's own L3 already
 * uses, since `op.sectionIndex`/`op.part`/`op.variant` share that family's own field names); the
 * target part's own paragraph COUNT and every OTHER paragraph's text/format/runs stay exactly as
 * they were -- only the last paragraph's drawing content grows by one image, never its text.
 *
 * GH #317 (step 3): the header/footer target's `selector` gains an optional
 * `paragraph: {index, of}` -- the SAME descriptor + mismatch-is-FAIL_CLOSED discipline
 * header-footer.cjs's own paragraph mode already uses (`paragraphDescriptor`, reused, not
 * re-derived), nested inside `selector` like this family's own section/part/variant. `of` is the
 * caller's OWN SEEN paragraph count; the part actually having a different one at dispatch time is
 * a FAIL_CLOSED refusal, never a silent fallback to the last paragraph. Absent `paragraph` keeps
 * the pre-existing behaviour (always the part's own last paragraph) exactly as it was before this
 * card. MECHANISM: unchanged from the LAST-paragraph case -- `AddDrawing` on the public paragraph
 * object at the named index instead of `n - 1` (no NEW paragraph is created here either way; see
 * the MECHANISM note above for why).
 *
 * GH #317 (step 4): the header/footer target's `selector` gains an optional
 * `position: 'start'|'end'` (default `'end'` = the pre-existing, unchanged behaviour). LIVE-measured
 * (docx-gh317-image-start-of-paragraph-probe.live.cjs -- the probe's own result picked candidate A):
 * `AddDrawing` always appends to a paragraph's EXISTING run flow -- never lands a drawing before
 * text already there. `'start'` instead wraps the image in a fresh run
 * (`Api.CreateRun().AddDrawing(img)`) and inserts it at index 0 via `targetPara.AddElement(run, 0)`
 * -- `AddElement(nPos, oElement)`, the OTHER argument order, was measured a silent no-op (no throw,
 * no effect). This field lives in `selector`, a SEPARATE schema location from the body target's own
 * step-level `position` (`'after'|'end'`, a DIFFERENT contract: anchor-relative paragraph placement,
 * not within-paragraph run order) -- the two never share a schema node, so widening one can never
 * silently widen the other.
 *
 * RETRY IDENTITY (`position: 'start'`): the pre-existing part+paragraph-index+size match (below)
 * is NECESSARY but not SUFFICIENT -- it would also SATISFY on an image placed with `'end'` at the
 * SAME paragraph (same part/paragraph-index/size), which is the WRONG retry answer for a `'start'`
 * caller. The read_headers_footers run-level `hasDrawing` BOOLEAN (docx-driver.cjs, GH #317 step 4
 * -- shipped as a primitive, NOT the inline-type array itself: that array ONE level deeper than the
 * paragraph's own already-landed `inline` field wiped the WHOLE callCommand result, LIVE-measured)
 * additionally requires paragraph `idx`'s OWN FIRST run to carry the drawing -- distinguishing a
 * drawing truly at run-index-0 from one merely present somewhere in the paragraph (a prior probe's
 * own flaw: mistaking an EMPTY-text run for a drawing, which this run-level signal avoids by
 * reading the run's own inline-content types, never its text emptiness).
 */

const TARGET_KINDS = Object.freeze(['paragraph', 'heading']);
const SOURCE_RE = /^data:image\/(png|jpeg);base64,[A-Za-z0-9+/]+=*$/;
const POSITIONS = Object.freeze(['after', 'end']);
const DEFAULT_POSITION = 'after';
// GH #306b 2. resz: the header/footer target -- part in ['header','footer']; GH #296 part 3: variant
// default | title | even (title/even only on an existing, already-active part -- header comment).
const HF_PARTS = Object.freeze(['header', 'footer']);
const HF_VARIANTS = Object.freeze(['default', 'title', 'even']);
// GH #317 step 4: selector.position -- a SEPARATE schema location + enum from the body target's own
// step-level `position` (POSITIONS/DEFAULT_POSITION above); see the header comment's own section.
const HF_POSITIONS = Object.freeze(['start', 'end']);
const HF_DEFAULT_POSITION = 'end';
/** Every valid selector key-set for the header/footer target, `part`+`section` required, `paragraph`/
 * `position`/`variant` each independently optional (2^3 = 8 combinations, alphabetically sorted to
 * match `Object.keys(sel).sort().join(',')`). */
const HF_SELECTOR_KEY_SETS = Object.freeze([
  'part,section', 'part,section,variant',
  'paragraph,part,section', 'paragraph,part,section,variant',
  'part,position,section', 'part,position,section,variant',
  'paragraph,part,position,section', 'paragraph,part,position,section,variant',
]);

function commonImageChecks(step, index) {
  if (typeof step.source !== 'string' || !SOURCE_RE.test(step.source)) {
    return `step ${index}: insert_image source must be a data:image/png;base64,... or data:image/jpeg;base64,... URI -- an external URL is never accepted (not fetched, not even at this layer)`;
  }
  if (!Number.isInteger(step.widthEmu) || step.widthEmu < 1) return `step ${index}: insert_image requires a positive integer widthEmu`;
  if (!Number.isInteger(step.heightEmu) || step.heightEmu < 1) return `step ${index}: insert_image requires a positive integer heightEmu`;
  return null;
}

/** GH #306b 2. resz: the header/footer target's own planning branch -- no `position` field (the
 * header comment's MECHANISM note: only an append to the part's own LAST paragraph is measured),
 * UNLESS `paragraph: {index, of}` names a specific one (GH #317, step 3). */
function planHeaderFooter(step, index, sel) {
  const s = sel.section;
  if (!s || typeof s !== 'object' || Object.keys(s).sort().join(',') !== 'index,of' || !Number.isInteger(s.index) || !Number.isInteger(s.of) || s.of < 1 || s.index < 0 || s.index >= s.of) {
    return { ok: false, detail: `step ${index}: insert_image header/footer target requires section: { index, of } (integers, 0 <= index < of = the section count)` };
  }
  if (!HF_PARTS.includes(sel.part)) return { ok: false, detail: `step ${index}: insert_image header/footer target requires part to be one of ${HF_PARTS.join(', ')}` };
  const variant = sel.variant === undefined ? 'default' : sel.variant;
  if (!HF_VARIANTS.includes(variant)) return { ok: false, detail: `step ${index}: insert_image header/footer target variant must be one of ${HF_VARIANTS.join(', ')}` };
  const commonErr = commonImageChecks(step, index);
  if (commonErr) return { ok: false, detail: commonErr };
  // GH #317 step 4: position lives in selector.position (NOT step.position -- that field is the
  // BODY target's own, a different contract; see the header comment's own section).
  if (step.position !== undefined) return { ok: false, detail: `step ${index}: insert_image header/footer target's position lives in selector.position ('start'|'end'), not the step's own position field` };
  const position = sel.position === undefined ? HF_DEFAULT_POSITION : sel.position;
  if (!HF_POSITIONS.includes(position)) return { ok: false, detail: `step ${index}: insert_image header/footer target selector.position must be one of ${HF_POSITIONS.join(', ')}` };
  // GH #317: `selector.paragraph: {index, of}` -- the SAME descriptor + validation header-footer.cjs's
  // own paragraph mode already uses (reused, not re-derived), nested inside `selector` like this
  // family's own section/part/variant (unlike header-footer.cjs, where section/part/variant are
  // step-level -- a different input shape per family, not a drifted convention). `of` is the
  // caller's OWN SEEN paragraph count, a different one -> FAIL_CLOSED at dispatch time
  // (classifyHeaderFooter/commandBodyHeaderFooter below), never guessed or silently clamped to the
  // last index.
  const para = paragraphDescriptor(sel.paragraph);
  if (para === false) return { ok: false, detail: `step ${index}: insert_image header/footer target selector.paragraph must be { index, of } (integers, 0 <= index < of = the part's paragraph count)` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: step.intent,
      selector: Object.freeze({ section: Object.freeze({ index: s.index, of: s.of }), part: sel.part, variant, position, ...(para ? { paragraph: para } : {}) }),
      preservationMode: 'PRIMITIVE_INTEGRAL', source: step.source, widthEmu: step.widthEmu, heightEmu: step.heightEmu,
      sectionIndex: s.index, sectionCount: s.of, part: sel.part, variant, position,
      ...(para ? { paragraphIndex: para.index, paragraphCount: para.of } : {}),
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

/** GH #306 part 4: the table-cell target -- the SAME { tableCell: { text, tableAnchor? } } cell identity replace_table_cell uses (table.cjs cellSelector). */
function planCell(step, index, sel) {
  const tc = sel.tableCell;
  const tcKeys = tc && typeof tc === 'object' && !Array.isArray(tc) ? Object.keys(tc).sort().join(',') : '';
  const nonEmpty = v => typeof v === 'string' && !!norm(v);
  const shapeOk = tcKeys === 'text' || (tcKeys === 'tableAnchor,text' && nonEmpty(tc.tableAnchor));
  if (!shapeOk || !nonEmpty(tc.text)) {
    return { ok: false, detail: `step ${index}: insert_image table-cell target requires selector.tableCell: { text, tableAnchor? } (non-empty strings: the cell's full text, and optionally the full text of another cell of the same table)` };
  }
  const commonErr = commonImageChecks(step, index);
  if (commonErr) return { ok: false, detail: commonErr };
  if (step.position !== undefined) return { ok: false, detail: `step ${index}: insert_image table-cell target takes no position (the image is appended to the cell's last paragraph)` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: step.intent,
      selector: Object.freeze({ tableCell: Object.freeze({ text: tc.text, ...(tcKeys === 'tableAnchor,text' ? { tableAnchor: tc.tableAnchor } : {}) }) }),
      preservationMode: 'PRIMITIVE_INTEGRAL', source: step.source, widthEmu: step.widthEmu, heightEmu: step.heightEmu,
      cellText: norm(tc.text), tableAnchor: tcKeys === 'tableAnchor,text' ? norm(tc.tableAnchor) : null,
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

/** GH #306 part 4: the optional floating placement of the body target -- only on the NEW paragraph of position 'after' (the measured route). */
function placementOf(step, position, index) {
  if (step.placement === undefined) return { ok: true, value: null };
  if (position !== 'after') return { ok: false, detail: `step ${index}: insert_image placement is only supported with position 'after' (a floating image in its own new paragraph -- the measured route)` };
  // required here, not at the top: shape-insert.cjs requires THIS module at load time (a top-level require back would hand it a partial export)
  const pl = require('./shape-insert.cjs').planPlacement(step.placement);
  return pl.ok ? pl : { ok: false, detail: `step ${index}: insert_image placement ${pl.detail}` };
}

/** The body target, { afterParagraph }. */
function planBody(step, index, sel) {
  const err = (typeof sel.afterParagraph !== 'string' || !norm(sel.afterParagraph) ? `step ${index}: insert_image selector.afterParagraph must be a non-empty string` : null)
    || commonImageChecks(step, index);
  if (err) return { ok: false, detail: err };
  const position = step.position === undefined ? DEFAULT_POSITION : step.position;
  if (!POSITIONS.includes(position)) return { ok: false, detail: `step ${index}: insert_image position must be one of ${POSITIONS.join(', ')}` };
  const placement = placementOf(step, position, index);
  if (!placement.ok) return placement;
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: step.intent, selector: Object.freeze({ afterParagraph: sel.afterParagraph }),
      preservationMode: 'PRIMITIVE_INTEGRAL', source: step.source, widthEmu: step.widthEmu, heightEmu: step.heightEmu, position, placement: placement.value,
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

function plan(step, index) {
  if (step.intent !== 'insert_image') return { ok: false, detail: `step ${index}: ${step.intent} is not an image-insert intent` };
  const sel = step.selector;
  if (!sel || typeof sel !== 'object') return { ok: false, detail: `step ${index}: insert_image requires a selector` };
  const keys = Object.keys(sel).sort().join(',');
  const planner = keys === 'afterParagraph' ? planBody : (HF_SELECTOR_KEY_SETS.includes(keys) ? planHeaderFooter : (keys === 'tableCell' ? planCell : null));
  if (!planner) return { ok: false, detail: `step ${index}: insert_image requires a selector of exactly { afterParagraph }, { section, part, variant?, paragraph?, position? } or { tableCell: { text, tableAnchor? } }` };
  // GH #306 part 4: placement is measured for the body target only (a header/footer part or a table cell would take it unmeasured)
  if (planner !== planBody && step.placement !== undefined) return { ok: false, detail: `step ${index}: insert_image placement is only supported for the body target (afterParagraph), not a ${planner === planCell ? 'table cell' : 'header/footer part'}` };
  return planner(step, index, sel);
}

// ---- the observation ------------------------------------------------------------------------------

const blocksOf = primary => (primary && primary.blocks) || [];
const fieldReadable = v => !(typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0);

function resolveAnchor(primary, text) {
  const r = resolveStableTarget(primary, { text });
  if (!r.ok) return r;
  if (!TARGET_KINDS.includes(r.target.kind)) {
    return { ok: false, target: null, selector: r.selector, ambiguity: 'UNSUPPORTED_TARGET_KIND', kind: r.target.kind };
  }
  return r;
}

/** ±1 EMU -- the measured DS rounding drift (gh306 probe: requested 2514600, read back 2514599),
 * the SAME tolerance the XLSX insert_image precedent already accepts for size, never exact-matched. */
const sizeMatches = (row, op) => fieldReadable(row.widthEmu) && fieldReadable(row.heightEmu)
  && Math.abs(row.widthEmu - op.widthEmu) <= 1 && Math.abs(row.heightEmu - op.heightEmu) <= 1;

/** position 'end': does the ANCHOR'S OWN paragraph already carry an inline image of the requested size? */
function matchingImageAtEnd(primary, anchorText, op) {
  const rows = imageRows(primary) || [];
  const want = norm(anchorText);
  return rows.some(r => fieldReadable(r.paragraph) && norm(r.paragraph) === want && r.placement === 'inline' && sizeMatches(r, op));
}

/** position 'after': see the header's own RETRY IDENTITY note -- position-based, not text-based (the new paragraph has none). */
function matchingImageAfter(primary, anchorIndex, op) {
  const blocks = blocksOf(primary);
  const next = blocks[anchorIndex + 1];
  if (!next || next.kind !== 'paragraph' || norm(next.text) !== '') return false;
  const rows = imageRows(primary) || [];
  return rows.some(r => fieldReadable(r.paragraph) && norm(r.paragraph) === '' && placementMatches(r, op) && sizeMatches(r, op));
}

/** GH #306 part 4: an inline request is matched only by an inline image, a floating one only by an image with exactly the requested wrap / behindDoc / position
 * (the SAME comparison shape-insert.cjs's placementMatches makes on the SAME ToJSON projection; inventory names the drawingType 'anchor' as placement 'floating'). */
const PLACEMENT_FIELDS = Object.freeze(['wrap', 'behindDoc', 'hRel', 'hAlign', 'hOffset', 'vRel', 'vAlign', 'vOffset']);
function placementMatches(r, op) {
  const p = op.placement || null;
  if (!p) return r.placement === 'inline';
  if (!PLACEMENT_FIELDS.every(k => fieldReadable(r[k])) || r.placement !== 'floating') return false;
  const want = { wrap: p.wrapType, behindDoc: p.behindDoc, hRel: p.hRel, hAlign: p.hAlign, hOffset: p.hAlign ? r.hOffset : p.hOffsetEmu, vRel: p.vRel, vAlign: null, vOffset: p.vOffsetEmu };
  return PLACEMENT_FIELDS.every(k => r[k] === want[k]);
}

// ---- GH #306b 2. resz: header/footer target helpers -----------------------------------------------

const isHF = op => !!(op.selector && op.selector.section);
const hfPartLabel = op => `${op.part}:${op.sectionIndex}:${op.variant}`;
const partAtHF = (sections, i, kind, variant) => (sections[i] && sections[i][kind] ? sections[i][kind][variant] : null);

/** The structural state the header/footer branch classifies/resolves/verifies from -- ONE observation. */
function hfState(primary, op) {
  const d = primary && primary.document;
  const sections = d ? d.sections : null;
  if (!Array.isArray(sections)) return { ok: false, reason: 'sections-unreadable' };
  if (sections.length !== op.sectionCount) return { ok: false, reason: 'section-descriptor-mismatch', sectionCount: sections.length };
  const inactive = variantInactiveReason(sections, op);
  if (inactive) return { ok: false, reason: inactive };
  return { ok: true, sections, part: partAtHF(sections, op.sectionIndex, op.part, op.variant) };
}

/** GH #296 part 3: why a title/even target is not writable, or null. Every section's variant state
 * must be KNOWN (an unread flag is never taken as a real one), and the target variant must ALREADY be
 * active -- title: the target section's own titlePage; even: evenAndOdd, document-wide, on every
 * section. 'default' is always shown, so it never reaches either check. */
function variantInactiveReason(sections, op) {
  if (op.variant === 'default') return null;
  const unknown = sections.findIndex(x => typeof x.titlePage !== 'boolean' || typeof x.evenAndOdd !== 'boolean');
  if (unknown >= 0) return `section-variant-state-unavailable-section-${unknown}`;
  const on = op.variant === 'title' ? sections[op.sectionIndex].titlePage === true : sections.every(x => x.evenAndOdd === true);
  return on ? null : `insert-image-${op.variant}-variant-inactive`;
}

/** GH #317: which paragraph this op targets -- the caller's named index when `paragraph` is present,
 * else the part's own LAST paragraph (the pre-existing, un-addressed behaviour). */
const targetIndexOf = (part, op) => (Number.isInteger(op.paragraphIndex) ? op.paragraphIndex : part.paras.length - 1);

/** GH #317 step 4: does this SPECIFIC run (read_headers_footers' own per-run `hasDrawing` boolean,
 * added alongside this candidate -- a PRIMITIVE; the inline-type ARRAY this started as wiped the
 * whole callCommand result LIVE, see docx-driver.cjs's own comment) carry a drawing? NOT the run's
 * text emptiness -- a prior probe's own measured flaw was mistaking an empty-text run for a
 * drawing; this reads the run's own page-side inline-content-type walk instead (computed there,
 * shipped here only as `true`/`false`). */
const runHasDrawing = run => !!run && run.hasDrawing === true;

/** Does the target paragraph (named by `paragraph`, or the part's OWN LAST one -- see targetIndexOf)
 * carry a matching inline image, LABELED as this exact section+kind+variant by the GH #306b docPr.id
 * part-join? See the header comment's RETRY IDENTITY note for why this is position-based, scoped by
 * part. GH #317 step 4: for `op.position === 'start'`, ALSO requires the paragraph's OWN FIRST run
 * to carry the drawing -- the part+paragraph-index+size match alone cannot tell a `'start'` placement
 * apart from an `'end'` one at the SAME paragraph (both would otherwise satisfy identically). */
function matchingImageInHeaderFooter(primary, partLabel, part, op) {
  if (!part || !Array.isArray(part.paras) || !part.paras.length) return false;
  const idx = targetIndexOf(part, op);
  if (idx < 0 || idx >= part.paras.length) return false;
  const targetPara = part.paras[idx];
  const rows = imageRows(primary) || [];
  // GH #317 follow-up: the paragraph is matched by its INDEX inside the part, not by text -- the row's
  // `paragraph` text is GetText(), which drops a real tab, while part.paras keeps it (GH #317).
  const matched = rows.some(r => fieldReadable(r.part) && r.part === partLabel && r.paragraphIndex === idx && r.placement === 'inline' && sizeMatches(r, op));
  if (!matched) return false;
  if (op.position === 'start') return runHasDrawing(Array.isArray(targetPara.runs) ? targetPara.runs[0] : null);
  return true;
}

// ---- classify / resolve ----------------------------------------------------------------------------

function classifyHeaderFooter(primary, op) {
  const s = hfState(primary, op);
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason, sectionCount: s.sectionCount } };
  if (!s.part) return { state: 'FAIL_CLOSED', evidence: { reason: 'insert-image-header-footer-part-missing' } };
  if (!s.part.simple) return { state: 'FAIL_CLOSED', evidence: { reason: 'header-footer-not-plain-paragraphs-of-runs' } };
  // GH #317: a named paragraph descriptor whose `of` no longer matches the part's actual paragraph
  // count -> FAIL_CLOSED, never falling back to "the nearest" or "the last" paragraph.
  if (Number.isInteger(op.paragraphIndex) && s.part.paragraphs !== op.paragraphCount) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'FAIL_CLOSED:paragraph-descriptor-mismatch', paragraphCount: s.part.paragraphs } };
  }
  const partLabel = hfPartLabel(op);
  if (matchingImageInHeaderFooter(primary, partLabel, s.part, op)) {
    return { state: 'SATISFIED', evidence: { section: op.sectionIndex, part: op.part, variant: op.variant, position: op.position, ...(Number.isInteger(op.paragraphIndex) ? { paragraph: op.paragraphIndex } : {}), widthEmu: op.widthEmu, heightEmu: op.heightEmu, proof: Number.isInteger(op.paragraphIndex) ? 'fresh-live-image-at-named-paragraph' : 'fresh-live-image-at-part-last-paragraph' } };
  }
  return { state: 'APPLY', evidence: { section: op.sectionIndex, part: op.part, variant: op.variant, position: op.position, ...(Number.isInteger(op.paragraphIndex) ? { paragraph: op.paragraphIndex } : {}) } };
}

// ---- GH #306 part 4: table-cell target helpers --------------------------------------------------

const isCell = op => !!(op.selector && op.selector.tableCell);
const IMAGE_INLINE_MARKERS = Object.freeze(['blipFill', 'image', 'paraDrawing']);

/** The ONE top-level table cell the op names (table.cjs's own read + scope, nested cells NOT included), with the text of its LAST paragraph. */
function cellTarget(primary, op) {
  const { tableReadWithVerticalMerges, scopeOf } = require('./table.cjs');
  const read = tableReadWithVerticalMerges(primary);
  if (!read.ok) return { ok: false, reason: read.reason };
  const scope = scopeOf(read, op.tableAnchor);
  if (!scope.ok) return { ok: false, reason: scope.reason };
  const hits = scope.cells.filter(c => c.cell.text === op.cellText);
  const cell = hits.length === 1 ? hits[0].cell : null;
  let reason = null;
  if (!cell) reason = hits.length ? 'insert-image-table-cell-ambiguous' : 'insert-image-table-cell-not-found';
  else if (!cell.simple) reason = 'insert-image-table-cell-not-plain-paragraphs-of-runs';
  else if (!Array.isArray(cell.paraFormats) || cell.paraFormats.length < 1 || !Array.isArray(cell.inline)) reason = 'insert-image-table-cell-content-unreadable';
  if (reason) return { ok: false, reason };
  return { ok: true, at: hits[0].at, cell, lastText: norm(cell.paraFormats[cell.paraFormats.length - 1].text), read };
}

/** Image rows that sit in a table-cell paragraph (paragraphIndex null -- docx-images-read.cjs's own walk) with the given text, inline, at the requested size. */
const inCellParagraph = (r, lastText) => r.paragraphIndex === null && r.part === 'body' && fieldReadable(r.paragraph) && norm(r.paragraph) === lastText;
const cellImageRows = (primary, lastText, op) => (imageRows(primary) || []).filter(r => inCellParagraph(r, lastText) && r.placement === 'inline' && sizeMatches(r, op));
const cellCarriesImage = cell => IMAGE_INLINE_MARKERS.some(t => cell.inline.includes(t));

function classifyCell(primary, op) {
  const t = cellTarget(primary, op);
  if (!t.ok) return { state: 'FAIL_CLOSED', evidence: { reason: t.reason } };
  if (cellImageRows(primary, t.lastText, op).length > 0 && cellCarriesImage(t.cell)) {
    return { state: 'SATISFIED', evidence: { cell: op.cellText, at: t.at, widthEmu: op.widthEmu, heightEmu: op.heightEmu, proof: 'fresh-live-image-in-the-cells-last-paragraph' } };
  }
  return { state: 'APPLY', evidence: { cell: op.cellText, at: t.at } };
}

function classify(primary, op) {
  const special = isHF(op) ? classifyHeaderFooter : (isCell(op) ? classifyCell : null);
  if (special) return special(primary, op);
  const anchor = resolveAnchor(primary, op.selector.afterParagraph);
  if (!anchor.ok) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `insert-image-anchor-${String(anchor.ambiguity).toLowerCase().replace(/_/g, '-')}` } };
  }
  if (op.position === 'end') {
    if (matchingImageAtEnd(primary, anchor.target.text, op)) {
      return { state: 'SATISFIED', evidence: { anchor: anchor.target.text, position: 'end', widthEmu: op.widthEmu, heightEmu: op.heightEmu, proof: 'fresh-live-image-at-anchor-end' } };
    }
    return { state: 'APPLY', evidence: { anchor: anchor.target.text, position: 'end' } };
  }
  const anchorIndex = blocksOf(primary).indexOf(anchor.target);
  if (matchingImageAfter(primary, anchorIndex, op)) {
    return { state: 'SATISFIED', evidence: { anchor: anchor.target.text, position: 'after', widthEmu: op.widthEmu, heightEmu: op.heightEmu, proof: 'fresh-live-empty-paragraph-image-immediately-after-anchor' } };
  }
  return { state: 'APPLY', evidence: { anchor: anchor.target.text, position: 'after' } };
}

function resolveHeaderFooter(primary, sel) {
  const pseudoOp = { sectionIndex: sel.section.index, sectionCount: sel.section.of, part: sel.part, variant: sel.variant };
  const s = hfState(primary, pseudoOp);
  if (!s.ok) return { ok: false, target: null, selector: sel, ambiguity: String(s.reason).toUpperCase().replace(/-/g, '_') };
  if (!s.part) return { ok: false, target: null, selector: sel, ambiguity: 'INSERT_IMAGE_HEADER_FOOTER_PART_MISSING' };
  if (!s.part.simple) return { ok: false, target: null, selector: sel, ambiguity: 'HEADER_FOOTER_NOT_PLAIN_PARAGRAPHS_OF_RUNS' };
  const para = sel.paragraph ? { index: sel.paragraph.index, of: sel.paragraph.of } : null;
  if (para && s.part.paragraphs !== para.of) return { ok: false, target: null, selector: sel, ambiguity: 'PARAGRAPH_DESCRIPTOR_MISMATCH' };
  return { ok: true, target: { kind: 'insert_image_header_footer_target', sectionIndex: sel.section.index, sectionCount: sel.section.of, part: sel.part, variant: sel.variant, ...(para ? { paragraphIndex: para.index, paragraphCount: para.of } : {}) }, selector: sel, ambiguity: null };
}

function resolve(primary, selector) {
  const sel = selector || {};
  if (sel.section) return resolveHeaderFooter(primary, sel);
  if (sel.tableCell) {
    const t = cellTarget(primary, { cellText: norm(sel.tableCell.text), tableAnchor: sel.tableCell.tableAnchor === undefined ? null : norm(sel.tableCell.tableAnchor) });
    if (!t.ok) return { ok: false, target: null, selector, ambiguity: String(t.reason).toUpperCase().replace(/-/g, '_') };
    return { ok: true, target: { kind: 'insert_image_table_cell', cellText: t.cell.text, tableAnchor: sel.tableCell.tableAnchor === undefined ? null : norm(sel.tableCell.tableAnchor), at: t.at }, selector, ambiguity: null };
  }
  const anchor = resolveAnchor(primary, sel.afterParagraph);
  if (!anchor.ok) return { ok: false, target: null, selector, ambiguity: anchor.ambiguity };
  return { ok: true, target: { kind: 'insert_image_anchor', anchorText: norm(anchor.target.text) }, selector, ambiguity: null };
}

/** GH #317: a named paragraph target names the SAME index + count; an un-addressed (last-paragraph) target names neither. */
const sameParagraphTarget = (op, target) => (Number.isInteger(op.paragraphIndex)
  ? target.paragraphIndex === op.paragraphIndex && target.paragraphCount === op.paragraphCount
  : target.paragraphIndex === undefined);

function preDispatchCheck(op, target) {
  if (isHF(op)) {
    return target && target.kind === 'insert_image_header_footer_target' && target.sectionIndex === op.sectionIndex
      && target.sectionCount === op.sectionCount && target.part === op.part && target.variant === op.variant && sameParagraphTarget(op, target)
      ? { ok: true } : { ok: false, dispatched: false, detail: 'insert-image-header-footer-target-unresolved' };
  }
  if (isCell(op)) {
    return target && target.kind === 'insert_image_table_cell' && target.cellText === op.cellText && target.tableAnchor === op.tableAnchor
      ? { ok: true } : { ok: false, dispatched: false, detail: 'insert-image-table-cell-target-unresolved' };
  }
  return target && target.kind === 'insert_image_anchor' && target.anchorText === norm(op.selector.afterParagraph)
    ? { ok: true }
    : { ok: false, dispatched: false, detail: 'insert-image-anchor-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  if (target.kind === 'insert_image_header_footer_target') {
    const meta = Object.freeze({ family: 'image-insert', action: 'insertHeaderFooter', sectionIndex: target.sectionIndex, sectionCount: target.sectionCount, part: target.part, variant: target.variant, source: op.source, widthEmu: op.widthEmu, heightEmu: op.heightEmu, position: op.position,
      ...(Number.isInteger(target.paragraphIndex) ? { paragraphIndex: target.paragraphIndex, paragraphCount: target.paragraphCount } : {}),
      // GH #296 part 3: title/even re-check their flag in-sandbox through the B21 bridge READ prelude
      // (docx-driver.cjs prepends it, and refuses the command off the word /3 runtime).
      ...(target.variant !== 'default' ? { bridgeSectionVariantRead: true } : {}) });
    return { kind: 'callCommand', body: commandBodyHeaderFooter(meta), meta };
  }
  if (target.kind === 'insert_image_table_cell') {
    const meta = Object.freeze({ family: 'image-insert', action: 'insertCell', cellText: target.cellText, tableAnchor: target.tableAnchor, source: op.source, widthEmu: op.widthEmu, heightEmu: op.heightEmu });
    return { kind: 'callCommand', body: commandBodyCell(meta), meta };
  }
  const meta = Object.freeze({ family: 'image-insert', action: 'insert', anchorText: target.anchorText, source: op.source, widthEmu: op.widthEmu, heightEmu: op.heightEmu, position: op.position, placement: op.placement || null });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (GH #306b 2. resz). Re-locates the section/part by index (the SAME GetSections() +
 * GetHeader/GetFooter(variant, false) call header-footer.cjs's own commandBody uses), re-checks the
 * section count and that the part exists and holds only plain paragraphs of runs (an unread or
 * unexpected shape returns unchanged), then either appends the new image to the target paragraph's
 * OWN LAST run via AddDrawing (`position: 'end'`, the EXACT object shape LIVE-measured,
 * docx-gh306b-header-image-probe.live.cjs's SETUP: `hdr.GetElement(0).AddDrawing(...)`), or -- GH
 * #317 step 4, LIVE-measured via docx-gh317-image-start-of-paragraph-probe.live.cjs --
 * wraps the image in a FRESH run and inserts it at index 0 via `AddElement(run, 0)` (`position:
 * 'start'`): the OTHER argument order, `AddElement(0, run)`, was measured a silent no-op. No NEW
 * paragraph is ever created either way -- see the family header comment's MECHANISM note for why.
 */
function commandBodyHeaderFooter(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  const secs = doc.GetSections() || [];
  if (secs.length !== m.sectionCount) return;
  const sec = secs[m.sectionIndex];
  if (!sec) return;
  // GH #296 part 3: a title/even target's flag, re-read through the B21 bridge (__eoB21Variant, the
  // driver's prelude) -- an unavailable or inactive state leaves the document unchanged.
  if (m.variant !== 'default') {
    const v = __eoB21Variant(m.sectionIndex);
    if (!v || (m.variant === 'title' ? v.titlePage : v.evenAndOdd) !== true) return;
  }
  const part = m.part === 'header' ? sec.GetHeader(m.variant, false) : sec.GetFooter(m.variant, false);
  if (!part) return;
  const n = part.GetElementsCount();
  if (n < 1) return;
  for (let i = 0; i < n; i += 1) {
    const e = part.GetElement(i);
    if (!e || e.GetClassType() !== 'paragraph') return;
    for (let j = 0; j < e.GetElementsCount(); j += 1) { const ch = e.GetElement(j); if (!ch || ch.GetClassType() !== 'run') return; }
  }
  // GH #317: a named paragraph index, re-checked against the part's ACTUAL paragraph count read
  // just now (fail-closed on a mismatch, never clamped); absent -> the pre-existing last-paragraph
  // behaviour.
  if (Number.isInteger(m.paragraphIndex) && n !== m.paragraphCount) return;
  const targetIndex = Number.isInteger(m.paragraphIndex) ? m.paragraphIndex : n - 1;
  const targetPara = part.GetElement(targetIndex);
  if (!targetPara || targetPara.GetClassType() !== 'paragraph') return;
  if (typeof Api.CreateImage !== 'function') return;
  const img = Api.CreateImage(m.source, m.widthEmu, m.heightEmu);
  if (!img) return;
  // GH #317 step 4: 'start' wraps the image in a FRESH run and inserts it at index 0 -- AddDrawing
  // always APPENDS to the paragraph's existing run flow, never lands before text already there.
  if (m.position === 'start') {
    if (typeof Api.CreateRun !== 'function' || typeof targetPara.AddElement !== 'function') return;
    const run = Api.CreateRun();
    if (typeof run.AddDrawing !== 'function') return;
    run.AddDrawing(img);
    targetPara.AddElement(run, 0);
    return;
  }
  if (typeof targetPara.AddDrawing !== 'function') return;
  targetPara.AddDrawing(img);
} catch (_) {}
`;
}

/**
 * One callCommand (GH #306 probe, commit 749a4f607 -- both routes LIVE PASS). Re-locates the
 * anchor paragraph by the SAME text-matching walk every other DOCX write family uses (an identity
 * walk never matches a new-per-call wrapper, measured three times now: #308/#313/#306), re-checked
 * for uniqueness inside the sandbox. `Api.CreateImage` then either `anchor.AddDrawing` (position
 * 'end') or a fresh paragraph + `doc.AddElement(index + 1, ...)` (position 'after', the SAME
 * mechanism #308 already measured for a table). Any guard miss returns unchanged.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  const textOpts = { Numbering: false };
  const lookupTextOf = p => String(p.GetText(textOpts) || '').replace(/[\\r\\n]+$/, '');
  const n = doc.GetElementsCount();
  let anchorIndex = -1;
  let anchorParagraph = null;
  for (let i = 0; i < n; i += 1) {
    const e = doc.GetElement(i);
    if (e && typeof e.GetClassType === 'function' && e.GetClassType() === 'paragraph' && lookupTextOf(e) === m.anchorText) {
      if (anchorIndex >= 0) return; // ambiguous at dispatch time: refuse rather than guess
      anchorIndex = i;
      anchorParagraph = e;
    }
  }
  if (anchorIndex < 0 || !anchorParagraph) return;
  if (typeof Api.CreateImage !== 'function') return;
  const img = Api.CreateImage(m.source, m.widthEmu, m.heightEmu);
  if (!img) return;
  if (m.position === 'end') {
    if (typeof anchorParagraph.AddDrawing !== 'function') return;
    anchorParagraph.AddDrawing(img);
  } else {
    if (typeof doc.AddElement !== 'function') return;
    // GH #306 part 4: a floating placement is set on the image BEFORE AddDrawing (LIVE-measured, each setter returned true); any refusal leaves the document unchanged
    const p = m.placement;
    if (p) {
      if (typeof img.SetWrappingStyle !== 'function' || typeof img.SetVerPosition !== 'function') return;
      if (p.hAlign ? typeof img.SetHorAlign !== 'function' : typeof img.SetHorPosition !== 'function') return;
      if (img.SetWrappingStyle(p.wrap) === false) return;
      if ((p.hAlign ? img.SetHorAlign(p.hRel, p.hAlign) : img.SetHorPosition(p.hRel, p.hOffsetEmu)) === false) return;
      if (img.SetVerPosition(p.vRel, p.vOffsetEmu) === false) return;
    }
    const np = Api.CreateParagraph();
    if (typeof np.AddDrawing !== 'function') return;
    np.AddDrawing(img);
    doc.AddElement(anchorIndex + 1, np);
  }
} catch (_) {}
`;
}

/**
 * One callCommand (GH #306 part 4, LIVE: docx-gh306-cell-floating-image-probe.live.cjs): the cell is found the way replace_table_cell's own
 * command finds it (table.cjs commandBody: every GetAllTables() table, an optional tableAnchor naming ONE table, the cell's FULL normalized
 * content text unique in that scope), re-checked to hold only paragraphs made only of runs, and the image is appended with AddDrawing to
 * the cell content's LAST paragraph (the measured reach: GetCell(...).GetContent().GetElement(n - 1).AddDrawing). Any guard miss returns
 * unchanged -- including a nested-table cell with the same text, which the planner's top-level scope does not see.
 */
function commandBodyCell(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
  const doc = Api.GetDocument();
  const cellsOf = t => { const out = []; for (let r = 0; r < t.GetRowsCount(); r += 1) { const row = t.GetRow(r); for (let c = 0; c < row.GetCellsCount(); c += 1) out.push(row.GetCell(c)); } return out; };
  const textOf = cell => norm(cell.GetContent().GetText(textOpts));
  let tables = doc.GetAllTables() || [];
  if (m.tableAnchor !== null) {
    tables = tables.filter(t => cellsOf(t).some(cell => textOf(cell) === m.tableAnchor));
    if (tables.length !== 1) return;
  }
  const hits = [];
  for (const t of tables) for (const cell of cellsOf(t)) if (textOf(cell) === m.cellText) hits.push(cell);
  if (hits.length !== 1) return;
  const content = hits[0].GetContent();
  const n = content.GetElementsCount();
  if (n < 1) return;
  for (let i = 0; i < n; i += 1) {
    const e = content.GetElement(i);
    if (!e || e.GetClassType() !== 'paragraph') return;
    for (let j = 0; j < e.GetElementsCount(); j += 1) { const ch = e.GetElement(j); if (!ch || ch.GetClassType() !== 'run') return; }
  }
  const last = content.GetElement(n - 1);
  if (typeof last.AddDrawing !== 'function' || typeof Api.CreateImage !== 'function') return;
  const img = Api.CreateImage(m.source, m.widthEmu, m.heightEmu);
  if (!img) return;
  last.AddDrawing(img);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** L1 (table-cell target): the SAME cell, found again, now carries the image marker, and exactly ONE more matching image row sits in its last paragraph. */
function verifyPrimitiveCell(op, before, after) {
  const tb = cellTarget(before && before.primary, op);
  const ta = cellTarget(after && after.primary, op);
  if (!tb.ok || !ta.ok) return { ok: false, evidence: { reason: !tb.ok ? `before-${tb.reason}` : `after-${ta.reason}` } };
  const nb = cellImageRows(before.primary, tb.lastText, op).length;
  const na = cellImageRows(after.primary, ta.lastText, op).length;
  const ok = ta.at === tb.at && na === nb + 1 && cellCarriesImage(ta.cell);
  return { ok, evidence: { cell: op.cellText, at: ta.at, matchingRowsBefore: nb, matchingRowsAfter: na, cellCarriesImage: cellCarriesImage(ta.cell) } };
}

/** Every table with the target cell's own `inline` taken out -- the one field the write may change (an image marker joins it). */
function tablesWithoutTargetInline(read, at) {
  const [ti, ri, ci] = at.split(':').map(Number);
  return JSON.stringify(read.tables.map((t, i) => (t.rows || []).map((row, r) => row.map((cell, c) => (i === ti && r === ri && c === ci ? { ...cell, inline: null } : cell)))));
}

/**
 * L3 (table-cell target): every image preserved (multiset, by key) and exactly one new; the body block sequence unchanged (an image adds no
 * text -- canonicalBlock compares the table cells' text and spans); every table, every cell -- the target included -- unchanged except the
 * target cell's own `inline` list; the target cell's text, paragraph count and per-paragraph text/format unchanged (part of that comparison).
 */
function verifyPreservationCell(op, before, after) {
  const tb = cellTarget(before && before.primary, op);
  const ta = cellTarget(after && after.primary, op);
  if (!tb.ok || !ta.ok) return { ok: false, evidence: { reason: !tb.ok ? `before-${tb.reason}` : `after-${ta.reason}` } };
  const beforeRows = imageRows(before.primary) || [];
  const afterRows = imageRows(after.primary) || [];
  const imagesPreserved = everyBeforeImagePreserved(beforeRows, afterRows);
  const countOk = afterRows.length === beforeRows.length + 1;
  const sameBlockSequence = JSON.stringify(blocksOf(before.primary).map(b => canonicalBlock(b))) === JSON.stringify(blocksOf(after.primary).map(b => canonicalBlock(b)));
  const sameCellAt = ta.at === tb.at;
  const tablesSame = sameCellAt && tablesWithoutTargetInline(tb.read, tb.at) === tablesWithoutTargetInline(ta.read, ta.at);
  const inlineGrewOnly = tb.cell.inline.every(t => ta.cell.inline.includes(t));
  const ok = [imagesPreserved, countOk, sameBlockSequence, tablesSame, inlineGrewOnly].every(Boolean);
  return { ok, evidence: { cell: op.cellText, at: ta.at, imagesPreserved, countOk, sameBlockSequence, tablesSame, inlineGrewOnly } };
}

/** L1 (header/footer target): the part's own LAST paragraph carries a matching, part-labeled inline
 * image, AND read_headers_footers' own `inline` array for that paragraph names it (the SAME three
 * markers LIVE-measured, docx-gh306b-header-image-probe.live.cjs: blipFill/image/paraDrawing). */
function verifyPrimitiveHeaderFooter(op, after) {
  const s = hfState(after && after.primary, op);
  if (!s.ok) return { ok: false, evidence: { reason: `after-${s.reason}` } };
  if (!s.part) return { ok: false, evidence: { reason: 'after-insert-image-header-footer-part-missing' } };
  const partLabel = hfPartLabel(op);
  const ok1 = matchingImageInHeaderFooter(after.primary, partLabel, s.part, op);
  const idx = targetIndexOf(s.part, op);
  const targetPara = s.part.paras[idx];
  const inline = targetPara && Array.isArray(targetPara.inline) ? targetPara.inline : null;
  const inlineHasImage = !!inline && ['blipFill', 'image', 'paraDrawing'].some(t => inline.includes(t));
  return { ok: ok1 && inlineHasImage, evidence: { section: op.sectionIndex, part: op.part, variant: op.variant, ...(Number.isInteger(op.paragraphIndex) ? { paragraph: op.paragraphIndex } : {}), partLabel, widthEmu: op.widthEmu, heightEmu: op.heightEmu, inlineHasImage } };
}

function verifyPrimitive(op, before, after) {
  if (isHF(op)) return verifyPrimitiveHeaderFooter(op, after);
  if (isCell(op)) return verifyPrimitiveCell(op, before, after);
  const anchorAfter = resolveAnchor(after && after.primary, op.selector.afterParagraph);
  if (!anchorAfter.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the after-state' } };
  if (op.position === 'end') {
    const ok = matchingImageAtEnd(after.primary, anchorAfter.target.text, op);
    return { ok, evidence: { anchor: anchorAfter.target.text, position: 'end', widthEmu: op.widthEmu, heightEmu: op.heightEmu } };
  }
  const anchorIndexAfter = blocksOf(after.primary).indexOf(anchorAfter.target);
  const ok = matchingImageAfter(after.primary, anchorIndexAfter, op);
  return { ok, evidence: { anchor: anchorAfter.target.text, position: 'after', anchorIndexAfter } };
}

/** Every image's own key (placement/wrap/size/paragraph) -- the SAME fields image.cjs's own imageKey uses, minus alt/title (this family never writes or reads them). */
const imageKey = r => JSON.stringify([r.placement, r.wrap, r.widthEmu, r.heightEmu, r.paragraph]);

/** Every BEFORE image still exists AFTER, by key (a multiset subtraction) -- insert never loses or mutates an existing image. */
function everyBeforeImagePreserved(beforeRows, afterRows) {
  const pool = afterRows.map(imageKey);
  for (const r of beforeRows) {
    const i = pool.indexOf(imageKey(r));
    if (i < 0) return false;
    pool.splice(i, 1);
  }
  return true;
}

/**
 * PRIMITIVE_INTEGRAL: every OTHER image preserved (multiset, above) AND exactly one new image
 * appears (the count check); the body block sequence is otherwise unchanged -- position 'end'
 * touches no block (an image is not part of a block's own text), position 'after' inserts EXACTLY
 * the one empty paragraph this write creates, immediately after the anchor, nothing else shifted
 * (the SAME block-sequence discipline `table-insert.cjs`'s own verifyPreservation already uses for
 * its table, `canonicalBlock` REUSED from there unchanged).
 */
/** Every image's own key INCLUDING its part label -- for the header/footer target ONLY, where part
 * now distinguishes a header/footer image from a same-shape body image. The body target's own
 * `imageKey`/`everyBeforeImagePreserved` above are UNTOUCHED (no part field, matching their existing
 * fixtures and tests exactly as before this card). */
const imageKeyWithPart = r => JSON.stringify([r.part, r.placement, r.wrap, r.widthEmu, r.heightEmu, r.paragraph]);

function everyBeforeImagePreservedWithPart(beforeRows, afterRows) {
  const pool = afterRows.map(imageKeyWithPart);
  for (const r of beforeRows) {
    const i = pool.indexOf(imageKeyWithPart(r));
    if (i < 0) return false;
    pool.splice(i, 1);
  }
  return true;
}

/** The body as the reads see it, WITHOUT the total image count (header-footer.cjs's own bodyKey
 * includes it -- this family's header/footer target deliberately changes that count by design;
 * image preservation is checked separately, by key WITH part, above). */
function bodyTextKey(primary) {
  const d = (primary && primary.document) || {};
  return JSON.stringify({ units: textUnits(primary).map(u => [u.kind, norm(u.text)]), tables: tablesContent(d.tables), footnotes: d.footnotes || null, endnotes: d.endnotes || null, bookmarks: d.bookmarks || null, inMainBody: d.inMainBody === undefined ? null : d.inMainBody });
}

/**
 * L3 (header/footer target): every image anywhere, by key WITH part, preserved except the one new
 * one; the body (text units, tables, notes, bookmarks) unchanged; every OTHER section/kind/variant
 * part unchanged (header-footer.cjs's own `otherParts`, REUSED as-is); the target part's own
 * paragraph count and every paragraph's TEXT unchanged (only the last paragraph's drawing content
 * may grow).
 */
function verifyPreservationHeaderFooter(op, before, after) {
  const sb = hfState(before && before.primary, op);
  const sa = hfState(after && after.primary, op);
  if (!sb.ok || !sa.ok) return { ok: false, evidence: { reason: !sb.ok ? `before-${sb.reason}` : `after-${sa.reason}` } };
  const beforeRows = imageRows(before.primary) || [];
  const afterRows = imageRows(after.primary) || [];
  const imagesPreserved = everyBeforeImagePreservedWithPart(beforeRows, afterRows);
  const countOk = afterRows.length === beforeRows.length + 1;
  const sameBody = bodyTextKey(before.primary) === bodyTextKey(after.primary);
  const pb = otherParts(sb.sections, op);
  const pa = otherParts(sa.sections, op);
  const changedParts = Object.keys(pb).filter(k => pb[k] !== pa[k]);
  const bp = sb.part;
  const ap = sa.part;
  const paragraphCountOk = !!bp && !!ap && bp.paragraphs === ap.paragraphs && bp.paras.length === ap.paras.length;
  const textsSame = paragraphCountOk && bp.paras.every((x, i) => norm(x.text) === norm(ap.paras[i].text));
  const ok = imagesPreserved && countOk && sameBody && changedParts.length === 0 && paragraphCountOk && textsSame;
  return { ok, evidence: { section: op.sectionIndex, part: op.part, variant: op.variant, imagesPreserved, countOk, bodyUnchanged: sameBody, changedParts, paragraphCountOk, textsSame } };
}

function verifyPreservation(op, before, after) {
  if (isHF(op)) return verifyPreservationHeaderFooter(op, before, after);
  if (isCell(op)) return verifyPreservationCell(op, before, after);
  const anchorBefore = resolveAnchor(before && before.primary, op.selector.afterParagraph);
  if (!anchorBefore.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the before-state' } };
  const beforeRows = imageRows(before.primary) || [];
  const afterRows = imageRows(after.primary) || [];
  const imagesPreserved = everyBeforeImagePreserved(beforeRows, afterRows);
  const countOk = afterRows.length === beforeRows.length + 1;

  const beforeBlocks = blocksOf(before.primary);
  const afterBlocks = blocksOf(after.primary);
  const anchorIndex = beforeBlocks.indexOf(anchorBefore.target);

  if (op.position === 'end') {
    const sameBlockSequence = JSON.stringify(beforeBlocks.map(b => canonicalBlock(b))) === JSON.stringify(afterBlocks.map(b => canonicalBlock(b)));
    const ok = imagesPreserved && countOk && sameBlockSequence;
    return { ok, evidence: { position: 'end', imagesPreserved, countOk, sameBlockSequence } };
  }

  const newBlock = afterBlocks[anchorIndex + 1];
  const newBlockOk = !!newBlock && newBlock.kind === 'paragraph' && norm(newBlock.text) === '';
  const beforePrefix = beforeBlocks.slice(0, anchorIndex + 1).map(b => canonicalBlock(b));
  const afterPrefix = afterBlocks.slice(0, anchorIndex + 1).map(b => canonicalBlock(b));
  const prefixSame = JSON.stringify(beforePrefix) === JSON.stringify(afterPrefix);
  const restBefore = beforeBlocks.slice(anchorIndex + 1).map(b => canonicalBlock(b));
  const restAfter = afterBlocks.slice(anchorIndex + 2).map(b => canonicalBlock(b));
  const restSame = JSON.stringify(restBefore) === JSON.stringify(restAfter);
  const lengthOk = afterBlocks.length === beforeBlocks.length + 1;
  const ok = imagesPreserved && countOk && newBlockOk && prefixSame && restSame && lengthOk;
  return { ok, evidence: { position: 'after', imagesPreserved, countOk, newBlockOk, prefixSame, restSame, lengthOk } };
}

function wholeTaskChecks(op, finalObservation) {
  if (isHF(op)) {
    const s = hfState(finalObservation.primary, op);
    const ok = s.ok && !!s.part && matchingImageInHeaderFooter(finalObservation.primary, hfPartLabel(op), s.part, op);
    const where = Number.isInteger(op.paragraphIndex) ? `at paragraph ${op.paragraphIndex}/${op.paragraphCount} of` : 'at the end of';
    return [{ name: `${op.id}: an image (${op.widthEmu}x${op.heightEmu} EMU) sits ${where} section ${op.sectionIndex}/${op.sectionCount} ${op.part} (${op.variant})`, ok }];
  }
  if (isCell(op)) {
    const t = cellTarget(finalObservation.primary, op);
    const ok = t.ok && cellCarriesImage(t.cell) && cellImageRows(finalObservation.primary, t.lastText, op).length > 0;
    return [{ name: `${op.id}: an image (${op.widthEmu}x${op.heightEmu} EMU) sits in the table cell "${op.selector.tableCell.text}"`, ok }];
  }
  const anchor = resolveAnchor(finalObservation.primary, op.selector.afterParagraph);
  let ok = false;
  if (anchor.ok) {
    if (op.position === 'end') ok = matchingImageAtEnd(finalObservation.primary, anchor.target.text, op);
    else { const idx = blocksOf(finalObservation.primary).indexOf(anchor.target); ok = matchingImageAfter(finalObservation.primary, idx, op); }
  }
  const where = (op.position === 'end' ? `at the end of "${op.selector.afterParagraph}"` : `immediately after "${op.selector.afterParagraph}"`) + (op.placement ? ` (floating, wrap ${op.placement.wrap})` : '');
  return [{ name: `${op.id}: an image (${op.widthEmu}x${op.heightEmu} EMU) sits ${where}`, ok }];
}

// refs: none -- an image is not a body text unit a later step can name
function producedText() {
  return null;
}

module.exports = {
  plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  resolveAnchor, matchingImageAtEnd, matchingImageAfter, sizeMatches, everyBeforeImagePreserved, imageKey, POSITIONS, DEFAULT_POSITION, SOURCE_RE,
  isHF, hfPartLabel, hfState, matchingImageInHeaderFooter, commandBodyHeaderFooter, everyBeforeImagePreservedWithPart, imageKeyWithPart, bodyTextKey, HF_PARTS, HF_VARIANTS, targetIndexOf,
  HF_POSITIONS, HF_DEFAULT_POSITION, HF_SELECTOR_KEY_SETS, runHasDrawing, placementMatches,
  isCell, cellTarget, cellImageRows, commandBodyCell, tablesWithoutTargetInline,
};
