'use strict';

const { norm } = require('../inventory.cjs');
const { resolveStableTarget } = require('../resolver.cjs');
const { imageRows } = require('./image.cjs');
const { canonicalBlock } = require('./table-insert.cjs');
const { everyBeforeImagePreserved } = require('./image-insert.cjs');

/**
 * DOCX `shape-insert` family (GH #309 part 1, NEW_CAPABILITY): insert_shape -- a NEW inline DrawingML shape with its own text in a NEW paragraph right after a named anchor paragraph.
 * Part 2 of the GH (a native diagram, #323) and a vector-image fallback (#306) are NOT this family.
 *
 * LIVE-measured (prod NC/DS, disposable copy; test/live/docx-gh309-shape-probe.live.cjs + docx-gh309-shape-read-probe.live.cjs, both PASS):
 *   - `Api.CreateShape(preset, wEmu, hEmu, fill, stroke)` with `Api.CreateSolidFill(Api.CreateRGBColor(..))` and `Api.CreateStroke(19050, solidFill)` exists; the shape's own text goes in through
 *     `shape.GetDocContent().GetElement(0).AddText(text)`; a NEW paragraph carrying `AddDrawing(shape)` is placed with `doc.AddElement(anchorIndex + 1, paragraph)` -- the SAME mechanism
 *     insert_image (#306) and insert_table (#308) use. Presets measured: 'rect' and 'flowChartProcess' (part 1), then 'roundRect', 'ellipse', 'triangle', 'trapezoid',
 *     'flowChartDecision', 'flowChartTerminator', 'rightArrow', 'chevron', 'homePlate' (the same probe, GH309_PRESETS, each stored with its own prst + size + fill + 19050 outline + text, inline,
 *     directly after its anchor) -- ONLY these are accepted. An unmeasured preset is refused by name, never tried: LIVE-measured, the engine does NOT refuse an unknown preset name -- it stores
 *     a 'rect' instead, silently (write ok, prst="rect"), and the L1 below reads text and size, not the preset, so it would not catch that.
 *   - stored (fresh reopen): `<w:drawing><wp:inline>` + `<wps:wsp>`, `a:prstGeom prst` = the preset, `wp:extent cx/cy` = the requested EMU BYTE-EXACT, `a:solidFill srgbClr` = the fill, `a:ln w` = 19050,
 *     the text in `w:txbxContent`; the shape paragraph has NO text of its own and sits DIRECTLY after the anchor; the anchor and every other paragraph are untouched.
 *   - the anchor is found by TEXT, never by wrapper identity (a body element wrapper is new per call -- measured in #308/#313/#306 and again here).
 *   - READ: the product's read_text / read_images do NOT see a shape (0/2). doc.GetAllShapes() does, with the shape's own text (GetDocContent) and size (GetWidth/GetHeight, EMU-exact); the driver's page-side
 *     walk (docx-shapes-read.cjs) turns that into `document.shapes` rows { text, widthEmu, heightEmu, paragraph, previous } -- the L1 below is built on THAT, never on the empty paragraph alone.
 *   NOT measured, and therefore NOT claimed: the shape's stroke COLOUR and the exact preset round-trip through the product read (the LIVE harness reads those from the stored XML), floating / wrapped
 *   placement (MEASURED since, see PLACEMENT below), header / footer / table-cell targets, shapes inside a text box, a stroke width other than 19050 (the one measured), multi-paragraph shape text.
 *
 * PLACEMENT (GH #309 part 3, LIVE: test/live/docx-gh309-floating-probe.live.cjs, five variants + an inline control): an optional shape.placement { wrap, horizontal, vertical } is applied
 *   on the shape BEFORE AddDrawing with SetWrappingStyle / SetHorAlign or SetHorPosition / SetVerPosition (each returned true); the stored drawing becomes a wp:anchor with the wrap element, behindDoc
 *   and positionH / positionV as requested (EMU byte-exact), and the shape still reads back from its own EMPTY paragraph right after the anchor (the retry identity below is unchanged). The read-back
 *   is the ToJSON projection in docx-shapes-read.cjs; the retry identity ALSO compares the placement (an inline request is not SATISFIED by a floating shape, nor the other way round).
 *
 * ADDRESSING: `{ afterParagraph }` -- the SAME unique body-level paragraph / heading contract insert_table and insert_image use.
 *
 * RETRY IDENTITY -- a NAMED, ACCEPTED LIMIT, not hidden (the SAME class insert_image's header states): SATISFIED when the block right after the anchor is an EMPTY paragraph AND a shape row sits in an
 * empty paragraph whose `previous` is the anchor with the requested text and size (+-1 EMU). That cannot tell two inserts of the SAME text and size at the SAME anchor from one call that already landed.
 * When the shape list is UNREADABLE the classification FAILS CLOSED: a retry that guessed 'no shape yet' could duplicate one.
 *
 * Preservation placement: PRIMITIVE_INTEGRAL. Every OTHER shape is preserved as a multiset, every image is unchanged (multiset), exactly ONE new shape appears, and the body block sequence is unchanged except the
 * ONE new empty paragraph immediately after the anchor. The anchor block alone gets the mso-border-*-alt:none equivalence (the #322 lesson: the DS rewrites that on a re-save) -- every other block, and every
 * other field of the anchor, compares byte-exact.
 */

const TARGET_KINDS = Object.freeze(['paragraph', 'heading']);
const PRESETS = Object.freeze(['rect', 'flowChartProcess', 'roundRect', 'ellipse', 'triangle', 'trapezoid', 'flowChartDecision', 'flowChartTerminator', 'rightArrow', 'chevron', 'homePlate']);
const EMU_PER_MM = 36000;
const MAX_MM = 400;
const STROKE_WIDTH_EMU = 19050; // the one measured stroke width
const MAX_TEXT = 500;
const HEX_RE = /^[0-9A-Fa-f]{6}$/;
const SHAPE_KEYS = Object.freeze(['fill', 'heightMm', 'stroke', 'text', 'type', 'widthMm']);
const OPTIONAL_SHAPE_KEYS = Object.freeze(['placement']);
// GH #309 part 3: the MEASURED floating placements (test/live/docx-gh309-floating-probe.live.cjs) -> what ToJSON reads back. 'tight' is NOT offered: LIVE-measured, the engine stores it with
// behindDoc="1" (the shape behind the text), which is not what a caller asking for a tight wrap means. 'through' was not measured.
const WRAPS = Object.freeze({
  square: Object.freeze({ wrapType: 'square', behindDoc: false }),
  topAndBottom: Object.freeze({ wrapType: 'top_and_bottom', behindDoc: false }),
  inFront: Object.freeze({ wrapType: 'none', behindDoc: false }),
  behind: Object.freeze({ wrapType: 'none', behindDoc: true }),
});
const H_ALIGNS = Object.freeze(['left', 'center', 'right']);
const V_RELATIVE = Object.freeze(['paragraph', 'page']);
const isOffsetMm = v => Number.isInteger(v) && v >= 0 && v <= MAX_MM;
const keysAre = (o, keys) => !!o && typeof o === 'object' && !Array.isArray(o) && Object.keys(o).sort().join(',') === [...keys].sort().join(',');

const isMm = v => Number.isInteger(v) && v >= 1 && v <= MAX_MM;
// a control character in the shape text is refused: only a single plain text run was measured
const PLAIN_TEXT_RE = /^[^\u0000-\u001f\u007f]+$/;

function plan(step, index) {
  if (step.intent !== 'insert_shape') return { ok: false, detail: `step ${index}: ${step.intent} is not a shape-insert intent` };
  const sel = step.selector;
  if (!sel || typeof sel !== 'object' || Object.keys(sel).sort().join(',') !== 'afterParagraph') {
    return { ok: false, detail: `step ${index}: insert_shape requires a selector of exactly { afterParagraph }` };
  }
  if (typeof sel.afterParagraph !== 'string' || !norm(sel.afterParagraph)) {
    return { ok: false, detail: `step ${index}: insert_shape selector.afterParagraph must be a non-empty string` };
  }
  const sh = step.shape;
  if (!sh || typeof sh !== 'object' || Array.isArray(sh)) return { ok: false, detail: `step ${index}: insert_shape requires a shape object` };
  const extra = Object.keys(sh).filter(k => !SHAPE_KEYS.includes(k));
  if (!SHAPE_KEYS.every(k => k in sh) || extra.some(k => !OPTIONAL_SHAPE_KEYS.includes(k))) {
    return { ok: false, detail: `step ${index}: insert_shape shape must have exactly: ${SHAPE_KEYS.join(', ')} (and optionally ${OPTIONAL_SHAPE_KEYS.join(', ')})` };
  }
  if (!PRESETS.includes(sh.type)) return { ok: false, detail: `step ${index}: insert_shape shape.type must be one of ${PRESETS.join(', ')} (only the measured presets)` };
  if (!isMm(sh.widthMm) || !isMm(sh.heightMm)) return { ok: false, detail: `step ${index}: insert_shape shape.widthMm / heightMm must be integers 1..${MAX_MM}` };
  if (typeof sh.fill !== 'string' || !HEX_RE.test(sh.fill)) return { ok: false, detail: `step ${index}: insert_shape shape.fill must be a 6-digit hex colour (RRGGBB)` };
  if (typeof sh.stroke !== 'string' || !HEX_RE.test(sh.stroke)) return { ok: false, detail: `step ${index}: insert_shape shape.stroke must be a 6-digit hex colour (RRGGBB)` };
  if (typeof sh.text !== 'string' || !norm(sh.text) || sh.text.length > MAX_TEXT || !PLAIN_TEXT_RE.test(sh.text)) {
    return { ok: false, detail: `step ${index}: insert_shape shape.text must be a non-empty single-line string of at most ${MAX_TEXT} characters (no control characters)` };
  }
  const placement = 'placement' in sh ? planPlacement(sh.placement) : { ok: true, value: null };
  if (!placement.ok) return { ok: false, detail: `step ${index}: insert_shape shape.placement ${placement.detail}` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: step.intent, selector: Object.freeze({ afterParagraph: sel.afterParagraph }),
      preservationMode: 'PRIMITIVE_INTEGRAL', preset: sh.type, widthEmu: sh.widthMm * EMU_PER_MM, heightEmu: sh.heightMm * EMU_PER_MM,
      fill: sh.fill.toUpperCase(), stroke: sh.stroke.toUpperCase(), text: sh.text, placement: placement.value,
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

/**
 * placement { wrap, horizontal, vertical } -> the frozen expected read-back, or a refusal. horizontal is { relativeTo: 'margin', align } or { relativeTo: 'page', offsetMm }; vertical is
 * { relativeTo: 'paragraph' | 'page', offsetMm } -- exactly the measured combinations (a vertical align, a margin offset or a column / character reference were not measured).
 */
function planPlacement(pl) {
  if (!keysAre(pl, ['wrap', 'horizontal', 'vertical'])) return { ok: false, detail: 'must be an object of exactly { wrap, horizontal, vertical }' };
  if (!Object.prototype.hasOwnProperty.call(WRAPS, pl.wrap)) return { ok: false, detail: `wrap must be one of ${Object.keys(WRAPS).join(', ')} (the measured wrap styles)` };
  const h = pl.horizontal;
  let hor;
  if (keysAre(h, ['relativeTo', 'align']) && h.relativeTo === 'margin' && H_ALIGNS.includes(h.align)) hor = { hRel: 'margin', hAlign: h.align, hOffsetEmu: null };
  else if (keysAre(h, ['relativeTo', 'offsetMm']) && h.relativeTo === 'page' && isOffsetMm(h.offsetMm)) hor = { hRel: 'page', hAlign: null, hOffsetEmu: h.offsetMm * EMU_PER_MM };
  else return { ok: false, detail: `horizontal must be { relativeTo: 'margin', align: ${H_ALIGNS.join(' | ')} } or { relativeTo: 'page', offsetMm: 0..${MAX_MM} }` };
  const v = pl.vertical;
  if (!keysAre(v, ['relativeTo', 'offsetMm']) || !V_RELATIVE.includes(v.relativeTo) || !isOffsetMm(v.offsetMm)) {
    return { ok: false, detail: `vertical must be { relativeTo: ${V_RELATIVE.join(' | ')}, offsetMm: 0..${MAX_MM} }` };
  }
  return { ok: true, value: Object.freeze({ wrap: pl.wrap, ...WRAPS[pl.wrap], ...hor, vRel: v.relativeTo, vOffsetEmu: v.offsetMm * EMU_PER_MM }) };
}

// ---- the observation ------------------------------------------------------------------------------

const blocksOf = primary => (primary && primary.blocks) || [];
const fieldReadable = v => !(typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0);

/** document.shapes as an array, or null when the driver could not read it (a string / absent). */
function shapeRows(primary) {
  const d = primary && primary.document;
  const rows = d ? d.shapes : undefined;
  return Array.isArray(rows) ? rows : null;
}

function resolveAnchor(primary, text) {
  const r = resolveStableTarget(primary, { text });
  if (!r.ok) return r;
  if (!TARGET_KINDS.includes(r.target.kind)) {
    return { ok: false, target: null, selector: r.selector, ambiguity: 'UNSUPPORTED_TARGET_KIND', kind: r.target.kind };
  }
  return r;
}

/** +-1 EMU on both sides: the SAME tolerance insert_image keeps (this DS read back the requested EMU exactly in the probe, so this only absorbs a rounding drift, never a wrong size). */
const sizeMatches = (row, op) => fieldReadable(row.widthEmu) && fieldReadable(row.heightEmu)
  && Math.abs(row.widthEmu - op.widthEmu) <= 1 && Math.abs(row.heightEmu - op.heightEmu) <= 1;

const placementReadable = row => ['drawingType', 'wrapType', 'behindDoc', 'hRel', 'hAlign', 'hOffset', 'vRel', 'vAlign', 'vOffset'].every(k => fieldReadable(row[k]));

/** No placement -> the shape must read back INLINE (drawingType is the only field that tells inline apart: an inline shape reads wrapType 'through'). A placement -> 'anchor' + every expected field. */
function placementMatches(row, op) {
  if (!placementReadable(row)) return false;
  const p = op.placement;
  if (!p) return row.drawingType === 'inline';
  return row.drawingType === 'anchor' && row.wrapType === p.wrapType && row.behindDoc === p.behindDoc && row.hRel === p.hRel
    && (p.hAlign ? row.hAlign === p.hAlign : row.hAlign === null && row.hOffset === p.hOffsetEmu)
    && row.vRel === p.vRel && row.vAlign === null && row.vOffset === p.vOffsetEmu;
}

/** The shape's own text and size, in an EMPTY paragraph whose preceding body element is the anchor -- placement NOT compared (see classify: an unreadable placement there fails closed). */
const rowAt = (row, anchorText, op) => fieldReadable(row.text) && row.text === op.text && fieldReadable(row.paragraph) && norm(row.paragraph) === ''
  && fieldReadable(row.previous) && row.previous !== null && norm(row.previous) === norm(anchorText) && sizeMatches(row, op);

/** The shape this op would create: rowAt + its placement (inline, or the requested floating placement). */
const rowMatches = (row, anchorText, op) => rowAt(row, anchorText, op) && placementMatches(row, op);

/** See the header's RETRY IDENTITY note. `rows` null (unreadable) is handled by the callers (fail closed), never treated as 'none'. */
function matchingShapeAfter(primary, anchorIndex, op, anchorText) {
  const blocks = blocksOf(primary);
  const next = blocks[anchorIndex + 1];
  if (!next || next.kind !== 'paragraph' || norm(next.text) !== '') return false;
  const rows = shapeRows(primary);
  return Array.isArray(rows) && rows.some(r => rowMatches(r, anchorText, op));
}

// ---- classify / resolve ----------------------------------------------------------------------------

function classify(primary, op) {
  const anchor = resolveAnchor(primary, op.selector.afterParagraph);
  if (!anchor.ok) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `insert-shape-anchor-${String(anchor.ambiguity).toLowerCase().replace(/_/g, '-')}` } };
  }
  if (shapeRows(primary) === null) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'insert-shape-shapes-unobservable', detail: 'the shape list could not be read, so a retry could not tell whether the shape already exists' } };
  }
  if (shapeRows(primary).some(r => rowAt(r, anchor.target.text, op) && !placementReadable(r))) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'insert-shape-placement-unobservable', detail: 'a shape with this text and size sits after the anchor, but its placement could not be read' } };
  }
  const anchorIndex = blocksOf(primary).indexOf(anchor.target);
  if (matchingShapeAfter(primary, anchorIndex, op, anchor.target.text)) {
    return { state: 'SATISFIED', evidence: { anchor: anchor.target.text, preset: op.preset, widthEmu: op.widthEmu, heightEmu: op.heightEmu, proof: 'fresh-live-shape-with-own-text-in-empty-paragraph-immediately-after-anchor' } };
  }
  return { state: 'APPLY', evidence: { anchor: anchor.target.text } };
}

function resolve(primary, selector) {
  const sel = selector || {};
  const anchor = resolveAnchor(primary, sel.afterParagraph);
  if (!anchor.ok) return { ok: false, target: null, selector, ambiguity: anchor.ambiguity };
  return { ok: true, target: { kind: 'insert_shape_anchor', anchorText: norm(anchor.target.text) }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && target.kind === 'insert_shape_anchor' && target.anchorText === norm(op.selector.afterParagraph)
    ? { ok: true }
    : { ok: false, dispatched: false, detail: 'insert-shape-anchor-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({
    family: 'shape-insert', action: 'insert', anchorText: target.anchorText, preset: op.preset, widthEmu: op.widthEmu, heightEmu: op.heightEmu,
    fill: op.fill, stroke: op.stroke, strokeWidthEmu: STROKE_WIDTH_EMU, text: op.text, placement: op.placement || null,
  });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (the LIVE probe, both presets PASS). Re-locates the anchor by the SAME text-matching walk every other DOCX write family uses (an identity walk never matches a new-per-call
 * wrapper), re-checked for uniqueness inside the sandbox. The shape and its text are built BEFORE anything touches the document: any guard miss or throw returns with the document unchanged (the
 * insert is the LAST step). Then a fresh paragraph carries `AddDrawing(shape)` and lands at `doc.AddElement(anchorIndex + 1, ...)`.
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
  for (let i = 0; i < n; i += 1) {
    const e = doc.GetElement(i);
    if (e && typeof e.GetClassType === 'function' && e.GetClassType() === 'paragraph' && lookupTextOf(e) === m.anchorText) {
      if (anchorIndex >= 0) return; // ambiguous at dispatch time: refuse rather than guess
      anchorIndex = i;
    }
  }
  if (anchorIndex < 0) return;
  if (typeof Api.CreateShape !== 'function' || typeof Api.CreateSolidFill !== 'function' || typeof Api.CreateRGBColor !== 'function' || typeof Api.CreateStroke !== 'function') return;
  if (typeof doc.AddElement !== 'function') return;
  const rgb = hex => Api.CreateRGBColor(parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16));
  const fill = Api.CreateSolidFill(rgb(m.fill));
  const stroke = Api.CreateStroke(m.strokeWidthEmu, Api.CreateSolidFill(rgb(m.stroke)));
  const shape = Api.CreateShape(m.preset, m.widthEmu, m.heightEmu, fill, stroke);
  if (!shape || typeof shape.GetDocContent !== 'function') return;
  const content = shape.GetDocContent();
  const para = content ? content.GetElement(0) : null;
  if (!para || typeof para.AddText !== 'function') return;
  para.AddText(m.text);
  const p = m.placement;
  if (p) {
    if (typeof shape.SetWrappingStyle !== 'function' || typeof shape.SetVerPosition !== 'function') return;
    if (p.hAlign ? typeof shape.SetHorAlign !== 'function' : typeof shape.SetHorPosition !== 'function') return;
    if (shape.SetWrappingStyle(p.wrap) === false) return;
    if ((p.hAlign ? shape.SetHorAlign(p.hRel, p.hAlign) : shape.SetHorPosition(p.hRel, p.hOffsetEmu)) === false) return;
    if (shape.SetVerPosition(p.vRel, p.vOffsetEmu) === false) return;
  }
  const np = Api.CreateParagraph();
  if (typeof np.AddDrawing !== 'function') return;
  np.AddDrawing(shape);
  doc.AddElement(anchorIndex + 1, np);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

function verifyPrimitive(op, before, after) {
  const anchorAfter = resolveAnchor(after && after.primary, op.selector.afterParagraph);
  if (!anchorAfter.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the after-state' } };
  if (shapeRows(after.primary) === null) return { ok: false, evidence: { reason: 'the shape list is unobservable in the after-state' } };
  const anchorIndexAfter = blocksOf(after.primary).indexOf(anchorAfter.target);
  const ok = matchingShapeAfter(after.primary, anchorIndexAfter, op, anchorAfter.target.text);
  return { ok, evidence: { anchor: anchorAfter.target.text, anchorIndexAfter, preset: op.preset } };
}

/** A shape's own key (text / size / own paragraph) -- NOT `previous`: inserting after the anchor moves the preceding-element text of an existing shape that followed the anchor, which is not a change to that shape. */
const shapeKey = r => JSON.stringify([r.text, r.widthEmu, r.heightEmu, r.paragraph, r.drawingType, r.wrapType, r.behindDoc, r.hRel, r.hAlign, r.hOffset, r.vRel, r.vAlign, r.vOffset]);

/** Every BEFORE shape still exists AFTER, by key (a multiset subtraction) -- insert never loses or mutates an existing shape. */
function everyBeforeShapePreserved(beforeRows, afterRows) {
  const pool = afterRows.map(shapeKey);
  for (const r of beforeRows) {
    const i = pool.indexOf(shapeKey(r));
    if (i < 0) return false;
    pool.splice(i, 1);
  }
  return true;
}

/** Same count and same members: images are not touched by this write at all. */
function imagesUnchanged(beforeRows, afterRows) {
  return beforeRows.length === afterRows.length && everyBeforeImagePreserved(beforeRows, afterRows);
}

/**
 * PRIMITIVE_INTEGRAL: every OTHER shape preserved (multiset) AND exactly one new shape appears; every image unchanged; the body block sequence is unchanged except the ONE empty paragraph this write creates,
 * immediately after the anchor (the SAME block-sequence discipline table-insert.cjs / image-insert.cjs use, `canonicalBlock` REUSED). Only the ANCHOR block gets the border-alt equivalence.
 */
function verifyPreservation(op, before, after) {
  const anchorBefore = resolveAnchor(before && before.primary, op.selector.afterParagraph);
  if (!anchorBefore.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the before-state' } };
  const beforeShapes = shapeRows(before.primary);
  const afterShapes = shapeRows(after.primary);
  if (beforeShapes === null || afterShapes === null) return { ok: false, evidence: { reason: 'the shape list is unobservable (before or after)' } };
  const shapesPreserved = everyBeforeShapePreserved(beforeShapes, afterShapes);
  const countOk = afterShapes.length === beforeShapes.length + 1;
  const imagesOk = imagesUnchanged(imageRows(before.primary) || [], imageRows(after.primary) || []);

  const beforeBlocks = blocksOf(before.primary);
  const afterBlocks = blocksOf(after.primary);
  const anchorIndex = beforeBlocks.indexOf(anchorBefore.target);

  const newBlock = afterBlocks[anchorIndex + 1];
  const newBlockOk = !!newBlock && newBlock.kind === 'paragraph' && norm(newBlock.text) === '';
  // the anchor is the LAST entry of each prefix slice -- only that one gets the border-alt equivalence
  const prefixKey = blocks => blocks.map((b, i) => canonicalBlock(b, i === blocks.length - 1));
  const prefixSame = JSON.stringify(prefixKey(beforeBlocks.slice(0, anchorIndex + 1))) === JSON.stringify(prefixKey(afterBlocks.slice(0, anchorIndex + 1)));
  const restSame = JSON.stringify(beforeBlocks.slice(anchorIndex + 1).map(b => canonicalBlock(b))) === JSON.stringify(afterBlocks.slice(anchorIndex + 2).map(b => canonicalBlock(b)));
  const lengthOk = afterBlocks.length === beforeBlocks.length + 1;
  const ok = shapesPreserved && countOk && imagesOk && newBlockOk && prefixSame && restSame && lengthOk;
  return { ok, evidence: { shapesPreserved, countOk, imagesOk, newBlockOk, prefixSame, restSame, lengthOk } };
}

function wholeTaskChecks(op, finalObservation) {
  const anchor = resolveAnchor(finalObservation.primary, op.selector.afterParagraph);
  let ok = false;
  if (anchor.ok) {
    const idx = blocksOf(finalObservation.primary).indexOf(anchor.target);
    ok = matchingShapeAfter(finalObservation.primary, idx, op, anchor.target.text);
  }
  const where = op.placement ? `floating (${op.placement.wrap})` : 'inline';
  return [{ name: `${op.id}: a ${where} ${op.preset} shape (${op.widthEmu}x${op.heightEmu} EMU) with the text "${op.text}" sits immediately after "${op.selector.afterParagraph}"`, ok }];
}

// refs: none -- the shape's text is not a body text unit a later step can name
function producedText() {
  return null;
}

module.exports = {
  plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  resolveAnchor, shapeRows, rowMatches, rowAt, placementMatches, planPlacement, WRAPS, matchingShapeAfter, sizeMatches, everyBeforeShapePreserved, shapeKey, PRESETS, EMU_PER_MM, MAX_MM, MAX_TEXT, STROKE_WIDTH_EMU,
};
