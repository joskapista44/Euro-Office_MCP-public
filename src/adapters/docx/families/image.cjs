'use strict';

const { norm, textUnits, tablesContent } = require('../inventory.cjs');

/**
 * DOCX `image` family (GH #268c, post-M7 NEW_CAPABILITY): set_image_size, set_image_wrap.
 *
 * Source: docs/GH268-DOCX-IMAGES-READ-INVENTORY-WRITE-RESEARCH.md section 3 (static sdkjs reading) +
 * test/live/docx-images-probe.live.cjs (the measured LIVE run, R3 SetSize / R4 SetWrappingStyle). No
 * joskapista44 source oracle exists for images (the source never covered this surface) -- this is a
 * NEW_CAPABILITY family, same placement rule as set_page_setup/format_table above it.
 *
 * ADDRESSING (owner decision, GH #268c's own decision doc section "Addressing"): the
 * read_images row's `docPr.id` / `name` are EDITOR-SESSION-LOCAL (g268v1 LIVE) and were dropped from
 * the row entirely (GH #268b's own drop-id-name commit) -- so neither this family nor its caller may
 * address an image by id/name. Two selector shapes instead, exactly one required:
 *   { paragraph, ordinal? }  the `ordinal`-th (1-based, default 1) image whose PARENT PARAGRAPH text
 *                            (exact, `norm()`-compared) equals `paragraph`, counted in document order
 *                            across every paragraph with that text (ordinal disambiguates images that
 *                            share one paragraph; it is not a claim that the paragraph itself is unique)
 *   { index }                the document-level index read_images itself returns (`GetAllImages()`
 *                            order) -- valid for ONE read only, same as read_images' own scope note
 * Both shapes are re-resolved against a FRESH read_images observation at classify/resolve time (the
 * adapter's own per-operation observation, never a cached one) -- "verified with a fresh read_images
 * before write" falls out of the framework's existing fresh-observation contract, the same way
 * header-footer.cjs's `section: {index, of}` re-checks the section count on every call.
 *
 * SCOPE (what this v1 ships, what it refuses):
 *   set_image_size   SetSize(widthEmu, heightEmu) -- LIVE-measured (R3, docx-images-probe.live.cjs):
 *                     applied to an INLINE image (imgs[0]), reopened, both widthEmu/heightEmu
 *                     reflect the new value, the getters agree with the ToJSON extent. The static
 *                     source shows ONE codepath (getParaDrawing().setExtent + the xfrm ext) with no
 *                     placement branch, so this release does not restrict by placement -- but only
 *                     the inline case is LIVE-proven; the decision doc section 2 states this plainly.
 *                     No aspect-ratio lock, no crop (research doc section 7 point 2).
 *   set_image_wrap    SetWrappingStyle(wrap) -- ONLY `wrap: 'square'` is accepted, and only FROM an
 *                     inline source (R4 measured exactly that one transition: inline -> anchor via
 *                     'square'). The other six source values (tight/through/topAndBottom/behind/
 *                     inFront, and re-wrapping an already-anchored image, or reverting anchor ->
 *                     inline) are NOT measured -- refused at plan time (schema) and, defensively,
 *                     refused again inside the callCommand (the `before !== 'inline'` guard) rather
 *                     than assumed. No position-setting (research doc section 7 point 3).
 */

const INTENTS = Object.freeze(['set_image_size', 'set_image_wrap']);
const SUPPORTED_WRAP = Object.freeze(['square']);

function plan(step, index) {
  if (!INTENTS.includes(step.intent)) return { ok: false, detail: `step ${index}: ${step.intent} is not an image intent` };
  const sel = step.selector;
  if (!sel || typeof sel !== 'object') return { ok: false, detail: `step ${index}: ${step.intent} requires a selector` };
  const hasIndex = sel.index !== undefined;
  const hasParagraph = sel.paragraph !== undefined;
  if (hasIndex === hasParagraph) return { ok: false, detail: `step ${index}: ${step.intent} selector must name exactly one of index or paragraph` };
  let selector;
  if (hasIndex) {
    if (sel.ordinal !== undefined) return { ok: false, detail: `step ${index}: ${step.intent} selector.ordinal is only valid with paragraph, not index` };
    if (!Number.isInteger(sel.index) || sel.index < 0) return { ok: false, detail: `step ${index}: ${step.intent} selector.index must be a non-negative integer` };
    selector = Object.freeze({ index: sel.index });
  } else {
    if (typeof sel.paragraph !== 'string' || !norm(sel.paragraph)) return { ok: false, detail: `step ${index}: ${step.intent} selector.paragraph must be a non-empty string` };
    const ordinal = sel.ordinal === undefined ? 1 : sel.ordinal;
    if (!Number.isInteger(ordinal) || ordinal < 1) return { ok: false, detail: `step ${index}: ${step.intent} selector.ordinal must be a positive integer` };
    selector = Object.freeze({ paragraph: sel.paragraph, ordinal });
  }
  const base = { id: step.id || `s${index}`, intent: step.intent, selector, preservationMode: 'PRIMITIVE_INTEGRAL', ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) };
  if (step.intent === 'set_image_size') {
    if (!Number.isInteger(step.widthEmu) || step.widthEmu < 1) return { ok: false, detail: `step ${index}: set_image_size requires a positive integer widthEmu` };
    if (!Number.isInteger(step.heightEmu) || step.heightEmu < 1) return { ok: false, detail: `step ${index}: set_image_size requires a positive integer heightEmu` };
    return { ok: true, operation: { ...base, widthEmu: step.widthEmu, heightEmu: step.heightEmu } };
  }
  if (!SUPPORTED_WRAP.includes(step.wrap)) return { ok: false, detail: `step ${index}: set_image_wrap only supports wrap: ${SUPPORTED_WRAP.join(', ')} in this release (the other SetWrappingStyle values are NEM-MÉRT)` };
  return { ok: true, operation: { ...base, wrap: step.wrap } };
}

// ---- the observation ----------------------------------------------------------------------------

const fieldReadable = v => !(typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0);

/** The images read, or null (unread / genuinely absent list). */
function imageRows(primary) {
  const d = primary && primary.document;
  const imgs = d ? d.images : undefined;
  return Array.isArray(imgs) ? imgs : null;
}

/** The ONE row op.selector names, from a FRESH observation -- never cached across calls. */
function targetRow(primary, op) {
  const rows = imageRows(primary);
  if (!rows) {
    const d = primary && primary.document;
    return { ok: false, reason: typeof (d && d.images) === 'string' ? d.images : 'images-unreadable' };
  }
  if (op.selector.index !== undefined) {
    const row = rows.find(r => r.index === op.selector.index);
    return row ? { ok: true, row, rows } : { ok: false, reason: 'image-index-not-found', rows };
  }
  const want = norm(op.selector.paragraph);
  const matches = rows.filter(r => fieldReadable(r.paragraph) && norm(r.paragraph) === want);
  const row = matches[op.selector.ordinal - 1];
  return row ? { ok: true, row, rows } : { ok: false, reason: 'image-paragraph-ordinal-not-found', matchCount: matches.length, rows };
}

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  const t = targetRow(primary, op);
  if (!t.ok) return { state: 'FAIL_CLOSED', evidence: { reason: t.reason, ...(t.matchCount !== undefined ? { matchCount: t.matchCount } : {}) } };
  const row = t.row;
  if (op.intent === 'set_image_size') {
    if (!fieldReadable(row.widthEmu) || !fieldReadable(row.heightEmu)) return { state: 'FAIL_CLOSED', evidence: { reason: 'image-size-unreadable', index: row.index } };
    if (row.widthEmu === op.widthEmu && row.heightEmu === op.heightEmu) return { state: 'SATISFIED', evidence: { index: row.index, widthEmu: row.widthEmu, heightEmu: row.heightEmu } };
    return { state: 'APPLY', evidence: { index: row.index } };
  }
  if (!fieldReadable(row.placement) || !fieldReadable(row.wrap)) return { state: 'FAIL_CLOSED', evidence: { reason: 'image-wrap-unreadable', index: row.index } };
  if (row.placement === 'floating' && row.wrap === op.wrap) return { state: 'SATISFIED', evidence: { index: row.index, placement: row.placement, wrap: row.wrap } };
  if (row.placement !== 'inline') return { state: 'FAIL_CLOSED', evidence: { reason: 'image-wrap-unmeasured-source-placement', index: row.index, placement: row.placement } };
  return { state: 'APPLY', evidence: { index: row.index } };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  if (!op || !INTENTS.includes(op.intent)) return { ok: false, target: null, selector, ambiguity: 'INVALID_SELECTOR' };
  const up = r => String(r).toUpperCase().replace(/[:-]/g, '_');
  const c = classify(primary, op);
  if (c.state === 'FAIL_CLOSED') return { ok: false, target: null, selector, ambiguity: up(c.evidence.reason) };
  const t = targetRow(primary, op);
  return { ok: true, target: { kind: 'image', imageIndex: t.row.index, expectedParagraph: op.selector.paragraph !== undefined ? (fieldReadable(t.row.paragraph) ? t.row.paragraph : null) : null }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && target.kind === 'image' && Number.isInteger(target.imageIndex) ? { ok: true } : { ok: false, dispatched: false, detail: 'image-target-unresolved' };
}

// ---- lower ----------------------------------------------------------------------------------------

/**
 * One callCommand, public API only. It re-locates the image by the SAME document-order index the
 * fresh resolve() just read, re-checks the parent paragraph text when the selector named one (the
 * fresh-read-before-write guarantee re-stated inside the sandbox, not only outside it), and -- for
 * set_image_wrap -- re-checks the source is still inline (the only LIVE-measured transition) before
 * calling SetWrappingStyle. Any guard miss returns unchanged.
 */
function lower(op, target) {
  const meta = Object.freeze({
    family: 'image', action: op.intent === 'set_image_size' ? 'imageSetSize' : 'imageSetWrap',
    imageIndex: target.imageIndex, expectedParagraph: target.expectedParagraph,
    ...(op.intent === 'set_image_size' ? { widthEmu: op.widthEmu, heightEmu: op.heightEmu } : { wrap: op.wrap }),
  });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  const imgs = doc.GetAllImages() || [];
  const im = imgs[m.imageIndex];
  if (!im) return;
  if (m.expectedParagraph !== null) {
    let pText = null;
    try {
      const p = im.GetParentParagraph();
      if (p) { const t = String(p.GetText({ Numbering: false })); let n = t.length; while (n > 0 && (t.charCodeAt(n - 1) === 10 || t.charCodeAt(n - 1) === 13)) n -= 1; pText = t.slice(0, n); }
    } catch (e) {}
    if (pText !== m.expectedParagraph) return;
  }
  if (m.action === 'imageSetSize') {
    im.SetSize(m.widthEmu, m.heightEmu);
  } else {
    let before = null;
    try { before = JSON.parse(im.ToJSON(false, false)).drawingType; } catch (e) {}
    if (before !== 'inline') return;
    im.SetWrappingStyle(m.wrap);
  }
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** Every image field the written intent does not itself control, as a comparable key. */
function imageKey(r) {
  return JSON.stringify([r.alt, r.title, r.placement, r.wrap, r.widthEmu, r.heightEmu, r.paragraph]);
}

function otherImages(rows, targetIndex) {
  const out = {};
  for (const r of rows || []) if (r.index !== targetIndex) out[r.index] = imageKey(r);
  return out;
}

/** The main body as the reads see it, EXCLUDING images (those have their own, per-image comparison above). */
function bodyKey(primary) {
  const d = (primary && primary.document) || {};
  return JSON.stringify({ units: textUnits(primary).map(u => [u.kind, norm(u.text)]), tables: tablesContent(d.tables), footnotes: d.footnotes || null, endnotes: d.endnotes || null, bookmarks: d.bookmarks || null, inMainBody: d.inMainBody === undefined ? null : d.inMainBody });
}

function verifyPrimitive(op, before, after) {
  const b = targetRow(before && before.primary, op);
  const a = targetRow(after && after.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const row = a.row;
  if (op.intent === 'set_image_size') {
    const ok = row.widthEmu === op.widthEmu && row.heightEmu === op.heightEmu;
    return { ok, evidence: { index: row.index, widthEmu: row.widthEmu, heightEmu: row.heightEmu } };
  }
  const ok = row.placement === 'floating' && row.wrap === op.wrap;
  return { ok, evidence: { index: row.index, placement: row.placement, wrap: row.wrap } };
}

/** PRIMITIVE_INTEGRAL: every OTHER image, and the rest of the body, unchanged; the image count unchanged. */
function verifyPreservation(op, before, after) {
  const b = targetRow(before && before.primary, op);
  const a = targetRow(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` } };
  const bRows = imageRows(before.primary) || [];
  const aRows = imageRows(after.primary) || [];
  const pb = otherImages(bRows, b.row.index);
  const pa = otherImages(aRows, a.row.index);
  const keys = new Set([...Object.keys(pb), ...Object.keys(pa)]);
  const changedImages = [...keys].filter(k => pb[k] !== pa[k]);
  const sameCount = bRows.length === aRows.length;
  const sameBody = bodyKey(before.primary) === bodyKey(after.primary);
  const ok = changedImages.length === 0 && sameCount && sameBody;
  return { ok, evidence: { changedImages, imageCountBefore: bRows.length, imageCountAfter: aRows.length, bodyUnchanged: sameBody } };
}

function wholeTaskChecks(op, finalObservation) {
  const t = targetRow(finalObservation.primary, op);
  if (op.intent === 'set_image_size') {
    return [{ name: `${op.id}: image ${t.ok ? t.row.index : '?'} size = ${op.widthEmu}x${op.heightEmu} EMU`, ok: t.ok && t.row.widthEmu === op.widthEmu && t.row.heightEmu === op.heightEmu }];
  }
  return [{ name: `${op.id}: image ${t.ok ? t.row.index : '?'} wrap = floating/${op.wrap}`, ok: t.ok && t.row.placement === 'floating' && t.row.wrap === op.wrap }];
}

// refs: none -- an image is not a body text unit a later step can name
function producedText() {
  return null;
}

module.exports = { plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, targetRow, imageRows, otherImages, bodyKey, SUPPORTED_WRAP, INTENTS };
