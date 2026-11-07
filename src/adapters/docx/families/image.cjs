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
 *
 * GH #296: delete_image and replace_image, on the SAME two selector shapes plus an optional
 * `selector.part` assertion ('body' | 'header:<s>:<default|title|even>' | 'footer:...', the
 * read_images row's own GH #306b `part` label -- `index` counts header/footer images too, so a caller
 * that wants to be sure which part it hits names it; a mismatch is FAIL_CLOSED, never re-targeted).
 * LIVE-measured (test/live/docx-gh296-image-replace-delete-probe.live.cjs, run p296-r1, 6/6 PASS, and
 * the body rounds of docx-images-probe.live.cjs R5/R6):
 *   delete_image    ApiDrawing.Delete() -- header default + even (inline) and a body floating image:
 *                   the drawing leaves the stored part, the part's own text stays, every other image
 *                   and the body are unchanged.
 *   replace_image   doc.ReplaceDrawing(old, Api.CreateImage(source, w, h), true) -- footer default
 *                   (PNG -> JPEG, old size), header title (new size), body inline and body floating
 *                   (the anchor/square placement kept): the stored media bytes sha256-match the new
 *                   input, the position, placement and every other image unchanged. `source` is a
 *                   data:image/png|jpeg base64 URI (insert_image's own contract); widthEmu/heightEmu
 *                   are both given or both omitted (omitted = the image keeps its current size).
 * Both re-check the image inside the sandbox (paragraph text, extent, drawingType -- the fingerprint the
 * fresh resolve just read) and return a RECEIPT ({ ok: true } only when every guard passed AND the
 * engine call returned true; the driver's `publicReceipt` rule). VERIFY LIMIT, named: the in-session
 * read carries no image bytes, so L1 for replace_image proves position, placement and size plus the
 * receipt -- the bytes themselves are proven by the LIVE harness from the stored file. A retry is never
 * SATISFIED: a landed delete_image retried is FAIL_CLOSED (the image is gone, the selector misses), a
 * landed replace_image retried replaces again with the same bytes (same end state).
 */

const INTENTS = Object.freeze(['set_image_size', 'set_image_wrap', 'delete_image', 'replace_image']);
const EDIT_INTENTS = Object.freeze(['delete_image', 'replace_image']);
const SUPPORTED_WRAP = Object.freeze(['square']);
// the SAME source contract as image-insert.cjs (that module requires this one, so the literal is repeated, not imported)
const SOURCE_RE = /^data:image\/(png|jpeg);base64,[A-Za-z0-9+/]+=*$/;
const PART_RE = /^(body|(header|footer):\d+:(default|title|even))$/;

/** The selector shape every image intent shares -> { selector } or { detail }. */
function planSelector(step, index) {
  const sel = step.selector;
  if (!sel || typeof sel !== 'object') return { detail: `step ${index}: ${step.intent} requires a selector` };
  if ((sel.index !== undefined) === (sel.paragraph !== undefined)) return { detail: `step ${index}: ${step.intent} selector must name exactly one of index or paragraph` };
  const located = sel.index !== undefined ? planIndexSelector(sel, step, index) : planParagraphSelector(sel, step, index);
  if (located.detail || sel.part === undefined) return located;
  // GH #296: the part assertion (header comment)
  if (!EDIT_INTENTS.includes(step.intent)) return { detail: `step ${index}: ${step.intent} selector.part is only valid with ${EDIT_INTENTS.join(' / ')}` };
  if (typeof sel.part !== 'string' || !PART_RE.test(sel.part)) return { detail: `step ${index}: ${step.intent} selector.part must be 'body' or '<header|footer>:<section>:<default|title|even>' (the read_images row's own part label)` };
  return { selector: Object.freeze({ ...located.selector, part: sel.part }) };
}

function planIndexSelector(sel, step, index) {
  if (sel.ordinal !== undefined) return { detail: `step ${index}: ${step.intent} selector.ordinal is only valid with paragraph, not index` };
  if (!Number.isInteger(sel.index) || sel.index < 0) return { detail: `step ${index}: ${step.intent} selector.index must be a non-negative integer` };
  return { selector: Object.freeze({ index: sel.index }) };
}

function planParagraphSelector(sel, step, index) {
  if (typeof sel.paragraph !== 'string' || !norm(sel.paragraph)) return { detail: `step ${index}: ${step.intent} selector.paragraph must be a non-empty string` };
  const ordinal = sel.ordinal === undefined ? 1 : sel.ordinal;
  if (!Number.isInteger(ordinal) || ordinal < 1) return { detail: `step ${index}: ${step.intent} selector.ordinal must be a positive integer` };
  return { selector: Object.freeze({ paragraph: sel.paragraph, ordinal }) };
}

/** GH #296: replace_image's own fields on top of the shared base. */
function planReplace(step, index, base) {
  if (typeof step.source !== 'string' || !SOURCE_RE.test(step.source)) return { ok: false, detail: `step ${index}: replace_image source must be a data:image/png;base64,... or data:image/jpeg;base64,... URI -- an external URL is never accepted` };
  const hasW = step.widthEmu !== undefined;
  if (hasW !== (step.heightEmu !== undefined)) return { ok: false, detail: `step ${index}: replace_image takes widthEmu and heightEmu together, or neither (neither = keep the current size)` };
  const positive = v => Number.isInteger(v) && v >= 1;
  if (hasW && !(positive(step.widthEmu) && positive(step.heightEmu))) return { ok: false, detail: `step ${index}: replace_image widthEmu / heightEmu must be positive integers` };
  return { ok: true, operation: { ...base, source: step.source, ...(hasW ? { widthEmu: step.widthEmu, heightEmu: step.heightEmu } : {}) } };
}

function plan(step, index) {
  if (!INTENTS.includes(step.intent)) return { ok: false, detail: `step ${index}: ${step.intent} is not an image intent` };
  const sel = planSelector(step, index);
  if (sel.detail) return { ok: false, detail: sel.detail };
  const base = { id: step.id || `s${index}`, intent: step.intent, selector: sel.selector, preservationMode: 'PRIMITIVE_INTEGRAL', ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) };
  if (step.intent === 'delete_image') return { ok: true, operation: base };
  if (step.intent === 'replace_image') return planReplace(step, index, base);
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
  let found;
  if (op.selector.index !== undefined) {
    const row = rows.find(r => r.index === op.selector.index);
    found = row ? { ok: true, row, rows } : { ok: false, reason: 'image-index-not-found', rows };
  } else {
    const want = norm(op.selector.paragraph);
    const matches = rows.filter(r => fieldReadable(r.paragraph) && norm(r.paragraph) === want);
    const row = matches[op.selector.ordinal - 1];
    found = row ? { ok: true, row, rows } : { ok: false, reason: 'image-paragraph-ordinal-not-found', matchCount: matches.length, rows };
  }
  // GH #296: selector.part is an assertion on the row the selector found, never a re-targeting
  if (found.ok && op.selector.part !== undefined && found.row.part !== op.selector.part) return { ok: false, reason: 'image-part-mismatch', rows };
  return found;
}

// ---- classify / resolve -------------------------------------------------------------------------

/** GH #296: the fingerprint the sandbox re-checks must be readable; a retry is never SATISFIED (header comment). */
function classifyEdit(row) {
  const readable = [row.paragraph, row.widthEmu, row.heightEmu, row.placement, row.part].every(fieldReadable) && (row.placement === 'inline' || row.placement === 'floating');
  return readable ? { state: 'APPLY', evidence: { index: row.index, part: row.part } } : { state: 'FAIL_CLOSED', evidence: { reason: 'image-fingerprint-unreadable', index: row.index } };
}

function classify(primary, op) {
  const t = targetRow(primary, op);
  if (!t.ok) return { state: 'FAIL_CLOSED', evidence: { reason: t.reason, ...(t.matchCount !== undefined ? { matchCount: t.matchCount } : {}) } };
  const row = t.row;
  if (EDIT_INTENTS.includes(op.intent)) return classifyEdit(row);
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
  if (EDIT_INTENTS.includes(op.intent)) {
    const r = t.row;
    return { ok: true, target: { kind: 'image', imageIndex: r.index, expectedParagraph: r.paragraph, expectedWidthEmu: r.widthEmu, expectedHeightEmu: r.heightEmu, expectedDrawingType: r.placement === 'inline' ? 'inline' : 'anchor' }, selector, ambiguity: null };
  }
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
  if (EDIT_INTENTS.includes(op.intent)) {
    const meta = Object.freeze({
      family: 'image', action: op.intent === 'delete_image' ? 'imageDelete' : 'imageReplace', publicReceipt: true,
      imageIndex: target.imageIndex, expectedParagraph: target.expectedParagraph,
      expectedWidthEmu: target.expectedWidthEmu, expectedHeightEmu: target.expectedHeightEmu, expectedDrawingType: target.expectedDrawingType,
      ...(op.intent === 'replace_image' ? { source: op.source, widthEmu: op.widthEmu === undefined ? target.expectedWidthEmu : op.widthEmu, heightEmu: op.heightEmu === undefined ? target.expectedHeightEmu : op.heightEmu } : {}),
    });
    return { kind: 'callCommand', body: editCommandBody(meta), meta };
  }
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

/**
 * GH #296: delete / replace. The same index re-location as above, then the WHOLE fingerprint the fresh
 * resolve read (paragraph text, extent, drawingType) re-checked in the sandbox -- an empty-text header
 * image has no paragraph text to tell it apart. The body RETURNS a receipt: { ok: true } only when every
 * guard passed and the engine call returned true (the driver's publicReceipt rule); anything else is a
 * failed write the core re-reads.
 */
function editCommandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  const imgs = doc.GetAllImages() || [];
  const im = imgs[m.imageIndex];
  if (!im) return { ok: false, reason: 'image-index-gone' };
  let pText = null;
  try {
    const p = im.GetParentParagraph();
    if (p) { const t = String(p.GetText({ Numbering: false })); let n = t.length; while (n > 0 && (t.charCodeAt(n - 1) === 10 || t.charCodeAt(n - 1) === 13)) n -= 1; pText = t.slice(0, n); }
  } catch (e) {}
  if (pText !== m.expectedParagraph) return { ok: false, reason: 'paragraph-mismatch' };
  let j = null;
  try { j = JSON.parse(im.ToJSON(false, false)); } catch (e) {}
  const ex = j && j.extent ? j.extent : null;
  if (!j || j.drawingType !== m.expectedDrawingType || !ex || ex.cx !== m.expectedWidthEmu || ex.cy !== m.expectedHeightEmu) return { ok: false, reason: 'fingerprint-mismatch' };
  if (m.action === 'imageDelete') {
    return im.Delete() === true ? { ok: true } : { ok: false, reason: 'delete-returned-false' };
  }
  if (typeof Api.CreateImage !== 'function') return { ok: false, reason: 'no-create-image' };
  const img = Api.CreateImage(m.source, m.widthEmu, m.heightEmu);
  if (!img) return { ok: false, reason: 'create-image-null' };
  return doc.ReplaceDrawing(im, img, true) === true ? { ok: true } : { ok: false, reason: 'replace-returned-false' };
} catch (e) { return { ok: false, reason: 'threw' }; }
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

/** GH #296: every image field the read sees, INCLUDING its part and paragraph index -- the multiset key for delete / replace. */
function placedImageKey(r) {
  return JSON.stringify([r.part, r.paragraphIndex, r.alt, r.title, r.placement, r.wrap, r.widthEmu, r.heightEmu, r.paragraph]);
}

/** before minus the target (one occurrence) vs after minus the replacement (when one is expected) -- equal multisets? */
function imagesAccountedFor(beforeRows, afterRows, removedKey, addedMatch) {
  const pool = beforeRows.map(placedImageKey);
  const i = pool.indexOf(removedKey);
  if (i < 0) return { ok: false, reason: 'before-target-key-absent' };
  pool.splice(i, 1);
  const rest = afterRows.map(placedImageKey);
  if (addedMatch) {
    const k = afterRows.findIndex(addedMatch);
    if (k < 0) return { ok: false, reason: 'replacement-not-found' };
    rest.splice(k, 1);
  }
  const a = [...pool].sort();
  const b = [...rest].sort();
  return { ok: a.length === b.length && a.every((x, n) => x === b[n]), reason: null, unexplained: b.filter(x => !a.includes(x)), missing: a.filter(x => !b.includes(x)) };
}

/** replace_image: the image at the target's own part + paragraph index, same placement/wrap/paragraph, at the expected size (±1 EMU, the measured DS rounding). */
function replacementMatcher(op, targetRowBefore) {
  const w = op.widthEmu === undefined ? targetRowBefore.widthEmu : op.widthEmu;
  const h = op.heightEmu === undefined ? targetRowBefore.heightEmu : op.heightEmu;
  return r => r.part === targetRowBefore.part && r.paragraphIndex === targetRowBefore.paragraphIndex && r.placement === targetRowBefore.placement
    && r.wrap === targetRowBefore.wrap && r.paragraph === targetRowBefore.paragraph && fieldReadable(r.widthEmu) && fieldReadable(r.heightEmu)
    && Math.abs(r.widthEmu - w) <= 1 && Math.abs(r.heightEmu - h) <= 1;
}

function verifyEdit(op, before, after) {
  const b = targetRow(before && before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const bRows = imageRows(before.primary) || [];
  const aRows = imageRows(after && after.primary);
  if (!aRows) return { ok: false, evidence: { reason: 'after-images-unreadable' } };
  const isDelete = op.intent === 'delete_image';
  const countOk = aRows.length === bRows.length - (isDelete ? 1 : 0);
  const acc = imagesAccountedFor(bRows, aRows, placedImageKey(b.row), isDelete ? null : replacementMatcher(op, b.row));
  return { ok: countOk && acc.ok, evidence: { index: b.row.index, part: b.row.part, imageCountBefore: bRows.length, imageCountAfter: aRows.length, ...(acc.reason ? { reason: acc.reason } : { unexplained: acc.unexplained, missing: acc.missing }) } };
}

function verifyPrimitive(op, before, after) {
  if (EDIT_INTENTS.includes(op.intent)) return verifyEdit(op, before, after);
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
  if (EDIT_INTENTS.includes(op.intent)) {
    // every other image is already accounted for one by one in L1 (verifyEdit); here the body itself
    const e = verifyEdit(op, before, after);
    const sameBody = bodyKey(before && before.primary) === bodyKey(after && after.primary);
    return { ok: e.ok && sameBody, evidence: { ...e.evidence, bodyUnchanged: sameBody } };
  }
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
  if (EDIT_INTENTS.includes(op.intent)) {
    // the per-operation L1 is the proof (a deleted image leaves no row to name; a replacement has no byte identity in the read)
    const rows = imageRows(finalObservation && finalObservation.primary);
    return [{ name: `${op.id}: ${op.intent} -- images readable after the task`, ok: Array.isArray(rows) }];
  }
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

module.exports = { plan, classify, resolve, preDispatchCheck, lower, commandBody, editCommandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, targetRow, imageRows, otherImages, bodyKey, placedImageKey, SUPPORTED_WRAP, INTENTS, EDIT_INTENTS };
