'use strict';

const { norm, resolveTextUnit, textUnits } = require('../inventory.cjs');

/**
 * DOCX `rich-inline` family (M4 Wave 3): add_bookmark (P0), add_link (P1, PRIMITIVE_INTEGRAL).
 *
 * Source oracle (joskapista44/euro-mcp f8c6d6cf, read-only):
 *   operations  docx-rich-structures.cjs normalizeRichOperation (bookmark name ^[A-Za-z_][A-Za-z0-9_]{0,39}$)
 *   target      docx-rich-structures.cjs:52-63 (paragraph / heading / list item, exactly one exact occurrence)
 *   lower       coedit-docx-rich-structures.cjs:44 (range.AddBookmark), docx-session-rich-structures.cjs:115
 *               (effective TextPr before -> paragraph.AddHyperlink -> SetTextPr(before)) -- the link
 *               is re-scoped to the target RANGE here, see below
 *   L1          docx-rich-structures.cjs:80-87; link P1 docx-session-rich-structures.cjs:129 +
 *               docx-session-basic-edit.cjs:20 comparableProperties (7 keys, underline included)
 *   retry / L2  docx-agent-idempotency.cjs:181-187, docx-agent-whole-verify.cjs:77
 *
 * Measured live before migrating (test/live/docx-rich-probe*.live.cjs, three runs 2026-09-24):
 *   - the source READS all of this through frame.evaluate (bookmark manager, getSelectedElements);
 *     the public Builder getters answer the same inside callCommand, so the driver's read-only query
 *     carries them: bookmarks (GetAllBookmarksNames + GetBookmarkRange text), hyperlink children
 *     (url, screen tip, displayed text), and the cursor context (GetCurrentFootEndnote() null);
 *   - a bookmark written while the cursor is inside a note lands in the wrong place (empty range,
 *     the name twice after reload) and Search does not reveal that state -- so a bookmark write
 *     requires the MAIN body, read before planning and re-checked inside the command;
 *   - after AddHyperlink the "Hyperlink" character style changes the effective underline/colour;
 *     SetTextPr(before) restores the seven compared properties (probe 3: all equal);
 *   - the source's ApiParagraph.AddHyperlink wraps the WHOLE paragraph: a note reference already in
 *     it moves inside the link and loses its reference style (blue, underlined, not superscript --
 *     the package keeps it), so a later read no longer finds it (probe 4; the first W3 LIVE run,
 *     w3_mixed). ApiRange.AddHyperlink on the exact target text leaves the reference outside with
 *     its style, keeps the seven properties, and links a plain paragraph the same way (probe 5).
 *
 * Where the target differs from the source, stated:
 *   enrichment  NONE. The source took the bookmark duplicate guard from the PACKAGE; the live
 *               public read is the authority here (the package lags the save by ~10 s, and the
 *               document-server package URL answered 403 inside a session -- rich probe 2).
 *   L1 bookmark the name exactly once AND its range is the target text (the source counted names only).
 *   link scope the exact target-text RANGE, not the paragraph (source: paragraph.AddHyperlink, which
 *               damages an existing note reference -- a source defect; same result on a plain
 *               paragraph). L3 also keeps the target's note references (the damage is caught at the
 *               primitive, not only by a later note's L2).
 *   retry       an existing bookmark / link on the target is SATISFIED only when it is exactly the
 *               requested one; anything else on the target fails closed (never a second write).
 *   selector    text or ref only.
 */

const BOOKMARK_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,39}$/;
const TEXT_KEYS = ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'language'];

function selectorCheck(step, index) {
  const sel = step.selector;
  if (!sel || typeof sel !== 'object') return { ok: false, detail: `step ${index}: ${step.intent} requires a selector` };
  const keys = Object.keys(sel);
  const textSel = keys.length === 1 && typeof sel.text === 'string' && sel.text.trim() !== '';
  const refSel = keys.length === 1 && typeof sel.ref === 'string' && sel.ref !== '';
  if (!textSel && !refSel) return { ok: false, detail: `step ${index}: a ${step.intent} selector is { text } or { ref } in M4 Wave 3 (got: ${keys.join(', ') || 'nothing'})` };
  return { ok: true };
}

function plan(step, index) {
  const s = selectorCheck(step, index);
  if (!s.ok) return s;
  const op = { id: step.id || `s${index}`, intent: step.intent, selector: step.selector, ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) };
  if (step.intent === 'add_bookmark') {
    const name = typeof step.name === 'string' ? step.name.trim() : '';
    if (!BOOKMARK_NAME.test(name)) return { ok: false, detail: `step ${index}: add_bookmark requires a name matching ${BOOKMARK_NAME}` };
    return { ok: true, operation: { ...op, name, preservationMode: 'NONE' } };
  }
  if (step.intent === 'add_link') {
    const url = typeof step.url === 'string' ? step.url.trim() : '';
    if (!url) return { ok: false, detail: `step ${index}: add_link requires a url` };
    if (step.tooltip !== undefined && typeof step.tooltip !== 'string') return { ok: false, detail: `step ${index}: add_link tooltip must be a string` };
    return { ok: true, operation: { ...op, url, tooltip: step.tooltip || '', preservationMode: 'PRIMITIVE_INTEGRAL' } };
  }
  return { ok: false, detail: `step ${index}: ${step.intent} is not a rich-inline intent` };
}

/** The target text an operation names (a ref resolves to the planned semantic text). */
function keyText(op) {
  if (op.selector && op.selector.ref !== undefined) return typeof op.refText === 'string' ? op.refText : null;
  return op.selector ? op.selector.text : null;
}

/**
 * The ONE paragraph / heading / list item with this text, and no table cell with it either (the
 * write finds its range by an exact document Search, which would see a cell too -- source
 * searchableExactTextCount). Anything else is not a target.
 */
function inlineTarget(primary, text) {
  if (typeof text !== 'string' || !norm(text)) return { ok: false, reason: 'target-text-unknown' };
  const r = resolveTextUnit(primary, text);
  let cells = 0;
  for (const b of (primary && primary.blocks) || []) if (b.kind === 'table') for (const row of b.rows || []) for (const c of row) if (norm(c.text) === norm(text)) cells += 1;
  if (!r.ok) return { ok: false, reason: r.reason === 'text-unit-ambiguous' ? 'target-ambiguous' : 'target-not-found', count: r.count + cells };
  if (cells) return { ok: false, reason: 'target-ambiguous', count: 1 + cells };
  return { ok: true, text: r.unit.text, kind: r.unit.kind, node: r.unit.node };
}

const docOf = primary => (primary && primary.document) || null;

function bookmarksNamed(primary, name) {
  const d = docOf(primary);
  if (!d || !Array.isArray(d.bookmarks)) return null;
  return d.bookmarks.filter(b => b.name === name);
}

/**
 * The screen tip the link carries. ApiRange.AddHyperlink with an EMPTY tip writes nothing (W3 LIVE
 * run 2, m4-wave1-1790254996349-d77605: no hyperlink); the source's paragraph method with an empty
 * tip leaves the editor default, the url (runs 1-2 source final state). Without a tooltip the url is
 * therefore passed explicitly -- the same durable state as the source.
 */
const tipOf = op => op.tooltip || op.url;
const linkMatches = (l, op, text) => l.url === op.url && l.tip === tipOf(op) && norm(l.text) === norm(text);

// ---- classify / resolve -------------------------------------------------------------------------

/**
 * Fresh LIVE classification:
 *   target absent / ambiguous / the needed read unread                    -> FAIL_CLOSED
 *   the requested bookmark / link is already there, exactly               -> SATISFIED
 *   a bookmark with this name elsewhere, or another link on the target    -> FAIL_CLOSED
 *   the editor cursor is not in the main body                             -> FAIL_CLOSED
 *   otherwise                                                              -> APPLY
 */
function classify(primary, op) {
  const t = inlineTarget(primary, keyText(op));
  if (!t.ok) return { state: 'FAIL_CLOSED', evidence: { reason: `${op.intent}-${t.reason}`, count: t.count } };
  const d = docOf(primary);
  if (op.intent === 'add_bookmark') {
    const same = bookmarksNamed(primary, op.name);
    if (!same) return { state: 'FAIL_CLOSED', evidence: { reason: 'bookmarks-unreadable', text: t.text } };
    if (same.length === 1 && norm(same[0].text) === norm(t.text)) return { state: 'SATISFIED', evidence: { text: t.text, name: op.name, proof: 'fresh-live-bookmark-range-is-target' } };
    if (same.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'bookmark-name-exists', name: op.name, ranges: same.map(b => b.text) } };
  } else {
    const links = t.node.hyperlinks;
    if (!Array.isArray(links)) return { state: 'FAIL_CLOSED', evidence: { reason: 'hyperlinks-unreadable', text: t.text } };
    if (links.length === 1 && linkMatches(links[0], op, t.text)) return { state: 'SATISFIED', evidence: { text: t.text, url: op.url, proof: 'fresh-live-hyperlink-is-requested' } };
    if (links.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'target-already-linked', text: t.text, links: links.map(l => l.url) } };
    if (!t.node.props) return { state: 'FAIL_CLOSED', evidence: { reason: 'run-properties-unreadable', text: t.text } };
  }
  if (!d || d.inMainBody !== true) return { state: 'FAIL_CLOSED', evidence: { reason: 'editor-not-in-main-body', text: t.text, inMainBody: d ? d.inMainBody : null } };
  return { state: 'APPLY', evidence: { text: t.text } };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  let text = selector && selector.text;
  if (selector && selector.ref !== undefined) {
    const ref = ctx.refs && ctx.refs[selector.ref];
    text = ref && ref.evidence && ref.evidence.text;
    if (typeof text !== 'string') return { ok: false, target: null, selector, ambiguity: 'REF_WITHOUT_TEXT' };
    if (op && typeof op.refText === 'string' && norm(op.refText) !== norm(text)) return { ok: false, target: null, selector, ambiguity: 'REF_RESULT_MISMATCH' };
  }
  const t = inlineTarget(primary, text);
  if (!t.ok) return { ok: false, target: null, selector, ambiguity: t.reason.replace(/^target-/, '').toUpperCase() };
  return { ok: true, target: { kind: t.kind, text: t.text }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && typeof target.text === 'string' ? { ok: true } : { ok: false, dispatched: false, detail: 'rich-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze(op.intent === 'add_bookmark'
    ? { family: 'rich-inline', action: 'bookmark', targetText: target.text, name: op.name }
    : { family: 'rich-inline', action: 'link', targetText: target.text, url: op.url, tooltip: tipOf(op) });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand. The target by its exact text; the cursor must be in the main body (a bookmark
 * written from inside a note lands wrong -- rich probe 1). A name that already exists -> no write.
 * The link covers the exact target text (the range, not the paragraph -- probe 5) and keeps its
 * effective text properties: read before, restored after (source parity, probe 3).
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const doc = Api.GetDocument();
  if (!doc || doc.GetCurrentFootEndnote() !== null) return;
  const exact = () => (doc.Search(m.targetText, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || '') === m.targetText);
  const hits = exact();
  if (hits.length !== 1) return;
  if (m.action === 'bookmark') {
    if ((doc.GetAllBookmarksNames() || []).indexOf(m.name) >= 0) return;
    hits[0].AddBookmark(m.name);
    return;
  }
  const before = hits[0].GetTextPr();
  if (!before || !hits[0].AddHyperlink(m.url, m.tooltip)) return;
  const again = exact();
  if (again.length !== 1) return;
  again[0].SetTextPr(before);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** L1 in the FRESH after-state. */
function verifyPrimitive(op, before, after) {
  const b = inlineTarget(before && before.primary, keyText(op));
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  if (op.intent === 'add_bookmark') {
    const same = bookmarksNamed(after.primary, op.name);
    const ok = !!same && same.length === 1 && norm(same[0].text) === norm(b.text);
    return { ok, evidence: { text: b.text, name: op.name, observed: same ? same.map(x => x.text) : 'unread' } };
  }
  const a = inlineTarget(after.primary, b.text);
  const links = a.ok ? a.node.hyperlinks : null;
  const ok = Array.isArray(links) && links.length === 1 && linkMatches(links[0], op, b.text);
  return { ok, evidence: { text: b.text, expected: { url: op.url, tooltip: tipOf(op) }, observed: links || (a.ok ? 'unread' : a.reason) } };
}

// ---- L3 round trip: everything the link must not touch, as far as the live read observes it ----

const count = list => { const m = new Map(); for (const k of list) m.set(k, (m.get(k) || 0) + 1); return m; };
function sameMultiset(a, b) {
  const x = count(a); const y = count(b);
  const diff = [];
  for (const k of new Set([...x.keys(), ...y.keys()])) if ((x.get(k) || 0) !== (y.get(k) || 0)) diff.push({ key: k, before: x.get(k) || 0, after: y.get(k) || 0 });
  return { ok: diff.length === 0, diff };
}
const unitKey = u => JSON.stringify([u.kind, norm(u.text), u.node.hyperlinks || 'unread', u.node.noteRefs || 'unread']);
const tableKey = b => JSON.stringify((b.rows || []).map(r => r.map(c => norm(c.text))));
const imageKey = i => JSON.stringify([i.src || null, i.alt || null, i.title || null]);

/**
 * Every OTHER text unit (text, kind, its hyperlinks and note references), every table shape and
 * image, and the document-level rich read (bookmarks with their range text, footnote and endnote
 * texts) -- unchanged. Only what the live read observes is claimed; an unread side is a failure.
 */
function roundTrip(before, after, targetText) {
  const want = norm(targetText);
  const others = p => textUnits(p).filter(u => norm(u.text) !== want).map(unitKey);
  const d = p => p.document || {};
  const list = (p, k) => (Array.isArray(d(p)[k]) ? d(p)[k] : null);
  const rich = ['bookmarks', 'footnotes', 'endnotes'].map(k => {
    const b = list(before, k); const a = list(after, k);
    if (!b || !a) return [k, { ok: false, reason: 'unread' }];
    const key = k === 'bookmarks' ? x => `${x.name}|${norm(x.text)}` : x => norm(x);
    return [k, sameMultiset(b.map(key), a.map(key))];
  });
  const parts = {
    otherUnits: sameMultiset(others(before), others(after)),
    tables: sameMultiset((before.blocks || []).filter(x => x.kind === 'table').map(tableKey), (after.blocks || []).filter(x => x.kind === 'table').map(tableKey)),
    images: sameMultiset((before.images || []).map(imageKey), (after.images || []).map(imageKey)),
    ...Object.fromEntries(rich),
  };
  return { ok: Object.values(parts).every(x => x.ok), parts };
}

/**
 * L3 for add_link, PRIMITIVE_INTEGRAL (source docx-session-rich-structures.cjs:129 with
 * comparableProperties, docx-session-basic-edit.cjs:20), owner review of Wave 3 point 3: the target
 * text exactly once; its seven effective text properties unchanged (the requested link is L1); the
 * target's note references unchanged -- counted by their 'footnote/endnote reference' style, so a
 * reference pulled into the link (probe 4) fails here; and the round trip above. NOT claimed: the
 * references' direct run properties (superscript, colour) -- the live read carries their STYLE only.
 * add_bookmark is P0 (no L3; the audit row says why).
 */
function verifyPreservation(op, before, after) {
  if (op.intent !== 'add_link') return null;
  const b = inlineTarget(before && before.primary, keyText(op));
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const a = inlineTarget(after.primary, b.text);
  if (!a.ok) return { ok: false, evidence: { targetText: b.text, reason: `after-${a.reason}` } };
  if (!b.node.props || !a.node.props) return { ok: false, evidence: { targetText: b.text, reason: 'run-properties-unreadable' } };
  const pick = p => Object.fromEntries(TEXT_KEYS.map(k => [k, p[k]]));
  const bp = pick(b.node.props);
  const ap = pick(a.node.props);
  const changed = TEXT_KEYS.filter(k => JSON.stringify(bp[k]) !== JSON.stringify(ap[k]));
  const refsKept = !!b.node.noteRefs && !!a.node.noteRefs && JSON.stringify(b.node.noteRefs) === JSON.stringify(a.node.noteRefs);
  const rt = roundTrip(before.primary, after.primary, b.text);
  return { ok: changed.length === 0 && refsKept && rt.ok, evidence: { targetText: b.text, before: bp, after: ap, changed, noteRefs: { before: b.node.noteRefs, after: a.node.noteRefs, kept: refsKept }, roundTrip: rt.parts, expectationSource: 'derived-from-fresh-pre-mutation-observation' } };
}

/** L2: the final state still carries the requested bookmark / link on the target. */
function wholeTaskChecks(op, finalObservation) {
  const t = inlineTarget(finalObservation.primary, keyText(op));
  if (op.intent === 'add_bookmark') {
    const same = bookmarksNamed(finalObservation.primary, op.name);
    return [{ name: `${op.id}: bookmark "${op.name}" exactly once, on "${keyText(op)}"`, ok: t.ok && !!same && same.length === 1 && norm(same[0].text) === norm(t.text) }];
  }
  const links = t.ok ? t.node.hyperlinks : null;
  return [{ name: `${op.id}: "${keyText(op)}" carries exactly the link ${op.url}`, ok: Array.isArray(links) && links.length === 1 && linkMatches(links[0], op, t.text) }];
}

// the target text is unchanged by both intents, so a later step may name it by ref
function producedText(op) {
  return keyText(op);
}

module.exports = {
  plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation,
  wholeTaskChecks, producedText, inlineTarget, keyText, selectorCheck, BOOKMARK_NAME,
};
