'use strict';

const { norm } = require('../inventory.cjs');
const { inlineTarget, keyText, selectorCheck } = require('./rich.cjs');

/**
 * DOCX `notes` family (M4 Wave 3): add_footnote, add_endnote (both P0).
 *
 * Source oracle (joskapista44/euro-mcp f8c6d6cf, read-only):
 *   target      docx-rich-structures.cjs:52-63 (paragraph / heading / list item, exactly one occurrence)
 *   lower       coedit-docx-rich-structures.cjs:46-47 (range.Select -> doc.AddFootnote/AddEndnote ->
 *               first paragraph AddText)
 *   L1          docx-rich-structures.cjs:89-93 (exactly one note with the text)
 *   retry       docx-agent-idempotency.cjs:189-200 (footnote: marker on the target + text once),
 *               :181-187 (endnote: text once)
 *   L2          docx-agent-whole-verify.cjs:77-78
 *
 * Measured live before migrating (test/live/docx-rich-probe*.live.cjs, 2026-09-24):
 *   - after AddFootnote the editor cursor stays INSIDE the note (GetCurrentFootEndnote() is an
 *     object), and in that state the next AddFootnote/AddEndnote returns null and a bookmark lands
 *     in the wrong place, while Search still sees the main body. The source restores the main body
 *     with SetDocPosType(0) through frame.evaluate -- the backdoor the owner brief forbids. Here the
 *     SAME callCommand ends with the public ApiDocument.MoveCursorToStart(), which returns the cursor
 *     to the main body (RemoveSelection() does not); both measured in probe 2, the final shape in
 *     probe 3 (backInMain after every note, the package carries every reference);
 *   - the reads (note texts, the reference markers on the target, the cursor context) are public
 *     getters in the driver's read-only query; a reference is an empty run styled
 *     'footnote reference' / 'endnote reference'.
 *
 * Where the target differs from the source, stated:
 *   L1          the note text exactly once AND the target carries one more reference of the kind AND
 *               the cursor is back in the main body (the source checked the text only).
 *   retry       SATISFIED only when the text is there exactly once and the target carries a
 *               reference of the kind (the source did this for footnotes; endnotes counted the text
 *               only). The text elsewhere without a reference on the target -> FAIL_CLOSED.
 *   selector    text or ref only.
 */

const KIND = Object.freeze({ add_footnote: 'footnote', add_endnote: 'endnote' });

function plan(step, index) {
  const s = selectorCheck(step, index);
  if (!s.ok) return s;
  if (!KIND[step.intent]) return { ok: false, detail: `step ${index}: ${step.intent} is not a notes intent` };
  if (typeof step.text !== 'string' || !norm(step.text)) return { ok: false, detail: `step ${index}: ${step.intent} requires a non-empty text` };
  return { ok: true, operation: { id: step.id || `s${index}`, intent: step.intent, selector: step.selector, text: step.text, preservationMode: 'NONE', ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) } };
}

/** The note texts of this kind, from the fresh read; null = unread. */
function noteTexts(primary, kind) {
  const d = primary && primary.document;
  const list = d ? d[`${kind}s`] : null;
  return Array.isArray(list) ? list : null;
}

/** The state that decides both classification and verification, from ONE observation. */
function noteState(primary, op) {
  const kind = KIND[op.intent];
  const t = inlineTarget(primary, keyText(op));
  if (!t.ok) return { ok: false, reason: `${op.intent}-${t.reason}`, count: t.count };
  const texts = noteTexts(primary, kind);
  if (!texts) return { ok: false, reason: `${kind}s-unreadable`, text: t.text };
  const refs = t.node.noteRefs;
  if (!refs || !Number.isInteger(refs[kind])) return { ok: false, reason: 'note-references-unreadable', text: t.text };
  const d = primary.document;
  return { ok: true, kind, text: t.text, count: texts.filter(x => norm(x) === norm(op.text)).length, refs: refs[kind], inMainBody: d.inMainBody };
}

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  const s = noteState(primary, op);
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason, count: s.count, text: s.text } };
  if (s.count === 1 && s.refs >= 1) return { state: 'SATISFIED', evidence: { text: s.text, note: op.text, proof: `fresh-live-${s.kind}-text-once-and-reference-on-target` } };
  if (s.count > 1) return { state: 'FAIL_CLOSED', evidence: { reason: `${s.kind}-duplicated`, text: s.text, count: s.count } };
  if (s.count === 1) return { state: 'FAIL_CLOSED', evidence: { reason: `${s.kind}-text-without-reference-on-target`, text: s.text } };
  if (s.inMainBody !== true) return { state: 'FAIL_CLOSED', evidence: { reason: 'editor-not-in-main-body', text: s.text, inMainBody: s.inMainBody } };
  return { state: 'APPLY', evidence: { text: s.text, refsBefore: s.refs } };
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
  return target && typeof target.text === 'string' ? { ok: true } : { ok: false, dispatched: false, detail: 'note-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({ family: 'notes', action: 'note', kind: KIND[op.intent], targetText: target.text, text: op.text });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand: the cursor must be in the main body; the note is attached at the end of the
 * target's selected range; the command ends by returning the cursor to the main body (measured,
 * see header) -- on every path after the note exists.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const doc = Api.GetDocument();
  if (!doc || doc.GetCurrentFootEndnote() !== null) return;
  const hits = (doc.Search(m.targetText, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || '') === m.targetText);
  if (hits.length !== 1) return;
  hits[0].Select();
  const note = m.kind === 'footnote' ? doc.AddFootnote() : doc.AddEndnote();
  if (!note) { doc.MoveCursorToStart(); return; }
  try { const first = note.GetElement(0); if (first) first.AddText(m.text); } finally { doc.MoveCursorToStart(); }
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** L1: the text once, one more reference of the kind on the target, the cursor back in the main body. */
function verifyPrimitive(op, before, after) {
  const b = noteState(before && before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const a = noteState(after.primary, { ...op, selector: { text: b.text } });
  if (!a.ok) return { ok: false, evidence: { text: b.text, reason: `after-${a.reason}` } };
  const ok = a.count === 1 && a.refs === b.refs + 1 && a.inMainBody === true;
  return { ok, evidence: { text: b.text, note: op.text, count: a.count, refsBefore: b.refs, refsAfter: a.refs, inMainBody: a.inMainBody } };
}

/** L2: the note text exactly once, the target carries a reference of the kind, the cursor in the main body. */
function wholeTaskChecks(op, finalObservation) {
  const s = noteState(finalObservation.primary, op);
  return [{ name: `${op.id}: ${KIND[op.intent]} "${op.text}" exactly once, referenced from "${keyText(op)}"`, ok: s.ok && s.count === 1 && s.refs >= 1 && s.inMainBody === true }];
}

// the target text is unchanged (the reference marker is not text), so a later step may name it
function producedText(op) {
  return keyText(op);
}

module.exports = { plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, wholeTaskChecks, producedText, noteState, KIND };
