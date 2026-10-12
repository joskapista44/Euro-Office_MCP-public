'use strict';

const { norm, blockText } = require('../inventory.cjs');
const { resolveStableTarget } = require('../resolver.cjs');

/**
 * DOCX `paragraph-structure` family (M4 Wave 1): insert_paragraph_before/after, delete_paragraph,
 * move_paragraph_before/after. Preservation placement: NONE (P0, preservation audit) -- but the
 * STRUCTURAL terminal semantics are strong (owner brief Wave 1 points 5-8).
 *
 * Source oracle (joskapista44/euro-mcp f8c6d6cf, read-only):
 *   lower       coedit-docx-paragraph-operations.cjs:15-18 (CreateParagraph + InsertParagraph,
 *               Delete, Copy + InsertParagraph + Delete; all through callCommand)
 *   classifier  docx-agent-idempotency.cjs:74-90
 *   verify      docx-session-paragraph-operations.cjs:33 (relation) / docx-agent-whole-verify.cjs
 *
 * Where the target is STRONGER than the source, stated:
 *   insert      the source classifier and whole-task verify were position-blind ("text present
 *               once"). Here SATISFIED, L1 and L2 all require the new paragraph exactly once AND
 *               immediately on the requested side of its anchor: before -> NEW TARGET, after -> TARGET NEW.
 *   delete      the source called ANY absent text SATISFIED (a typo'd text included). Here an absent
 *               target is SATISFIED only if the task's own context anchors (before + after) are now
 *               adjacent -- the surviving context. L1 checks the target is gone and its neighbours
 *               from the fresh pre-mutation state closed up; the deleted target is never used as a
 *               locator afterwards.
 *   move        target exactly once, destination exactly once, adjacent on the requested side, and
 *               the old position closed (the old neighbours now adjacent). "Both exist" is not a pass.
 *
 * Identity: a moved paragraph and an insert anchor are identified by their unique text (and kind /
 * headingPath when given) -- their positional context changes by design, so `before`/`after` are
 * not used for them. Only paragraphs and headings are targets in Wave 1 (list items: list family).
 */

const TARGET_KINDS = Object.freeze(['paragraph', 'heading']);
const INSERT = Object.freeze({ insert_paragraph_before: 'before', insert_paragraph_after: 'after' });
const MOVE = Object.freeze({ move_paragraph_before: 'before', move_paragraph_after: 'after' });

function plan(step, index) {
  const id = step.id || `s${index}`;
  const dependsOn = Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {};
  if (!step.selector || typeof step.selector !== 'object') return { ok: false, detail: `step ${index}: ${step.intent} requires a selector` };
  const base = { id, intent: step.intent, selector: step.selector, preservationMode: 'NONE', ...dependsOn };
  if (INSERT[step.intent]) {
    if (typeof step.text !== 'string' || !step.text.trim()) return { ok: false, detail: `step ${index}: ${step.intent} requires a non-empty text` };
    return { ok: true, operation: { ...base, text: step.text } };
  }
  if (step.intent === 'delete_paragraph') return { ok: true, operation: base };
  if (MOVE[step.intent]) {
    const d = step.destinationSelector;
    if (!d || typeof d !== 'object') return { ok: false, detail: `step ${index}: ${step.intent} requires a destinationSelector` };
    if (d.ref !== undefined) return { ok: false, detail: `step ${index}: a destination ref is not part of the M3 ref contract (only selector.ref); use a semantic destinationSelector with dependsOn` };
    return { ok: true, operation: { ...base, destinationSelector: d } };
  }
  return { ok: false, detail: `step ${index}: ${step.intent} is not a paragraph-structure intent` };
}

// ---- identity ----------------------------------------------------------------------------------

/** A selector without positional context (before/after), for identities whose context moves. */
function contextFree(selector) {
  if (!selector) return selector;
  const { before: _b, after: _a, ...rest } = selector;
  return rest;
}

/** The selector an operation resolves its SUBJECT by: a ref resolves to the planned semantic text. */
function subjectSelector(op, { positional = true } = {}) {
  if (op.selector && op.selector.ref !== undefined) return typeof op.refText === 'string' ? { text: op.refText } : null;
  return positional ? op.selector : contextFree(op.selector);
}

function resolveKind(primary, selector) {
  if (!selector) return { ok: false, target: null, selector, ambiguity: 'REF_WITHOUT_TEXT' };
  const r = resolveStableTarget(primary, selector);
  if (r.ok && !TARGET_KINDS.includes(r.target.kind)) return { ok: false, target: null, selector, ambiguity: 'UNSUPPORTED_TARGET_KIND', kind: r.target.kind };
  return r;
}

const blocksOf = primary => (primary && primary.blocks) || [];
const sameBlock = (a, b) => !!a && !!b && a.kind === b.kind && norm(blockText(a)) === norm(blockText(b));
const indexesOf = (primary, text) => blocksOf(primary).map((b, i) => (TARGET_KINDS.includes(b.kind) && norm(b.text) === norm(text) ? i : -1)).filter(i => i >= 0);

/** Exactly-once position of a paragraph/heading text, or the count when it is not exactly one. */
function uniqueIndex(primary, text) {
  const idx = indexesOf(primary, text);
  return idx.length === 1 ? { ok: true, index: idx[0] } : { ok: false, count: idx.length };
}

/** `side`: the subject must sit immediately before / after the anchor. */
function adjacent(subjectIndex, anchorIndex, side) {
  return side === 'before' ? subjectIndex === anchorIndex - 1 : subjectIndex === anchorIndex + 1;
}

/** The neighbours of a block in a state (whole blocks, any kind), for the surviving-context proof. */
function neighbours(primary, index) {
  const blocks = blocksOf(primary);
  return { prev: index > 0 ? blocks[index - 1] : null, next: index + 1 < blocks.length ? blocks[index + 1] : null };
}

/** Is block `a` immediately followed by block `b` in `primary`? (null = document edge.) */
function closedUp(primary, a, b) {
  const blocks = blocksOf(primary);
  if (!a && !b) return blocks.length === 0;
  if (!a) return sameBlock(blocks[0], b);
  if (!b) return sameBlock(blocks[blocks.length - 1], a);
  for (let i = 0; i + 1 < blocks.length; i += 1) if (sameBlock(blocks[i], a) && sameBlock(blocks[i + 1], b)) return true;
  return false;
}

// ---- classify ----------------------------------------------------------------------------------

function classify(primary, op) {
  if (INSERT[op.intent]) return classifyInsert(primary, op, INSERT[op.intent]);
  if (op.intent === 'delete_paragraph') return classifyDelete(primary, op);
  return classifyMove(primary, op, MOVE[op.intent]);
}

const usesOrdinal = sel => !!sel && (sel.sectionHeading !== undefined || sel.paragraphOrdinal !== undefined);

/**
 * M5 gate 3 (source proven LIVE: m5-srcdiff run ...61251a): an insert anchored by sectionHeading +
 * paragraphOrdinal N. The anchor's ordinal SHIFTS with the write, so the terminal position is stated
 * as ordinals of the same section: before -> the new paragraph IS ordinal N and the anchor N+1; after
 * -> the anchor stays N and the new paragraph is N+1. Stronger than the source ("the text is present
 * once"): the position is proven, and a selector precondition is re-checked on the anchor.
 */
function ordinalTerminal(primary, op, side) {
  const n = op.selector.paragraphOrdinal;
  const at = k => resolveStableTarget(primary, { sectionHeading: op.selector.sectionHeading, paragraphOrdinal: k });
  const produced = at(side === 'before' ? n : n + 1);
  const anchor = at(side === 'before' ? n + 1 : n);
  const pre = !op.precondition || (anchor.ok && norm(anchor.target.text) === norm(op.precondition.text));
  const ok = produced.ok && norm(produced.target.text) === norm(op.text) && anchor.ok && pre;
  return { ok, anchorText: anchor.ok ? anchor.target.text : null, preconditionHolds: pre };
}

function classifyInsert(primary, op, side) {
  if (usesOrdinal(op.selector)) {
    const produced = indexesOf(primary, op.text);
    if (produced.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'insert-produced-text-ambiguous', count: produced.length } };
    if (produced.length === 0) {
      const a = resolveKind(primary, contextFree(op.selector));
      if (!a.ok) return { state: 'FAIL_CLOSED', evidence: { reason: `insert-anchor-${String(a.ambiguity).toLowerCase()}` } };
      return { state: 'APPLY', evidence: { anchor: a.target.text, side, ordinal: op.selector.paragraphOrdinal } };
    }
    const t = ordinalTerminal(primary, op, side);
    if (t.ok) return { state: 'SATISFIED', evidence: { text: op.text, anchor: t.anchorText, side, proof: 'fresh-live-ordinal-terminal-position' } };
    return { state: 'FAIL_CLOSED', evidence: { reason: t.preconditionHolds ? 'insert-produced-text-in-wrong-position' : 'insert-anchor-precondition-mismatch-at-terminal', side } };
  }
  const anchor = resolveKind(primary, subjectSelector(op, { positional: false }));
  if (!anchor.ok) return { state: 'FAIL_CLOSED', evidence: { reason: `insert-anchor-${String(anchor.ambiguity).toLowerCase()}` } };
  const produced = indexesOf(primary, op.text);
  if (produced.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'insert-produced-text-ambiguous', count: produced.length } };
  if (produced.length === 0) return { state: 'APPLY', evidence: { anchor: anchor.target.text, side } };
  const anchorIndex = blocksOf(primary).indexOf(anchor.target);
  if (adjacent(produced[0], anchorIndex, side)) {
    return { state: 'SATISFIED', evidence: { text: op.text, anchor: anchor.target.text, side, proof: 'fresh-live-directional-adjacency' } };
  }
  return { state: 'FAIL_CLOSED', evidence: { reason: 'insert-produced-text-in-wrong-position', anchor: anchor.target.text, side } };
}

function classifyDelete(primary, op) {
  const r = resolveKind(primary, subjectSelector(op));
  if (r.ok) return { state: 'APPLY', evidence: { text: r.target.text } };
  if (r.ambiguity === 'AMBIGUOUS') return { state: 'FAIL_CLOSED', evidence: { reason: 'delete-target-ambiguous', candidates: r.candidates } };
  if (r.ambiguity !== 'NOT_FOUND') return { state: 'FAIL_CLOSED', evidence: { reason: `delete-target-${String(r.ambiguity).toLowerCase()}` } };
  // Absent. SATISFIED only if the task's own context proves the deletion: the neighbours it named are
  // adjacent now. Without both anchors an absent text proves nothing (a typo is absent too).
  const sel = op.selector || {};
  if (typeof sel.before !== 'string' || typeof sel.after !== 'string') {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'delete-terminal-identity-insufficient-selector' } };
  }
  const blocks = blocksOf(primary);
  const snip = b => norm(blockText(b)).slice(0, 80);
  const matchBefore = b => (sel.before === '' ? b === null : !!b && snip(b).includes(norm(sel.before)));
  const matchAfter = b => (sel.after === '' ? b === null : !!b && snip(b).includes(norm(sel.after)));
  let found = 0;
  for (let i = -1; i < blocks.length; i += 1) {
    const a = i >= 0 ? blocks[i] : null;
    const b = i + 1 < blocks.length ? blocks[i + 1] : null;
    if (a === null && b === null) continue;
    if (matchBefore(a) && matchAfter(b)) found += 1;
  }
  if (found === 1) return { state: 'SATISFIED', evidence: { text: sel.text, proof: 'fresh-live-surviving-context-adjacent' } };
  return { state: 'FAIL_CLOSED', evidence: { reason: found ? 'delete-surviving-context-ambiguous' : 'delete-surviving-context-not-found' } };
}

function classifyMove(primary, op, side) {
  const src = resolveKind(primary, subjectSelector(op, { positional: false }));
  if (!src.ok) return { state: 'FAIL_CLOSED', evidence: { reason: `move-target-${String(src.ambiguity).toLowerCase()}` } };
  const dst = resolveKind(primary, contextFree(op.destinationSelector));
  if (!dst.ok) return { state: 'FAIL_CLOSED', evidence: { reason: `move-destination-${String(dst.ambiguity).toLowerCase()}` } };
  if (sameBlock(src.target, dst.target)) return { state: 'FAIL_CLOSED', evidence: { reason: 'move-source-equals-destination' } };
  const blocks = blocksOf(primary);
  if (adjacent(blocks.indexOf(src.target), blocks.indexOf(dst.target), side)) {
    return { state: 'SATISFIED', evidence: { text: src.target.text, destination: dst.target.text, side, proof: 'fresh-live-directional-adjacency' } };
  }
  return { state: 'APPLY', evidence: { text: src.target.text, destination: dst.target.text, side } };
}

// ---- resolve / lower -----------------------------------------------------------------------------

/**
 * Resolve against the FRESH targeting observation. A `selector.ref` names an earlier operation's
 * semantic result: its text comes from the ref evidence and must equal the planned `op.refText`.
 * For a move, the destination is resolved in the same observation and returned with the target.
 */
function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation || {};
  let sel = selector;
  if (selector && selector.ref !== undefined) {
    const ref = ctx.refs && ctx.refs[selector.ref];
    const text = ref && ref.evidence && ref.evidence.text;
    if (typeof text !== 'string') return { ok: false, target: null, selector, ambiguity: 'REF_WITHOUT_TEXT' };
    if (typeof op.refText === 'string' && norm(op.refText) !== norm(text)) return { ok: false, target: null, selector, ambiguity: 'REF_RESULT_MISMATCH' };
    sel = { text };
  } else if (INSERT[op.intent] || MOVE[op.intent]) {
    sel = contextFree(selector);
  }
  const r = resolveKind(primary, sel);
  if (!r.ok || !MOVE[op.intent]) return r;
  const dst = resolveKind(primary, contextFree(op.destinationSelector));
  if (!dst.ok) return { ok: false, target: null, selector, ambiguity: `DESTINATION_${dst.ambiguity}` };
  if (sameBlock(r.target, dst.target)) return { ok: false, target: null, selector, ambiguity: 'SOURCE_EQUALS_DESTINATION' };
  return { ok: true, target: { ...r.target, destination: dst.target }, selector, ambiguity: null };
}

function lower(op, target) {
  const sourceText = String(target.text || '');
  if (INSERT[op.intent]) {
    return { kind: 'callCommand', body: insertBody(sourceText, INSERT[op.intent], op.text),
      meta: Object.freeze({ family: 'paragraph', action: 'insert', sourceText, side: INSERT[op.intent], text: op.text }) };
  }
  if (op.intent === 'delete_paragraph') {
    return { kind: 'callCommand', body: deleteBody(sourceText), meta: Object.freeze({ family: 'paragraph', action: 'delete', sourceText }) };
  }
  const destinationText = String(target.destination.text || '');
  return { kind: 'callCommand', body: moveBody(sourceText, MOVE[op.intent], destinationText),
    meta: Object.freeze({ family: 'paragraph', action: 'move', sourceText, side: MOVE[op.intent], destinationText }) };
}

// The three editor commands re-express coedit-docx-paragraph-operations.cjs:15-18 (same API calls).
const exactParagraph = `
  const exactParagraph = t => {
    const rs = (doc.Search(t, { MatchCase: true }) || []).filter(r => String(r.GetText({ Numbering: false }) || '') === t);
    return rs.length === 1 ? rs[0].GetParagraph(0) : null;
  };`;

function insertBody(sourceText, side, text) {
  return `
try {
  const doc = Api.GetDocument();
  if (!doc) return;${exactParagraph}
  const anchor = exactParagraph(${JSON.stringify(sourceText)});
  if (!anchor) return;
  const paragraph = Api.CreateParagraph();
  paragraph.AddText(${JSON.stringify(text)});
  anchor.InsertParagraph(paragraph, ${JSON.stringify(side)}, true);
} catch (_) {}
`;
}

function deleteBody(sourceText) {
  return `
try {
  const doc = Api.GetDocument();
  if (!doc) return;${exactParagraph}
  const target = exactParagraph(${JSON.stringify(sourceText)});
  if (!target) return;
  target.Delete();
} catch (_) {}
`;
}

function moveBody(sourceText, side, destinationText) {
  return `
try {
  const doc = Api.GetDocument();
  if (!doc) return;${exactParagraph}
  const source = exactParagraph(${JSON.stringify(sourceText)});
  const destination = exactParagraph(${JSON.stringify(destinationText)});
  if (!source || !destination) return;
  const moved = source.Copy();
  if (!moved) return;
  if (!destination.InsertParagraph(moved, ${JSON.stringify(side)}, true)) return;
  source.Delete();
} catch (_) {}
`;
}

// ---- verification --------------------------------------------------------------------------------

/**
 * L1 from the FRESH after-state. The before-state is the fresh pre-mutation targeting observation
 * (the same one the operation was resolved in); nothing is carried from an earlier state.
 */
function verifyPrimitive(op, before, after) {
  if (INSERT[op.intent]) return verifyInsert(op, before, after, INSERT[op.intent]);
  if (op.intent === 'delete_paragraph') return verifyDelete(op, before, after);
  return verifyMove(op, before, after, MOVE[op.intent]);
}

function verifyInsert(op, before, after, side) {
  const anchor = resolveKind(before.primary, subjectSelector(op, { positional: false }));
  if (!anchor.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the before-state' } };
  const n = uniqueIndex(after.primary, op.text);
  const a = uniqueIndex(after.primary, anchor.target.text);
  const directional = n.ok && a.ok && adjacent(n.index, a.index, side);
  // M5 gate 3: an ordinal anchor also proves the terminal ORDINAL position in its section
  const ordinal = usesOrdinal(op.selector) ? ordinalTerminal(after.primary, op, side).ok : null;
  const ok = directional && ordinal !== false;
  return { ok, evidence: { text: op.text, anchor: anchor.target.text, side, insertedCount: n.ok ? 1 : n.count, anchorCount: a.ok ? 1 : a.count, directional, ...(ordinal === null ? {} : { ordinalTerminal: ordinal }) } };
}

function verifyDelete(op, before, after) {
  const r = resolveKind(before.primary, subjectSelector(op));
  if (!r.ok) return { ok: false, evidence: { reason: 'target not resolvable in the before-state' } };
  const index = blocksOf(before.primary).indexOf(r.target);
  const { prev, next } = neighbours(before.primary, index);
  const remaining = indexesOf(after.primary, r.target.text).length;
  const survived = closedUp(after.primary, prev, next);
  return { ok: remaining === 0 && survived, evidence: { deleted: r.target.text, remaining, survivingContext: { before: prev ? blockText(prev) : null, after: next ? blockText(next) : null, adjacent: survived } } };
}

function verifyMove(op, before, after, side) {
  const src = resolveKind(before.primary, subjectSelector(op, { positional: false }));
  const dst = resolveKind(before.primary, contextFree(op.destinationSelector));
  if (!src.ok || !dst.ok) return { ok: false, evidence: { reason: 'target or destination not resolvable in the before-state' } };
  const oldIndex = blocksOf(before.primary).indexOf(src.target);
  const { prev, next } = neighbours(before.primary, oldIndex);
  const s = uniqueIndex(after.primary, src.target.text);
  const d = uniqueIndex(after.primary, dst.target.text);
  const directional = s.ok && d.ok && adjacent(s.index, d.index, side);
  const oldPositionClosed = closedUp(after.primary, prev, next);
  return {
    ok: directional && oldPositionClosed,
    evidence: { text: src.target.text, destination: dst.target.text, side, targetCount: s.ok ? 1 : s.count, destinationCount: d.ok ? 1 : d.count, directional, oldPositionClosed },
  };
}

/** L2, read-only against the final observation. */
function wholeTaskChecks(op, finalObservation) {
  const f = finalObservation.primary;
  if (INSERT[op.intent] && usesOrdinal(op.selector)) {
    const n = uniqueIndex(f, op.text);
    return [
      { name: `${op.id}: inserted text present once`, ok: n.ok },
      { name: `${op.id}: inserted ${INSERT[op.intent]} ordinal ${op.selector.paragraphOrdinal} of its section (terminal ordinal position)`, ok: n.ok && ordinalTerminal(f, op, INSERT[op.intent]).ok },
    ];
  }
  if (INSERT[op.intent]) {
    const anchorSel = subjectSelector(op, { positional: false });
    const anchorText = anchorSel && anchorSel.text;
    const n = uniqueIndex(f, op.text);
    const a = typeof anchorText === 'string' ? uniqueIndex(f, anchorText) : { ok: false };
    return [
      { name: `${op.id}: inserted text present once`, ok: n.ok },
      { name: `${op.id}: inserted ${INSERT[op.intent]} its anchor (directional adjacency)`, ok: n.ok && a.ok && adjacent(n.index, a.index, INSERT[op.intent]) },
    ];
  }
  if (op.intent === 'delete_paragraph') {
    const sel = subjectSelector(op);
    const text = sel && sel.text;
    const checks = [{ name: `${op.id}: deleted text absent`, ok: typeof text === 'string' && indexesOf(f, text).length === 0 }];
    if (op.selector && typeof op.selector.before === 'string' && typeof op.selector.after === 'string') {
      checks.push({ name: `${op.id}: surviving context adjacent`, ok: classifyDelete(f, op).state === 'SATISFIED' });
    }
    return checks;
  }
  const srcSel = subjectSelector(op, { positional: false });
  const srcText = srcSel && srcSel.text;
  const dstText = op.destinationSelector && op.destinationSelector.text;
  const s = typeof srcText === 'string' ? uniqueIndex(f, srcText) : { ok: false };
  const d = typeof dstText === 'string' ? uniqueIndex(f, dstText) : { ok: false };
  return [
    { name: `${op.id}: moved paragraph present once`, ok: s.ok },
    { name: `${op.id}: destination present once`, ok: d.ok },
    { name: `${op.id}: moved ${MOVE[op.intent]} its destination (directional adjacency)`, ok: s.ok && d.ok && adjacent(s.index, d.index, MOVE[op.intent]) },
  ];
}

/** The semantic text this operation leaves behind (for a later operation's ref). */
function producedText(op) {
  if (INSERT[op.intent]) return op.text;
  if (MOVE[op.intent]) { const s = subjectSelector(op, { positional: false }); return s && typeof s.text === 'string' ? s.text : null; }
  return null;
}

module.exports = {
  plan, classify, resolve, lower, verifyPrimitive, wholeTaskChecks, producedText,
  insertBody, deleteBody, moveBody, adjacent, closedUp,
};
