'use strict';

const { norm, textUnits } = require('./inventory.cjs');

/**
 * M5 migration of the source task constraints and selector.precondition (owner, owner channel).
 * Source f8c6d6c: docx-agent-planner.cjs:12-17, docx-agent-preflight.cjs:6-110,
 * docx-agent-whole-verify.cjs:37-56. The source semantics are kept; where the source is weaker or
 * silent, the target is STRICTER and says so (never silently weaker):
 *
 *   constraints.section              the heading (normalized text, unique) and its section: the blocks
 *                                    after it up to the next heading of the same or a higher level.
 *                                    Every target (and a move's destination) must be INSIDE -- the
 *                                    heading itself is not -- checked before each dispatch on the fresh
 *                                    targeting observation; after EVERY write, every block outside the
 *                                    section (the heading included) must be unchanged. Source: checked
 *                                    before the writes and once at the end. Header/footer steps: the
 *                                    source skips them silently; here they are refused at planning.
 *   constraints.doNotTouchTables     table intents refused at planning; a target that is a table
 *                                    refused before dispatch; after every write the tables unchanged
 *                                    (the projected grid as the source compares it, plus the driver's
 *                                    public table read when both states carry it).
 *   constraints.preserveFormattingExceptRequested
 *                                    after every write, every text unit identifiable in both states
 *                                    (unique text) keeps its run properties and paragraph format; the
 *                                    format_text target may change ONLY the requested keys. Source: one
 *                                    end-of-task envelope check on the sectionHeading/paragraphOrdinal
 *                                    rewrite targets only, failing when there is none.
 *   selector.precondition            { text } only: the resolved target's normalized text must equal it,
 *                                    checked before each dispatch. A SATISFIED step is never resolved,
 *                                    so a satisfied retry stays 0 write / 0 command (the source refuses
 *                                    such a retry: identity-insufficient-selector / precondition mismatch).
 *
 * Unknown constraint keys are refused (the source copies them along and ignores them).
 */

const CONSTRAINT_KEYS = Object.freeze(['section', 'doNotTouchTables', 'preserveFormattingExceptRequested']);
const HF_INTENTS = Object.freeze(['edit_header', 'edit_footer']);
const isObject = v => v !== null && typeof v === 'object' && !Array.isArray(v);

/** Planning: the task.constraints shape (source planner.cjs:12-16 plus the unknown-key refusal). */
function validateConstraints(task) {
  if (!isObject(task) || task.constraints === undefined) return { ok: true, constraints: null };
  const c = task.constraints;
  if (!isObject(c)) return { ok: false, detail: 'task-constraints-invalid: task.constraints must be an object' };
  const unknown = Object.keys(c).filter(k => !CONSTRAINT_KEYS.includes(k));
  if (unknown.length) return { ok: false, detail: `task-constraints-unsupported: ${unknown.join(', ')} (supported: ${CONSTRAINT_KEYS.join(', ')})` };
  if (c.section !== undefined && (typeof c.section !== 'string' || !c.section.trim())) return { ok: false, detail: 'task-section-constraint-invalid: a non-blank heading text' };
  if (c.doNotTouchTables !== undefined && typeof c.doNotTouchTables !== 'boolean') return { ok: false, detail: 'task-table-constraint-invalid: a boolean' };
  if (c.preserveFormattingExceptRequested !== undefined && typeof c.preserveFormattingExceptRequested !== 'boolean') return { ok: false, detail: 'task-formatting-constraint-invalid: a boolean' };
  const out = {};
  if (c.section !== undefined) out.section = c.section;
  if (c.doNotTouchTables === true) out.doNotTouchTables = true;
  if (c.preserveFormattingExceptRequested === true) out.preserveFormattingExceptRequested = true;
  return { ok: true, constraints: Object.keys(out).length ? Object.freeze(out) : null };
}

/** Planning, per step: the steps a constraint forbids or cannot enforce. */
function checkStepAgainst(constraints, step, index, isTableIntent) {
  if (!constraints) return { ok: true };
  if (constraints.doNotTouchTables && isTableIntent) return { ok: false, detail: `task-table-write-forbidden-by-constraint: step ${index} (${step.intent})` };
  if (constraints.section !== undefined && HF_INTENTS.includes(step.intent)) {
    return { ok: false, detail: `constraint-section-not-enforceable: step ${index} (${step.intent}) writes a header/footer part, outside every body section` };
  }
  return { ok: true };
}

/** Planning: selector.precondition -> { text } exactly. */
function validatePrecondition(pre, index) {
  if (!isObject(pre)) return { ok: false, detail: `step ${index}: precondition-invalid: an object { text }` };
  const extra = Object.keys(pre).filter(k => k !== 'text');
  if (extra.length) return { ok: false, detail: `step ${index}: precondition-unsupported: ${extra.join(', ')} (supported: text)` };
  if (typeof pre.text !== 'string') return { ok: false, detail: `step ${index}: precondition-text-required` };
  return { ok: true, precondition: Object.freeze({ text: pre.text }) };
}

/** The section: the unique heading with this text and the half-open range of its content blocks. */
function sectionBounds(blocks, heading) {
  const hits = [];
  blocks.forEach((b, i) => { if (b.kind === 'heading' && norm(b.text) === norm(heading)) hits.push(i); });
  if (hits.length !== 1) return { ok: false, reason: hits.length ? 'section-ambiguous' : 'section-not-found', candidates: hits.length };
  const h = hits[0];
  let end = blocks.length;
  for (let i = h + 1; i < blocks.length; i += 1) if (blocks[i].kind === 'heading' && blocks[i].level <= blocks[h].level) { end = i; break; }
  return { ok: true, heading: h, start: h + 1, end };
}

/**
 * The body block a resolved target lives in: a block itself, the block of the unique text unit with
 * the target's text, or the table of a cell/row/column locator (`at` = "table:row[:col]"). Anything
 * else is unprovable -- and refused, never assumed inside.
 */
function blockIndexOf(blocks, primary, target) {
  if (!target) return null;
  const own = blocks.indexOf(target);
  if (own >= 0) return own;
  if (typeof target.at === 'string' && /^\d+:/.test(target.at)) {
    const ti = Number(target.at.split(':')[0]);
    const tables = blocks.map((b, i) => (b.kind === 'table' ? i : -1)).filter(i => i >= 0);
    return tables[ti] === undefined ? null : tables[ti];
  }
  if (typeof target.text === 'string' && ['paragraph', 'heading', 'list_item'].includes(target.kind)) {
    const hits = textUnits(primary).filter(u => norm(u.text) === norm(target.text));
    return hits.length === 1 ? blocks.indexOf(hits[0].block) : null;
  }
  return null;
}

/** The text a precondition is compared with: the resolved target's own text; none -> unobservable. */
function targetText(target) {
  if (!target) return null;
  if (typeof target.text === 'string') return target.text;
  if (typeof target.targetText === 'string') return target.targetText;
  return null;
}

/**
 * Before dispatch, against the fresh targeting observation: precondition, table target, section.
 * Returns { ok } or { ok:false, ambiguity, detail } -- the executor stops before any dispatch.
 */
function preDispatch(primary, op, target) {
  const blocks = (primary && primary.blocks) || [];
  if (op.precondition) {
    const actual = targetText(target);
    if (actual === null) return { ok: false, ambiguity: 'PRECONDITION_TARGET_TEXT_UNOBSERVABLE', detail: { expected: op.precondition.text } };
    if (norm(actual) !== norm(op.precondition.text)) return { ok: false, ambiguity: 'PRECONDITION_TEXT_MISMATCH', detail: { expected: op.precondition.text, actual } };
  }
  const c = op.constraints;
  if (!c) return { ok: true };
  if (c.doNotTouchTables && target && typeof target.kind === 'string' && target.kind.startsWith('table')) {
    return { ok: false, ambiguity: 'TABLE_TARGET_FORBIDDEN_BY_CONSTRAINT', detail: { targetKind: target.kind } };
  }
  if (c.section !== undefined) {
    const s = sectionBounds(blocks, c.section);
    if (!s.ok) return { ok: false, ambiguity: `CONSTRAINT_${s.reason.toUpperCase().replace(/-/g, '_')}`, detail: { section: c.section, candidates: s.candidates } };
    for (const [role, t] of [['target', target], ...(target && target.destination ? [['destination', target.destination]] : [])]) {
      const i = blockIndexOf(blocks, primary, t);
      if (i === null || i < 0) return { ok: false, ambiguity: 'CONSTRAINT_TARGET_UNPROVABLE', detail: { section: c.section, role } };
      if (i < s.start || i >= s.end) return { ok: false, ambiguity: 'TARGET_OUTSIDE_CONSTRAINED_SECTION', detail: { section: c.section, role, targetKind: t.kind, targetText: targetText(t) } };
    }
  }
  return { ok: true };
}

// ---- after every write: the constraint preservation (L3) -------------------------------------------

const canonicalTableBlocks = blocks => blocks.filter(b => b.kind === 'table').map(b => b.rows.map(r => r.map(c => [norm(c.text), c.span, c.rowspan])));
const canonicalBlock = b => (b.kind === 'table' ? ['table', canonicalTableBlocks([b])[0]]
  : b.kind === 'list' ? ['list', b.ordered, (b.items || []).map(x => norm(x.text))]
    : [b.kind, b.level || null, norm(b.text), b.style || null]);

function tablesUnchanged(before, after) {
  const grid = JSON.stringify(canonicalTableBlocks(before.blocks || [])) === JSON.stringify(canonicalTableBlocks(after.blocks || []));
  const bt = before.document && before.document.tables;
  const at = after.document && after.document.tables;
  const read = bt && at ? JSON.stringify(bt) === JSON.stringify(at) : null;
  // one state read and the other not: the driver read cannot be compared -> unprovable, fail closed
  const readOk = (bt == null) === (at == null) ? read !== false : false;
  return { ok: grid && readOk, grid, driverRead: read === null && (bt == null) !== (at == null) ? 'one-side-unread' : read };
}

function outsideSectionUnchanged(before, after, section) {
  const b = sectionBounds(before.blocks || [], section);
  const a = sectionBounds(after.blocks || [], section);
  if (!b.ok || !a.ok) return { ok: false, reason: 'section-heading-not-unique', before: b.ok || b.reason, after: a.ok || a.reason };
  const outside = (blocks, s) => blocks.filter((_, i) => i < s.heading || i >= s.end).map(canonicalBlock);
  const ob = outside(before.blocks, b);
  const oa = outside(after.blocks, a);
  return { ok: JSON.stringify(ob) === JSON.stringify(oa), outsideBlocks: [ob.length, oa.length] };
}

function formattingUnchangedExceptRequested(op, before, after, beforeTargetText = null) {
  const unitsOf = p => {
    const m = new Map();
    for (const u of textUnits(p)) { const k = norm(u.text); m.set(k, m.has(k) ? null : u); }
    return m;
  };
  const ub = unitsOf(before);
  const ua = unitsOf(after);
  // "requested": format_text may change ONLY its requested keys on its own target; every other intent
  // changes its own target by definition (a link, a note reference, a list, a new text) -- that target
  // is the family's own L1/L3 authority, and the constraint guards every OTHER text unit.
  const requested = op.intent === 'format_text' && op.format ? Object.keys(op.format) : null;
  const targetKey = typeof beforeTargetText === 'string' ? norm(beforeTargetText) : null;
  if (requested && targetKey === null) return { ok: false, reason: 'format-target-unresolved-in-before-state', requested };
  const strip = (obj, keys) => { if (!obj) return obj; const o = { ...obj }; for (const k of keys || []) delete o[k]; return o; };
  const changed = [];
  let compared = 0;
  for (const [k, u] of ub) {
    const v = ua.get(k);
    if (!u || !v) continue; // not identifiable in both states: text changed, inserted, deleted or duplicated
    const isTarget = targetKey !== null && k === targetKey;
    if (isTarget && !requested) continue; // the operation's own target (not format_text)
    compared += 1;
    const keys = isTarget ? requested : [];
    const same = JSON.stringify(strip(u.node.props, keys)) === JSON.stringify(strip(v.node.props, keys))
      && JSON.stringify(strip(u.node.format, keys)) === JSON.stringify(strip(v.node.format, keys))
      && (isTarget || JSON.stringify(u.node.runs && u.node.runs.map(r => r.format)) === JSON.stringify(v.node.runs && v.node.runs.map(r => r.format)));
    if (!same) changed.push(u.text);
  }
  return { ok: changed.length === 0, compared, changed: changed.slice(0, 10), requested, ownTarget: targetKey };
}

/** L3 for a constrained operation: every active constraint, on the fresh before/after pair. */
function verifyConstraintPreservation(op, before, after, beforeTargetText = null) {
  const c = op.constraints;
  if (!c) return null;
  if (!before || !after || !before.primary || !after.primary) return { ok: false, evidence: { reason: 'constraint-preservation-unobservable' } };
  const evidence = {};
  if (c.doNotTouchTables) evidence.tablesUnchanged = tablesUnchanged(before.primary, after.primary);
  if (c.section !== undefined) evidence.outsideSectionUnchanged = outsideSectionUnchanged(before.primary, after.primary, c.section);
  if (c.preserveFormattingExceptRequested) evidence.formattingUnchangedExceptRequested = formattingUnchangedExceptRequested(op, before.primary, after.primary, beforeTargetText);
  return { ok: Object.values(evidence).every(e => e.ok === true), evidence };
}

module.exports = {
  CONSTRAINT_KEYS, validateConstraints, checkStepAgainst, validatePrecondition, sectionBounds, blockIndexOf, targetText,
  preDispatch, verifyConstraintPreservation, tablesUnchanged, outsideSectionUnchanged, formattingUnchangedExceptRequested, canonicalBlock,
};
