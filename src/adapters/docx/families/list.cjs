'use strict';

const { norm } = require('../inventory.cjs');

/** M11.5d.4: B18 -- the numbering removal through the word /2 Bridge WRITE (the driver's __eoB18Remove; no direct internal path, no fallback). */
const NUMBERING_REMOVE_WRITE_ID = 'docx.paragraph.numbering.remove/v1';

/**
 * DOCX `list` family (M4 Wave 2): apply_list, set_list_level, set_list_type, indent_list_item,
 * outdent_list_item, remove_list. Preservation placement: PRIMITIVE_INTEGRAL (P1, preservation audit:
 * docx-session-lists.cjs:59-66, target text unchanged + live round-trip inside the success branch).
 *
 * Source oracle (joskapista44/euro-mcp f8c6d6cf, read-only):
 *   operations  docx-lists.cjs normalizeListOperation (types bullet|numbered, levels 0-8)
 *   lower       coedit-docx-lists.cjs:7 (CreateNumbering / SetNumbering / SetTemplateType)
 *   observe     docx-session-structure.cjs readStructureInSession (type from the live ul/ol, level)
 *   L1          docx-lists.cjs expectedFromObservation / verifyListObservation
 *   retry       docx-agent-idempotency.cjs:103-116 (the task's PROJECTED terminal state per target)
 *   L2          docx-agent-whole-verify.cjs:70/80/81 (per item; grouped apply_list continuity)
 *
 * Measured live before migrating (test/live/docx-list-probe.live.cjs, three runs 2026-09-24):
 *   - no public Builder call removes numbering (SetNumbering(null), SetNumPr(null),
 *     ParaPr.SetNumPr(null), SetBullet(null), ParaPr.SetBullet(null): all without effect);
 *   - the internal CParagraph.RemoveNumPr(), called INSIDE callCommand, removes it (2/2 runs).
 *   The source removed a list (and outdented at level 0) with the OUTER editor's put_ListType through
 *   frame.evaluate -- the backdoor the owner brief forbids. Here every mutation is ONE callCommand
 *   through the common dispatch; the removal is an EO_INTERNAL_DEPENDENCY (owner decision, Talk
 *   owner channel after 28596): the driver's read carries whether the editor exposes RemoveNumPr, and a
 *   removal step on an editor without it fails closed BEFORE any dispatch
 *   (eo-internal-remove-numbering-unavailable). Re-measure after every Euro-Office/ONLYOFFICE release.
 *   M11.5d.4 (M11.5 continuous mandate, internal tracker): the removal is CUT OVER to the word /2 Bridge WRITE
 *   docx.paragraph.numbering.remove/v1 (the same editor routine behind the Bridge, PUBLIC_FIRST re-measured 2026-10-11);
 *   the paragraph above is history. Any runtime but word /2 refuses BEFORE the write
 *   (docx-numbering-remove-bridge-write-runtime-not-proven); there is no direct internal path and no fallback.
 *
 * Where the target differs from the source, stated:
 *   relative    indent/outdent need an absolute anchor (apply_list or set_list_level) for the same
 *               target EARLIER in the task; without it the terminal level is not provable and the
 *               operation fails closed (the source silently assumed a base level of 0).
 *   L3          besides the round-trip, every OTHER list item keeps its type and level (a shared
 *               numbering definition changed through one item is a collateral change, not a pass).
 *   selector    text or ref only (the source's headingPath/before/after filters are not migrated).
 *
 * M5 gate 7 (owner decision 7A after 28880). A numbering's list TYPE is per LEVEL: set_list_type
 * re-templates the item's own level only (source and target alike), so a later level move lands on a
 * level that may keep another type (list-level probe 2026-09-25: bullet@0 -> level 1 -> numbered ->
 * indent = bullet@2; the companion item at level 0 kept bullet). The type of the DESTINATION level
 * has no public getter; the driver reads every level's format READ ONLY from the internal CNum
 * (EO_INTERNAL_DEPENDENCY READ). Runtime capability guard, before any dispatch:
 *   - reader absent / unreadable                     -> eo-internal-list-level-format-unavailable
 *   - the item's OWN level format disagrees with the
 *     projection's ul/ol                             -> eo-internal-list-level-format-disagrees
 *   - the destination level's format is not one of
 *     the MEASURED formats below                     -> list-destination-level-type-unknown
 * The expected type of a level move is the destination level's FRESH LIVE type; the task projection
 * does not claim a type after a level move (it is known only live). L3 additionally proves that the
 * numbering's other levels keep their format. Re-measure the format table and the reader after every
 * Euro-Office / ONLYOFFICE release (test/live/docx-list-level-probe.live.cjs).
 */

// MEASURED live (list-level probe, 2026-09-25): bullet template levels -> 5; numbered templates -> 13, 46, 47.
// Any other value is unknown (fail closed), never guessed.
const LEVEL_FORMAT_TYPE = Object.freeze({ 5: 'bullet', 13: 'numbered', 46: 'numbered', 47: 'numbered' });
const LEVEL_MOVES = Object.freeze(['set_list_level', 'indent_list_item', 'outdent_list_item']);

const LIST_TYPES = Object.freeze(['bullet', 'numbered']);
const TEXT_KINDS = Object.freeze(['paragraph', 'heading']);

const validLevel = l => Number.isInteger(l) && l >= 0 && l <= 8;

function plan(step, index) {
  const id = step.id || `s${index}`;
  const sel = step.selector;
  if (!sel || typeof sel !== 'object') return { ok: false, detail: `step ${index}: ${step.intent} requires a selector` };
  const keys = Object.keys(sel);
  const textSel = keys.length === 1 && typeof sel.text === 'string' && sel.text.trim() !== '';
  const refSel = keys.length === 1 && typeof sel.ref === 'string' && sel.ref !== '';
  if (!textSel && !refSel) return { ok: false, detail: `step ${index}: a list selector is { text } or { ref } in M4 Wave 2 (got: ${keys.join(', ') || 'nothing'})` };
  const op = { id, intent: step.intent, selector: sel, preservationMode: 'PRIMITIVE_INTEGRAL', ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) };
  if (step.intent === 'apply_list') {
    if (!LIST_TYPES.includes(step.type)) return { ok: false, detail: `step ${index}: apply_list requires type bullet|numbered` };
    const level = step.level == null ? 0 : step.level;
    if (!validLevel(level)) return { ok: false, detail: `step ${index}: level must be an integer 0-8` };
    return { ok: true, operation: { ...op, type: step.type, level } };
  }
  if (step.intent === 'set_list_level') {
    if (!validLevel(step.level)) return { ok: false, detail: `step ${index}: set_list_level requires an integer level 0-8` };
    return { ok: true, operation: { ...op, level: step.level } };
  }
  if (step.intent === 'set_list_type') {
    if (!LIST_TYPES.includes(step.type)) return { ok: false, detail: `step ${index}: set_list_type requires type bullet|numbered` };
    return { ok: true, operation: { ...op, type: step.type } };
  }
  if (['indent_list_item', 'outdent_list_item', 'remove_list'].includes(step.intent)) return { ok: true, operation: op };
  return { ok: false, detail: `step ${index}: ${step.intent} is not a list intent` };
}

/** The target text an operation names (a ref resolves to the planned semantic text). */
function keyText(op) {
  if (op.selector && op.selector.ref !== undefined) return typeof op.refText === 'string' ? op.refText : null;
  return op.selector ? op.selector.text : null;
}

/**
 * One projection step (source projectedListStates, without the silent level-0 assumption).
 * M5 gate 7: the projection tracks the numbering's per-LEVEL types when the task itself created the
 * numbering (a new apply_list: every level of CreateNumbering(type) has that type -- measured, list-level
 * probe); set_list_type re-templates one level; a level move takes the destination level's type. A
 * continued or pre-existing numbering has unknown levels: a level move then projects NO type (null =
 * not determined by the task; L1 checks the destination level's LIVE type).
 */
function project(state, op) {
  const s = state || { type: null, level: null, anchored: false, levels: null };
  const moved = d => ({ ...s, type: s.levels ? s.levels[d] : null, level: d });
  switch (op.intent) {
    case 'apply_list': return { type: op.type, level: op.level, anchored: true, levels: op.continueFromText ? null : Object.freeze(Array(9).fill(op.type)) };
    case 'set_list_level': return { ...moved(op.level), anchored: true };
    case 'set_list_type': return { type: op.type, level: s.level, anchored: s.anchored, levels: s.levels && validLevel(s.level) ? Object.freeze(s.levels.map((t, k) => (k === s.level ? op.type : t))) : null };
    case 'remove_list': return { type: 'normal', level: null, anchored: true, levels: null };
    case 'indent_list_item': return s.anchored && s.level !== null && s.level < 8 ? moved(s.level + 1) : { ...s, anchored: false, level: null, unprovable: true };
    case 'outdent_list_item':
      if (!s.anchored || s.level === null) return { ...s, anchored: false, level: null, unprovable: true };
      return s.level > 0 ? moved(s.level - 1) : { type: 'normal', level: null, anchored: true, levels: null };
    default: return s;
  }
}

/**
 * Task-level annotation (called by the adapter planner): per target, the TERMINAL projected state
 * after every list step of the task, whether a relative step lacks an absolute anchor, and the
 * continuation of consecutive apply_list steps with the same type and level (source
 * docx-agent-executor.cjs:61 continueFromText).
 */
function annotateTask(operations) {
  const states = new Map();
  const unprovable = new Set();
  const lastApply = new Map();
  const out = operations.map(op => {
    if (!isList(op)) return op;
    const k = norm(keyText(op));
    let cont = null;
    if (op.intent === 'apply_list') {
      const g = `${op.type}|${op.level}`;
      const prev = lastApply.get(g);
      lastApply.set(g, keyText(op));
      if (prev && norm(prev) !== k) cont = prev;
    }
    const next = project(states.get(k), cont ? { ...op, continueFromText: cont } : op);
    if (next.unprovable) unprovable.add(op.id);
    states.set(k, next);
    return cont ? { ...op, continueFromText: cont } : op;
  });
  return out.map(op => {
    if (!isList(op)) return op;
    const t = states.get(norm(keyText(op)));
    return { ...op, terminal: Object.freeze({ type: t.type, level: t.level }), ...(unprovable.has(op.id) ? { relativeUnprovable: true } : {}) };
  });
}

const isList = op => ['apply_list', 'set_list_level', 'set_list_type', 'indent_list_item', 'outdent_list_item', 'remove_list'].includes(op && op.intent);

// ---- observation --------------------------------------------------------------------------------

/**
 * The list state of the ONE paragraph / list item with this text, from the fresh live inventory:
 * `normal` for a paragraph whose live numbering read says "not in a list"; `bullet|numbered` + level
 * + numbering identity for a list item. Anything unread or inconsistent is UNKNOWN (fail closed).
 */
function listState(primary, text) {
  if (typeof text !== 'string') return { ok: false, reason: 'list-target-text-unknown' };
  const want = norm(text);
  const hits = [];
  for (const b of (primary && primary.blocks) || []) {
    if (TEXT_KINDS.includes(b.kind) && norm(b.text) === want) hits.push({ kind: b.kind, block: b, numbering: b.numbering });
    if (b.kind === 'list') for (const item of b.items || []) if (norm(item.text) === want) hits.push({ kind: 'list_item', block: b, item, numbering: item.numbering });
  }
  if (hits.length === 0) return { ok: false, reason: 'list-target-not_found' };
  if (hits.length > 1) return { ok: false, reason: 'list-target-ambiguous', count: hits.length };
  const h = hits[0];
  const numberingRemoveBridge = runtimeVerified(primary, 'numberingRemoveBridge');
  if (h.kind !== 'list_item') {
    if (!h.numbering) return { ok: false, reason: 'list-state-unreadable' };
    if (h.numbering.present !== false) return { ok: false, reason: 'list-state-inconsistent' };
    return { ok: true, kind: h.kind, text: h.block.text, type: 'normal', level: null, numId: null, node: h.block, numberingRemoveBridge };
  }
  const n = h.numbering;
  if (!n || n.present !== true || !validLevel(n.level)) return { ok: false, reason: 'list-state-unreadable' };
  const type = h.block.ordered ? 'numbered' : 'bullet';
  return { ok: true, kind: 'list_item', text: h.item.text, type, level: n.level, numId: n.numId || null, block: h.block, node: h.item, numberingRemoveBridge, levels: gatedLevelsOf(primary, n, type) };
}

/** M8 runtime gate (owner decision A; identity = core/runtime-identity.cjs since decision (B)): is THIS session's runtime verified for the internal dependency? (absent -> no) */
const runtimeVerified = (primary, dependency) => !!(primary && primary.internalGate && primary.internalGate[dependency] === true);

/**
 * The per-level read (levelsOf). W-B1920-C1: the build-gate check this used to run
 * FIRST is gone -- the list-level format now comes from the Word bridge (docx.numbering.levelFormat.read/v1),
 * which is fail-closed at the SOURCE (a bridge failure already leaves n.levelFormats null before this
 * function is even reached, via listState's own `!h.numbering` refusal); levelsOf's own
 * `!Array.isArray(formats)` check below is what now reports an unread read.
 */
function gatedLevelsOf(primary, n, type) {
  return levelsOf(n.levelFormats, n.level, type);
}

/**
 * M5 gate 7: the numbering's per-level types from the internal format read, cross-checked on the
 * item's own level against the projection. { ok: false, reason } when unread or disagreeing.
 */
function levelsOf(formats, ownLevel, ownType) {
  if (!Array.isArray(formats) || formats.length !== 9) return Object.freeze({ ok: false, reason: 'eo-internal-list-level-format-unavailable' });
  const types = formats.map(f => (Number.isInteger(f) && LEVEL_FORMAT_TYPE[f]) || null);
  if (types[ownLevel] !== ownType) return Object.freeze({ ok: false, reason: 'eo-internal-list-level-format-disagrees', ownLevel, ownFormat: formats[ownLevel], projection: ownType });
  return Object.freeze({ ok: true, formats: Object.freeze(formats.slice()), types: Object.freeze(types) });
}

/** The level a level move lands on (null when the step is not a level move within the list). */
function destinationLevel(s, op) {
  if (op.intent === 'set_list_level') return op.level;
  if (op.intent === 'indent_list_item') return s.level + 1;
  if (op.intent === 'outdent_list_item' && s.level > 0) return s.level - 1;
  return null;
}

/**
 * The capability guard for a step on a list item that needs the per-level read: a level move (the
 * destination type) and set_list_type (the other levels' preservation proof). null = no need / fine.
 */
function levelReadRefusal(s, op) {
  const isListed = s.type === 'bullet' || s.type === 'numbered';
  const needs = isListed && (op.intent === 'set_list_type' || (LEVEL_MOVES.includes(op.intent) && destinationLevel(s, op) !== null));
  if (!needs) return null;
  if (!s.levels || !s.levels.ok) return { reason: (s.levels && s.levels.reason) || 'eo-internal-list-level-format-unavailable', dependency: 'docx.numbering.levelFormat.read/v1', ...(s.levels && s.levels.ownFormat !== undefined ? { ownFormat: s.levels.ownFormat, projection: s.levels.projection } : {}) };
  const d = destinationLevel(s, op);
  if (d !== null && validLevel(d) && !s.levels.types[d]) return { reason: 'list-destination-level-type-unknown', destinationLevel: d, format: s.levels.formats[d] };
  return null;
}

/** Does THIS step's write remove numbering (B18: remove_list, outdent at level 0)? */
const usesNumberingRemoval = (s, op) => op.intent === 'remove_list' ? s.type !== 'normal' : op.intent === 'outdent_list_item' && s.type !== 'normal' && s.level === 0;

/** M11.5d.4: the removal runs ONLY through the word /2 Bridge WRITE -- any other runtime refuses BEFORE the write. null = usable / not needed. */
function removalRefusal(s, op) {
  if (!usesNumberingRemoval(s, op)) return null;
  if (s.numberingRemoveBridge !== true) return 'docx-numbering-remove-bridge-write-runtime-not-proven';
  return null;
}

// A terminal field the task does not determine (null: e.g. the type of a standalone set_list_level)
// is not constrained; a determined one must match exactly.
const matches = (s, t) => !!s && !!t && (t.type === null || t.type === undefined || s.type === t.type) && (t.level === null || t.level === undefined || s.level === t.level);

/** The state THIS step must leave, from the fresh pre-mutation state (source expectedFromObservation). */
function stepExpected(s, op) {
  const isListed = s.type === 'bullet' || s.type === 'numbered';
  switch (op.intent) {
    case 'apply_list': return { type: op.type, level: op.level };
    case 'remove_list': return { type: 'normal', level: null };
    case 'set_list_type': return isListed ? { type: op.type, level: s.level } : null;
    case 'set_list_level':
    case 'indent_list_item':
    case 'outdent_list_item': {
      if (!isListed) return null;
      if (op.intent === 'outdent_list_item' && s.level === 0) return { type: 'normal', level: null };
      const d = destinationLevel(s, op);
      if (!validLevel(d)) return null;
      // M5 gate 7: the destination level's FRESH LIVE type (unknown -> refused before dispatch)
      const t = s.levels && s.levels.ok ? s.levels.types[d] : null;
      return t ? { type: t, level: d } : null;
    }
    default: return null;
  }
}

// ---- classify / resolve -------------------------------------------------------------------------

/**
 * Fresh LIVE classification, retry-safe by the TASK projection (source idempotency :103-116):
 *   target absent / ambiguous / unreadable              -> FAIL_CLOSED
 *   relative step without an absolute anchor in the task -> FAIL_CLOSED (terminal not provable)
 *   the target already holds the task's terminal state   -> SATISFIED (0 command, 0 write)
 *   precondition not met (e.g. set_list_level on a non-list) -> FAIL_CLOSED
 *   otherwise                                              -> APPLY
 */
function classify(primary, op) {
  const s = listState(primary, keyText(op));
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason, count: s.count } };
  if (op.relativeUnprovable) return { state: 'FAIL_CLOSED', evidence: { reason: 'list-relative-level-unprovable', text: s.text } };
  if (matches(s, op.terminal)) return { state: 'SATISFIED', evidence: { text: s.text, terminal: op.terminal, observed: { type: s.type, level: s.level }, proof: 'fresh-live-list-state-is-task-terminal' } };
  const levelRefusal = levelReadRefusal(s, op);
  if (levelRefusal) return { state: 'FAIL_CLOSED', evidence: { ...levelRefusal, text: s.text, observed: { type: s.type, level: s.level } } };
  const expected = stepExpected(s, op);
  if (!expected) return { state: 'FAIL_CLOSED', evidence: { reason: 'list-precondition-not-met', text: s.text, observed: { type: s.type, level: s.level } } };
  const removal = removalRefusal(s, op);
  if (removal) return { state: 'FAIL_CLOSED', evidence: { reason: removal, text: s.text, dependency: `BRIDGE_WRITE: ${NUMBERING_REMOVE_WRITE_ID}` } };
  return { state: 'APPLY', evidence: { text: s.text, observed: { type: s.type, level: s.level }, expected } };
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
  const s = listState(primary, text);
  if (!s.ok) return { ok: false, target: null, selector, ambiguity: s.reason.replace(/^list-/, '').toUpperCase() };
  return { ok: true, target: { kind: s.kind, text: s.text, state: s }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  if (!target || !target.state || !target.state.ok) return { ok: false, dispatched: false, detail: 'list-state-unreadable' };
  const levelRefusal = levelReadRefusal(target.state, op);
  if (levelRefusal) return { ok: false, dispatched: false, detail: levelRefusal.reason };
  if (!stepExpected(target.state, op)) return { ok: false, dispatched: false, detail: 'list-precondition-not-met' };
  const removal = removalRefusal(target.state, op);
  if (removal) return { ok: false, dispatched: false, detail: removal };
  return { ok: true };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const s = target.state;
  const meta = Object.freeze({
    family: 'list', action: 'list', listAction: op.intent, targetText: s.text,
    ...(op.type ? { type: op.type } : {}), ...(op.level !== undefined ? { level: op.level } : {}),
    beforeType: s.type, beforeLevel: s.level, ...(op.continueFromText ? { continueFromText: op.continueFromText } : {}),
    ...(usesNumberingRemoval(s, op) ? { bridgeNumberingRemove: true } : {}),
  });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (source coedit-docx-lists.cjs:7, re-expressed on the public numbering API where
 * one exists). The target by its exact text; the live level is re-checked against the fresh
 * observation before any mutation (a moved state -> no mutation, L1 then fails closed). Removal
 * (remove_list, outdent at level 0): M11.5d.4 -- ONE word /2 Bridge WRITE docx.paragraph.numbering.remove/v1 (the editor's
 * own paragraph routine behind the Bridge, the same target + live level re-check) through the driver's in-editor
 * client __eoB18Remove (prepended ONLY to a removal command); its receipt is returned, never trusted: the adapter's fresh
 * read verifies. A gate/refusal leaves the document untouched (no fallback).
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const doc = Api.GetDocument();
  if (!doc) return;
  const hits = (doc.Search(m.targetText, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || '') === m.targetText);
  if (hits.length !== 1) return;
  const paragraph = hits[0].GetParagraph(0);
  if (!paragraph) return;
  const current = typeof paragraph.GetNumbering === 'function' ? paragraph.GetNumbering() : null;
  const liveLevel = current && typeof current.GetLevelIndex === 'function' ? current.GetLevelIndex() : null;
  if (m.beforeType === 'normal' ? !!current : liveLevel !== m.beforeLevel) return;
  const removeNumbering = () => (m.bridgeNumberingRemove === true ? __eoB18Remove(m.targetText, m.beforeLevel) : { ok: false, stage: 'not-a-removal-command', mutation: 'NONE' });
  if (m.listAction === 'apply_list') {
    let numbering = null;
    if (m.continueFromText) {
      const prev = (doc.Search(m.continueFromText, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || '') === m.continueFromText);
      if (prev.length !== 1) return;
      const pl = prev[0].GetParagraph(0).GetNumbering();
      numbering = pl && typeof pl.GetNumbering === 'function' ? pl.GetNumbering() : null;
      if (!numbering) return;
    } else numbering = doc.CreateNumbering(m.type);
    if (!numbering) return;
    paragraph.SetNumbering(numbering.GetLevel(m.level));
    return;
  }
  if (m.listAction === 'remove_list') return removeNumbering();
  if (!current) return;
  const numbering = current.GetNumbering();
  if (!numbering) return;
  if (m.listAction === 'set_list_type') {
    const level = numbering.GetLevel(m.beforeLevel);
    if (!level || typeof level.SetTemplateType !== 'function') return;
    if (m.type === 'bullet') level.SetTemplateType('bullet', '\\u2022'); else level.SetTemplateType('1.');
    paragraph.SetNumbering(level);
    return;
  }
  let next = m.beforeLevel;
  if (m.listAction === 'set_list_level') next = m.level;
  else if (m.listAction === 'indent_list_item') next = m.beforeLevel + 1;
  else if (m.listAction === 'outdent_list_item') { if (m.beforeLevel === 0) return removeNumbering(); next = m.beforeLevel - 1; }
  if (next < 0 || next > 8) return;
  paragraph.SetNumbering(numbering.GetLevel(next));
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** L1: the target, found again by its text in the FRESH after-state, holds the step's expected state. */
function verifyPrimitive(op, before, after) {
  const b = listState(before && before.primary, keyText(op));
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const expected = stepExpected(b, op);
  const a = listState(after.primary, b.text);
  if (!expected || !a.ok) return { ok: false, evidence: { text: b.text, expected, reason: a.ok ? 'list-precondition-not-met' : `after-${a.reason}` } };
  return { ok: matches(a, expected), evidence: { text: b.text, expected, observed: { type: a.type, level: a.level } } };
}

const imageKey = i => JSON.stringify({ src: i.src || null, alt: i.alt || null, title: i.title || null });
const tableKey = b => JSON.stringify((b.rows || []).map(r => r.map(c => norm(c.text))));

/** Every text unit of a state (paragraphs, headings and list items), with its list state when an item. */
function units(primary) {
  const out = [];
  for (const b of (primary && primary.blocks) || []) {
    if (TEXT_KINDS.includes(b.kind)) out.push({ text: norm(b.text), state: 'normal' });
    if (b.kind === 'list') for (const item of b.items || []) out.push({ text: norm(item.text), state: item.numbering && item.numbering.present ? `${b.ordered ? 'numbered' : 'bullet'}|${item.numbering.level}` : 'unknown' });
  }
  return out;
}

function multisetContains(have, need) {
  const h = new Map();
  for (const v of have) h.set(v, (h.get(v) || 0) + 1);
  const missing = [];
  const n = new Map();
  for (const v of need) n.set(v, (n.get(v) || 0) + 1);
  for (const [k, c] of n) if ((h.get(k) || 0) < c) missing.push({ key: k, expected: c, actual: h.get(k) || 0 });
  return { ok: missing.length === 0, missing };
}

/**
 * L3, PRIMITIVE_INTEGRAL (source docx-session-lists.cjs:59-66 + verifyLiveRoundTrip): the target text
 * exactly once; every OTHER text unit survives; every other LIST ITEM keeps its type and level;
 * table shapes and images survive.
 */
function verifyPreservation(op, before, after) {
  const b = listState(before && before.primary, keyText(op));
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const want = norm(b.text);
  const ua = units(after.primary);
  const targetCount = ua.filter(u => u.text === want).length;
  const others = u => u.text !== want;
  const beforeUnits = units(before.primary).filter(others);
  const text = multisetContains(ua.filter(others).map(u => u.text), beforeUnits.map(u => u.text));
  const listItems = multisetContains(ua.filter(others).map(u => `${u.text}#${u.state}`), beforeUnits.filter(u => u.state !== 'normal').map(u => `${u.text}#${u.state}`));
  const tables = multisetContains((after.primary.blocks || []).filter(x => x.kind === 'table').map(tableKey), (before.primary.blocks || []).filter(x => x.kind === 'table').map(tableKey));
  const images = multisetContains((after.primary.images || []).map(imageKey), (before.primary.images || []).map(imageKey));
  const a = listState(after.primary, b.text);
  const own = ownFormatting(b, a);
  const levels = otherLevelsKept(op, b, a);
  const ok = targetCount === 1 && text.ok && listItems.ok && tables.ok && images.ok && own.ok && levels.ok;
  return { ok, evidence: { targetText: b.text, targetCount, untouchedText: text, otherListItems: listItems, ownFormatting: own, otherLevels: levels, tables, images, expectationSource: 'derived-from-fresh-pre-mutation-observation' } };
}

/**
 * M5 gate 7: set_list_type and a level move change NO level format of the numbering except, for
 * set_list_type, the item's own level. Same numbering before and after; the read on both sides
 * (the guard required it before dispatch) -- unreadable after -> not provable -> fail.
 */
function otherLevelsKept(op, b, a) {
  if (!(op.intent === 'set_list_type' || (LEVEL_MOVES.includes(op.intent) && destinationLevel(b, op) !== null))) return { ok: true, applies: false };
  if (!a.ok || !b.levels || !b.levels.ok || !a.levels || !a.levels.ok) return { ok: false, reason: 'level-formats-unreadable' };
  if (b.numId !== a.numId) return { ok: false, reason: 'numbering-identity-changed', before: b.numId, after: a.numId };
  const changed = [];
  for (let k = 0; k < 9; k += 1) if (!(op.intent === 'set_list_type' && k === b.level) && b.levels.formats[k] !== a.levels.formats[k]) changed.push({ level: k, before: b.levels.formats[k], after: a.levels.formats[k] });
  return { ok: changed.length === 0, applies: true, changed };
}

/**
 * W2.1: a list step changes the target's LIST state only. Its own text formatting -- per character
 * from the runs when both states have them, else the seven text properties of the whole-paragraph
 * read -- must be unchanged, so an earlier rewrite's properties or an earlier format_text survive
 * the conversion (owner brief W2.1 point 6). Unreadable on either side -> not provable -> fail.
 */
const TEXT_KEYS = ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'language'];
function ownFormatting(b, a) {
  if (!a.ok || !b.node || !a.node) return { ok: false, reason: 'target-node-unavailable' };
  const perChar = n => { if (!Array.isArray(n.runs)) return null; const c = []; for (const r of n.runs) for (let i = r.start; i < r.end; i += 1) c[i] = TEXT_KEYS.map(k => r.format[k]); return c; };
  const cb = perChar(b.node);
  const ca = perChar(a.node);
  if (cb && ca) return { ok: JSON.stringify(cb) === JSON.stringify(ca), level: 'runs' };
  if (!b.node.format || !a.node.format) return { ok: false, reason: 'formatting-unreadable' };
  const pick = f => TEXT_KEYS.map(k => f[k]);
  return { ok: JSON.stringify(pick(b.node.format)) === JSON.stringify(pick(a.node.format)), level: 'paragraph' };
}

/** L2: the final state of the target is the task's terminal projection; a continued apply_list shares its numbering. */
function wholeTaskChecks(op, finalObservation) {
  const s = listState(finalObservation.primary, keyText(op));
  const checks = [{ name: `${op.id}: list target holds the task's terminal state ${JSON.stringify(op.terminal)}`, ok: s.ok && matches(s, op.terminal) }];
  if (op.continueFromText) {
    const prev = listState(finalObservation.primary, op.continueFromText);
    checks.push({ name: `${op.id}: continued list shares the numbering of "${op.continueFromText}"`, ok: s.ok && prev.ok && !!s.numId && s.numId === prev.numId && s.block === prev.block });
  }
  return checks;
}

function producedText(op) {
  return op.intent === 'remove_list' ? null : keyText(op);
}

module.exports = {
  // W2.1: a list step turns a paragraph into a list item -- a declared representation transform.
  transformsRepresentation: true,
  plan, annotateTask, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation,
  wholeTaskChecks, producedText, listState, stepExpected, project, levelsOf, LEVEL_FORMAT_TYPE,
};
