'use strict';

const { normalizeInventory, norm } = require('./inventory.cjs');
const { resolveStableTarget } = require('./resolver.cjs');
const { INTENTS, STATUS } = require('./intents.cjs');
// W11 (option b): post-M7 capabilities with no source intent live in their OWN, disjoint
// list (capabilities.cjs's NEW_INTENTS) -- NOT in intents.cjs's INTENTS, so the 28/28 migration
// accounting gate (test/static/docx-intent-accounting.test.cjs) stays untouched. familyOf() below
// looks a row up across BOTH lists; the merged BY_INTENT is the only change to this file's own
// routing logic.
const { NEW_INTENTS } = require('./capabilities.cjs');
const text = require('./families/text.cjs');
const formatting = require('./families/formatting.cjs');
const paragraph = require('./families/paragraph.cjs');
const list = require('./families/list.cjs');
// W-07 (M7 backlog, numbering pattern/start): placed next to `list` (its closest
// sibling, importing list.cjs's own listState -- list.cjs itself untouched), deliberately NOT
// next to `pageSetup` below, where the concurrent W-04 table-format work is expected to insert
// its own new require/FAMILIES entry (the reviewer's own merge-collision warning).
const numbering = require('./families/numbering.cjs');
const rich = require('./families/rich.cjs');
const notes = require('./families/notes.cjs');
const table = require('./families/table.cjs');
const tableRow = require('./families/table-row.cjs');
const tableColumn = require('./families/table-column.cjs');
const tableMergeSplit = require('./families/table-merge-split.cjs');
const headerFooter = require('./families/header-footer.cjs');
const pageSetup = require('./families/page-setup.cjs');
const tableFormat = require('./families/table-format.cjs');
// W-10  -- placed AFTER
// page-setup, not next to list/table, to avoid a concurrent-edit collision on those lines.
const styleDefinition = require('./families/style-definition.cjs');
// GH #268c: placed AFTER style-definition, deliberately NOT next to header-footer/page-setup, to
// avoid a concurrent-edit collision on those lines (same reviewer note as numbering's own placement above).
const image = require('./families/image.cjs');
// GH #269c: a new family, placed after style-definition for the same merge-collision reason.
const comments = require('./families/comments.cjs');
const commentWrite = require('./families/comment-write.cjs'); // GH #269d
const fieldContract = require('./field-contract.cjs');
const constraints = require('./constraints.cjs');

/**
 * THE DOCX adapter — one surface toward the M0 section 5 contract (M4 brief points 6-18).
 *
 * Owns Word SEMANTICS only. It does not open, close, persist, forceSave or hold a lifecycle
 * (M1 owns those; validateAdapter + the static gates prove the absence). It mutates ONLY through
 * the `dispatch` capability the M3 executor hands to executePrimitive, and it reads ONLY the
 * observations the core hands it -- it imports nothing outside src/adapters/docx, so it has no other
 * route to the editor (static gate `docx-adapter-boundary`).
 *
 * Families live in ./families/*. Migrated: `text` (rewrite_paragraph, first checkpoint), in
 * Wave 1 `formatting` (format_text) and `paragraph-structure` (insert/delete/move), in Wave 2 `list`, in
 * Wave 3 `rich-inline` (add_bookmark, add_link) and `notes` (add_footnote, add_endnote), in Wave 4a
 * `table-cell` (replace_table_cell), in Wave 4b `table-row` (insert before/after, delete), in Wave 4c
 * `table-column` (insert before/after, delete), in Wave 4d `table-merge-split` (merge, split), in Wave 5 `header-footer` (edit_header, edit_footer). GH #268c adds
 * `image` (set_image_size, set_image_wrap), a NEW_CAPABILITY family (no source intent, like page-setup/table-format/style-definition above it). Every other
 * registered intent is refused at PLANNING with an explicit reason, before any dispatch.
 */

const FAMILIES = Object.freeze({ text, formatting, 'paragraph-structure': paragraph, list, numbering, 'rich-inline': rich, notes, 'table-cell': table, 'table-row': tableRow, 'table-column': tableColumn, 'table-merge-split': tableMergeSplit, 'header-footer': headerFooter, 'page-setup': pageSetup, 'table-format': tableFormat, 'style-definition': styleDefinition, image, comments, 'comment-write': commentWrite });
// Merged from the 28/28 migration inventory AND the disjoint new-capability list -- see the
// comment above the capabilities.cjs import. duplicates: [] is asserted by
// test/static/docx-new-capabilities.test.cjs (the two lists are disjoint by construction and by gate).
const BY_INTENT = new Map([...INTENTS, ...NEW_INTENTS].map(i => [i.intent, i]));

function familyOf(intent) {
  const row = BY_INTENT.get(intent);
  if (!row) return { ok: false, detail: `intent-unsupported: ${intent}` };
  if (row.status === STATUS.NOT_MIGRATED || !FAMILIES[row.family]) {
    return { ok: false, detail: `intent-not-migrated-in-m4: ${intent} (family ${row.family})` };
  }
  return { ok: true, family: FAMILIES[row.family], row };
}

/**
 * Explicit recovery context (owner decision R2, point 9). The framework keeps NO task state across
 * restarts, so the only way a later invocation can know that an operation previously ended in
 * unresolved uncertainty is for the CALLER to say so in the same invocation:
 *
 *   task.recovery = { unresolvedUncertainOperations: ['<operation id>', ...] }
 *
 * The ids are the operation ids of the earlier ExecutionResult (step.id, or `s<index>`). It can only
 * make classification STRICTER (a P1 terminal state without a provable baseline -> FAIL_CLOSED); it
 * never grants SATISFIED or APPLY. Malformed, or naming an operation this task does not plan ->
 * the plan fails (fail closed, never ignored).
 */
function recoveryIds(task, operations) {
  if (task.recovery === undefined) return { ok: true, ids: new Set() };
  const list = task.recovery && task.recovery.unresolvedUncertainOperations;
  if (!Array.isArray(list) || list.some(x => typeof x !== 'string' || x === '')) {
    return { ok: false, detail: 'task.recovery.unresolvedUncertainOperations must be a list of operation ids' };
  }
  const planned = new Set(operations.map(o => o.id));
  const unknown = list.filter(id => !planned.has(id));
  if (unknown.length) return { ok: false, detail: `task.recovery names operations this task does not plan: ${unknown.join(', ')}` };
  return { ok: true, ids: new Set(list) };
}

/**
 * Agent task -> common M3 operations. Declares, per operation: dependencies/refs (none in the
 * slice), preservation placement, and batch independence -- which is ALWAYS left unproven here
 * (no batchSafeGroup): M4 does not optimize for batching (brief point 11).
 */
function planOperations(task) {
  const steps = (task && task.steps) || [];
  if (!Array.isArray(steps) || steps.length === 0) return { ok: false, operations: [], detail: 'task has no steps' };
  // M5 fail-closed field contract: a field the target does not read and enforce refuses the WHOLE
  // task here, before any dispatch -- never silently dropped (field-contract.cjs).
  const shape = fieldContract.checkTask(task);
  if (!shape.ok) return { ok: false, operations: [], detail: shape.detail };
  // M5: task.constraints + selector.precondition (constraints.cjs). Each operation carries the task's
  // constraints; they are enforced before its dispatch and after its write, never only at the end.
  const cons = constraints.validateConstraints(task);
  if (!cons.ok) return { ok: false, operations: [], detail: cons.detail };
  const operations = [];
  for (const [index, step] of steps.entries()) {
    const f = familyOf(step && step.intent);
    if (!f.ok) return { ok: false, operations: [], detail: f.detail, index };
    const fields = fieldContract.checkStep(step, index);
    if (!fields.ok) return { ok: false, operations: [], detail: fields.detail, index };
    const allowed = constraints.checkStepAgainst(cons.constraints, step, index, TABLE_FAMILIES.includes(f.row.family));
    if (!allowed.ok) return { ok: false, operations: [], detail: allowed.detail, index };
    let familyStep = step;
    let precondition = null;
    if (step.selector && typeof step.selector === 'object' && step.selector.precondition !== undefined) {
      const pre = constraints.validatePrecondition(step.selector.precondition, index);
      if (!pre.ok) return { ok: false, operations: [], detail: pre.detail, index };
      precondition = pre.precondition;
      const rest = { ...step.selector };
      delete rest.precondition;
      familyStep = { ...step, selector: rest };
    }
    const p = f.family.plan(familyStep, index);
    if (!p.ok) return { ok: false, operations: [], detail: p.detail, index };
    let op = p.operation;
    if (precondition) op = { ...op, precondition };
    if (cons.constraints) {
      // the constraint L3 runs after every write: an operation with no preservation contract of its own
      // gets the POST_PRIMITIVE placement; an integral one keeps its placement and adds the constraints
      op = { ...op, constraints: cons.constraints, ...(!op.preservationMode || op.preservationMode === 'NONE' ? { preservationMode: 'POST_PRIMITIVE' } : {}) };
    }
    operations.push(op);
  }
  const ordinal = ordinalAfterStructuralStep(operations);
  if (!ordinal.ok) return { ok: false, operations: [], detail: ordinal.detail, index: ordinal.index };
  const recovery = recoveryIds(task, operations);
  if (!recovery.ok) return { ok: false, operations: [], detail: recovery.detail };
  const withRefs = attachRefText(operations);
  if (!withRefs.ok) return { ok: false, operations: [], detail: withRefs.detail };
  // A family may annotate its operations with task-level facts (list: the projected terminal state
  // per target, the continuation of consecutive apply_list). Pure; the operation ids never change.
  let annotated = withRefs.operations;
  for (const family of new Set(annotated.map(familyFor))) if (family.annotateTask) annotated = family.annotateTask(annotated);
  annotated = attachSuccessors(annotated);
  const marked = annotated.map(op => (recovery.ids.has(op.id) ? { ...op, recovery: text.UNRESOLVED } : op));
  return { ok: true, operations: marked, dependencies: [] };
}

/**
 * A `selector.ref` names an EARLIER operation's semantic result (M3). The planner records that
 * operation's planned semantic text as `op.refText`, so classification and verification -- which
 * receive no refs -- know what the ref names; resolveTarget cross-checks it against the ref's actual
 * evidence in the session (REF_RESULT_MISMATCH fails closed). Never a position.
 */
function attachRefText(operations) {
  const out = [];
  const byId = new Map();
  for (const op of operations) {
    let current = op;
    const ref = op.selector && op.selector.ref;
    if (ref !== undefined) {
      const earlier = byId.get(ref);
      if (!earlier) return { ok: false, detail: `operation ${op.id}: ref ${ref} names no earlier operation of this task` };
      const produced = familyFor(earlier).producedText ? familyFor(earlier).producedText(earlier) : null;
      if (typeof produced !== 'string') return { ok: false, detail: `operation ${op.id}: ref ${ref} names an operation with no semantic text result (${earlier.intent})` };
      current = { ...op, refText: produced };
    }
    out.push(current);
    byId.set(current.id, current);
  }
  return { ok: true, operations: out };
}

/**
 * W2.1 successor-aware final verification (owner brief W2.1 point 3). The planner records, for an
 * operation A, the LATER operations of this task that DECLAREDLY transform the representation of
 * A's semantic result (a family with `transformsRepresentation`, today: list -- paragraph -> list
 * item). A later operation targets A's result when it names A by ref, or names A's produced text.
 * A's final check may then accept the transformed representation -- and ONLY then: an undeclared
 * change keeps the strict check. Pure planner metadata; the common verifier API is unchanged.
 */
function attachSuccessors(operations) {
  const targetText = op => (op.selector && op.selector.ref !== undefined ? op.refText : op.selector && op.selector.text);
  return operations.map((a, i) => {
    const produced = familyFor(a).producedText ? familyFor(a).producedText(a) : null;
    if (typeof produced !== 'string') return a;
    const later = operations.slice(i + 1).filter(b => familyFor(b).transformsRepresentation === true
      && ((b.selector && b.selector.ref === a.id) || (typeof targetText(b) === 'string' && norm(targetText(b)) === norm(produced))));
    return later.length ? { ...a, successorTransforms: Object.freeze(later.map(b => b.id)) } : a;
  });
}

// W04 (table-format, read while studying this registration path, not previously named in v1/v2 of
// the design): 'table-format' writes a table by construction, so a task's doNotTouchTables
// constraint must refuse it too -- the same reason the other 4 table families are here.
const TABLE_FAMILIES = Object.freeze(['table-cell', 'table-row', 'table-column', 'table-merge-split', 'table-format']);
// The steps that change how many paragraph blocks a section holds (or which ones). An ordinal names a
// position counted on the state it is resolved against; with such a step ANYWHERE in the same task the
// count differs between the first run and a reopened retry (a LATER step has already run then), so the
// retry could name a different paragraph and write again. The task is refused at planning.
const COUNT_CHANGING = Object.freeze(['insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'apply_list', 'remove_list']);
const changesCount = op => COUNT_CHANGING.includes(op.intent) || (op.intent === 'rewrite_paragraph' && !op.text);
const usesOrdinal = op => !!(op.selector && (op.selector.sectionHeading !== undefined || op.selector.paragraphOrdinal !== undefined));
// M5 gate 3: an ordinal INSERT is the one count-changing step whose own retry is provable -- its terminal
// position is an ordinal too (the new paragraph IS ordinal N before / N+1 after). It is accepted ONLY as
// the task's single ordinal and single count-changing step; every other combination stays refused.
const ORDINAL_INSERT = Object.freeze(['insert_paragraph_before', 'insert_paragraph_after']);
function ordinalAfterStructuralStep(operations) {
  const ordinal = operations.findIndex(usesOrdinal);
  if (ordinal < 0) return { ok: true };
  const op = operations[ordinal];
  if (ORDINAL_INSERT.includes(op.intent)) {
    const other = operations.find(o => o !== op && (changesCount(o) || usesOrdinal(o)));
    if (other) return { ok: false, index: ordinal, detail: `ordinal-insert-with-another-ordinal-or-count-changing-step: operation ${op.id} and ${other.id} (${other.intent}) in one task; the insert's terminal ordinal would be counted on a different paragraph set` };
    return { ok: true };
  }
  if (changesCount(op)) return { ok: false, index: ordinal, detail: `ordinal-selector-on-count-changing-step: operation ${op.id} (${op.intent}) would not name the same paragraph on a retry` };
  const other = operations.find(changesCount);
  if (other) return { ok: false, index: ordinal, detail: `ordinal-selector-with-count-changing-step: operation ${op.id} and ${other.id} (${other.intent}) in one task; on a retry the ordinal would be counted on a different paragraph set` };
  return { ok: true };
}

function familyFor(op) {
  const f = familyOf(op && op.intent);
  if (!f.ok) throw new Error(f.detail);
  return f.family;
}

const docxAdapter = {
  kind: 'docx',

  /**
   * Semantic readiness. The DOCX transport opens the editor only once its live model answers a
   * non-empty projection (driver.open), and the adapter has no read capability of its own before
   * READY. The probe therefore reports that fact rather than pretending to measure it again.
   */
  async readinessProbe() {
    return { ok: true, detail: 'transport-level readiness: the live editor model answered at open' };
  },

  normalizeInventory,

  // No package enrichment in the M4 slice: rewrite_paragraph needs none (source BE:39). The M2
  // transport exists (core readPackage); families that need package evidence add a hook later.
  enrichObservation: null,

  resolveTarget(primary, selector, ctx = {}) {
    const op = ctx.operation;
    if (!(op && op.intent)) return resolveStableTarget(primary, selector);
    const r = familyFor(op).resolve(primary, selector, ctx);
    if (!r || r.ok !== true || (!op.precondition && !op.constraints)) return r;
    // M5: precondition + constraints on the FRESH targeting observation, before any dispatch
    const pre = constraints.preDispatch(primary, op, r.target);
    return pre.ok ? r : { ok: false, target: null, selector, ambiguity: pre.ambiguity, detail: pre.detail };
  },

  planOperations,

  /**
   * W-10-8: the executor's OPTIONAL hook (src/core/execution/executor.cjs) --
   * derives the task's own observation request, so a caller does not need to know that
   * style-definition's targeted read exists at all (a caller-side `options.observation` still
   * overrides this entirely, for the rare case that needs to). Only `define_style`/`modify_style`
   * steps contribute style names (`styleNamesOf`, families/style-definition.cjs); every other
   * task shape returns `{}`, the exact same request every other family already got before this
   * hook existed.
   */
  // W-10-10: format_text{style} needs the SAME targeted-read coverage
  // define_style/modify_style already get (W10-8) -- doc.GetAllStyles() does not see ANY custom
  // style, not only a freshly-created one (measured, LIVE regression against the W10-9 fix) --
  // so this family's own requested style names are merged in too, from the SAME task.
  observationRequestFor(task) {
    const names = [...new Set([...styleDefinition.styleNamesOf(task), ...formatting.styleNamesOf(task)])];
    // W-07: a set_numbering step needs the PUBLIC per-level numbering format read (opt-in; every other read stays
    // Bridge-only and byte-identical).
    // GH #293 L7: set_numbering / set_list_type also ask for the word /5 Bridge's whole affected set of their target item
    const members = numberingMembersWanted(task);
    const live = { ...(names.length ? { styleNames: names } : {}), ...(numberingFormatWanted(task) ? { numberingFormat: true } : {}), ...(members.length ? { numberingMembers: members } : {}), ...(commentsWanted(task) ? { includeComments: true } : {}) };
    return Object.keys(live).length ? { live } : {};
  },

  classifyOperation(primary, op) {
    return familyFor(op).classify(primary, op);
  },

  // Task-level form required by the M0 contract; the executor uses classifyOperation when present.
  classifyIdempotency(primary, task) {
    const planned = planOperations(task);
    if (!planned.ok) return { state: 'FAIL_CLOSED', evidence: { reason: planned.detail } };
    const states = planned.operations.map(op => familyFor(op).classify(primary, op).state);
    if (states.includes('FAIL_CLOSED')) return { state: 'FAIL_CLOSED', evidence: { states } };
    return { state: states.every(s => s === 'SATISFIED') ? 'SATISFIED' : 'APPLY', evidence: { states } };
  },

  async executePrimitive(_handle, op, { dispatch, target }) {
    const family = familyFor(op);
    const pre = family.preDispatchCheck ? family.preDispatchCheck(op, target) : { ok: true };
    if (!pre.ok) return pre; // refused before the dispatch surface: certain, no mutation
    return dispatch(family.lower(op, target));
  },

  async verifyPrimitive(_handle, op, before, after) {
    return familyFor(op).verifyPrimitive(op, before, after);
  },

  async verifyPreservation(_handle, op, before, after) {
    const family = familyFor(op);
    const own = family.verifyPreservation && op.preservationMode !== 'POST_PRIMITIVE' ? family.verifyPreservation(op, before, after) : null;
    if (!op.constraints) return own;
    // the operation's own target in the before-state: format_text may change its requested keys only,
    // any other intent changes its own target by definition (the family's L1/L3 judge it)
    let ownTarget = null;
    if (before && before.primary) {
      if (op.selector && op.selector.ref !== undefined) ownTarget = op.refText;
      else {
        let r = null;
        try { r = family.resolve(before.primary, op.selector, { operation: op }); } catch { r = null; }
        ownTarget = r && r.ok ? constraints.targetText(r.target) : null;
      }
    }
    const cons = constraints.verifyConstraintPreservation(op, before, after, ownTarget);
    if (!own) return cons;
    return { ok: own.ok === true && cons.ok === true, evidence: { family: own, constraints: cons.evidence } };
  },

  async verifyWholeTask(_handle, task, finalObservation) {
    const planned = planOperations(task);
    if (!planned.ok) return { ok: false, checks: [{ name: 'task plans', ok: false, detail: planned.detail }] };
    const checks = planned.operations.flatMap(op => familyFor(op).wholeTaskChecks(op, finalObservation));
    const c = planned.operations[0] && planned.operations[0].constraints;
    if (c && c.section !== undefined) {
      const s = constraints.sectionBounds((finalObservation.primary && finalObservation.primary.blocks) || [], c.section);
      checks.push({ name: `constraint: the section heading "${c.section}" still exists once`, ok: s.ok, detail: s.ok ? null : s.reason });
    }
    return { ok: checks.length > 0 && checks.every(c => c.ok === true), checks };
  },

  terminalContract: null,
};

/** W-07: set_numbering asks for the PUBLIC per-level numbering format read (`request.numberingFormat`); GH #293 N6: set_list_type
 * too -- its shared-level guard reads the deeper levels' lvlText (does a label quote the target level?). */
function numberingFormatWanted(task) {
  const steps = task && Array.isArray(task.steps) ? task.steps : [];
  return steps.some(st => st && (st.intent === 'set_numbering' || st.intent === 'set_list_type'));
}

/** GH #293 L7: the target texts of the set_numbering / set_list_type steps (a ref selector by the text the plan resolved for it); a task
 * that does not plan asks for none -- it fails at planning anyway. */
function numberingMembersWanted(task) {
  const planned = planOperations(task);
  if (!planned.ok) return [];
  return [...new Set(planned.operations.filter(op => op.intent === 'set_numbering' || op.intent === 'set_list_type').map(op => numbering.keyTextOf(op)).filter(t => typeof t === 'string' && t))];
}

/** GH #269c/#269d: only the comment intents ask for the main document's comments (`request.includeComments`, GH #269b). */
function commentsWanted(task) {
  const steps = task && Array.isArray(task.steps) ? task.steps : [];
  return steps.some(st => st && ['resolve_comment', 'reply_comment', 'add_comment_on_match'].includes(st.intent));
}

module.exports = { docxAdapter, planOperations };
