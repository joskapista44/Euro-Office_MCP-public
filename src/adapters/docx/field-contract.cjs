'use strict';

/**
 * M5 fail-closed FIELD CONTRACT for the DOCX agent task (owner decision, owner channel, after 28799,
 * point 1). Before this guard the target silently IGNORED fields it does not implement: a task
 * carrying `constraints` or a `selector.precondition` was executed as if they were absent (measured
 * offline: a mismatching precondition -> APPLIED, 1 write; the source refuses with 0 writes). A
 * caller-requested protection that is dropped without a word is fail-OPEN.
 *
 * The rule: a task is accepted only if EVERY field it carries is one the target reads and enforces
 * for that intent. Anything else is refused at PLANNING, for the whole task, before the first
 * dispatch (0 dispatch / 0 write). This is a TRANSITIONAL protection, not the parity: the source
 * semantics (constraints.section / doNotTouchTables / preserveFormattingExceptRequested,
 * selector.precondition, sectionHeading + paragraphOrdinal) are to be migrated in M5; each migrated
 * field then moves into the accepted sets below with its enforcement and its tests.
 *
 * The accepted sets are the fields each family's plan()/resolve() actually reads (M5 audit of
 * families/*.cjs). A field valid for one intent is refused on another that would drop it (e.g.
 * `level` on set_list_type, `dependsOn` on rewrite_paragraph, whose plan does not carry it).
 */

// M5: `constraints` migrated (constraints.cjs validates its shape and enforces it).
const TASK_KEYS = Object.freeze(['id', 'steps', 'recovery', 'constraints']);

// Source (f8c6d6cf) fields the target does not enforce yet -- named, so the refusal says what is
// missing instead of a generic "unknown field". Their migration is the M5 parity work.
const NOT_ENFORCED_YET = Object.freeze({
  task: Object.freeze({}),
  step: Object.freeze({ targetRef: 'step.targetRef (source reference syntax; the target reads selector.ref)', destinationRef: 'step.destinationRef', ref: 'step.ref (source producer id; the target reads step.id)', terminalTargetRowCellCount: 'step.terminalTargetRowCellCount (source split terminal assertion)' }),
  // M5: precondition migrated on `selector` (not on destinationSelector / endSelector); sectionHeading +
  // paragraphOrdinal migrated for rewrite_paragraph and format_text (the intents that keep the section's
  // paragraph count, so an ordinal names the same paragraph on a retry), and -- M5 gate 3, the source
  // proven LIVE (m5-srcdiff run ...61251a) -- for insert_paragraph_before / _after, whose TERMINAL
  // position is itself an ordinal (paragraph.cjs ordinalTerminal).
  selector: Object.freeze({ precondition: 'selector.precondition (on this selector field)', sectionHeading: 'selector.sectionHeading (on this intent)', paragraphOrdinal: 'selector.paragraphOrdinal (on this intent)' }),
});

// The keys the shared resolver READS (resolver.cjs resolveStableTarget): the stable semantic
// selector, plus the keys it REFUSES EXPLICITLY with a named reason at resolution (physical
// locators -> PHYSICAL_LOCATOR_REFUSED; bookmark / tableCell -> SELECTOR_FAMILY_NOT_MIGRATED). The
// latter are not silently dropped, so they stay on the resolver's own, more specific path.
const RESOLVER_KEYS = Object.freeze(['fingerprint', 'heading', 'level', 'text', 'contains', 'kind', 'headingPath', 'before', 'after',
  'index', 'position', 'offset', 'paragraphIndex', 'bookmark', 'tableCell']);
const WITH_REF = Object.freeze([...RESOLVER_KEYS, 'ref']);
const ORDINAL = Object.freeze(['sectionHeading', 'paragraphOrdinal']);
// `precondition` is branch-neutral: it constrains the resolved target, whichever branch resolved it.
const BRANCH_NEUTRAL = Object.freeze(['precondition']);
// The resolver picks ONE branch and reads only that branch's keys (fingerprint > heading > text);
// a ref resolves to the referenced operation's planned text alone. A key outside the chosen branch
// would be silently ignored, so it is refused. The explicitly refused keys stay allowed in every
// branch: the resolver refuses them first, with their own reason.
const EXPLICITLY_REFUSED = Object.freeze(['index', 'position', 'offset', 'paragraphIndex', 'bookmark', 'tableCell']);
function branchOf(sel) {
  if (sel.ref !== undefined) return { name: 'ref', keys: ['ref'] };
  if (sel.sectionHeading !== undefined || sel.paragraphOrdinal !== undefined) return { name: 'sectionHeading', keys: ORDINAL };
  if (sel.fingerprint) return { name: 'fingerprint', keys: ['fingerprint'] };
  if (sel.heading) return { name: 'heading', keys: ['heading', 'level'] };
  if (sel.text !== undefined) return { name: 'text', keys: ['text', 'contains', 'kind', 'headingPath', 'before', 'after'] };
  return null; // no branch: the resolver refuses it (UNSUPPORTED_SELECTOR)
}

const C = ['intent', 'id'];
const D = ['dependsOn'];
const STEP_KEYS = Object.freeze({
  rewrite_paragraph: [...C, 'selector', 'text', 'expectedPreservation'],
  format_text: [...C, ...D, 'selector', 'format', 'matchText'],
  insert_paragraph_before: [...C, ...D, 'selector', 'text'],
  insert_paragraph_after: [...C, ...D, 'selector', 'text'],
  delete_paragraph: [...C, ...D, 'selector'],
  move_paragraph_before: [...C, ...D, 'selector', 'destinationSelector'],
  move_paragraph_after: [...C, ...D, 'selector', 'destinationSelector'],
  apply_list: [...C, ...D, 'selector', 'type', 'level'],
  set_list_level: [...C, ...D, 'selector', 'level'],
  set_list_type: [...C, ...D, 'selector', 'type'],
  indent_list_item: [...C, ...D, 'selector'],
  outdent_list_item: [...C, ...D, 'selector'],
  remove_list: [...C, ...D, 'selector'],
  add_bookmark: [...C, ...D, 'selector', 'name'],
  add_link: [...C, ...D, 'selector', 'url', 'tooltip'],
  add_footnote: [...C, ...D, 'selector', 'text'],
  add_endnote: [...C, ...D, 'selector', 'text'],
  replace_table_cell: [...C, ...D, 'selector', 'text'],
  insert_table_row_before: [...C, ...D, 'selector', 'terminalRowCount', 'terminalTableAnchor'],
  insert_table_row_after: [...C, ...D, 'selector', 'terminalRowCount', 'terminalTableAnchor'],
  delete_table_row: [...C, ...D, 'selector', 'terminalRowCount', 'terminalTableAnchor'],
  insert_table_column_before: [...C, ...D, 'selector', 'terminalColumnCount', 'terminalTableAnchor'],
  insert_table_column_after: [...C, ...D, 'selector', 'terminalColumnCount', 'terminalTableAnchor'],
  delete_table_column: [...C, ...D, 'selector', 'terminalColumnCount', 'terminalTableAnchor'],
  merge_table_cells: [...C, ...D, 'selector', 'endSelector', 'terminalMergedText', 'terminalMergedSpan', 'terminalTableAnchor'],
  split_table_cell: [...C, ...D, 'selector', 'splitRows', 'splitCols', 'terminalTableAnchor'],
  // `selector` stays accepted here so the family keeps refusing it with its own, precise reason.
  edit_header: [...C, ...D, 'selector', 'section', 'text', 'variant'],
  edit_footer: [...C, ...D, 'selector', 'section', 'text', 'variant'],
});

// Selector key sets for the families that resolve through the shared resolver and would otherwise
// ignore an unknown key. list / rich / notes / the table families already accept only exact key
// sets in their own plan() and are not listed (their refusal is the family's).
const SELECTOR_KEYS = Object.freeze({
  rewrite_paragraph: { selector: [...RESOLVER_KEYS, ...ORDINAL, 'precondition'] },
  format_text: { selector: [...WITH_REF, ...ORDINAL, 'precondition'] },
  insert_paragraph_before: { selector: [...WITH_REF, ...ORDINAL, 'precondition'] },
  insert_paragraph_after: { selector: [...WITH_REF, ...ORDINAL, 'precondition'] },
  delete_paragraph: { selector: [...WITH_REF, 'precondition'] },
  move_paragraph_before: { selector: [...WITH_REF, 'precondition'], destinationSelector: WITH_REF },
  move_paragraph_after: { selector: [...WITH_REF, 'precondition'], destinationSelector: WITH_REF },
});

const isObject = v => v !== null && typeof v === 'object' && !Array.isArray(v);

function refuse(where, key, table) {
  const named = table && table[key];
  return { ok: false, detail: named ? `field-not-enforced-yet: ${where}: ${named} is not enforced by the target (M5 fail-closed guard)` : `field-unsupported: ${where}.${key}` };
}

function checkTask(task) {
  if (!isObject(task)) return { ok: true }; // the planner reports the missing task/steps itself
  for (const key of Object.keys(task)) if (!TASK_KEYS.includes(key)) return refuse('task', key, NOT_ENFORCED_YET.task);
  return { ok: true };
}

function checkStep(step, index) {
  if (!isObject(step)) return { ok: true }; // the family plan reports a malformed step
  const allowed = STEP_KEYS[step.intent];
  if (!allowed) return { ok: true }; // unknown intents are refused by familyOf with their own reason
  for (const key of Object.keys(step)) {
    if (!allowed.includes(key)) return refuse(`step ${index} (${step.intent})`, key, NOT_ENFORCED_YET.step);
  }
  for (const [field, keys] of Object.entries(SELECTOR_KEYS[step.intent] || {})) {
    const sel = step[field];
    if (!isObject(sel)) continue; // presence/shape is the family's check
    for (const key of Object.keys(sel)) {
      if (!keys.includes(key)) return refuse(`step ${index} (${step.intent}) ${field}`, key, NOT_ENFORCED_YET.selector);
    }
    const branch = branchOf(sel);
    if (branch) {
      const ignored = Object.keys(sel).filter(k => !branch.keys.includes(k) && !EXPLICITLY_REFUSED.includes(k) && !BRANCH_NEUTRAL.includes(k));
      if (ignored.length) return { ok: false, detail: `field-unsupported: step ${index} (${step.intent}) ${field}: ${ignored.join(', ')} would be ignored beside the ${branch.name} selector` };
    }
  }
  return { ok: true };
}

module.exports = { TASK_KEYS, STEP_KEYS, SELECTOR_KEYS, RESOLVER_KEYS, NOT_ENFORCED_YET, checkTask, checkStep };
