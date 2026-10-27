'use strict';

const { normalizeInventory } = require('./inventory.cjs');
const { parseRange, overlaps } = require('./a1.cjs');
const protectionGate = require('./protection-gate.cjs');
const cells = require('./families/cells.cjs');
const clear = require('./families/clear.cjs');
const sheets = require('./families/sheets.cjs');
const merge = require('./families/merge.cjs');
const layout = require('./families/layout.cjs');
const format = require('./families/format.cjs');
const formatRun = require('./families/format-run.cjs');
const freeze = require('./families/freeze.cjs');
const structure = require('./families/structure.cjs');
const deleteCells = require('./families/delete-cells.cjs');
const range = require('./families/range.cjs');
const names = require('./families/names.cjs');
const sort = require('./families/sort.cjs');
const validation = require('./families/validation.cjs');
const cf = require('./families/cf.cjs');
const filter = require('./families/filter.cjs');
const charts = require('./families/charts.cjs');
const drawingsInsert = require('./families/drawings-insert.cjs');
const shapeTextFamily = require('./families/shape-text.cjs');
const chartPointLabels = require('./families/chart-point-labels.cjs');
const shapeRotation = require('./families/shape-rotation.cjs');
const sheetVisibility = require('./families/sheet-visibility.cjs');
const page = require('./families/page.cjs');
const print = require('./families/print.cjs');
const pivots = require('./families/pivots.cjs');
const pivotRefresh = require('./families/pivot-refresh.cjs');
const arrayFormula = require('./families/array-formula.cjs');
const comments = require('./families/comments.cjs');

/**
 * THE XLSX adapter -- one surface toward the M0 section 5 contract (M6).
 *
 * Owns spreadsheet SEMANTICS only: no open, close, persist or forceSave (the M1 core owns the
 * lifecycle; validateAdapter proves the absence). It mutates only through the `dispatch` capability
 * the M3 executor hands to executePrimitive, and reads only the observations the core hands it.
 *
 * Migrated so far: `cells` (write_range with values -- wave 1 -- and formulas -- wave 2), `clear` (clear_range, content only -- wave 1b),
 * `sheets` (create/rename/delete/move_sheet -- wave 3; copy_sheet through the build-gated internal write, refused
 * while the validation / conditional-formatting rules cannot be observed), `merge` (merge_range / unmerge_range -- wave 4),
 * `layout` (layout_range: width / height / hide / show, no AutoFit -- wave 4), `format` (format_range: fill, wrap,
 * number format, font, alignment and borders through the build-gated internal reader -- wave 4), `freeze`
 * (freeze_panes / unfreeze_panes: activate + restore, the proof through the build-gated own-pane reader -- wave 4),
 * `structure` (insert / delete rows / columns: anchor + terminal position, reference-aware shift -- wave 5), `range`
 * (copy_range / move_range: occupancy guard, the move's anchor retry -- wave 5), `names` (set / delete a workbook-level
 * defined name through the build-gated names reader -- wave 6), `sort` (sort_range: one key, row integrity -- wave 6),
 * `validation` (set / clear_validation: the measured whole-number rule, exact ranges -- wave 6), `cf` (add / delete a
 * conditional format: full rule identity, the measured cellValue > rule -- wave 6), `filter` (filter_range /
 * clear_filter: the build-gated field identity, proven by W6 / W6b -- wave 6), `charts` (set_chart create /
 * rename_chart: the full identity derived from the data range, public getters -- wave 7), `drawings-insert`
 * (insert_shape / insert_image: cell-aligned position, size never exact-matched (the SAME ±1 EMU owner-tolerance
 * question as the chart-size one), name as retry identity -- X-O08, FORMATTING COMPLETION backlog's own standing
 * mandate, docs/X-O08-DRAWING-INSERT-DESIGN-20261007.md), `page` (set_page_layout:
 * orientation / gridlines / headings / margins in mm -- wave 8), `print` (fit + flag, odd header, row break, first page
 * number, horizontal centering on the sheet model, build-gated -- wave 8).
 * Every other intent is refused at
 * PLANNING, before any dispatch. The inventory of what remains is docs/evidence/M6-XLSX-CAPABILITY-MATRIX.md.
 */

// write_range's own L3 adds the observed style axis (M7 S3); cells.verifyPreservation stays the content L3 every other
// family builds on
const write = Object.freeze({ ...cells, verifyPreservation: cells.writePreservation });

const FAMILIES = Object.freeze({ write_range: write, clear_range: clear, ...Object.fromEntries(sheets.INTENTS.map(i => [i, sheets])), ...Object.fromEntries(merge.INTENTS.map(i => [i, merge])), layout_range: layout, format_range: format, ...Object.fromEntries(formatRun.INTENTS.map(i => [i, formatRun])), ...Object.fromEntries(freeze.INTENTS.map(i => [i, freeze])), ...Object.fromEntries(structure.INTENTS.map(i => [i, structure])), ...Object.fromEntries(deleteCells.INTENTS.map(i => [i, deleteCells])), ...Object.fromEntries(range.INTENTS.map(i => [i, range])), ...Object.fromEntries(names.INTENTS.map(i => [i, names])), sort_range: sort, ...Object.fromEntries(validation.INTENTS.map(i => [i, validation])), ...Object.fromEntries(cf.INTENTS.map(i => [i, cf])), ...Object.fromEntries(filter.INTENTS.map(i => [i, filter])), ...Object.fromEntries(charts.INTENTS.map(i => [i, charts])), ...Object.fromEntries(drawingsInsert.INTENTS.map(i => [i, drawingsInsert])), set_page_layout: page, ...Object.fromEntries(print.INTENTS.map(i => [i, print])), ...Object.fromEntries(pivots.INTENTS.map(i => [i, pivots])), ...Object.fromEntries(pivotRefresh.INTENTS.map(i => [i, pivotRefresh])), ...Object.fromEntries(arrayFormula.INTENTS.map(i => [i, arrayFormula])), ...Object.fromEntries(shapeTextFamily.INTENTS.map(i => [i, shapeTextFamily])), ...Object.fromEntries(chartPointLabels.INTENTS.map(i => [i, chartPointLabels])), ...Object.fromEntries(shapeRotation.INTENTS.map(i => [i, shapeRotation])), ...Object.fromEntries(sheetVisibility.INTENTS.map(i => [i, sheetVisibility])), ...Object.fromEntries(comments.INTENTS.map(i => [i, comments])) });

function familyOf(intent) {
  const f = FAMILIES[intent];
  return f ? { ok: true, family: f } : { ok: false, detail: `intent-not-migrated-in-m6: ${intent}` };
}

function familyFor(op) {
  const f = familyOf(op && op.intent);
  if (!f.ok) throw new Error(f.detail);
  return f.family;
}

// GH #76 MODEL 1 (owner decision 2026-10-18): verifyStoredContent is a task-level field the planner must
// tolerate (it carries no planning meaning here -- write-tools.cjs reads it after the plan has already run).
const TASK_FIELDS = Object.freeze(['id', 'steps', 'verifyStoredContent']);
// GH #326 (U4b): every reason protection-gate.cjs's gate() can return -- preflightDecisive below matches on this.
const PROTECTION_REFUSAL = /^xlsx-protect(ed-sheet-forbids|ed-locked-cells|ion-unobservable|ed-sheet-intent-not-measured):/;

/**
 * Agent task -> common M3 operations. Two writes whose ranges overlap on the same sheet are refused:
 * the terminal state of the overlap would depend on the order, and a reopened retry could not tell
 * which write it is looking at. A step on a sheet-level setting or a named object (a page layout, a chart name)
 * has no cell footprint (`selector.cellFree`) and names what it owns in `selector.claims`: two steps that claim
 * one object are refused the same way, and so is a step that USES what another claims (`selector.uses`: a formula
 * using a defined name another step deletes; 'name:*' = an unparseable formula may use any name) (W9 composition).
 *
 * M7 S3 (owner GO 2026-10-01; review of 93ff8899): a step may declare what it modifies PER AXIS, with the area on
 * that axis (`operation.footprint`): value / style -> an A1 range; columns / rows -> [from, to] or '*' (every one).
 * write_range { value }; format_range { style } + its rows for a wrap; layout_range a row height / hidden { rows }, a
 * hidden column { columns }, a column width { columns, rows: '*' } (it re-fits a wrapped row's AUTO height anywhere).
 * Two steps on one sheet that both declare are refused when they meet on ANY shared axis -- cells or not (A1 wrap +
 * column B width meet on row 1). Steps that meet on no axis may share cells; the later then depends on the earlier (a
 * barrier: its L1 / L3 run on the state the earlier produced). A step that declares nothing (an unknown or
 * multi-axis effect: merge, clear, sort, ...) gets no exemption: any cell overlap is refused.
 */
const BATCH_KEY = 'xlsx-write-values';
const batchProvable = op => op.intent === 'write_range' && Array.isArray(op.cells) && op.cells.every(c => c.formula === undefined && c.nf === undefined);

function planOperations(task) {
  if (!task || typeof task !== 'object') return { ok: false, operations: [], detail: 'task is not an object' };
  const extra = Object.keys(task).filter(k => !TASK_FIELDS.includes(k));
  if (extra.length) return { ok: false, operations: [], detail: `task field(s) not supported: ${extra.join(', ')}` };
  const steps = task.steps;
  if (!Array.isArray(steps) || steps.length === 0) return { ok: false, operations: [], detail: 'task has no steps' };
  const operations = [];
  for (const [index, step] of steps.entries()) {
    if (!step || typeof step !== 'object') return { ok: false, operations: [], detail: `step ${index} is not an object`, index };
    const f = familyOf(step.intent);
    if (!f.ok) return { ok: false, operations: [], detail: f.detail, index };
    const p = f.family.plan(step, index);
    if (!p.ok) return { ok: false, operations: [], detail: p.detail, index };
    operations.push(p.operation);
  }
  // a sheet operation changes the sheet graph every other step would be resolved against: it is the
  // only step of its task (combinations belong to composition, W9)
  if (operations.length > 1 && operations.some(o => o.family === 'sheets')) {
    return { ok: false, operations: [], detail: 'a sheet operation must be the only step of its task' };
  }
  // a freeze moves the active sheet during its command (activate + restore). M7 S5 S23 (owner M7 GO 2026-10-03; audit
  // S23; the source ran every freeze goal last): ONE freeze / unfreeze per task, and in a mixed task its LAST step --
  // every earlier step is applied and verified before the activation, its L3 and the whole-task checks after it
  const freezes = operations.filter(o => o.family === 'freeze').length;
  if (freezes > 1) return { ok: false, operations: [], detail: 'at most one freeze / unfreeze per task' };
  if (freezes === 1 && operations[operations.length - 1].family !== 'freeze') {
    return { ok: false, operations: [], detail: 'a freeze / unfreeze in a mixed task must be its LAST step' };
  }
  // a structural step shifts every coordinate the other steps would name. L-19 C (owner-decision C,
  // 2026-10-17): delete_rows is exempt when it is EVERY step of the task -- each repeat is classified /
  // resolved against the live state the task's own earlier steps already produced (executor.cjs carries
  // the post-dispatch observation forward between groups; delete_rows never earns a batchSafeGroup, so
  // plan.cjs's NOT_PROVEN rule forces each one into its own singleton group regardless of strategy).
  if (operations.length > 1 && operations.some(o => o.family === 'structure')) {
    if (!operations.every(o => o.intent === 'delete_rows')) {
      return { ok: false, operations: [], detail: 'an insert / delete of rows or columns must be the only step of its task, unless every step of the task is delete_rows' };
    }
  }
  // delete_cells shifts every coordinate in its own band -- same reasoning as `structure`, v1 (L-18/1)
  if (operations.length > 1 && operations.some(o => o.family === 'delete-cells')) {
    return { ok: false, operations: [], detail: 'delete_cells must be the only step of its task' };
  }
  // GH #301 D4: an array formula's guard reads the area's one-cell RING too -- another step of the same task could
  // change what that read saw, so v1 keeps it the only step of its task (array-formula.cjs)
  if (operations.length > 1 && operations.some(o => o.family === 'array-formula')) {
    return { ok: false, operations: [], detail: 'write_array_formula must be the only step of its task' };
  }
  // GH #300 K2: a comment's guard reads the sheet's whole comment list (a second comment on a cell is a silent no-op),
  // and its L3 compares every other comment -- another step could change both, so v1 keeps it the only step of its task
  if (operations.length > 1 && operations.some(o => o.family === 'comments')) {
    return { ok: false, operations: [], detail: 'add_comment must be the only step of its task' };
  }
  // a copy / move touches two blocks (and a move re-points references): the only step of its task
  if (operations.length > 1 && operations.some(o => o.family === 'range')) {
    return { ok: false, operations: [], detail: 'a copy_range / move_range must be the only step of its task' };
  }
  const ids = operations.map(o => o.id);
  if (new Set(ids).size !== ids.length) return { ok: false, operations: [], detail: 'duplicate operation ids' };
  for (let i = 0; i < operations.length; i += 1) {
    for (let j = i + 1; j < operations.length; j += 1) {
      const pair = pairOf(operations[i], operations[j]);
      if (pair.clash) return { ok: false, operations: [], detail: `overlapping-writes-refused: ${operations[i].id} and ${operations[j].id} on ${pair.clash}`, index: j };
      if (pair.overlap) operations[j].dependsOn = [...new Set([...(operations[j].dependsOn || []), operations[i].id])];
    }
  }
  // GH #99 (owner GO 2026-10-14): the batch proof -- a plain-value write_range (no formula, no typed value) only
  // writes its own window, so consecutive such steps may share one targeting read and one post-read (the executor's
  // batch strategy); an overlap is already a dependsOn above, and a dependency always cuts the group (plan.cjs).
  for (const op of operations) if (batchProvable(op)) op.batchSafeGroup = BATCH_KEY;
  return { ok: true, operations, dependencies: [] };
}

/**
 * Two planned steps: { clash } = the refusal text (a cell overlap on a shared or undeclared axis, a claimed object, a
 * used name); { overlap } = their cells overlap on disjoint axes (the later one depends on the earlier).
 */
function pairOf(x, y) {
  // L-19 C: two structure-family steps of the SAME task (both delete_rows, by the homogeneity rule above) name
  // their ranges in two DIFFERENT points of the task's own timeline -- the second step's `at` already accounts
  // for the shift the first one produced. A literal-range comparison would compare across that shift and could
  // refuse a valid pair (e.g. step 1 deletes rows 10-14; step 2's post-shift `at`=12 names what was originally
  // rows 17-21, and 12-16 literally overlaps 10-14 without being the same rows). Correctness here is the live,
  // per-step anchor read (structure.classify/gate): a wrong `at` fails closed, it does not silently clash.
  if (x.family === 'structure' && y.family === 'structure') return {};
  const a = x.selector;
  const b = y.selector;
  const usesHit = (u = [], c = []) => c.find(t => u.includes(t) || (u.includes('name:*') && t.startsWith('name:')));
  const claimsClash = (a.claims || []).find(t => (b.claims || []).includes(t)) || usesHit(a.uses, b.claims) || usesHit(b.uses, a.claims);
  const overlap = a.sheet === b.sheet && !a.cellFree && !b.cellFree && overlaps(parseRange(a.range), parseRange(b.range));
  const declared = !!x.footprint && !!y.footprint;
  const shared = a.sheet === b.sheet && declared ? sharedAreas(x.footprint, y.footprint) : [];
  if (shared.length) return { clash: `${a.sheet} (same axis: ${shared.join(', ')})` };
  if (overlap && !declared) return { clash: `${a.sheet} (axes not declared)` };
  return claimsClash ? { clash: claimsClash } : { overlap };
}

/** The axes on which two footprints MEET: a cell area overlaps, or a [from, to] / '*' span intersects. */
function sharedAreas(fx, fy) {
  const spans = (p, q) => p === '*' || q === '*' || (p[0] <= q[1] && q[0] <= p[1]);
  const meet = (axis, p, q) => (axis === 'value' || axis === 'style' ? overlaps(parseRange(p), parseRange(q)) : spans(p, q));
  return Object.keys(fx).filter(axis => axis in fy && meet(axis, fx[axis], fy[axis]));
}

/**
 * What the live read must cover for this task, for the caller's `observation.live` option: the
 * target windows (a write may land outside every used range). The used ranges themselves are always
 * read by the driver. Invalid tasks give no windows; planning refuses them anyway.
 */
// M7 S32 slice 2 (owner GO 2026-10-03): copy_range and the two name intents, on the T3 formula index; move_range is NOT
// here (its Cut re-points formulas outside the windows, which the digest cannot compare)
// X-L13/3 (matrix category (a) MIGRALHATO -- no liveExtras of its own, a plain own-window request): merge_range /
// unmerge_range. The digest (T2) carries no merge-area field at all, so a cell's merge area OUTSIDE the task's windows
// is never re-proven from a tiered observation -- merge.cjs's own verifyPreservation() states this explicitly instead
// of silently comparing nothing (see its own comment); the invariant it relies on is the pre-write gate's own refusal
// (xlsx-merge-area-reaches-outside-the-target), which already keeps a target cell's merge area from reaching outside it.
// X-L13/5 (matrix category (a) FULL -- fdb66c26's code-level survey): sort_range / clear_filter / set_print_area /
// clear_print_area have NO extra axis of their own, or (print area) share the ALREADY-migrated set_defined_name's
// exact mechanism (dependents, workbook-wide, unbounded by MAX_READ_CELLS) -- the digest/dependents proof is the
// whole of their FULL-tier preservation, nothing is reduced. filter_range (needs `layout`, not in TIERED_EXTRAS)
// and format_run/freeze/CF/validation (need extras outside TIERED_EXTRAS too) stay OUT of this list -- separate
// card, an eligibility gap, not a preservation gap (see docs/X-L13-LARGE-BOOK-MATRIX-20261013.md).
// X-L13/6: format_run (runCells), freeze_panes/unfreeze_panes (panes+sheetVisibility), add_/delete_conditional_format
// (cfRanges), set_/clear_validation (validationRanges) -- each extra is now in TIERED_EXTRAS (xlsx-driver.cjs),
// request-sized, not a used-range scan. filter_range stays OUT: its `layout` extra's cost grows with the sheet's
// own used range (not in TIERED_EXTRAS) -- a separate, unresolved eligibility gap, not migrated here.
const TIERED_INTENTS = Object.freeze(['write_range', 'clear_range', 'copy_range', 'set_defined_name', 'delete_defined_name', 'merge_range', 'unmerge_range', 'sort_range', 'clear_filter', 'set_print_area', 'clear_print_area', 'format_run', 'freeze_panes', 'unfreeze_panes', 'add_conditional_format', 'delete_conditional_format', 'set_validation', 'clear_validation']);

// GH #107 (b): the per-step extras union, shared by liveScope (the request) and boundedVariant (the merged variant of a
// mixed tiered + bounded task) so the two cannot drift apart.
const liveExtrasOf = op => (familyFor(op).liveExtras ? familyFor(op).liveExtras(op) : {});
const extraKey = (ops, key) => [...new Set(ops.flatMap(op => liveExtrasOf(op)[key] || []))];
const wantsDependents = ops => ops.some(op => !!liveExtrasOf(op).dependents);
// GH #96 follow-up: the T3 formula index rides with the dependency readers unless EVERY dependents step opts out
const skipsFormulaIndex = ops => wantsDependents(ops) && ops.every(op => !liveExtrasOf(op).dependents || liveExtrasOf(op).formulaIndex === false);

function liveScope(task) {
  const planned = planOperations(task);
  if (!planned.ok) return { windows: [] };
  const ops = planned.operations;
  const extra = key => extraKey(ops, key);
  const dependents = wantsDependents(ops);
  const noFormulaIndex = skipsFormulaIndex(ops);
  const layoutExtents = extra('layout');
  const styleSheets = extra('styleSheets');
  const validationRanges = extra('validationRanges');
  const cfRanges = extra('cfRanges');
  const runCells = extra('runCells');
  const arrayCells = extra('arrayCells');
  const commentSheets = extra('commentSheets');
  const sessionUser = ops.some(op => !!liveExtrasOf(op).sessionUser);
  const filterSheets = extra('filterSheets');
  const chartSheets = extra('chartSheets');
  const pageSheets = extra('pageSheets');
  const printSheets = extra('printSheets');
  const drawingSheets = extra('drawingSheets');
  const colBreaks = ops.some(op => !!(familyFor(op).liveExtras && familyFor(op).liveExtras(op).colBreaks));
  const drawingAnchor = ops.some(op => !!(familyFor(op).liveExtras && familyFor(op).liveExtras(op).drawingAnchor));
  // M7 S32: only a task of TIERED_INTENTS steps may be observed in tiers above 45000 used cells
  const tiered = ops.length > 0 && ops.every(op => TIERED_INTENTS.includes(op.intent));
  const panes = ops.some(op => !!(familyFor(op).liveExtras && familyFor(op).liveExtras(op).panes));
  const sheetVisibility = ops.some(op => !!(familyFor(op).liveExtras && familyFor(op).liveExtras(op).sheetVisibility));
  const bounded = tiered ? null : boundedVariant(ops);
  // GH #107 (a): a multi-step task with neither path names its blockers in the refusal (the driver's refusalOf)
  const mixedTask = !tiered && !bounded && ops.length > 1 ? mixedTaskNote(ops) : null;
  // GH #326 (U4b): the protection gate's own targets, for EVERY op (not family liveExtras -- it is a
  // cross-cutting concern, not a per-family one) -- a range-level target (write_range/clear_range/a copy|move
  // target) through protectionRanges (bounded per-request, same TIERED_EXTRAS group as cfRanges/validationRanges);
  // every other op with a real sheet target through protectionSheets (U4a's cheaper, no-cell-walk axis).
  const protectionTargets = protectionTargetsOf(ops);
  const protectionRanges = protectionTargets.filter(t => t.range !== undefined);
  const protectionSheets = [...new Set(protectionTargets.filter(t => t.range === undefined).map(t => t.sheet))];
  return { windows: ops.flatMap(op => familyFor(op).windowsOf(op)), formatSheets: extra('formatSheets'), objectSheets: extra('objectSheets'), ...(dependents ? { dependents: true } : {}), ...(noFormulaIndex ? { noFormulaIndex: true } : {}), ...(layoutExtents.length ? { layout: layoutExtents } : {}), ...(styleSheets.length ? { styleSheets } : {}), ...(panes ? { panes: true } : {}), ...(sheetVisibility ? { sheetVisibility: true } : {}), ...(validationRanges.length ? { validationRanges } : {}), ...(cfRanges.length ? { cfRanges } : {}), ...(runCells.length ? { runCells } : {}), ...(arrayCells.length ? { arrayCells } : {}), ...(commentSheets.length ? { commentSheets } : {}), ...(sessionUser ? { sessionUser: true } : {}), ...(filterSheets.length ? { filterSheets } : {}), ...(chartSheets.length ? { chartSheets } : {}), ...(pageSheets.length ? { pageSheets } : {}), ...(printSheets.length ? { printSheets } : {}), ...(drawingSheets.length ? { drawingSheets } : {}), ...(colBreaks ? { colBreaks: true } : {}), ...(drawingAnchor ? { drawingAnchor: true } : {}), ...(protectionRanges.length ? { protectionRanges } : {}), ...(protectionSheets.length ? { protectionSheets } : {}), ...(tiered ? { tiered: true } : {}), ...(bounded ? { boundedVariant: bounded } : {}), ...(mixedTask ? { mixedTask } : {}) };
}

/** The protection gate's own (sheet[, range]) read targets, deduped by sheet+range -- shared by liveScope and
 * boundedVariant so the two cannot drift apart (the SAME reason extraKey/liveExtrasOf is shared above). */
function protectionTargetsOf(ops) {
  const seen = new Set();
  const out = [];
  for (const t of ops.flatMap(protectionGate.targetsOf)) {
    const k = `${t.sheet}\u0000${t.range || ''}`;
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(t);
  }
  return out;
}

/**
 * GH #107 (a): the tiered observation is TASK-level -- one step outside TIERED_INTENTS makes the whole task non-tiered,
 * and the bounded variant needs a variant for every step. This names the steps that blocked each path, and whether every
 * step alone has one of the two (then one task per family is the way through). A message only: nothing is tiered here.
 */
function mixedTaskNote(ops) {
  const label = op => `${op.id} (${op.intent})`;
  // GH #107 (b): a step has a path when it is tiered or has a bounded variant of its own (a mix of the two is observed together)
  const hasBounded = op => TIERED_INTENTS.includes(op.intent) || !!(familyFor(op).boundedExtras && familyFor(op).boundedExtras(op));
  return {
    notTiered: ops.filter(op => !TIERED_INTENTS.includes(op.intent)).map(label),
    noBoundedVariant: ops.filter(op => !hasBounded(op)).map(label),
    splittable: ops.every(op => TIERED_INTENTS.includes(op.intent) || hasBounded(op)),
  };
}

/**
 * X-N01b: the reduced request a task may be observed with ONLY above the read bound (the driver decides; below it the
 * full request runs unchanged). The windows stay as they are (a format step's target + ring keeps its style read), the
 * sheet-wide style read is replaced by the digest's number formats. The observation it yields says so (tier 'tiered',
 * bounded) -- never a full proof.
 * X-L13/8 (the (b)-group BOUNDED classification, owner decision): `layout` is merged through the same way as
 * `formatSheets` -- `layout_range`'s own `boundedExtras()` needs it carried into the bounded request (its
 * gate()/verifyPrimitive() read the op's own target through it; dropping it would break L1, not just demote
 * L3 -- see layout.cjs's boundedExtras comment).
 * GH #107 (b) (design docs/GH107B-MIXED-TIERED-BOUNDED-DESIGN-20261014.md): a TIERED_INTENTS step used to need a
 * bounded variant of its own (write_range only), so format_range + set_validation had no path. Now a tiered step
 * contributes its own liveExtras (every key of which is in the driver's TIERED_EXTRAS -- pinned by a test), a
 * non-tiered step its boundedExtras; the variant is the union. Eligible when every step is tiered or has boundedExtras.
 * Each step keeps reporting the scope its own family reports (the tiered families read `tier`, only format_range /
 * layout_range read `bounded`), so no step is granted more than in its single-family task, and none loses an extra.
 */
function boundedVariant(ops) {
  if (!ops.length) return null;
  const tieredOps = ops.filter(op => TIERED_INTENTS.includes(op.intent));
  const own = ops.filter(op => !TIERED_INTENTS.includes(op.intent)).map(op => (familyFor(op).boundedExtras ? familyFor(op).boundedExtras(op) : null));
  if (own.some(x => x === null)) return null;
  const merge = key => [...new Set([...extraKey(tieredOps, key), ...own.flatMap(x => x[key] || [])])];
  const arrays = Object.fromEntries(['formatSheets', 'layout', 'validationRanges', 'cfRanges', 'runCells', 'filterSheets'].map(k => [k, merge(k)]).filter(([, v]) => v.length));
  const dependents = wantsDependents(tieredOps);
  const noFormulaIndex = skipsFormulaIndex(tieredOps);
  const flag = key => tieredOps.some(op => !!liveExtrasOf(op)[key]);
  // GH #326 (U4b): the protection gate's targets, over EVERY op (it is cross-cutting, not tiered-vs-bounded).
  const protectionTargets = protectionTargetsOf(ops);
  const protectionRanges = protectionTargets.filter(t => t.range !== undefined);
  const protectionSheets = [...new Set(protectionTargets.filter(t => t.range === undefined).map(t => t.sheet))];
  return {
    windows: ops.flatMap(op => familyFor(op).windowsOf(op)),
    formatSheets: arrays.formatSheets || [], tiered: true, windowStyle: true,
    ...Object.fromEntries(Object.entries(arrays).filter(([k]) => k !== 'formatSheets')),
    ...(dependents ? { dependents: true } : {}), ...(noFormulaIndex ? { noFormulaIndex: true } : {}),
    ...(flag('panes') ? { panes: true } : {}), ...(flag('sheetVisibility') ? { sheetVisibility: true } : {}),
    ...(protectionRanges.length ? { protectionRanges } : {}), ...(protectionSheets.length ? { protectionSheets } : {}),
  };
}

const xlsxAdapter = {
  kind: 'xlsx',

  async readinessProbe() {
    return { ok: true, detail: 'transport-level readiness: the spreadsheet model answered at open' };
  },

  normalizeInventory,

  enrichObservation: null,

  resolveTarget(primary, selector, ctx = {}) {
    return familyFor(ctx.operation).resolve(primary, selector, ctx);
  },

  planOperations,

  // L-19 C: in a batched delete_rows task a later step's anchor can
  // coincidentally already sit at its OWN terminal position while an earlier step of the SAME task
  // still needed a fresh APPLY -- e.g. step 1 at=10 count=5 anchor R15, step 2 given in TASK-START
  // coordinates at=17 count=5 anchor R22: after step 1, R22 (unmoved by step 1, since it was always
  // past the cut) now sits at row 17, which IS step 2's terminal -- SATISFIED, 0 writes, and rows
  // 17-21 are never deleted. SATISFIED operations of one run must form a PREFIX (a retry can only
  // ever find an EARLIER-run prefix already done); a SATISFIED verdict arriving after this SAME
  // run already had to APPLY an earlier step breaks that prefix and is refused, named, instead.
  // GH #326 (U4b): the sheet-protection gate runs FIRST, before any family-specific classification -- a
  // protected sheet's write is refused on the gate's own evidence, never reaching the family's own
  // SATISFIED/APPLY logic. `gate()`'s raceChecks (empty when nothing is restricted) ride on `op` itself
  // (mutated here, consumed by executePrimitive below) -- the SAME op-mutation convention range.cjs's own
  // classify() already uses for move_range's `op.expect` (built fresh, right before dispatch, the SAME `p`).
  classifyOperation(primary, op, context) {
    const protection = protectionGate.gate(primary, op);
    if (protection.refusal) return { state: 'FAIL_CLOSED', evidence: protection.refusal };
    if (protection.raceChecks.length) op._protectionRaceCheck = Object.freeze(protection.raceChecks.map(c => Object.freeze(c)));
    const out = familyFor(op).classify(primary, op);
    if (out.state === 'SATISFIED' && op.family === 'structure' && context && context.priorAppliedThisRun) {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-structure-satisfied-after-applied-in-batch', wouldHaveBeen: out.evidence } };
    }
    return out;
  },

  /**
   * X-M01 (contract: an unsupported shape is refused pre-dispatch / pre-mutation): is a refusal of a write_range /
   * unmerge_range / set_chart_title (the product's multistep intents) measured on the PLANNING observation already final,
   * i.e. can no EARLIER step of the task change it? True only when every earlier step is one that cannot create or move a
   * merge, a formula cell, a sheet or a chart: write_range / format_range / set_chart_title, or an unmerge_range that does
   * not intersect the refused cell's merged area. Anything else -> false (the step is classified at its turn, as before).
   *
   * GitHub #79: a merge_range / sort_range earlier step is ALSO safe for these same two refusal reasons
   * (xlsx-write-into-merged-area-refused / xlsx-write-over-formula-refused) when its own range does not reach the
   * refused region (the merge area, or the single refused cell) -- neither family can create or move a formula, and
   * a merge_range that does not touch the area cannot be the one that produced it. Before this, ANY earlier
   * merge_range/sort_range made the whole check non-decisive, so a task whose later step was certain to be refused
   * dispatched its earlier merge/sort steps anyway, leaving the document in an intermediate state. For every OTHER
   * refusal reason, merge_range / sort_range still disqualify (unchanged) -- only the two reasons named above widen.
   */
  preflightDecisive(earlier, op, evidence) {
    // GH #326 (U4b): a protection-gate refusal is decisive REGARDLESS of `earlier` -- no intent in this
    // product sets or clears sheet protection (grep: nothing under families/ touches sheetProtection), so no
    // earlier step of the SAME task can ever cure it. The whole task is refused before the first dispatch.
    if (evidence && PROTECTION_REFUSAL.test(evidence.reason || '')) return true;
    if (!op || !['write_range', 'unmerge_range', 'set_chart_title'].includes(op.intent)) return false;
    const area = evidence && evidence.reason === 'xlsx-write-into-merged-area-refused' && typeof evidence.mergeArea === 'string' ? parseRange(evidence.mergeArea) : null;
    const refusedCell = evidence && evidence.reason === 'xlsx-write-over-formula-refused' && typeof evidence.cell === 'string' ? parseRange(evidence.cell) : null;
    const region = area || refusedCell;
    return earlier.every(e => {
      if (['write_range', 'format_range', 'set_chart_title'].includes(e.intent)) return true;
      if (e.intent === 'unmerge_range') {
        if (!area) return true; // the refusal is not a merge refusal: an unmerge cannot cure it
        if (e.selector.sheet !== op.selector.sheet) return true;
        const g = parseRange(e.selector.range);
        return !!g && !overlaps(g, area);
      }
      if (['merge_range', 'sort_range'].includes(e.intent) && region) {
        if (e.selector.sheet !== op.selector.sheet) return true;
        const g = parseRange(e.selector.range);
        return !!g && !overlaps(g, region);
      }
      return false;
    });
  },

  classifyIdempotency(primary, task) {
    const planned = planOperations(task);
    if (!planned.ok) return { state: 'FAIL_CLOSED', evidence: { reason: planned.detail } };
    const states = planned.operations.map(op => familyFor(op).classify(primary, op).state);
    if (states.includes('FAIL_CLOSED')) return { state: 'FAIL_CLOSED', evidence: { states } };
    return { state: states.every(s => s === 'SATISFIED') ? 'SATISFIED' : 'APPLY', evidence: { states } };
  },

  // GH #326 (U4b), the race condition: classifyOperation's gate ran against the PLANNING observation; by the
  // time this dispatches, the document may have moved. `op._protectionRaceCheck` (empty/absent when nothing
  // was restricted) rides into the command's own `meta`, for the driver's dispatch() to re-verify, unchanged,
  // in the SAME editor session, immediately before the first setter of `command.body` runs.
  async executePrimitive(_handle, op, { dispatch, target }) {
    const command = familyFor(op).lower(op, target);
    if (op._protectionRaceCheck) command.meta = Object.freeze({ ...command.meta, protectionRaceCheck: op._protectionRaceCheck });
    return dispatch(command);
  },

  async verifyPrimitive(_handle, op, before, after) {
    return familyFor(op).verifyPrimitive(op, before, after);
  },

  async verifyPreservation(_handle, op, before, after, context = {}) {
    // GH #99: batch-round siblings exist only for planner-proven write_range groups (BATCH_KEY below); any other
    // family never receives one, so a sibling reaching a family that does not exclude it would be a planner bug
    const siblings = (context && context.siblings) || [];
    if (siblings.length && !siblings.every(o => o.batchSafeGroup === BATCH_KEY) ) return { ok: false, evidence: { reason: 'xlsx-preservation-sibling-not-batch-proven' } };
    return familyFor(op).verifyPreservation(op, before, after, { siblings });
  },

  async verifyWholeTask(_handle, task, finalObservation) {
    const planned = planOperations(task);
    if (!planned.ok) return { ok: false, checks: [{ name: 'task plans', ok: false, detail: planned.detail }] };
    const checks = planned.operations.flatMap(op => familyFor(op).wholeTaskChecks(op, finalObservation));
    return { ok: checks.length > 0 && checks.every(c => c.ok === true), checks };
  },

  terminalContract: null,
};

module.exports = { xlsxAdapter, planOperations, liveScope };
