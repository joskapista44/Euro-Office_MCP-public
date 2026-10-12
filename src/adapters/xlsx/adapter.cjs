'use strict';

const { normalizeInventory } = require('./inventory.cjs');
const { parseRange, overlaps } = require('./a1.cjs');
const cells = require('./families/cells.cjs');
const clear = require('./families/clear.cjs');
const sheets = require('./families/sheets.cjs');
const merge = require('./families/merge.cjs');
const layout = require('./families/layout.cjs');
const format = require('./families/format.cjs');
const freeze = require('./families/freeze.cjs');
const structure = require('./families/structure.cjs');
const range = require('./families/range.cjs');
const names = require('./families/names.cjs');
const sort = require('./families/sort.cjs');
const validation = require('./families/validation.cjs');
const cf = require('./families/cf.cjs');
const filter = require('./families/filter.cjs');
const charts = require('./families/charts.cjs');
const page = require('./families/page.cjs');
const print = require('./families/print.cjs');
const pivots = require('./families/pivots.cjs');
const pivotRefresh = require('./families/pivot-refresh.cjs');

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
 * rename_chart: the full identity derived from the data range, public getters -- wave 7), `page` (set_page_layout:
 * orientation / gridlines / headings / margins in mm -- wave 8), `print` (fit + flag, odd header, row break, first page
 * number, horizontal centering on the sheet model, build-gated -- wave 8).
 * Every other intent is refused at
 * PLANNING, before any dispatch. The inventory of what remains is docs/evidence/M6-XLSX-CAPABILITY-MATRIX.md.
 */

// write_range's own L3 adds the observed style axis (M7 S3); cells.verifyPreservation stays the content L3 every other
// family builds on
const write = Object.freeze({ ...cells, verifyPreservation: cells.writePreservation });

const FAMILIES = Object.freeze({ write_range: write, clear_range: clear, ...Object.fromEntries(sheets.INTENTS.map(i => [i, sheets])), ...Object.fromEntries(merge.INTENTS.map(i => [i, merge])), layout_range: layout, format_range: format, ...Object.fromEntries(freeze.INTENTS.map(i => [i, freeze])), ...Object.fromEntries(structure.INTENTS.map(i => [i, structure])), ...Object.fromEntries(range.INTENTS.map(i => [i, range])), ...Object.fromEntries(names.INTENTS.map(i => [i, names])), sort_range: sort, ...Object.fromEntries(validation.INTENTS.map(i => [i, validation])), ...Object.fromEntries(cf.INTENTS.map(i => [i, cf])), ...Object.fromEntries(filter.INTENTS.map(i => [i, filter])), ...Object.fromEntries(charts.INTENTS.map(i => [i, charts])), set_page_layout: page, ...Object.fromEntries(print.INTENTS.map(i => [i, print])), ...Object.fromEntries(pivots.INTENTS.map(i => [i, pivots])), ...Object.fromEntries(pivotRefresh.INTENTS.map(i => [i, pivotRefresh])) });

function familyOf(intent) {
  const f = FAMILIES[intent];
  return f ? { ok: true, family: f } : { ok: false, detail: `intent-not-migrated-in-m6: ${intent}` };
}

function familyFor(op) {
  const f = familyOf(op && op.intent);
  if (!f.ok) throw new Error(f.detail);
  return f.family;
}

const TASK_FIELDS = Object.freeze(['id', 'steps']);

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
  // a structural step shifts every coordinate the other steps would name
  if (operations.length > 1 && operations.some(o => o.family === 'structure')) {
    return { ok: false, operations: [], detail: 'an insert / delete of rows or columns must be the only step of its task' };
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
  return { ok: true, operations, dependencies: [] };
}

/**
 * Two planned steps: { clash } = the refusal text (a cell overlap on a shared or undeclared axis, a claimed object, a
 * used name); { overlap } = their cells overlap on disjoint axes (the later one depends on the earlier).
 */
function pairOf(x, y) {
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
const TIERED_INTENTS = Object.freeze(['write_range', 'clear_range', 'copy_range', 'set_defined_name', 'delete_defined_name']);

function liveScope(task) {
  const planned = planOperations(task);
  if (!planned.ok) return { windows: [] };
  const ops = planned.operations;
  const extra = key => [...new Set(ops.flatMap(op => (familyFor(op).liveExtras ? familyFor(op).liveExtras(op)[key] || [] : [])))];
  const dependents = ops.some(op => !!(familyFor(op).liveExtras && familyFor(op).liveExtras(op).dependents));
  const layoutExtents = extra('layout');
  const styleSheets = extra('styleSheets');
  const validationRanges = extra('validationRanges');
  const cfRanges = extra('cfRanges');
  const filterSheets = extra('filterSheets');
  const chartSheets = extra('chartSheets');
  const pageSheets = extra('pageSheets');
  const printSheets = extra('printSheets');
  const drawingSheets = extra('drawingSheets');
  const colBreaks = ops.some(op => !!(familyFor(op).liveExtras && familyFor(op).liveExtras(op).colBreaks));
  const drawingAnchor = ops.some(op => !!(familyFor(op).liveExtras && familyFor(op).liveExtras(op).drawingAnchor));
  // M7 S32: only a task of TIERED_INTENTS steps may be observed in tiers above 5000 used cells
  const tiered = ops.length > 0 && ops.every(op => TIERED_INTENTS.includes(op.intent));
  const panes = ops.some(op => !!(familyFor(op).liveExtras && familyFor(op).liveExtras(op).panes));
  return { windows: ops.flatMap(op => familyFor(op).windowsOf(op)), formatSheets: extra('formatSheets'), objectSheets: extra('objectSheets'), ...(dependents ? { dependents: true } : {}), ...(layoutExtents.length ? { layout: layoutExtents } : {}), ...(styleSheets.length ? { styleSheets } : {}), ...(panes ? { panes: true } : {}), ...(validationRanges.length ? { validationRanges } : {}), ...(cfRanges.length ? { cfRanges } : {}), ...(filterSheets.length ? { filterSheets } : {}), ...(chartSheets.length ? { chartSheets } : {}), ...(pageSheets.length ? { pageSheets } : {}), ...(printSheets.length ? { printSheets } : {}), ...(drawingSheets.length ? { drawingSheets } : {}), ...(colBreaks ? { colBreaks: true } : {}), ...(drawingAnchor ? { drawingAnchor: true } : {}), ...(tiered ? { tiered: true } : {}) };
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

  classifyOperation(primary, op) {
    return familyFor(op).classify(primary, op);
  },

  classifyIdempotency(primary, task) {
    const planned = planOperations(task);
    if (!planned.ok) return { state: 'FAIL_CLOSED', evidence: { reason: planned.detail } };
    const states = planned.operations.map(op => familyFor(op).classify(primary, op).state);
    if (states.includes('FAIL_CLOSED')) return { state: 'FAIL_CLOSED', evidence: { states } };
    return { state: states.every(s => s === 'SATISFIED') ? 'SATISFIED' : 'APPLY', evidence: { states } };
  },

  async executePrimitive(_handle, op, { dispatch, target }) {
    return dispatch(familyFor(op).lower(op, target));
  },

  async verifyPrimitive(_handle, op, before, after) {
    return familyFor(op).verifyPrimitive(op, before, after);
  },

  async verifyPreservation(_handle, op, before, after) {
    return familyFor(op).verifyPreservation(op, before, after);
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
