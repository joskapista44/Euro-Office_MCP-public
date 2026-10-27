'use strict';

const { parseRange, rangeText, addr } = require('../a1.cjs');
const { cellAt, isBlank } = require('../inventory.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `filter` family: filter_range (one field, one value, xlOr -- the measured shape, B-D06) / clear_filter (B-D13)
 * (M6 wave 6).
 *
 * The reader: the PUBLIC AutoFilter getters plus the field identity through the DS Bridge
 * (sheet.autoFilter.fieldIdentity.read/v1, B07 cutover; the W6 / W6b probes proved the 0-based ColId the bridge maps;
 * a missing bridge or another build is refused before the write; an unobservable filter is never "no filter").
 * Refused before the write: an AutoFilter on ANOTHER range; an active filter on another field (B-D06's guards).
 * filter: SATISFIED when the AutoFilter is on exactly the range and the field's active filter is exactly the requested
 * one. clear: SATISFIED when there is no AutoFilter or none of its filters is on (ShowAllData keeps the controls,
 * B-D13/D14). L1 the filter state; L3 every cell. Which rows are hidden is NOT compared (B-D06: not measured).
 *
 * M7 R19 (tier D probe, m7-d-probe1-20261003.json FL, session = reopen): besides the one value, a field may take
 * - `values: [a, b]` -- exactly two plain values, SetAutoFilter(field, [a, b], 'xlFilterValues'); reads back xlOr
 *   `=a` / `=b` (a longer list is NOT measured);
 * - `and: ['>n', '<m']` -- two numeric comparisons joined by xlAnd, SetAutoFilter(field, c1, 'xlAnd', c2); reads back
 *   exactly (only `>` and `<` were measured).
 * Several fields on ONE AutoFilter range: the probe set a second field while the first was active and both persisted,
 * so another active field no longer refuses; L1 proves every OTHER field's filter exactly as before.
 * A one-value filter now also requires an empty criteria2 (W6 read it back null) -- otherwise a two-value list holding
 * the value would read as that one-value filter.
 *
 * GH #301 D2 (AF-02 LIVE on 9.3.5, session = reopen = stored), three more criteria, each read back as noted:
 * - `values: [3..10 different plain values]` -- SetAutoFilter(field, [..], 'xlFilterValues'); reads back operator
 *   xlFilterValues with criteria1 the ARRAY (compared as a set); the two-value shape above is unchanged;
 * - `top: n` (1..500) -- SetAutoFilter(field, n, 'xlTop10Items'); reads back xlTop10Items, criteria1 n;
 * - `dynamic: 'aboveAverage' | 'belowAverage'` -- SetAutoFilter(field, 'xlFilter<Above|Below>Average',
 *   'xlFilterDynamic'); reads back xlFilterDynamic, criteria1 that constant.
 * The cell-colour filter is NOT a writer shape: GetCriteria1 reads it back null, so no write of it could be verified.
 *
 * X-L05 (S26 probe m7-s26-probe1, LIVE MEASURED): applying a filter_range with NO active AutoFilter yet
 * UNHIDES every data row of the range the engine judges matching -- INCLUDING a row a human hid by hand before the
 * filter ever ran (the probe's own risk case: a manually-hidden, matching row came back visible the instant
 * SetAutoFilter ran). filter_range is therefore refused BEFORE the write whenever the range has NO active filter AND
 * any data row (below the header) already reads hidden -- that hidden state can only be manual, and applying the
 * filter would silently erase it. An unreadable hidden state on any data row refuses the SAME way, fail-closed (we
 * cannot prove it is safe to apply). A row OUTSIDE the range is untouched by this check, and so is a range that
 * ALREADY carries an active filter on it (there, a hidden row may be filter's own doing -- R19's multi-field path is
 * unchanged). clear_filter is NOT touched by this gate.
 */

const INTENTS = Object.freeze(['filter_range', 'clear_filter']);

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const set = step.intent === 'filter_range';
  const fields = set ? ['intent', 'id', 'sheet', 'range', 'field', 'value', 'values', 'and', 'top', 'dynamic'] : ['intent', 'id', 'sheet', 'range'];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g || g.r2 - g.r1 < 1) return { ok: false, detail: `${at} requires a plain A1 range with a header row and data` };
  if (set) {
    if (!Number.isInteger(step.field) || step.field < 1 || step.field > g.c2 - g.c1 + 1) return { ok: false, detail: `${at}: field must be a column number inside the range (1-based)` };
    const problem = criteriaProblem(step);
    if (problem) return { ok: false, detail: `${at}: ${problem}` };
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'filter',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      field: set ? step.field : null,
      value: set && 'value' in step ? String(step.value) : null,
      values: set && 'values' in step ? Object.freeze(step.values.map(String)) : null,
      and: set && 'and' in step ? Object.freeze([...step.and]) : null,
      top: set && 'top' in step ? step.top : null,
      dynamic: set && 'dynamic' in step ? step.dynamic : null,
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const plainValue = v => (typeof v === 'string' && v !== '' && !/^[=<>]/.test(v)) || Number.isFinite(v);
const COMPARE = /^[<>]-?\d+(\.\d+)?$/;

const MAX_VALUES = 10;
const MAX_TOP = 500;
const DYNAMIC = Object.freeze({ aboveAverage: 'xlFilterAboveAverage', belowAverage: 'xlFilterBelowAverage' });

function criteriaProblem(step) {
  const given = ['value', 'values', 'and', 'top', 'dynamic'].filter(k => k in step);
  if (given.length !== 1) return 'give exactly one of value / values / and / top / dynamic';
  if (given[0] === 'value') return plainValue(step.value) ? null : 'value must be one plain value (no operator)';
  if (given[0] === 'values') {
    const ok = Array.isArray(step.values) && step.values.length >= 2 && step.values.length <= MAX_VALUES && step.values.every(plainValue) && new Set(step.values.map(String)).size === step.values.length;
    return ok ? null : `values must be 2 to ${MAX_VALUES} different plain values`;
  }
  if (given[0] === 'top') return Number.isInteger(step.top) && step.top >= 1 && step.top <= MAX_TOP ? null : `top must be a whole number of items, 1 to ${MAX_TOP}`;
  if (given[0] === 'dynamic') return Object.hasOwn(DYNAMIC, step.dynamic) ? null : `dynamic must be one of ${Object.keys(DYNAMIC).join(', ')} (other dynamic filters are not measured)`;
  const ok = Array.isArray(step.and) && step.and.length === 2 && step.and.every(c => typeof c === 'string' && COMPARE.test(c));
  return ok ? null : "and must be two numeric comparisons, each '>n' or '<n' (other operators are not measured)";
}

const bare = x => String(x === null || x === undefined ? '' : x).replace(/^=/, '');
const active = st => st.filters.filter(f => f.on !== false);

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const st = p.filterReads ? p.filterReads[op.selector.sheet] : undefined;
  if (!st || typeof st !== 'object') return { ok: false, reason: 'xlsx-filter-unobservable', why: st === undefined ? 'not-read' : st };
  if (st.range !== null && String(st.range).replace(/\$/g, '') !== op.selector.range) return { ok: false, reason: 'xlsx-filter-range-conflict', autoFilter: st.range };
  const region = op.intent === 'filter_range' && st.range === null ? adjacentData(p, op) : null;
  if (region) return { ok: false, ...region };
  const manual = op.intent === 'filter_range' && !active(st).length ? manualHiddenRow(p, op) : null;
  return manual ? { ok: false, ...manual } : { ok: true, st };
}

/**
 * LIVE (release-l11-3 harness, prod): with no AutoFilter on the sheet, SetAutoFilter on a range that has data right
 * next to it applies the filter to the whole current region (A1:B5 asked -> A1:B1877 applied) -- rows outside the
 * request get filtered / hidden, and the X-L05 check above never saw them. A non-blank cell in the ring around the
 * range (the rows above / below, the columns on either side, the corners) refuses before the write. An absent cell
 * is outside the used range: empty.
 */
function adjacentData(p, op) {
  const g = parseRange(op.selector.range);
  const filled = (c, r) => { if (c < 1 || r < 1) return false; const cell = cellAt(p, op.selector.sheet, addr(c, r)); return !!cell && !isBlank(cell); };
  const ring = [];
  for (let c = g.c1 - 1; c <= g.c2 + 1; c += 1) ring.push([c, g.r1 - 1], [c, g.r2 + 1]);
  for (let r = g.r1; r <= g.r2; r += 1) ring.push([g.c1 - 1, r], [g.c2 + 1, r]);
  const hits = ring.filter(([c, r]) => filled(c, r)).map(([c, r]) => addr(c, r));
  return hits.length ? { reason: 'xlsx-filter-range-not-the-current-region', cells: hits.slice(0, 20), count: hits.length } : null;
}

/** X-L05: no ACTIVE filter on the range (none, or a cleared one that kept its range) -> an already-hidden data row is
 * manual, and applying the filter would unhide it (measured, m7-s26-probe1). Fail-closed on an unreadable row. */
function manualHiddenRow(p, op) {
  const g = parseRange(op.selector.range);
  const axis = p.layout && p.layout[op.selector.sheet];
  if (!axis || typeof axis !== 'object') return { reason: 'xlsx-filter-rows-unobservable', detail: axis || 'not-read' };
  for (let r = g.r1 + 1; r <= g.r2; r += 1) {
    const e = axis.rows[String(r)];
    if (!e || typeof e.hidden !== 'boolean') return { reason: 'xlsx-filter-rows-unobservable', row: r };
    if (e.hidden === true) return { reason: 'xlsx-filter-would-unhide-manual-rows', row: r };
  }
  return null;
}

const empty = x => x === null || x === undefined || x === '';
const sameSet = (a, b) => a.length === b.length && [...a].sort().join('\u0000') === [...b].sort().join('\u0000');

/** The requested filter of op.field, as the reader shows it. */
function matches(f, op) {
  if (op.values && op.values.length > 2) return f.operator === 'xlFilterValues' && Array.isArray(f.criteria1) && empty(f.criteria2) && sameSet(f.criteria1.map(String), op.values);
  if (op.values) return f.operator === 'xlOr' && bare(f.criteria1) === op.values[0] && bare(f.criteria2) === op.values[1];
  if (op.top !== null) return f.operator === 'xlTop10Items' && Number(f.criteria1) === op.top && empty(f.criteria2);
  if (op.dynamic !== null) return f.operator === 'xlFilterDynamic' && f.criteria1 === DYNAMIC[op.dynamic] && empty(f.criteria2);
  if (op.and) return f.operator === 'xlAnd' && String(f.criteria1) === op.and[0] && String(f.criteria2) === op.and[1];
  return bare(f.criteria1) === op.value && empty(f.criteria2);
}

const mine = (st, op) => active(st).filter(f => f.field === op.field);
const holds = (st, op) => st.range !== null && mine(st, op).length === 1 && matches(mine(st, op)[0], op);
const others = (st, op) => (st.range === null ? [] : st.filters.filter(f => f.field !== op.field));

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, range: op.selector.range };
  if (op.intent === 'clear_filter') return active(g.st).length === 0 ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-no-active-filter' } } : { state: 'APPLY', evidence: where };
  if (holds(g.st, op)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-the-filter' } };
  return { state: 'APPLY', evidence: { ...where, field: op.field, value: op.value, values: op.values, and: op.and, top: op.top, dynamic: op.dynamic, otherFields: active(g.st).filter(f => f.field !== op.field).map(f => f.field) } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range }) };
}

function setArgs(op) {
  if (op.values) return `${JSON.stringify(op.values)}, 'xlFilterValues'`;
  if (op.and) return `${JSON.stringify(op.and[0])}, 'xlAnd', ${JSON.stringify(op.and[1])}`;
  if (op.top !== null) return `${op.top}, 'xlTop10Items'`;
  if (op.dynamic !== null) return `${JSON.stringify(DYNAMIC[op.dynamic]).replace(/"/g, "'")}, 'xlFilterDynamic'`;
  return `${JSON.stringify(op.value)}, 'xlOr'`;
}

function lower(op) {
  const { sheet, range } = op.selector;
  const call = op.intent === 'filter_range' ? `ws.GetRange(${JSON.stringify(range)}).SetAutoFilter(${op.field}, ${setArgs(op)});` : 'const af = ws.GetAutoFilter(); if (!af) return { ok: false, stage: \'no-autofilter\' }; af.ShowAllData();';
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ${call}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'filter', action: op.intent, sheet, range, field: op.field, value: op.value, values: op.values, and: op.and, top: op.top, dynamic: op.dynamic }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  // every OTHER field's filter exactly as before (the whole-task check, with no separate before, compares to itself)
  const b = before === after ? g : gate(before && before.primary, op);
  const othersKept = !b.ok ? false : JSON.stringify(others(b.st, op)) === JSON.stringify(others(g.st, op));
  const ok = op.intent === 'filter_range' ? holds(g.st, op) && othersKept : active(g.st).length === 0;
  return { ok, evidence: { sheet: op.selector.sheet, range: op.selector.range, state: g.st, proof: 'fresh-live-autofilter-model' } };
}

/** X-L13/5: clear_filter has NO extra axis of its own -- the shared digest proof is the whole of it, so a tiered
 * observation reports preservationScope FULL. filter_range stays OUT of TIERED_INTENTS (its own `layout` extra is
 * not in TIERED_EXTRAS -- a separate eligibility gap), so `tier` can never be 'tiered' for it; this check is safe
 * for both intents without branching on which one it is. Below 45000 this stays byte-identical to before. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  if (before.primary.tier === 'tiered') return { ok: rest.ok, evidence: { ...rest.evidence, preservationScope: 'FULL' } };
  return rest;
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, finalObservation, finalObservation);
  return [{ name: `${op.selector.sheet}!${op.selector.range} ${op.intent === 'filter_range' ? `filtered on field ${op.field} = ${JSON.stringify(op.value !== null ? op.value : op.values || op.and || (op.top !== null ? { top: op.top } : { dynamic: op.dynamic }))}` : 'has no active filter'}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
// X-L05: filter_range also needs the range's own row-hidden state (the new pre-write gate);
// clear_filter does not touch that gate, so it stays filterSheets-only.
const liveExtras = op => ({ filterSheets: [op.selector.sheet], ...(op.intent === 'filter_range' ? { layout: [{ sheet: op.selector.sheet, maxCol: 0, maxRow: parseRange(op.selector.range).r2 }] } : {}) });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
