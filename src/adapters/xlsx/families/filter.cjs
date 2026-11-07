'use strict';

const { parseRange, rangeText, addr } = require('../a1.cjs');
const { cellAt, isBlank } = require('../inventory.cjs');
const cells = require('./cells.cjs');
const { colourOf } = require('./filter-read.cjs');
const { tableOverlap, anyAutoFilterTable } = require('../table-guard.cjs');

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
 *
 * fa7244ac/2 (internal tracker; the LIVE measurement is c19b1065, docs/evidence/fa7244ac-w2fix-20261103/
 * fa7244ac-autofilter-object-prod24-a033de78.json, Q1/Q2/Q3): the AutoFilter OBJECT itself has two more writer
 * shapes, both through the SAME bare `SetAutoFilter()` call (no field, no criteria) -- the engine toggles it:
 * `set_autofilter_range` (Q1 MEASURED: `<autoFilter ref>` with NO filterColumn, i.e. the header's filter buttons
 * with nothing active yet) applies ONLY when the sheet carries no AutoFilter object at all; if one already
 * exists -- on the SAME range with active criteria, or on a DIFFERENT range -- it refuses named rather than
 * silently reusing the toggle's OTHER measured effect (Q3: the SAME bare call REMOVES an existing one). An exact
 * repeat (already present, exactly this range, no active criteria) is SATISFIED, not a refusal: the state asked
 * for already holds. `remove_autofilter` (Q3 MEASURED: the stored `<autoFilter>` element is gone, not just its
 * criteria the way clear_filter's `ShowAllData()` leaves it) is the toggle's removal direction; no AutoFilter on
 * the sheet is SATISFIED (nothing to remove).
 *
 * fa7244ac/2 Q2 (MEASURED): `values` equal to EVERY distinct value already in the field's own data (read from the
 * SAME live session, never assumed) is engine-side a no-op -- the motor writes no filterColumn at all (on:false in
 * session, absent from the saved file). filter_range therefore refuses this shape PRE-MUTATION, named, pointing at
 * the criterion-free shape above instead of attempting a write the engine would silently drop.
 *
 * fa7244ac/3 (LIVE finding, 8f23f67b review, prod): set_autofilter_range on a range carrying NO data at
 * all -- not even a header -- is another engine no-op the same bare SetAutoFilter() call cannot report: the
 * write dispatches, no <autoFilter> is created, and verify is left with VERIFY_CANNOT_PROVE_TERMINAL_STATE.
 * Refused PRE-MUTATION, named (xlsx-autofilter-range-empty), when every cell in the range reads blank. A
 * header-only or data-only (no header) range is NOT MEASURED and is untouched by this guard.
 *
 * GH #361 2. egyseg (the mert alap: the
 * READ-only measurement unit, docs/evidence/gh361-table-autofilter-live-20261104/, GH #361 komment 6643346450):
 * an Excel Table's OWN AutoFilter is invisible to this family's sheet-level reader (filterReads.range stays
 * null -- read_tables/tables-read is the only axis that sees it, as a presence-only boolean). Without a guard,
 * filter_range / set_autofilter_range on (or overlapping) a table's own range dispatched into
 * DISPATCHED_OUTCOME_UNKNOWN (MEASURED: T_FILTER_RANGE_WRITE / T_SET_AUTOFILTER_RANGE), and remove_autofilter
 * reported a FALSE SATISFIED while the table's own AutoFilter stayed completely untouched (MEASURED: R_EFFECT
 * SHEET_LEVEL_NOOP_TABLE_FILTER_KEPT) -- refused PRE-MUTATION, named (xlsx-filter-table-owned-autofilter),
 * instead, via table-guard.cjs (the merge.cjs/sort.cjs pattern). filter_range / set_autofilter_range: the
 * TARGET RANGE overlapping a table whose own AutoFilter flag is true (tableOverlap()). remove_autofilter /
 * clear_filter: a BLANKET, sheet-wide check (anyAutoFilterTable()) -- these two address the ONE sheet-level
 * AutoFilter the bare SetAutoFilter() toggle reaches, so their risk is not limited to the range asked. A sheet
 * with no table, or a table that does not overlap the target range (filter_range/set_autofilter_range) and
 * does not carry its own AutoFilter (remove_autofilter/clear_filter), is untouched -- same behaviour as before
 * this guard. Tables that cannot be read (CAPABILITY_MISSING or any other read gap) keep the pre-GH#361
 * behaviour too (same convention as table-guard.cjs's merge/sort consumers: an unobservable gap is never read
 * as "no table").
 */

const INTENTS = Object.freeze(['filter_range', 'clear_filter', 'set_autofilter_range', 'remove_autofilter']);

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const set = step.intent === 'filter_range';
  const fields = set ? ['intent', 'id', 'sheet', 'range', 'field', 'value', 'values', 'and', 'top', 'dynamic', 'colour'] : ['intent', 'id', 'sheet', 'range'];
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
      colour: set && 'colour' in step ? Object.freeze({ kind: step.colour.kind, rgb: Object.freeze([...step.colour.rgb]) }) : null,
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
// GH #344 colour filter WRITE (internal tracker, docs/DS-9.3.5-COLOR-FILTER-WRITE-PROPOSAL.md §3-4): a sixth,
// mutually-exclusive criteria alternative. kind maps 1:1 to the measured SetAutoFilter operator strings
// (xlsx-gh344-gh345-filter-color-empty-public-probe.live.cjs:118-119); rgb is three 0-255 integers, no
// theme/tint/indexed input in v1 (the research doc's own scope line; not measured as a write).
const COLOUR_OPERATOR_OF = Object.freeze({ cellColor: 'xlFilterCellColor', fontColor: 'xlFilterFontColor' });

function colourProblem(colour) {
  if (!colour || typeof colour !== 'object' || !Object.hasOwn(COLOUR_OPERATOR_OF, colour.kind)) return 'colour-kind-invalid';
  if (!Array.isArray(colour.rgb) || colour.rgb.length !== 3 || !colour.rgb.every(n => Number.isInteger(n) && n >= 0 && n <= 255)) return 'colour-rgb-out-of-range';
  return null;
}

function criteriaProblem(step) {
  const given = ['value', 'values', 'and', 'top', 'dynamic', 'colour'].filter(k => k in step);
  if (given.length === 0) return 'give exactly one of value / values / and / top / dynamic / colour';
  if (given.length > 1) return given.includes('colour') ? 'colour-and-other-criteria' : 'give exactly one of value / values / and / top / dynamic / colour';
  if (given[0] === 'colour') return colourProblem(step.colour);
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

/**
 * fa7244ac/3: true only when EVERY cell in the requested range (header row included) is blank. `cellAt` returns
 * null for a cell with no entry at all, and `isBlank` is false for null (adjacentData's own convention: an
 * absent cell reads as no data, not as "unreadable" -- p.complete was already checked above this call).
 */
function rangeIsEmpty(p, op) {
  const g = parseRange(op.selector.range);
  for (let r = g.r1; r <= g.r2; r += 1) {
    for (let c = g.c1; c <= g.c2; c += 1) {
      const cell = cellAt(p, op.selector.sheet, addr(c, r));
      if (cell && !isBlank(cell)) return false;
    }
  }
  return true;
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  // GH #361 2. egyseg: a Table-owned AutoFilter, checked BEFORE the sheet-level filterReads axis even matters --
  // see this file's own header comment for the measured risk and the range-vs-sheet split between intents.
  const tableBlock = op.intent === 'filter_range' || op.intent === 'set_autofilter_range'
    ? tableOverlap(p, op.selector.sheet, op.selector.range)
    : anyAutoFilterTable(p, op.selector.sheet);
  if (tableBlock.table && tableBlock.table.autoFilter) return { ok: false, reason: 'xlsx-filter-table-owned-autofilter', table: tableBlock.table.name, ref: tableBlock.table.ref };
  const st = p.filterReads ? p.filterReads[op.selector.sheet] : undefined;
  if (!st || typeof st !== 'object') return { ok: false, reason: 'xlsx-filter-unobservable', why: st === undefined ? 'not-read' : st };
  if (st.range !== null && String(st.range).replace(/\$/g, '') !== op.selector.range) return { ok: false, reason: 'xlsx-filter-range-conflict', autoFilter: st.range };
  // fa7244ac/2: the toggle's OTHER measured direction (Q3) removes a whole AutoFilter -- set_autofilter_range may
  // only reach for the create direction when none exists yet; an existing one with active criteria on the SAME
  // range refuses named rather than being silently cleared by the bare SetAutoFilter() call. (A different range
  // already refused above, same reason as filter_range/clear_filter.)
  if (op.intent === 'set_autofilter_range' && st.range !== null && active(st).length > 0) return { ok: false, reason: 'xlsx-autofilter-already-present', autoFilter: st.range };
  // fa7244ac/3 LIVE finding (8f23f67b review, gh141-final-readwrite-acceptance.live.cjs, prod): on a
  // range that carries NO data at all -- not even a header -- the engine's bare SetAutoFilter() does not create
  // the <autoFilter> object, but the write still dispatches and then VERIFY_CANNOT_PROVE_TERMINAL_STATE (the
  // same seed WITH a header + two data rows applied cleanly). Refused before the write, named, rather than
  // dispatching a write the engine silently drops. Only the fully-empty case is MEASURED -- a header-only or
  // data-only (no header) range is untouched by this guard and behaves as before (NOT MEASURED).
  if (op.intent === 'set_autofilter_range' && st.range === null && rangeIsEmpty(p, op)) return { ok: false, reason: 'xlsx-autofilter-range-empty' };
  const region = op.intent === 'filter_range' && st.range === null ? adjacentData(p, op) : null;
  if (region) return { ok: false, ...region };
  const manual = op.intent === 'filter_range' && !active(st).length ? manualHiddenRow(p, op) : null;
  if (manual) return { ok: false, ...manual };
  // GH #344 G2's other-field data-loss risk outranks fa7244ac/2's "this is a no-op" classification when a
  // values set happens to be both 3+ AND the column's full distinct set (the two axes are independent, but the
  // measured LIVE danger -- an existing LIVE finding -- is the one named when they coincide).
  const risk = op.intent === 'filter_range' && !holds(st, op) ? values3PlusOtherFieldAtRisk(st, op) : null;
  if (risk) return { ok: false, ...risk };
  const everything = op.intent === 'filter_range' && !holds(st, op) ? allValuesInColumn(p, op) : null;
  return everything ? { ok: false, ...everything } : { ok: true, st };
}

/**
 * fa7244ac/2 Q2 (MEASURED, prod /24): `values` naming EVERY distinct value already in the field's column data is
 * a no-op to the engine (no filterColumn written at all) -- refused PRE-MUTATION, named, rather than dispatching a
 * write the motor silently drops. The distinct set is read from THIS session's own cells (never assumed); a cell
 * the observation cannot read (outside a tiered read's windows) makes the set incomplete, so this check steps
 * aside (returns null, same as "not this shape") rather than guessing either way -- the write proceeds normally,
 * and if it IS the no-op shape the generic not-applied verify failure still catches it, just without this name.
 * A blank data cell is excluded from the set: `values` has no blank-selection shape to ever equal it against
 * (plainValue() above), so including blanks could never turn a genuine filter into a false "select everything".
 */
function allValuesInColumn(p, op) {
  if (!(op.values && op.values.length >= 2)) return null;
  const g = parseRange(op.selector.range);
  const col = g.c1 + op.field - 1;
  const seen = new Set();
  for (let r = g.r1 + 1; r <= g.r2; r += 1) {
    const cell = cellAt(p, op.selector.sheet, addr(col, r));
    if (!cell) return null; // unreadable: cannot determine, do not refuse pre-mutation on a guess
    if (isBlank(cell)) continue;
    seen.add(String(cell.v));
  }
  if (seen.size === 0) return null;
  const requested = new Set(op.values.map(String));
  if (requested.size !== seen.size) return null;
  for (const v of requested) if (!seen.has(v)) return null;
  return { reason: 'xlsx-filter-values-select-everything-is-no-filter', distinct: [...seen].slice(0, 20), count: seen.size };
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

/**
 * LIVE finding (GH #344 G2 regression probe, prod 1.1.0-44fe2c4e8639, internal tracker): a values list of 3 or
 * more (xlFilterValues) written to one field while ANOTHER field already carries an active filter made the OTHER
 * field's filter disappear from the read-back, and the write itself came back DISPATCHED_OUTCOME_UNKNOWN.
 * Reproduced on the installed product WITHOUT any colour filter (a plain one-value filter on the other field is
 * enough) -- the 3+-value path is the trigger, not GH #344. A two-value list (xlOr) and every other criteria form
 * do NOT show this (R19's own test proves several active fields normally coexist, xlsx-m7-d-filter.test.cjs).
 * Refused BEFORE the write (X-L05's own pattern, adjacentData/manualHiddenRow above) rather than letting a
 * destructive write run and reporting the loss only after the fact. Only gates an ACTUAL write (holds() already
 * false): a values3+ filter that is already exactly applied, with another field active alongside it, is left
 * alone -- this gate only stops US from TRIGGERING the loss, not an already-settled state we did not create.
 */
function values3PlusOtherFieldAtRisk(st, op) {
  if (!(op.values && op.values.length > 2)) return null;
  const otherActive = active(st).filter(f => f.field !== op.field);
  return otherActive.length ? { reason: 'xlsx-filter-values3plus-would-drop-other-field', otherFields: otherActive.map(f => f.field) } : null;
}

const empty = x => x === null || x === undefined || x === '';
const sameSet = (a, b) => a.length === b.length && [...a].sort().join('\u0000') === [...b].sort().join('\u0000');

const hexOfRgb = ([r, g, b]) => [r, g, b].map(n => n.toString(16).padStart(2, '0')).join('').toUpperCase();

/**
 * GH #344: the public getters read a colour filter's criteria1/criteria2 as null (filter-read.cjs's own comment,
 * AF-02 LIVE) -- so unlike every other shape, `matches()` cannot compare against f.criteria1/criteria2 directly.
 * The ONLY in-family resolution is filter-read's `colourOf()` (Bridge /21), reused here rather than duplicated
 * (design doc §3, option (b)). `colourOf()` always returns `unsupported: true`; a successful resolution carries a
 * non-null `.color` and NO `.colorReason` -- anything else (UNOBSERVABLE/kind-mismatch/COLOR_NONE) is the
 * `colour-unverifiable` refusal (design doc §4): the write is not reported APPLIED even though SetAutoFilter did
 * not throw. When the colour resolves as `form: 'rgb'`, the exact hex IS compared (we can prove it exactly); for
 * `indexed`/`theme` forms there is, as of this unit, no palette-to-RGB table to compare against (the measured
 * asymmetry the design doc names: write any RGB, but only provably verify the colour the Bridge resolves) -- a
 * matching `kind` on a successful resolution is accepted. Whether a non-palette RGB round-trips exactly is G1,
 * not run here.
 */
function matchesColour(f, op, colours) {
  if (f.operator !== COLOUR_OPERATOR_OF[op.colour.kind] || !empty(f.criteria1) || !empty(f.criteria2)) return false;
  const resolved = colourOf(op.field, f, colours);
  if (resolved.colorReason || !resolved.color) return false; // colour-unverifiable
  if (resolved.kind !== op.colour.kind) return false;
  if (resolved.color.form === 'rgb') return String(resolved.color.hex).toUpperCase() === hexOfRgb(op.colour.rgb);
  return true; // indexed/theme: kind-match accepted, exact-RGB not provable yet (G1)
}

/** The requested filter of op.field, as the reader shows it. `colours` is only used by the colour branch. */
function matches(f, op, colours) {
  if (op.colour) return matchesColour(f, op, colours);
  if (op.values && op.values.length > 2) return f.operator === 'xlFilterValues' && Array.isArray(f.criteria1) && empty(f.criteria2) && sameSet(f.criteria1.map(String), op.values);
  if (op.values) return f.operator === 'xlOr' && bare(f.criteria1) === op.values[0] && bare(f.criteria2) === op.values[1];
  if (op.top !== null) return f.operator === 'xlTop10Items' && Number(f.criteria1) === op.top && empty(f.criteria2);
  if (op.dynamic !== null) return f.operator === 'xlFilterDynamic' && f.criteria1 === DYNAMIC[op.dynamic] && empty(f.criteria2);
  if (op.and) return f.operator === 'xlAnd' && String(f.criteria1) === op.and[0] && String(f.criteria2) === op.and[1];
  return bare(f.criteria1) === op.value && empty(f.criteria2);
}

const mine = (st, op) => active(st).filter(f => f.field === op.field);
const holds = (st, op) => st.range !== null && mine(st, op).length === 1 && matches(mine(st, op)[0], op, st.colours);
const others = (st, op) => (st.range === null ? [] : st.filters.filter(f => f.field !== op.field));

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, range: op.selector.range };
  if (op.intent === 'clear_filter') return active(g.st).length === 0 ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-no-active-filter' } } : { state: 'APPLY', evidence: where };
  // fa7244ac/2: gate() above already enforced range equality (when non-null) and the no-active-criteria
  // precondition for set_autofilter_range, so a non-null st.range here IS exactly the requested, button-only state.
  if (op.intent === 'set_autofilter_range') return g.st.range !== null ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-autofilter-buttons-present' } } : { state: 'APPLY', evidence: where };
  if (op.intent === 'remove_autofilter') return g.st.range === null ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-no-autofilter' } } : { state: 'APPLY', evidence: where };
  if (holds(g.st, op)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-the-filter' } };
  return { state: 'APPLY', evidence: { ...where, field: op.field, value: op.value, values: op.values, and: op.and, top: op.top, dynamic: op.dynamic, otherFields: active(g.st).filter(f => f.field !== op.field).map(f => f.field) } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range }) };
}

function setArgs(op) {
  // GH #344: SetAutoFilter(field, Api.CreateColorFromRGB(r,g,b), 'xlFilter<Cell|Font>Color') -- the SAME positional
  // (criteria1, operator) shape as every other branch, measured in xlsx-gh344-gh345-filter-color-empty-public-
  // probe.live.cjs:118-119. Api.CreateColorFromRGB is the SAME builder format.cjs/cf.cjs/format-remove.cjs already
  // call for a plain RGB colour -- no new helper.
  if (op.colour) return `Api.CreateColorFromRGB(${op.colour.rgb[0]}, ${op.colour.rgb[1]}, ${op.colour.rgb[2]}), ${JSON.stringify(COLOUR_OPERATOR_OF[op.colour.kind]).replace(/"/g, "'")}`;
  if (op.values) return `${JSON.stringify(op.values)}, 'xlFilterValues'`;
  if (op.and) return `${JSON.stringify(op.and[0])}, 'xlAnd', ${JSON.stringify(op.and[1])}`;
  if (op.top !== null) return `${op.top}, 'xlTop10Items'`;
  if (op.dynamic !== null) return `${JSON.stringify(DYNAMIC[op.dynamic]).replace(/"/g, "'")}, 'xlFilterDynamic'`;
  return `${JSON.stringify(op.value)}, 'xlOr'`;
}

function callOf(op, range) {
  if (op.intent === 'filter_range') return `ws.GetRange(${JSON.stringify(range)}).SetAutoFilter(${op.field}, ${setArgs(op)});`;
  if (op.intent === 'clear_filter') return 'const af = ws.GetAutoFilter(); if (!af) return { ok: false, stage: \'no-autofilter\' }; af.ShowAllData();';
  // fa7244ac/2: set_autofilter_range / remove_autofilter are the SAME bare SetAutoFilter() call, no field, no
  // criteria -- the engine toggles it (Q1 MEASURED: creates buttons-only when none exists; Q3 MEASURED: removes
  // the whole object when one does). gate()/classify() already decided which direction this dispatch is.
  return `ws.GetRange(${JSON.stringify(range)}).SetAutoFilter();`;
}

function lower(op) {
  const { sheet, range } = op.selector;
  const call = callOf(op, range);
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
    meta: Object.freeze({ family: 'filter', action: op.intent, sheet, range, field: op.field, value: op.value, values: op.values, and: op.and, top: op.top, dynamic: op.dynamic, colour: op.colour }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  // every OTHER field's filter exactly as before (the whole-task check, with no separate before, compares to itself)
  const b = before === after ? g : gate(before && before.primary, op);
  const othersKept = !b.ok ? false : JSON.stringify(others(b.st, op)) === JSON.stringify(others(g.st, op));
  // fa7244ac/2: set_autofilter_range needs the OBJECT present on exactly this range with no active criteria yet
  // (the buttons-only state); remove_autofilter needs the object entirely gone (stronger than clear_filter's
  // "no active filter", which still leaves an inactive <autoFilter> behind).
  const ok = op.intent === 'filter_range' ? holds(g.st, op) && othersKept
    : op.intent === 'set_autofilter_range' ? g.st.range !== null && active(g.st).length === 0
    : op.intent === 'remove_autofilter' ? g.st.range === null
    : active(g.st).length === 0; // clear_filter
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

const criteriaLabel = op => (op.value !== null ? op.value : op.values || op.and || (op.top !== null ? { top: op.top } : op.dynamic !== null ? { dynamic: op.dynamic } : { colour: op.colour }));

function labelOf(op) {
  if (op.intent === 'filter_range') return `filtered on field ${op.field} = ${JSON.stringify(criteriaLabel(op))}`;
  if (op.intent === 'set_autofilter_range') return 'has AutoFilter buttons with no active criteria';
  if (op.intent === 'remove_autofilter') return 'has no AutoFilter';
  return 'has no active filter'; // clear_filter
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, finalObservation, finalObservation);
  return [{ name: `${op.selector.sheet}!${op.selector.range} ${labelOf(op)}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
// X-L05: filter_range also needs the range's own row-hidden state (the new pre-write gate);
// clear_filter does not touch that gate, so it stays filterSheets-only.
// GH #344: a colour criteria asks the driver for the /21 colour axis (filterColor) -- without it the observation carries
// no `colours` and the L1 check could never resolve the colour it wrote
// GH #361 2. egyseg: every intent now also asks for its own sheet's tables (read_tables' own axis, the
// merge.cjs/sort.cjs tableSheets pattern) -- the table-owned-AutoFilter guard in gate() needs it regardless of
// intent; a sheet with no table (or where the read cannot be done) keeps the pre-GH#361 behaviour unchanged.
const liveExtras = op => ({ filterSheets: [op.selector.sheet], tableSheets: [op.selector.sheet], ...(op.colour ? { filterColor: true } : {}), ...(op.intent === 'filter_range' ? { layout: [{ sheet: op.selector.sheet, maxCol: 0, maxRow: parseRange(op.selector.range).r2 }] } : {}) });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
