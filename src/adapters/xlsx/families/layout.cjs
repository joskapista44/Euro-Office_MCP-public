'use strict';

const { colNum, colName, parseCell, MAX_COL, MAX_ROW } = require('../a1.cjs');
const cells = require('./cells.cjs');
const { observedStyleChanges } = require('../style-axis.cjs');

/**
 * XLSX `layout` family: layout_range -- column width, row height, hide / show columns or rows (M6 wave 4),
 * AutoFit (M7 backlog X-L01/X-L02).
 *
 * Measured (W4b LIVE probe, docs/evidence/M6-W4-APPEARANCE-AUDIT-AND-CONTRACT.md section 5): the PUBLIC
 * GetEntireColumn()/GetEntireRow() setters and getters, per column / row; a set width / height reads back
 * exactly and survives a reopen; a hidden column reads width 0, a hidden row height null; an unset row
 * height is unobservable (null in the session, 15 after a reopen) -- only a SET height is verified.
 *
 * AutoFit (S30/S30b/S30c, docs/M7-AUTOFIT-CONTRACT-DRAFT.md; owner option B, owner channel,
 * 2026-10-05 -- this REPLACES the prior #ref refusal below, kept as the narrower width/height/hidden
 * decision it always was): the end state still cannot be read without re-applying it, so a retry is
 * NEVER treated as already-satisfied -- classify() always returns APPLY for `autofit`, re-dispatching the
 * SAME command every time. The write (MEASURED): row -> SetRowHeight(<the row's OWN current GetRowHeight(),
 * read live inside the command>) then AutoFit(true, false) -- in that order, so a SHRINK persists through a
 * reopen (S30c: AutoFit alone does not persist a row shrink; marking the row custom first does). Column ->
 * AutoFit(false, true) alone. Looped PER row / column inside one command (never one AutoFit call across a
 * multi-item span): only a single-row / single-column AutoFit was ever measured, and SetRowHeight(current)
 * needs each row's OWN height, not one value forced onto every target row. Refused before the write: a
 * merge anywhere among the target rows / columns' used cells (`mergedInRange` below -- the family's own
 * default used-range read already carries every cell's `.merge`; layout_range is not a TIERED_INTENTS
 * member, so this is never a partial read). L1 is deliberately weak and HONEST: only that the row / column
 * now reads a numeric size and is not hidden -- never that the size is "the correct fit" (there is no
 * reference to compare against; a reference write was the REJECTED option C). L3 (no other row / column /
 * cell changes) is the family's existing, unmodified verifyPreservation.
 *
 * Also refused before the write, stated: a width / height on a hidden column / row (its size reads 0 /
 * null, the set value would be unverifiable) -- `autofit` shares this same check (classify()'s existing
 * hidden-target branch runs for every prop other than `hidden` itself). L1 per column / row; L3 every
 * other column / row of every sheet (over its used range and the requested extent) keeps width / height /
 * hidden, every cell its content. Retry (width / height / hidden) = the terminal state on a fresh read ->
 * SATISFIED, 0 write / 0 command. Retry (autofit) = ALWAYS APPLY, 1 write / 1 command again (owner option B).
 */

const INTENTS = Object.freeze(['layout_range']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'columns', 'rows', 'width', 'height', 'hidden', 'autofit']);
const MAX_ITEMS = 500;
const TOL = 0.05;

function span(text, kind) {
  const m = kind === 'columns' ? /^([A-Z]{1,3})(?::([A-Z]{1,3}))?$/.exec(String(text)) : /^([1-9]\d{0,6})(?::([1-9]\d{0,6}))?$/.exec(String(text));
  if (!m) return null;
  const n = kind === 'columns' ? colNum : Number;
  const from = n(m[1]);
  const to = n(m[2] || m[1]);
  const max = kind === 'columns' ? MAX_COL : MAX_ROW;
  return from <= to && to <= max ? { from, to } : null;
}

function plan(step, index) {
  const at = `step ${index}: layout_range`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  if (('columns' in step) === ('rows' in step)) return { ok: false, detail: `${at} requires exactly one of columns ('B' / 'B:D') or rows ('2' / '2:5')` };
  const kind = 'columns' in step ? 'columns' : 'rows';
  const props = ['width', 'height', 'hidden', 'autofit'].filter(k => k in step);
  if (props.length !== 1) return { ok: false, detail: `${at} requires exactly one of width, height, hidden, autofit` };
  const prop = props[0];
  const value = step[prop];
  if (prop === 'width' && kind !== 'columns') return { ok: false, detail: `${at}: width applies to columns` };
  if (prop === 'height' && kind !== 'rows') return { ok: false, detail: `${at}: height applies to rows` };
  if (prop === 'hidden' && typeof value !== 'boolean') return { ok: false, detail: `${at}: hidden must be true or false` };
  if (prop === 'autofit' && value !== true) return { ok: false, detail: `${at}: autofit must be true` };
  if (prop === 'width' && !(Number.isFinite(value) && value > 0 && value <= 255)) return { ok: false, detail: `${at}: width must be a number in (0, 255]` };
  if (prop === 'height' && !(Number.isFinite(value) && value > 0 && value <= 409)) return { ok: false, detail: `${at}: height must be a number in (0, 409]` };
  const g = span(step[kind], kind);
  if (!g) return { ok: false, detail: `${at} requires plain ${kind === 'columns' ? "column letters ('B' or 'B:D')" : "row numbers ('2' or '2:5')"}` };
  const count = g.to - g.from + 1;
  if (count > MAX_ITEMS) return { ok: false, detail: `${at} names ${count} ${kind}, over the ${MAX_ITEMS} ceiling` };
  const items = [];
  for (let i = g.from; i <= g.to; i += 1) items.push(kind === 'columns' ? colName(i) : String(i));
  const text = kind === 'columns' ? `${colName(g.from)}:${colName(g.to)}` : `${g.from}:${g.to}`;
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'layout_range',
      family: 'layout',
      // the selector range is the A1 area the columns / rows cover (the planner's overlap check uses it)
      selector: Object.freeze({ sheet: step.sheet, range: kind === 'columns' ? `${colName(g.from)}1:${colName(g.to)}${MAX_ROW}` : `A${g.from}:${colName(MAX_COL)}${g.to}`, [kind]: text }),
      kind,
      prop,
      value,
      items: Object.freeze(items),
      extent: Object.freeze({ maxCol: kind === 'columns' ? g.to : 0, maxRow: kind === 'rows' ? g.to : 0 }),
      cells: Object.freeze([]),
      footprint: footprintOf(kind, prop, g),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/**
 * M7 S3: what the step modifies, per axis. A row height / hidden: its rows; a hidden column: its columns. A column
 * WIDTH also re-fits the AUTO height of a wrapped row ANYWHERE in those columns (W4e LIVE), so it claims EVERY row
 * ('*') -- not only the rows of cells it overlaps.
 */
function footprintOf(kind, prop, g) {
  const span = Object.freeze([g.from, g.to]);
  if (kind === 'rows') return Object.freeze({ rows: span });
  return Object.freeze((prop === 'width' || prop === 'autofit') ? { columns: span, rows: '*' } : { columns: span });
}

const entriesOf = (axis, op) => (op.kind === 'columns' ? axis.cols : axis.rows);
const near = (a, b) => typeof a === 'number' && Math.abs(a - b) <= TOL;
/**
 * OD-2a / D-W (owner GO, owner channel, 2026-10-05): a column WIDTH is verified against the engine's CANONICAL persisted
 * value of the request (the driver's build-gated pure conversion, `axis.canon`), EXACTLY -- never against the requested
 * number with a tolerance (20.5 persists as 20.43; claiming 20.5 would be false). A height keeps TOL (X-L04, measured
 * row-axis drift). The canonical value unobservable -> the width step is refused before the write (fail-closed).
 */
const canonOf = (axis, op) => (op.prop === 'width' && axis && axis.canon && typeof axis.canon === 'object' ? axis.canon[String(op.value)] : undefined);
const targetOf = (axis, op) => (op.prop === 'width' ? canonOf(axis, op) : op.value);
const holds = (e, op, target) => {
  if (op.prop === 'hidden') return e.hidden === op.value;
  if (e.hidden !== false) return false;
  return op.prop === 'width' ? typeof target === 'number' && e.width === target : near(e[op.prop], op.value);
};

/** M7 X-L01/X-L02: a merge anywhere among the target rows' / columns' USED cells refuses autofit before
 * the write. layout_range is not a TIERED_INTENTS member, so the family's own default used-range read
 * (never a partial one) already carries every cell's `.merge` -- no extra window is requested. */
function mergedInRange(primary, sheet, kind, items) {
  const want = new Set(items);
  for (const [k, c] of Object.entries(primary.cells || {})) {
    const sep = k.indexOf('\u0000');
    if (k.slice(0, sep) !== sheet) continue;
    const a = k.slice(sep + 1);
    const pc = parseCell(a);
    if (!pc) continue;
    const id = kind === 'columns' ? colName(pc.c) : String(pc.r);
    if (!want.has(id)) continue;
    if (c.merge && c.merge !== a) return { cell: a, mergeArea: c.merge };
  }
  return null;
}

function gate(primary, op) {
  if (!primary || primary.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!primary.complete) return { ok: false, reason: primary.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: primary.unsupported };
  const s = primary.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: primary.sheets.map(x => x.name) };
  const axis = primary.layout && primary.layout[op.selector.sheet];
  if (!axis || typeof axis !== 'object') return { ok: false, reason: 'xlsx-layout-unobservable', detail: axis || 'not-read' };
  const found = {};
  for (const i of op.items) {
    const e = entriesOf(axis, op)[i];
    if (!e || typeof e.hidden !== 'boolean') return { ok: false, reason: 'xlsx-layout-target-unread', item: i };
    found[i] = e;
  }
  return propRefusal(primary, axis, op) || { ok: true, entries: found, target: targetOf(axis, op) };
}

/** The prop-specific refusals before the write: a width whose canonical value is unobservable (D-W, fail-closed); an
 * autofit over a merge (X-L01/X-L02). null = none. */
function propRefusal(primary, axis, op) {
  if (op.prop === 'width' && typeof canonOf(axis, op) !== 'number') return { ok: false, reason: 'xlsx-layout-width-canonical-unobservable', detail: axis.canon === undefined ? 'not-read' : axis.canon };
  const merged = op.prop === 'autofit' ? mergedInRange(primary, op.selector.sheet, op.kind, op.items) : null;
  return merged ? { ok: false, reason: 'xlsx-layout-autofit-merged-range', ...merged } : null;
}

function classify(primary, op) {
  const g = gate(primary, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, [op.kind]: op.selector[op.kind], [op.prop]: op.value, ...(op.prop === 'width' ? { canonical: g.target } : {}) };
  if (op.prop !== 'hidden') {
    const hidden = op.items.filter(i => g.entries[i].hidden === true);
    if (hidden.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-layout-target-hidden', items: hidden.slice(0, 20), count: hidden.length } };
  }
  // M7 X-L01/X-L02, owner option B: autofit NEVER reads as already-satisfied -- it always re-dispatches.
  if (op.prop === 'autofit') return { state: 'APPLY', evidence: { ...where, note: 'autofit always re-applies (owner option B): no state is treated as satisfied' } };
  if (op.items.every(i => holds(g.entries[i], op, g.target))) return { state: 'SATISFIED', evidence: { ...where, proof: `fresh-live-every-${op.kind === 'columns' ? 'column' : 'row'}` } };
  return { state: 'APPLY', evidence: where };
}

function resolve(primary, selector, ctx = {}) {
  const g = gate(primary, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, before: g.entries }) };
}

/** One callCommand: an exact-name sheet check, then the public setter on the entire columns / rows.
 * `autofit` loops PER item inside the SAME command -- a row's SetRowHeight(current) step needs ITS OWN
 * current height, not one value forced onto every target row, and only a single-row / single-column
 * AutoFit was ever measured (docs/M7-AUTOFIT-CONTRACT-DRAFT.md). */
function lower(op) {
  const { sheet } = op.selector;
  if (op.prop === 'autofit') {
    const steps = op.kind === 'rows'
      ? op.items.map(i => `{ const r = ws.GetRange(${JSON.stringify(`A${i}`)}).GetEntireRow(); r.SetRowHeight(r.GetRowHeight()); r.AutoFit(true, false); }`)
      // column (p56 LIVE, 2026-10-05): the engine's column AutoFit returns WITHOUT a change when the width is already
      // the fit (WorksheetView _autoFitColumnWidth: cc === oldColWidth) -- a retry then dispatches a write with no save
      // traffic and the save contract (mutated-without-observed-save) fails it. One width step away first, in the SAME
      // command, makes the AutoFit always write; the end state is the plain AutoFit's (the fit width, BestFit).
      : op.items.map(i => `{ const c = ws.GetRange(${JSON.stringify(`${i}1`)}).GetEntireColumn(); const w = c.GetColumnWidth(); c.SetColumnWidth(w >= 254 ? w - 1 : w + 1); c.AutoFit(false, true); }`);
    return {
      kind: 'callCommand',
      body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ${steps.join('\n  ')}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
      meta: Object.freeze({ family: 'layout', sheet, kind: op.kind, items: op.items, prop: op.prop, value: op.value }),
    };
  }
  const first = op.items[0];
  const last = op.items[op.items.length - 1];
  const a1 = op.kind === 'columns' ? `${first}1:${last}1` : `A${first}:A${last}`;
  const entire = op.kind === 'columns' ? 'GetEntireColumn()' : 'GetEntireRow()';
  const setter = { width: 'SetColumnWidth', height: 'SetRowHeight', hidden: 'SetHidden' }[op.prop];
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.GetRange(${JSON.stringify(a1)}).${entire}.${setter}(${JSON.stringify(op.value)});
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'layout', sheet, kind: op.kind, items: op.items, prop: op.prop, value: op.value }),
  };
}

/** L1: every target column / row holds the value (a size also needs the column / row visible).
 * `autofit` is deliberately weaker and honest: only that the row / column now reads a numeric size and
 * is not hidden -- never that the size is "the correct fit" (there is no reference to compare against;
 * a reference write was the rejected option C, docs/M7-AUTOFIT-CONTRACT-DRAFT.md). */
/** X-L01/X-L02 L1: the row / column now reads a numeric size and is not hidden -- never a "correct fit" claim.
 * `sizes` is echoed back for a caller (the LIVE harness) that wants to compare a retry's result to the first
 * call's -- NOT itself a correctness claim; verifyPrimitive never knows the "right" fitted value. */
function verifyAutofit(op, axis) {
  const size = op.kind === 'columns' ? 'width' : 'height';
  const entry = i => entriesOf(axis, op)[i];
  const wrong = op.items.filter(i => !entry(i) || entry(i).hidden !== false || typeof entry(i)[size] !== 'number');
  const sizes = Object.fromEntries(op.items.map(i => [i, entry(i) ? entry(i)[size] : null]));
  return { ok: wrong.length === 0, evidence: { sheet: op.selector.sheet, [op.kind]: op.selector[op.kind], prop: op.prop, wrong: wrong.slice(0, 20), sizes, proof: 'autofit-in-session-size-readable-not-hidden' } };
}

function verifyPrimitive(op, _before, after) {
  const a = after && after.primary;
  const axis = a && a.layout && a.layout[op.selector.sheet];
  if (!axis || typeof axis !== 'object') return { ok: false, evidence: { reason: 'xlsx-layout-unobservable-after', detail: axis || 'not-read' } };
  if (op.prop === 'autofit') return verifyAutofit(op, axis);
  const target = targetOf(axis, op);
  const wrong = op.items.filter(i => { const e = entriesOf(axis, op)[i]; return !e || !holds(e, op, target); });
  return { ok: wrong.length === 0, evidence: { sheet: op.selector.sheet, [op.kind]: op.selector[op.kind], prop: op.prop, value: op.value, ...(op.prop === 'width' ? { requested: op.value, canonical: target } : {}), wrong: wrong.slice(0, 20), proof: op.prop === 'width' ? 'fresh-live-per-column-equals-canonical' : 'fresh-live-per-column-or-row' } };
}

/**
 * L3: every cell keeps its content (cells.verifyPreservation, no target cell), every other column / row of every sheet its
 * layout, and (M7 S3) every cell of the target sheet its style -- a size never changes a format (the style is read for
 * the target sheet; a key unobservable on either side makes no claim).
 */
/**
 * An unset (auto) row READS null or its single-line default depending on when the engine materialised it (W4e LIVE,
 * 2026-09-29, docs/evidence/m6-w4e-probe-20260929.json: 14.25 appeared on a CONTROL sheet with no layout write at all;
 * after reopen the rows read null again and a wrapped row still grew on its own -- they stayed AUTO). So null <-> one of
 * these defaults is the same state; any other height change on a row IS an L3 change.
 */
const AUTO_ROW_DEFAULTS = Object.freeze([14.25, 15]);
const sameAutoRow = (e, n) => e.hidden === n.hidden && ((e.height === null && AUTO_ROW_DEFAULTS.includes(n.height)) || (n.height === null && AUTO_ROW_DEFAULTS.includes(e.height)));

function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary.layout;
  const a = after.primary.layout;
  if (!b || !a) return { ok: false, evidence: { reason: 'xlsx-layout-unobservable', before: !!b, after: !!a } };
  const changed = [];
  for (const [sheet, ax] of Object.entries(b)) {
    const now = a[sheet];
    if (typeof ax !== 'object' || typeof now !== 'object' || now === null) { changed.push({ sheet, change: 'layout-unobservable' }); continue; }
    for (const kind of ['cols', 'rows']) {
      const size = kind === 'cols' ? 'width' : 'height';
      for (const [i, e] of Object.entries(ax[kind])) {
        const isTarget = sheet === op.selector.sheet && (kind === 'cols') === (op.kind === 'columns') && op.items.includes(i);
        if (isTarget) continue;
        const n = now[kind][i];
        if (!n) continue; // outside the after read's extent (a used range can shrink); the before side decides
        if (kind === 'rows' && sameAutoRow(e, n)) continue;
        if (n.hidden !== e.hidden || n[size] !== e[size]) changed.push({ sheet, [kind === 'cols' ? 'column' : 'row']: i, before: e, after: n });
      }
    }
  }
  const style = observedStyleChanges(before.primary, after.primary);
  return { ok: rest.ok && changed.length === 0 && style.changed.length === 0, evidence: { ...rest.evidence, layoutChanged: changed.slice(0, 20), layoutChangedCount: changed.length, styleCompared: style.compared, styleChanged: style.changed.slice(0, 20), styleChangedCount: style.changed.length } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, finalObservation, finalObservation);
  return [{ name: `${op.selector.sheet} ${op.kind} ${op.selector[op.kind]} ${op.prop} = ${op.value}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

/** No cell windows (a whole column is a million cells); the layout extent instead. */
const windowsOf = () => [];
const liveExtras = op => ({ layout: [{ sheet: op.selector.sheet, ...op.extent, ...(op.prop === 'width' ? { canon: [op.value] } : {}) }], styleSheets: [op.selector.sheet] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, TOL };
