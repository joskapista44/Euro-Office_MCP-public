'use strict';

const { parseRange, rangeText, addr, MAX_COL, MAX_ROW } = require('../a1.cjs');
const { shiftFormula, referencesBlock } = require('../shift.cjs');
const { cellAt, formulaEntries } = require('../inventory.cjs');

/**
 * XLSX `table-structure` family (GH #336 T05a rows, GH #337 T05b columns): insert_table_rows / delete_table_rows /
 * insert_table_columns / delete_table_columns {sheet, table, at, count, currentRows | currentColumns}.
 *
 * ENGINE PATH (measured LIVE, prod DS 9.3.5, baseline docs/evidence/gh336-gh337-baseline-20261030.json): the PUBLIC
 * ApiRange.Insert('down'|'right') / Delete('up'|'left') over a range that is the table's FULL width (a row op) or FULL height,
 * header included (a column op). Those are right (table ref, autoFilter, column names, anchors, a formula that uses A1 references).
 * Everything else the engine does is a data hazard (a partial-width row shears the table, a header delete eats the first data
 * row, an all-rows delete leaves a header-only table, a delete leaves #REF! in formulas outside) -- so the family only ever emits
 * the full-width / full-height range and refuses, before the write, each named:
 *   a structured reference to the table anywhere in the workbook     xlsx-table-structured-reference
 *       (LIVE: the engine leaves the STORED value of Table[Col] stale after an insert/delete, in-session AND after a reopen, and no
 *        public recalculation fixes the saved cache -- docs/evidence/gh336-gh337-recalc|reopen|reopen-save-20261030.json)
 *   a table without a header row / with a totals row                 xlsx-table-no-header-row / xlsx-table-totals-row-unmeasured
 *   deleting EVERY data row (a header-only table) / every column     xlsx-table-delete-all-data-rows / xlsx-table-delete-all-columns
 *   a formula the delete would BREAK (#REF!)                         xlsx-table-block-referenced
 *   a merge, another table or a pivot in the SHIFTED BAND            xlsx-table-band-has-merge / -table / -pivot
 *       (rows: the table's columns, from the cut down; columns: the table's rows, from the cut right)
 *   a validation / conditional format / chart / defined name that reaches the sheet (conservative; the structure family's reader)
 *                                                                    xlsx-table-structure-dependent-not-measured
 * `at` is the 1-based index WITHIN the table: a data row (1 = the first data row) or a column (1 = the first column), so the header
 * row is not addressable at all. An insert goes BEFORE `at`; the table cannot grow by an append (LIVE: an insert under the last row
 * does not enlarge it). The new columns' headers are the engine's own (Column1, ...): v1 reports them, never sets them.
 *
 * GH #366 6th-unit follow-up (LIVE, prod, SMALL copy): `insert_table_columns` AT THE TABLE'S OWN FIRST COLUMN
 * (`at: 1`) does not grow the table -- the engine shifts the WHOLE TABLE one column right instead (no new column,
 * no data change) -- xlsx-table-insert-before-first-column-unsupported, refused PRE-dispatch, 0 writes. Rows have
 * no equivalent hazard (the header row, the table's anchor, is excluded from `at` entirely for row ops); columns
 * address the anchor column too, so `at: 1` there coincides with it. No clean engine path is known yet.
 *
 * RETRY: a range Insert/Delete is not idempotent, so every step carries the table's CURRENT size (currentRows | currentColumns): the
 * fresh read at that size -> APPLY; at the size the step would produce -> SATISFIED (0 write); any other size -> refused,
 * xlsx-table-size-unexpected. L1: the table's ref, columns (the untouched ones, in order), headerRow, totalsRow, autoFilter and style
 * as the shift implies, the shifted band's literal cells where the shift puts them, the new cells blank. L3: everything outside the
 * band unchanged (every other sheet too); no #REF! appears.
 * Needs a FULL-tier observation (every cell of the sheet read): a windowed / tiered read is refused, named.
 */

const INTENTS = Object.freeze(['insert_table_rows', 'delete_table_rows', 'insert_table_columns', 'delete_table_columns']);
const SPEC = Object.freeze({
  insert_table_rows: { axis: 'rows', sign: 1, size: 'currentRows' },
  delete_table_rows: { axis: 'rows', sign: -1, size: 'currentRows' },
  insert_table_columns: { axis: 'columns', sign: 1, size: 'currentColumns' },
  delete_table_columns: { axis: 'columns', sign: -1, size: 'currentColumns' },
});
// our implementation ceiling, not an engine limit
const MAX_COUNT = 100;
const positive = v => Number.isSafeInteger(v) && v >= 1 && v <= 1000000;

function plan(step, index) {
  const intent = step.intent;
  const spec = SPEC[intent];
  const at = `step ${index}: ${intent}`;
  if (!spec) return { ok: false, detail: `${at} is not a table-structure intent` };
  const fields = ['intent', 'id', 'sheet', 'table', 'at', 'count', spec.size];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  if (typeof step.table !== 'string' || step.table.trim() === '') return { ok: false, detail: `${at} requires the table name` };
  if (!positive(step.at)) return { ok: false, detail: `${at} requires at: a positive integer (the 1-based ${spec.axis === 'rows' ? 'data row' : 'column'} index within the table)` };
  if (!positive(step.count) || step.count > MAX_COUNT) return { ok: false, detail: `${at} requires count: an integer 1..${MAX_COUNT}` };
  if (!positive(step[spec.size])) return { ok: false, detail: `${at} requires ${spec.size}: the table's current ${spec.axis === 'rows' ? 'data row' : 'column'} count (the retry guard)` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent,
      family: 'table-structure',
      selector: Object.freeze({ sheet: step.sheet, table: step.table }),
      shift: Object.freeze({ axis: spec.axis, sign: spec.sign, at: step.at, count: step.count, size: step[spec.size] }),
      // the family owns its own cell L1 / L3; the generic cell claim is empty
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** The sheet's tables and the named one, or the named reason they cannot be told. */
function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  if (p.tier && p.tier !== 'full') return { ok: false, reason: 'xlsx-table-structure-needs-a-full-read', tier: p.tier };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const read = p.tableReads ? p.tableReads[op.selector.sheet] : undefined;
  if (read === undefined) return { ok: false, reason: 'xlsx-tables-unobservable', why: 'not-read' };
  if (typeof read === 'string') return { ok: false, reason: 'xlsx-tables-unobservable', why: read };
  if (!read || !Array.isArray(read.tables)) return { ok: false, reason: 'xlsx-tables-unobservable', why: 'malformed-read' };
  const named = read.tables.filter(t => String(t.name).toLowerCase() === op.selector.table.toLowerCase());
  if (named.length !== 1) return { ok: false, reason: named.length > 1 ? 'xlsx-table-name-ambiguous' : 'xlsx-table-not-found', available: read.tables.map(t => t.name) };
  const g = parseRange(String(named[0].ref));
  if (!g) return { ok: false, reason: 'xlsx-table-ref-unparsable', table: named[0].name, ref: named[0].ref };
  return { ok: true, tables: read.tables, table: named[0], g };
}

/** the table's current size on the step's axis (rows: the data rows under the header; columns: all columns) */
const sizeOf = (g, axis) => (axis === 'rows' ? g.r2 - g.r1 : g.c2 - g.c1 + 1);

/** the cells the engine call acts on (inserted / deleted), and the band it shifts, from the table's geometry */
function geometry(g, shift) {
  if (shift.axis === 'rows') {
    const r1 = g.r1 + shift.at; // the sheet row of data row `at`
    return { act: { r1, r2: r1 + shift.count - 1, c1: g.c1, c2: g.c2 }, band: { r1, r2: MAX_ROW, c1: g.c1, c2: g.c2 } };
  }
  const c1 = g.c1 + shift.at - 1;
  return { act: { r1: g.r1, r2: g.r2, c1, c2: c1 + shift.count - 1 }, band: { r1: g.r1, r2: g.r2, c1, c2: MAX_COL } };
}

const within = (b, r, c) => r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2;
const cellsOfSheet = (p, sheet) => Object.entries(p.cells || {}).filter(([k]) => k.startsWith(`${sheet}\u0000`)).map(([k, c]) => [k.slice(sheet.length + 1), c]);
const escapeRx = s => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** a structured reference to the table in ANY formula of the workbook (or a this-row reference, whose table name the read may drop) */
function structuredRefs(p, name) {
  const named = new RegExp(`(^|[^A-Za-z0-9_.])${escapeRx(name)}\\s*\\[`, 'i');
  return formulaEntries(p).filter(([, f]) => named.test(f) || /\[@|\[#(This Row|Headers|Data|Totals|All)\]/i.test(f)).map(([k]) => k.replace('\u0000', '!'));
}

/** a merge / another table / a pivot in the shifted band, or a read that cannot tell */
function bandConflict(p, op, t, band) {
  const sheet = op.selector.sheet;
  for (const [a, c] of cellsOfSheet(p, sheet)) {
    const g = parseRange(a);
    if (g && within(band, g.r1, g.c1) && typeof c.merge === 'string' && c.merge !== a) return { reason: 'xlsx-table-band-has-merge', cell: a, mergeArea: c.merge };
  }
  for (const o of t.tables) {
    if (o === t.table) continue;
    const og = parseRange(String(o.ref));
    if (!og) return { reason: 'xlsx-table-ref-unparsable', table: o.name, ref: o.ref };
    if (og.r2 >= band.r1 && og.r1 <= band.r2 && og.c2 >= band.c1 && og.c1 <= band.c2) return { reason: 'xlsx-table-band-has-table', table: o.name, ref: o.ref };
  }
  const objs = p.objects ? p.objects[sheet] : undefined;
  if (!objs || !Array.isArray(objs.pivots)) return { reason: 'xlsx-pivots-unobservable', why: objs ? objs.pivots : 'objects-not-read' };
  for (const pv of objs.pivots) {
    const pg = typeof pv.table === 'string' ? parseRange(pv.table.replace(/\$/g, '')) : null;
    if (!pg) return { reason: 'xlsx-pivots-unobservable', pivot: pv.name, why: pv.table };
    if (pg.r2 >= band.r1 && pg.r1 <= band.r2 && pg.c2 >= band.c1 && pg.c1 <= band.c2) return { reason: 'xlsx-table-band-has-pivot', pivot: pv.name, table: pv.table };
  }
  return null;
}

/**
 * the structure family's own dependency reader: complete, and no validation / rule / chart / name reaching the sheet (conservative).
 * The TARGET table's own name is not a dependent: every table carries one in the defined-name read (it points at its data body,
 * GH #331 T06 LIVE), and the engine resizes the table with it -- LIVE, without this the product refused every table (first T05 run).
 */
/**
 * `ignoreUnobservable(entry)`, if given, narrows ONLY the incompleteness branch below (an unobservable entry it
 * accepts no longer makes the dependency set 'incomplete' on its own) -- it never narrows the reach scan: that
 * still runs, unconditionally, on whatever refs ARE known, so a real dependent elsewhere is still caught even
 * when every unobservable gap has been forgiven (GH #335 remainder, internal tracker).
 */
function dependencyConflict(p, sheet, tableName, ignoreUnobservable) {
  const own = r => r.kind === 'name' && typeof r.name === 'string' && r.name.toUpperCase() === String(tableName).toUpperCase();
  if (!p.dependencies || !p.dependencies.complete) {
    const unobservable = p.dependencies ? p.dependencies.unobservable : null;
    const stillIncomplete = !unobservable || !ignoreUnobservable || unobservable.some(u => !ignoreUnobservable(u));
    if (stillIncomplete) {
      // GH #332/#333 (owner review): `unobservable` is the FULL list, forgiven axes included -- an axis
      // appearing there is not proof it blocked the refusal. `blocking` is the subset the ignore
      // predicate did NOT forgive (or the whole list, when there is no predicate): the actual cause.
      const blocking = Array.isArray(unobservable) ? unobservable.filter(u => !ignoreUnobservable || !ignoreUnobservable(u)) : unobservable;
      return { reason: 'xlsx-structure-dependency-observation-incomplete', unobservable: unobservable ? unobservable.slice(0, 20) : 'not-read', blocking: Array.isArray(blocking) ? blocking.slice(0, 20) : blocking };
    }
  }
  const refs = p.dependencies && Array.isArray(p.dependencies.refs) ? p.dependencies.refs : [];
  const reach = refs.filter(r => r.kind !== 'formula' && !own(r) && (r.sheets.includes(sheet) || r.host === sheet));
  return reach.length ? { reason: 'xlsx-table-structure-dependent-not-measured', dependents: reach.slice(0, 20) } : null;
}

/** a formula (outside the deleted cells) that the delete would BREAK: it overlaps the deleted cells and the whole-axis shift breaks it (conservative) */
function brokenByDelete(p, op, act) {
  const sheet = op.selector.sheet;
  const block = { sheet, ...act };
  const shift = { sheet, axis: op.shift.axis, at: op.shift.axis === 'rows' ? act.r1 : act.c1, count: -op.shift.count };
  const hits = [];
  for (const [k, f] of formulaEntries(p)) {
    const host = k.split('\u0000')[0];
    const a = parseRange(k.split('\u0000')[1]);
    if (host === sheet && a && within(act, a.r1, a.c1)) continue; // a formula inside the deleted cells goes with them
    if (referencesBlock(f, host, block) && shiftFormula(f, host, shift).broken) hits.push(k.replace('\u0000', '!'));
  }
  return hits;
}

function classify(p, op) {
  const t = gate(p, op);
  if (!t.ok) return { state: 'FAIL_CLOSED', evidence: t };
  const where = { sheet: op.selector.sheet, table: t.table.name };
  const { shift } = op;
  const now = sizeOf(t.g, shift.axis);
  if (now === shift.size + shift.sign * shift.count) return { state: 'SATISFIED', evidence: { ...where, observation: 'table-already-at-the-target-size', size: now, proof: 'fresh-live-read_tables' } };
  if (now !== shift.size) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-table-size-unexpected', expected: shift.size, current: now } };
  if (t.table.headerRow !== true) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-table-no-header-row' } };
  if (t.table.totalsRow !== false) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-table-totals-row-unmeasured' } };
  const last = shift.at + shift.count - 1;
  if (shift.sign > 0 ? shift.at > now : last > now) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-table-at-out-of-range', at: shift.at, count: shift.count, size: now } };
  if (shift.sign < 0 && shift.at === 1 && last === now) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: shift.axis === 'rows' ? 'xlsx-table-delete-all-data-rows' : 'xlsx-table-delete-all-columns' } };
  // GH #366 6th-unit follow-up (LIVE t6, prod, SMALL copy): insert_table_columns at the table's own
  // FIRST column (`at: 1`) does not grow the table -- the engine's Range.Insert('right') call acts
  // over that column's own cells, which ARE the table's top-left anchor, and the anchor-tracking
  // SHIFTS the whole table instead (ref moved right by one column, no new column, no data change).
  // `at: 1` for ROWS has no equivalent hazard: `at` addresses DATA rows only (the header, g.r1, is
  // excluded from the range entirely), so a row insert at `at:1` never touches the anchor cell. For
  // COLUMNS every column including the first (the anchor) is addressable, so this check is
  // columns-only. No clean engine path for this shape is known today -- refused named, PRE-dispatch,
  // 0 writes, until one is measured (a separate unit, not assumed here).
  if (shift.axis === 'columns' && shift.sign > 0 && shift.at === 1) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-table-insert-before-first-column-unsupported' } };
  const refs = structuredRefs(p, t.table.name);
  if (refs.length) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-table-structured-reference', formulas: refs.slice(0, 20) } };
  const { act, band } = geometry(t.g, shift);
  const b = bandConflict(p, op, t, band);
  if (b) return { state: 'FAIL_CLOSED', evidence: { ...where, ...b } };
  const d = dependencyConflict(p, op.selector.sheet, t.table.name);
  if (d) return { state: 'FAIL_CLOSED', evidence: { ...where, ...d } };
  if (shift.sign < 0) {
    const broken = brokenByDelete(p, op, act);
    if (broken.length) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-table-block-referenced', formulas: broken.slice(0, 20) } };
  }
  return { state: 'APPLY', evidence: { ...where, ref: t.table.ref, act: rangeText(act) } };
}

function resolve(p, selector, ctx = {}) {
  const t = gate(p, ctx.operation);
  if (!t.ok) return { ok: false, target: null, selector, ambiguity: t.reason, detail: t };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, table: t.table.name, ref: t.table.ref }) };
}

/** One callCommand: the exact-name sheet check, then the public Insert / Delete over the FULL-width / FULL-height range of the table. */
function lower(op, target) {
  const g = parseRange(String(target && target.ref));
  const { act } = geometry(g, op.shift);
  const range = rangeText(act);
  const call = { insert_table_rows: "Insert('down')", delete_table_rows: "Delete('up')", insert_table_columns: "Insert('right')", delete_table_columns: "Delete('left')" }[op.intent];
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(op.selector.sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.GetRange(${JSON.stringify(range)}).${call};
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'table-structure', intent: op.intent, sheet: op.selector.sheet, table: op.selector.table, range, ref: String(target.ref), axis: op.shift.axis, sign: op.shift.sign, at: op.shift.at, count: op.shift.count }),
  };
}

/** the table as the shift implies: the ref grows / shrinks on the axis, the kept columns are the same in the same order */
function expectedTable(was, shift, g) {
  const r2 = shift.axis === 'rows' ? g.r2 + shift.sign * shift.count : g.r2;
  const c2 = shift.axis === 'columns' ? g.c2 + shift.sign * shift.count : g.c2;
  const ref = rangeText({ r1: g.r1, c1: g.c1, r2, c2 });
  let kept = was.columns;
  if (shift.axis === 'columns') kept = shift.sign > 0 ? was.columns : was.columns.filter((_, i) => i < shift.at - 1 || i >= shift.at - 1 + shift.count);
  return { ref, kept };
}

function verifyPrimitive(op, before, after) {
  const b = gate(before && before.primary, op);
  const a = gate(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-table-structure-unobservable', before: b.ok ? null : b.reason, after: a.ok ? null : a.reason } };
  const exp = expectedTable(b.table, op.shift, b.g);
  const t = a.table;
  const cols = Array.isArray(t.columns) ? t.columns : [];
  const insertedCols = op.shift.axis === 'columns' && op.shift.sign > 0 ? cols.slice(op.shift.at - 1, op.shift.at - 1 + op.shift.count) : [];
  const keptNow = op.shift.axis === 'columns' && op.shift.sign > 0 ? cols.filter((_, i) => i < op.shift.at - 1 || i >= op.shift.at - 1 + op.shift.count) : cols;
  const refOk = String(t.ref).toUpperCase() === exp.ref;
  const colsOk = JSON.stringify(keptNow) === JSON.stringify(exp.kept) && (op.shift.axis !== 'columns' || op.shift.sign < 0 || insertedCols.length === op.shift.count);
  const flagsOk = t.headerRow === b.table.headerRow && t.totalsRow === b.table.totalsRow && t.autoFilter === b.table.autoFilter && JSON.stringify(t.style) === JSON.stringify(b.table.style);
  const cellsNow = bandCellsOk(before.primary, after.primary, op, b.g);
  return { ok: refOk && colsOk && flagsOk && cellsNow.ok, evidence: { sheet: op.selector.sheet, table: t.name, expectedRef: exp.ref, ref: t.ref, refOk, columnsOk: colsOk,
    newColumnHeaders: insertedCols, flagsKept: flagsOk, band: cellsNow.evidence, proof: 'fresh-live-read_tables-and-cells' } };
}

/** the band's LITERAL cells where the shift puts them, the new cells blank, a deleted cell gone */
function bandCellsOk(pb, pa, op, g) {
  const { act, band } = geometry(g, op.shift);
  const sheet = op.selector.sheet;
  const rows = op.shift.axis === 'rows';
  const wrong = [];
  const place = (r, c) => {
    if (op.shift.sign > 0) return rows ? [r + op.shift.count, c] : [r, c + op.shift.count];
    return rows ? [r - op.shift.count, c] : [r, c - op.shift.count];
  };
  for (const [a, c] of cellsOfSheet(pb, sheet)) {
    const p0 = parseRange(a);
    if (!p0 || !within(band, p0.r1, p0.c1) || c.f || c.v === null || c.v === '') continue; // only the non-blank LITERAL cells are tracked (a blank may lie outside the read area)
    if (op.shift.sign < 0 && within(act, p0.r1, p0.c1)) continue; // deleted
    const [r, cc] = place(p0.r1, p0.c1);
    const now = cellAt(pa, sheet, addr(cc, r));
    if (!now || now.f !== '' || String(now.v) !== String(c.v)) wrong.push(`${a}->${addr(cc, r)}`);
  }
  if (op.shift.sign > 0) {
    for (let r = act.r1; r <= act.r2; r += 1) for (let c = act.c1; c <= act.c2; c += 1) {
      if (!rows && r === g.r1) continue; // an inserted column's HEADER cell holds the engine's own ColumnN
      const now = cellAt(pa, sheet, addr(c, r));
      if (now && (now.f !== '' || (now.v !== null && now.v !== ''))) wrong.push(`new:${addr(c, r)}`);
    }
  }
  return { ok: wrong.length === 0, evidence: { moved: 'literal cells', wrong: wrong.slice(0, 20) } };
}

/** everything outside the band unchanged (literal cells; other sheets' formulas too), and no #REF! appears anywhere */
function verifyPreservation(op, before, after) {
  const b = gate(before && before.primary, op);
  const a = gate(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-table-preservation-unobservable', before: b.ok ? null : b.reason, after: a.ok ? null : a.reason } };
  const { band } = geometry(b.g, op.shift);
  const sheet = op.selector.sheet;
  const changed = [];
  for (const [k, c] of Object.entries(before.primary.cells || {})) {
    const [s, a1] = k.split('\u0000');
    const g = parseRange(a1);
    if (s === sheet && g && within(band, g.r1, g.c1)) continue;
    const now = after.primary.cells ? after.primary.cells[k] : undefined;
    if (!now) { if (!(c.f === '' && (c.v === null || c.v === ''))) changed.push(`${s}!${a1}:gone`); continue; } // a blank cell may fall outside the (shrunk) read area
    if (c.f === '' ? (now.f !== '' || String(now.v) !== String(c.v)) : (s !== sheet && now.f !== c.f)) changed.push(`${s}!${a1}`);
  }
  // a non-blank cell that APPEARED outside the band (a band's moved cells land inside the band) is a change too
  for (const [k, c] of Object.entries(after.primary.cells || {})) {
    if (before.primary.cells && before.primary.cells[k]) continue;
    if (c.f === '' && (c.v === null || c.v === '')) continue;
    const [s, a1] = k.split('\u0000');
    const g = parseRange(a1);
    if (!(s === sheet && g && within(band, g.r1, g.c1))) changed.push(`${s}!${a1}:appeared`);
  }
  const ref = formulaEntries(after.primary).filter(([k, f]) => /#REF!/i.test(f) || /^#REF!/i.test(String((after.primary.cells[k] || {}).v))).map(([k]) => k.replace('\u0000', '!'));
  const wasRef = new Set(formulaEntries(before.primary).filter(([, f]) => /#REF!/i.test(f)).map(([k]) => k.replace('\u0000', '!')));
  const newRef = ref.filter(x => !wasRef.has(x));
  const otherTables = list => JSON.stringify(list.filter(t => String(t.name).toLowerCase() !== op.selector.table.toLowerCase()));
  const tablesKept = otherTables(b.tables) === otherTables(a.tables);
  return { ok: changed.length === 0 && newRef.length === 0 && tablesKept, evidence: { cellsChangedOutsideTheBand: changed.slice(0, 20), newRefErrors: newRef.slice(0, 20), otherTablesKept: tablesKept } };
}

function wholeTaskChecks(op, finalObservation) {
  const t = gate(finalObservation && finalObservation.primary, op);
  const now = t.ok ? sizeOf(t.g, op.shift.axis) : null;
  const ok = t.ok && now === op.shift.size + op.shift.sign * op.shift.count;
  return [{ name: `${op.selector.sheet} table ${op.selector.table} has ${op.shift.size + op.shift.sign * op.shift.count} ${op.shift.axis === 'rows' ? 'data rows' : 'columns'}`, ok, detail: ok ? null : { reason: t.reason || 'size-differs', current: now } }];
}

const windowsOf = () => [];
const liveExtras = op => ({ tableSheets: [op.selector.sheet], objectSheets: [op.selector.sheet], dependents: true });

module.exports = { INTENTS, MAX_COUNT, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, geometry, sizeOf, cellsOfSheet, structuredRefs, bandConflict, dependencyConflict };
