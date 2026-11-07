'use strict';

const { parseRange, rangeText, cellsOf, overlaps } = require('../a1.cjs');
const { cellAt, isBlank } = require('../inventory.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `tables` write family: create_table (GH #298 D, owner decision 2026-10-27 "3: A" -- read_tables first, then
 * create_table; read_tables is released with the cell Bridge /18). The verifier is read_tables' own observation
 * (`tableReads`, families/tables-read.cjs), reused unchanged.
 *
 * ENGINE PATH: the PUBLIC `ApiWorksheet.FormatAsTable(sRange)` through callCommand. Measured LIVE (GH #298 matrix,
 * docs/GH298-ST-LIVE-MATRIX-20261025.md, ST-03 / ST-12, session + reopen + OOXML): it creates one Excel Table over the
 * range -- fixed style TableStyleLight9, the FIRST row always the header, an automatic name, no return value; on the
 * same range twice it is a silent no-op. The engine does NOT refuse what Excel refuses, so these are refused here,
 * before the write, each named:
 *   overlap with an existing table      -> xlsx-table-overlaps-table   (ST-12: the engine stores an overlapping table)
 *   a merge inside the range            -> xlsx-table-over-merge       (ST-12: accepted by the engine, Excel refuses it)
 *   overlap with a pivot table          -> xlsx-table-overlaps-pivot
 *   a single row                        -> plan: at least a header and one data row (ST-12: one row becomes a 2-row table)
 *   a blank / formula / duplicate header cell -> xlsx-table-header-*  (a table column needs a unique, literal name;
 *                                          a formula header was never measured and Excel turns it into a constant)
 * Name and style are not parameters in v1: the public call has neither; they are the table-edit writes (option E).
 *
 * Retry = the terminal state on a fresh read: a table whose ref IS the range, its header row on and its columns the
 * header texts -> SATISFIED, 0 write.
 */

const INTENTS = Object.freeze(['create_table']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'range']);
// our implementation ceiling (every cell is read for the merge / header gates), the same as write_range's
const MAX_CELLS = 1000;

function plan(step, index) {
  const at = `step ${index}: create_table`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `${at} requires a plain A1 range` };
  if (g.r2 - g.r1 < 1) return { ok: false, detail: `${at} range needs a header row and at least one data row` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `${at} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'create_table',
      family: 'tables',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      // a table owns no cell content: every cell's value and formula is preserved, not claimed
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** The sheet's tables (read_tables' own `tableReads`, DS Bridge /18) and pivots, or the named reason they cannot be told. */
function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const read = p.tableReads ? p.tableReads[op.selector.sheet] : undefined;
  if (read === undefined) return { ok: false, reason: 'xlsx-tables-unobservable', why: 'not-read' };
  if (typeof read === 'string') return { ok: false, reason: 'xlsx-tables-unobservable', why: read };
  if (!read || !Array.isArray(read.tables)) return { ok: false, reason: 'xlsx-tables-unobservable', why: 'malformed-read' };
  return { ok: true, tables: read.tables, own: read.tables.find(t => String(t.ref).toUpperCase() === op.selector.range) || null };
}

/** the header texts of the range's first row, or the named reason they are not a valid set of column names */
function headerOf(p, op) {
  const g = parseRange(op.selector.range);
  const names = [];
  for (const a of cellsOf({ ...g, r2: g.r1 })) {
    const cell = cellAt(p, op.selector.sheet, a);
    if (!cell) return { ok: false, reason: 'xlsx-target-cell-unread', cell: a };
    if (isBlank(cell)) return { ok: false, reason: 'xlsx-table-header-blank', cell: a };
    if (cell.f !== '') return { ok: false, reason: 'xlsx-table-header-formula-unmeasured', cell: a };
    const name = String(cell.v);
    if (name.trim() === '') return { ok: false, reason: 'xlsx-table-header-blank', cell: a };
    if (names.some(n => n.toLowerCase() === name.toLowerCase())) return { ok: false, reason: 'xlsx-table-header-duplicate', cell: a, name };
    names.push(name);
  }
  return { ok: true, names };
}

/** a table over the range */
function tableConflict(tables, g) {
  const hit = tables.find(t => { const tg = parseRange(String(t.ref)); return !tg || overlaps(tg, g); });
  return hit ? { reason: 'xlsx-table-overlaps-table', table: hit.name, ref: hit.ref } : null;
}

/** a merge inside the range (a cell whose merge area is not itself), or a cell not read */
function mergeConflict(p, sheet, g) {
  for (const a of cellsOf(g)) {
    const cell = cellAt(p, sheet, a);
    if (!cell || typeof cell.merge !== 'string') return { reason: 'xlsx-target-cell-unread', cell: a };
    if (cell.merge !== a) return { reason: 'xlsx-table-over-merge', cell: a, mergeArea: cell.merge };
  }
  return null;
}

/** a pivot table over the range, or pivots that cannot be read */
function pivotConflict(p, sheet, g) {
  const o = p.objects ? p.objects[sheet] : undefined;
  if (!o || !Array.isArray(o.pivots)) return { reason: 'xlsx-pivots-unobservable', why: o ? o.pivots : 'objects-not-read' };
  for (const pv of o.pivots) {
    const pg = typeof pv.table === 'string' ? parseRange(pv.table.replace(/\$/g, '')) : null;
    if (!pg) return { reason: 'xlsx-pivots-unobservable', pivot: pv.name, why: pv.table };
    if (overlaps(pg, g)) return { reason: 'xlsx-table-overlaps-pivot', pivot: pv.name, table: pv.table };
  }
  return null;
}

/** a table over the range, a merge inside it, or a pivot over it -- each refused, named */
function conflictOf(p, op, tables) {
  const g = parseRange(op.selector.range);
  return tableConflict(tables, g) || mergeConflict(p, op.selector.sheet, g) || pivotConflict(p, op.selector.sheet, g);
}

/** exactly the created table: its ref is the range, the header row on, its columns the header texts */
const isCreated = (t, names) => !!t && t.headerRow === true && Array.isArray(t.columns) && JSON.stringify(t.columns) === JSON.stringify(names);

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, range: op.selector.range };
  const h = headerOf(p, op);
  if (g.own) {
    if (h.ok && isCreated(g.own, h.names)) return { state: 'SATISFIED', evidence: { ...where, table: g.own.name, proof: 'fresh-live-a-table-with-this-range' } };
    return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-table-range-holds-a-different-table', table: g.own.name } };
  }
  const c = conflictOf(p, op, g.tables);
  if (c) return { state: 'FAIL_CLOSED', evidence: { ...where, ...c } };
  if (!h.ok) return { state: 'FAIL_CLOSED', evidence: { ...where, ...h } };
  return { state: 'APPLY', evidence: { ...where, columns: h.names } };
}

function resolve(p, selector, ctx = {}) {
  const op = ctx.operation;
  const g = gate(p, op);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  const h = headerOf(p, op);
  if (!h.ok) return { ok: false, target: null, selector, ambiguity: h.reason, detail: h };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range, columns: Object.freeze(h.names) }) };
}

/** One callCommand: exact-name sheet check, then the public FormatAsTable. */
function lower(op) {
  const { sheet, range } = op.selector;
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.FormatAsTable(${JSON.stringify(range)});
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'tables', action: 'create', sheet, range }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason, why: g.why || null } };
  const h = headerOf(after.primary, op);
  const ok = h.ok && isCreated(g.own, h.names);
  return { ok, evidence: { sheet: op.selector.sheet, range: op.selector.range, found: !!g.own, table: g.own ? g.own.name : null,
    headerRow: g.own ? g.own.headerRow : null, columns: g.own ? g.own.columns : null, proof: 'fresh-live-read_tables' } };
}

const sameTable = (x, y) => JSON.stringify(x) === JSON.stringify(y);

/** every OTHER table of the sheet unchanged, and every cell's value and formula (a table owns none) */
function verifyPreservation(op, before, after) {
  const b = gate(before && before.primary, op);
  const a = gate(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-table-preservation-unobservable', before: b.ok ? null : b.reason, after: a.ok ? null : a.reason } };
  const others = list => list.filter(t => String(t.ref).toUpperCase() !== op.selector.range);
  const was = others(b.tables);
  const now = others(a.tables);
  const changed = was.filter(t => !now.some(x => sameTable(x, t))).map(t => t.name);
  const added = now.filter(t => !was.some(x => x.name === t.name)).map(t => t.name);
  const rest = cells.verifyPreservation(op, before, after);
  return { ok: changed.length === 0 && added.length === 0 && now.length === was.length && rest.ok, evidence: { ...rest.evidence, tablesCompared: was.length, tablesChanged: changed.slice(0, 20), tablesAdded: added.slice(0, 20) } };
}

function wholeTaskChecks(op, finalObservation) {
  const p = finalObservation && finalObservation.primary;
  const g = gate(p, op);
  const h = g.ok ? headerOf(p, op) : { ok: false };
  const ok = g.ok && h.ok && isCreated(g.own, h.names);
  return [{ name: `${op.selector.sheet}!${op.selector.range} is one table with its header row`, ok, detail: ok ? null : { reason: g.reason || h.reason || 'not-the-expected-state' } }];
}

const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.range }];
const liveExtras = op => ({ tableSheets: [op.selector.sheet], objectSheets: [op.selector.sheet] });

module.exports = { INTENTS, MAX_CELLS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
