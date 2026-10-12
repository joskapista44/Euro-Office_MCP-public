'use strict';

const { parseRange, cellsOf, rangeText, colNum, colName, addr } = require('../a1.cjs');
const { cellAt } = require('../inventory.cjs');
const cells = require('./cells.cjs');
const formula = require('../formula.cjs');

/**
 * XLSX `sort` family: sort_range, one key column (M6 wave 6). Audit B-D01 / D02 MIGRATE (public SetSort); the
 * source verified only that the KEY column was ordered -- not that the other columns moved with their keys (row
 * integrity) nor the row multiset; B-D03 its blank handling was asymmetric (a correct sort failed closed).
 *
 * Contract (the W6 probe: numeric-aware, stable case-insensitive, blanks last descending, a formula outside the block
 * is not re-pointed -- the refusals below stay, they are the engine-dependent shapes):
 *   refused before the write: a blank key in an ASCENDING sort (blank placement not measured), a mix of numeric and
 *   text keys, a formula in the KEY column, a formula in the block that is not a same-row formula (see sameRowShape),
 *   a merge in the block, and TEXT keys whose order differs between code-point, case-insensitive and numeric-aware
 *   comparison (so the verdict does not hang on the collation).
 *   M7 S5 R21 (owner M7 GO 2026-10-03; audit R21; W6 probe m6-w6-probe-20260929 SO, session + reopen): a DESCENDING
 *   sort with blank keys -- the blank rows go last; a formula whose every reference is a plain relative cell of its
 *   OWN row inside the block moves with its row and points at its new row (H2..H5 = G2*10..G5*10).
 * L1: the header row (hasHeaders) unchanged; the key column ordered (blanks last); the ROW MULTISET preserved (each
 * row's cells moved together; a same-row formula compared by its row-free shape, not by its value -- formula values
 * are stale in session, W6 probe). L3: every other cell (a formula outside referencing the block is compared by its
 * text -- its value changes with the sort by design). Retry: the key already ordered -> SATISFIED, 0 write / 0 command.
 */

const INTENTS = Object.freeze(['sort_range']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'range', 'key', 'order', 'hasHeaders']);
const MAX_CELLS = 1000;

function plan(step, index) {
  const at = `step ${index}: sort_range`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `${at} requires a plain A1 range` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `${at} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  if (typeof step.key !== 'string' || !/^[A-Z]{1,3}$/.test(step.key) || colNum(step.key) < g.c1 || colNum(step.key) > g.c2) return { ok: false, detail: `${at}: key must be a column letter inside the range` };
  if (!['asc', 'desc'].includes(step.order)) return { ok: false, detail: `${at}: order must be 'asc' or 'desc'` };
  if (typeof step.hasHeaders !== 'boolean') return { ok: false, detail: `${at}: hasHeaders must be true or false (no default -- a guessed header row is a sorted data row)` };
  const first = g.r1 + (step.hasHeaders ? 1 : 0);
  if (g.r2 - first < 1) return { ok: false, detail: `${at}: fewer than two data rows` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'sort_range',
      family: 'sort',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      grid: Object.freeze({ c1: g.c1, c2: g.c2, r1: g.r1, r2: g.r2, first }),
      key: colNum(step.key),
      order: step.order,
      hasHeaders: step.hasHeaders,
      cells: Object.freeze(cellsOf(g).map(a => Object.freeze({ a }))),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const cmpCode = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
const cmpCase = (a, b) => cmpCode(a.toLowerCase(), b.toLowerCase());
const cmpNum = (a, b) => a.localeCompare(b, 'en', { numeric: true, sensitivity: 'base' });

/** The data rows as tuples of { v, f } and their keys; or why the block cannot be sorted under the contract. */
function rowsOf(p, op) {
  const { sheet } = op.selector;
  const rows = [];
  for (let r = op.grid.first; r <= op.grid.r2; r += 1) {
    const tuple = [];
    for (let c = op.grid.c1; c <= op.grid.c2; c += 1) {
      const e = entryOf(cellAt(p, sheet, addr(c, r)), c, r, op);
      if (e.problem) return { problem: e.problem, cell: addr(c, r) };
      tuple.push(e);
    }
    rows.push(tuple);
  }
  return { rows };
}

/** One block cell as a row-tuple entry { v, n } -- a same-row formula by its row-free shape -- or why it cannot sort. */
function entryOf(cell, c, r, op) {
  if (!cell) return { problem: 'xlsx-sort-cell-unread' };
  if (cell.merge && cell.merge !== addr(c, r)) return { problem: 'xlsx-sort-merge-in-block' };
  if (cell.f !== '') {
    if (c === op.key) return { problem: 'xlsx-sort-formula-key-unmeasured' };
    const shape = sameRowShape(cell.f, r, op.grid);
    return shape === null ? { problem: 'xlsx-sort-formula-in-block-unmeasured' } : { v: shape, n: false };
  }
  if (cell.count === null) return { problem: 'xlsx-cell-type-unreadable' };
  return { v: cell.v === null ? '' : String(cell.v), n: cell.count === 1 };
}

/**
 * M7 S5 R21: a formula's row-free shape when EVERY reference is a plain relative single cell of its own row `r` inside
 * the block's columns (the measured shape: moves with its row, points at its new row); otherwise null. Refused
 * (null): '$', ranges, sheet references, another row, a column outside the block, any other identifier (a defined
 * name), an error literal. String literals are kept verbatim.
 */
function sameRowShape(f, r, grid) {
  const c = formula.canonical(f);
  const toks = c === null ? null : formula.tokenize(c.slice(1));
  if (!toks || toks.some(t => t.kind === 'qname')) return null;
  let ok = true;
  const body = toks.map(t => (t.kind === 'str' ? t.text : t.text.replace(/(?<![0-9.])[A-Z_][A-Z0-9_.]*(\(?)|[$!:#[]/g, (m, call) => {
    if (call || m === 'TRUE' || m === 'FALSE') return m;
    const ref = /^([A-Z]{1,3})(\d+)$/.exec(m);
    if (!ref || Number(ref[2]) !== r || colNum(ref[1]) < grid.c1 || colNum(ref[1]) > grid.c2) { ok = false; return m; }
    return `${ref[1]}@`;
  }))).join('');
  return ok ? `=${body}` : null;
}

/** The comparator for this block's keys, or why the collation would be a guess. */
function comparatorOf(rows, op) {
  const k = op.key - op.grid.c1;
  const all = rows.map(t => t[k]);
  if (op.order === 'asc' && all.some(x => x.v === '')) return { problem: 'xlsx-sort-blank-key-unmeasured' };
  const keys = all.filter(x => x.v !== '');
  const numeric = keys.every(x => x.n);
  if (!numeric && keys.some(x => x.n)) return { problem: 'xlsx-sort-mixed-key-types-unmeasured' };
  if (numeric) return { cmp: (a, b) => Number(a) - Number(b) };
  const vals = [...new Set(keys.map(x => x.v))];
  const orderBy = f => vals.slice().sort(f).join('\u0000');
  const code = orderBy(cmpCode);
  if (orderBy(cmpCase) !== code || orderBy(cmpNum) !== code || vals.some(v => /[^\x20-\x7e]/.test(v))) return { problem: 'xlsx-sort-collation-unmeasured', keys: vals.slice(0, 20) };
  return { cmp: cmpCode };
}

/** The key ordered; a blank key (desc only, R21) after every non-blank one. */
function ordered(rows, op, cmp) {
  const k = op.key - op.grid.c1;
  return rows.every((t, i) => {
    if (i === 0) return true;
    const [x, y] = [rows[i - 1][k].v, t[k].v];
    if (x === '' || y === '') return y === '';
    return op.order === 'asc' ? cmp(x, y) <= 0 : cmp(x, y) >= 0;
  });
}
const multiset = rows => rows.map(t => JSON.stringify(t.map(x => x.v))).sort().join('\n');

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const r = rowsOf(p, op);
  if (r.problem) return { ok: false, reason: r.problem, cell: r.cell };
  const c = comparatorOf(r.rows, op);
  if (c.problem) return { ok: false, reason: c.problem, keys: c.keys };
  return { ok: true, rows: r.rows, cmp: c.cmp };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, range: op.selector.range, key: colName(op.key), order: op.order };
  if (ordered(g.rows, op, g.cmp)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-key-already-ordered' } };
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range }) };
}

/** One callCommand: exact-name sheet check, then the public SetSort over the block (the source's measured call). */
function lower(op) {
  const { sheet, range } = op.selector;
  const keyRange = `${colName(op.key)}${op.grid.r1}:${colName(op.key)}${op.grid.r2}`;
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.GetRange(${JSON.stringify(range)}).SetSort(ws.GetRange(${JSON.stringify(keyRange)}), ${JSON.stringify(op.order === 'asc' ? 'xlAscending' : 'xlDescending')}, null, null, null, null, ${JSON.stringify(op.hasHeaders ? 'xlYes' : 'xlNo')}, 'xlSortColumns');
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'sort', sheet, range, key: op.key, order: op.order, hasHeaders: op.hasHeaders, grid: op.grid }),
  };
}

/** L1: header unchanged, key ordered, the row multiset preserved (rows moved whole). */
function verifyPrimitive(op, before, after) {
  const b = gate(before && before.primary, op);
  const a = gate(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-sort-block-unreadable', before: b.ok || b.reason, after: a.ok || a.reason } };
  const headerKept = !op.hasHeaders || cellsOf({ c1: op.grid.c1, c2: op.grid.c2, r1: op.grid.r1, r2: op.grid.r1 }).every(x => {
    const w = cellAt(before.primary, op.selector.sheet, x);
    const n = cellAt(after.primary, op.selector.sheet, x);
    return !!w && !!n && w.f === n.f && JSON.stringify(w.v) === JSON.stringify(n.v);
  });
  const isOrdered = ordered(a.rows, op, b.cmp);
  const rowsKept = multiset(a.rows) === multiset(b.rows);
  return { ok: headerKept && isOrdered && rowsKept, evidence: { sheet: op.selector.sheet, range: op.selector.range, headerKept, ordered: isOrdered, rowsKept, proof: 'fresh-live-rows-and-order' } };
}

/** L3: every other cell (the block is L1's); a formula elsewhere is compared by its text. */
const verifyPreservation = (op, before, after) => cells.verifyPreservation(op, before, after);

function wholeTaskChecks(op, finalObservation) {
  const g = gate(finalObservation && finalObservation.primary, op);
  const ok = g.ok && ordered(g.rows, op, g.cmp);
  return [{ name: `${op.selector.sheet}!${op.selector.range} ordered by ${colName(op.key)} ${op.order}`, ok, detail: ok ? null : { reason: g.reason || 'not-ordered' } }];
}

const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.range }];

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf };
