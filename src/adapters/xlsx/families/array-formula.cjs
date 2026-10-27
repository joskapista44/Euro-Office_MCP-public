'use strict';

const { parseRange, cellsOf, rangeText, MAX_COL, MAX_ROW } = require('../a1.cjs');
const { cellAt, isBlank } = require('../inventory.cjs');
const cells = require('./cells.cjs');
const formula = require('../formula.cjs');

/**
 * XLSX `array-formula` family: write_array_formula (GH #301 D4, public SetFormulaArray).
 *
 * The LIVE design probe (test/live/xlsx-gh301-d4-array-formula-probe.live.cjs, docs/evidence/gh301-d4-array-formula-
 * probe-20261025.json) measured what the public API gives a verifier, and every rule below answers one finding:
 *   - GetFormulaArray is NOT an extent reader: on any cell or part of an array it returns the array's formula, on a
 *     superset null, on a plain formula null, on a value the value, on an empty cell ''. So a cell is an array MEMBER
 *     when GetFormulaArray starts with '=' AND its own GetFormula does too; the extent is proven by the area's every
 *     cell being a member while no cell of its ONE-CELL RING is.
 *   - SetFormulaArray silently overwrites values, silently builds OVERLAPPING arrays (two <f t="array"> refs over the
 *     same cells in the saved file) and a later single-cell write breaks an array in the saved file. So the write is
 *     refused before it runs unless every area cell is EMPTY and no cell of the area or its ring is a member.
 *   - In session the computed values stay '' until a reopen (the saved file carries them): L1 compares formulas, never
 *     values.
 * Refused before the write (named): a single cell (not measured), a merge in the area, an unread / unobservable cell,
 * a non-empty area cell, an array member in the area or the ring, a protected sheet (protection-gate's default).
 * L1: every area cell a member whose formula is the written one (formula.sameFormula), no ring cell a member.
 * L3: every other cell unchanged (cells.verifyPreservation; the area cells are the op's own). Retry: the terminal state
 * on a fresh read -> SATISFIED, 0 write / 0 command. The release referee is the saved <f t="array" ref="<area>">.
 */

const INTENTS = Object.freeze(['write_array_formula']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'range', 'formula']);
// our implementation ceiling, not an engine limit: the probe measured 3- and 4-cell areas only
const MAX_CELLS = 100;

function plan(step, index) {
  const at = `step ${index}: write_array_formula`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `${at} requires a plain A1 range` };
  if (g.count < 2) return { ok: false, detail: `${at}: the range must span at least two cells (a single-cell array formula is not measured)` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `${at} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  const problem = cells.formulaProblem(step.formula);
  if (problem) return { ok: false, detail: `${at}: ${problem}` };
  const ring = { c1: Math.max(1, g.c1 - 1), c2: Math.min(MAX_COL, g.c2 + 1), r1: Math.max(1, g.r1 - 1), r2: Math.min(MAX_ROW, g.r2 + 1) };
  const inArea = a => cellsOf(g).includes(a);
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'write_array_formula',
      family: 'array-formula',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      formula: step.formula,
      ring: Object.freeze(cellsOf(ring).filter(a => !inArea(a))),
      outer: rangeText(ring),
      cells: Object.freeze(cellsOf(g).map(a => Object.freeze({ a }))),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** One cell's array membership: true | false, or the reason it cannot be told. */
function memberOf(p, sheet, a) {
  const cell = cellAt(p, sheet, a);
  if (!cell) return { problem: 'xlsx-array-formula-cell-unread' };
  const raw = p.arrayReads ? p.arrayReads[`${sheet}!${a}`] : undefined;
  if (raw === undefined) return { problem: 'xlsx-array-formula-unobservable', why: 'not-read' };
  if (typeof raw === 'string' && raw.startsWith('UNOBSERVABLE')) return { problem: 'xlsx-array-formula-unobservable', why: raw };
  const member = typeof raw === 'string' && raw.startsWith('=') && cell.f !== '';
  return { cell, member, raw: member ? raw : null };
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const { sheet } = op.selector;
  const area = [];
  for (const { a } of op.cells) {
    const m = memberOf(p, sheet, a);
    if (m.problem) return { ok: false, reason: m.problem, cell: a, why: m.why };
    if (m.cell.merge && m.cell.merge !== a) return { ok: false, reason: 'xlsx-array-formula-merge-in-area', cell: a };
    area.push({ a, ...m });
  }
  const ringMembers = [];
  for (const a of op.ring) {
    const m = memberOf(p, sheet, a);
    if (m.problem) return { ok: false, reason: m.problem, cell: a, why: m.why };
    if (m.member) ringMembers.push(a);
  }
  return { ok: true, area, ringMembers };
}

/** the written array exactly over the area: every area cell a member with the written formula, no ring member */
const written = (g, op) => g.ringMembers.length === 0 && g.area.every(x => x.member && formula.sameFormula(op.formula, x.raw));

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, range: op.selector.range };
  if (written(g, op)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-every-area-cell-a-member-no-ring-member' } };
  const touching = [...g.area.filter(x => x.member).map(x => x.a), ...g.ringMembers];
  if (touching.length) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-array-formula-touches-an-existing-array', cells: touching.slice(0, 20) } };
  const filled = g.area.filter(x => !isBlank(x.cell)).map(x => x.a);
  if (filled.length) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-array-formula-area-not-empty', cells: filled.slice(0, 20) } };
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range }) };
}

/** One callCommand: exact-name sheet check, then the public SetFormulaArray over the area. */
function lower(op) {
  const { sheet, range } = op.selector;
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.GetRange(${JSON.stringify(range)}).SetFormulaArray(${JSON.stringify(op.formula)});
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'array-formula', sheet, range, formula: op.formula }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason, cell: g.cell || null, why: g.why || null } };
  const ok = written(g, op);
  return { ok, evidence: { sheet: op.selector.sheet, range: op.selector.range, members: g.area.filter(x => x.member).length, cells: g.area.length, ringMembers: g.ringMembers.slice(0, 20), mismatched: g.area.filter(x => !x.member || !formula.sameFormula(op.formula, x.raw)).map(x => x.a).slice(0, 20), proof: 'fresh-live-every-area-cell-a-member-no-ring-member' } };
}

function verifyPreservation(op, before, after) {
  return cells.verifyPreservation(op, before, after);
}

function wholeTaskChecks(op, finalObservation) {
  const g = gate(finalObservation && finalObservation.primary, op);
  const ok = g.ok && written(g, op);
  return [{ name: `${op.selector.sheet}!${op.selector.range} array formula ${op.formula}`, ok, detail: ok ? null : { reason: g.reason || 'not-the-written-array' } }];
}

const windowsOf = op => [{ sheet: op.selector.sheet, range: op.outer }];
const liveExtras = op => ({ arrayCells: [...op.cells.map(c => c.a), ...op.ring].map(cell => ({ sheet: op.selector.sheet, cell })) });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
