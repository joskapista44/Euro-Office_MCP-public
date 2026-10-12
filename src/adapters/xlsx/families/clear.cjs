'use strict';

const { parseRange, cellsOf, rangeText } = require('../a1.cjs');
const { cellAt } = require('../inventory.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `clear` family: clear_range -- CONTENT ONLY (M6 wave 1b).
 *
 * The primitive is MEASURED, not assumed (wave-1b LIVE probe, run m6-w1b-1790580791644-57a127 at 594439a,
 * docs/evidence/M6-W1B-CLEAR-PROBE.md): on a cell seeded with text + number format + fill,
 *   ApiRange.Clear()          -> value gone AND number format / fill reset
 *   ApiRange.ClearContents()  -> value gone AND number format / fill reset
 *   ApiRange.SetValue('')     -> value gone, number format and fill KEPT (same session and reopened)
 * Owner rule (W1b GO): a call that removes formatting it was not asked to remove is not a
 * content-only clear. So the write is SetValue('') per cell. The source used Clear() (B-C11: its
 * formatting effect was NOT_MEASURED) -- it silently resets formatting.
 *
 * Proof, per target cell, on a fresh read: blank (value '', no formula, COUNT 0) AND the number
 * format and fill read before are unchanged (L1).
 *
 * CELL-level BOLD (owner decision, owner channel 2026-09-28; the reader is EO_INTERNAL_DEPENDENCY READ,
 * build-gated in the driver):
 *   - not observable BEFORE the write on any target cell -> refused before dispatch (0 write);
 *   - PRESERVATION (L3) compares it per target cell: an observed change -> FAIL; unobservable after
 *     the write -> not proven. Both leave the core's uncertain outcome (INTEGRAL_PRESERVATION_NOT_PROVEN):
 *     never APPLIED, never an automatic rewrite. null / a missing getter / a read error is never false.
 * RICH-TEXT (character-level) formatting belongs to the deleted text; its disappearance with the content
 * is not a side effect and is not compared -- the cell's own formatting is what must stay.
 * A formula cell is cleared too (clear means the content, a formula included); a merged area is
 * refused, as in write_range. The retry is the terminal state: every cell already blank = SATISFIED.
 */

const MAX_CELLS = 1000;
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'range']);

function plan(step, index) {
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `step ${index}: clear_range does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `step ${index}: clear_range requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `step ${index}: clear_range requires a plain A1 range on the named sheet (no Sheet!A1)` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `step ${index}: clear_range range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'clear_range',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      cells: Object.freeze(cellsOf(g).map(a => Object.freeze({ a }))),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** Blank on a fresh read: no value, no formula, not a number. null = COUNT unread. */
function isCleared(cell) {
  if (!cell) return false;
  if (cell.count === null) return null;
  return cell.f === '' && (cell.v === null || cell.v === '') && cell.count === 0;
}

function gate(primary, op) {
  if (!primary || primary.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!primary.complete) return { ok: false, reason: primary.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: primary.unsupported };
  const s = primary.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: primary.sheets.map(x => x.name) };
  const found = {};
  for (const { a } of op.cells) {
    const cell = cellAt(primary, op.selector.sheet, a);
    if (!cell) return { ok: false, reason: 'xlsx-target-cell-unread', cell: a };
    if (cell.merge && cell.merge !== a) return { ok: false, reason: 'xlsx-clear-in-merged-area-refused', cell: a, mergeArea: cell.merge };
    if (!('nf' in cell) || !('fill' in cell)) return { ok: false, reason: 'xlsx-target-format-unread', cell: a };
    if (typeof cell.bold !== 'boolean') return { ok: false, reason: 'xlsx-target-bold-unobservable', cell: a, bold: 'bold' in cell ? cell.bold : 'UNOBSERVABLE:not-read' };
    found[a] = cell;
  }
  return { ok: true, cells: found };
}

function classify(primary, op) {
  const g = gate(primary, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const verdicts = op.cells.map(({ a }) => isCleared(g.cells[a]));
  if (verdicts.includes(null)) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-cell-type-unreadable' } };
  if (verdicts.every(Boolean)) return { state: 'SATISFIED', evidence: { sheet: op.selector.sheet, range: op.selector.range, proof: 'fresh-live-cells-already-blank' } };
  return { state: 'APPLY', evidence: { sheet: op.selector.sheet, range: op.selector.range, nonBlank: verdicts.filter(x => !x).length } };
}

function resolve(primary, selector, ctx = {}) {
  const g = gate(primary, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range, before: g.cells }) };
}

/** One callCommand: an exact-name sheet check, then SetValue('') per cell (the measured content-only path). */
function lower(op) {
  const sheet = op.selector.sheet;
  const addrs = op.cells.map(({ a }) => a);
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const plan = ${JSON.stringify(addrs)};
  for (const a of plan) ws.GetRange(a).SetValue('');
  return { ok: true, cleared: plan.length };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'clear', action: 'clear-content', sheet, cells: Object.freeze(addrs) }),
  };
}

/** L1: every target cell blank AND its number format and fill equal to the read before the write. */
function verifyPrimitive(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  const notBlank = [];
  const formatChanged = [];
  for (const { a: addr } of op.cells) {
    const now = cellAt(a, op.selector.sheet, addr);
    if (isCleared(now) !== true) { notBlank.push(addr); continue; }
    const was = cellAt(b, op.selector.sheet, addr);
    if (!was || !('nf' in was) || !('nf' in now) || was.nf !== now.nf || JSON.stringify(was.fill) !== JSON.stringify(now.fill)) formatChanged.push(addr);
  }
  return {
    ok: notBlank.length === 0 && formatChanged.length === 0,
    evidence: { sheet: op.selector.sheet, range: op.selector.range, cells: op.cells.length, notBlank, formatChanged, proof: 'fresh-live-blank-and-format-kept' },
  };
}

/**
 * L3: nothing outside the target changed (as write_range) AND every target cell keeps its cell-level
 * bold. An observed change is a FAIL; a bold that cannot be read after the write leaves preservation
 * NOT proven -- both keep the operation out of APPLIED (the core's uncertain rule).
 */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  const b = before && before.primary;
  const a = after && after.primary;
  const boldChanged = [];
  const boldUnobservable = [];
  for (const { a: addr } of op.cells) {
    const was = cellAt(b, op.selector.sheet, addr);
    const now = cellAt(a, op.selector.sheet, addr);
    if (!was || typeof was.bold !== 'boolean' || !now || typeof now.bold !== 'boolean') { boldUnobservable.push({ cell: addr, before: was ? was.bold : null, after: now ? now.bold : null }); continue; }
    if (was.bold !== now.bold) boldChanged.push({ cell: addr, before: was.bold, after: now.bold });
  }
  const ok = rest.ok && boldChanged.length === 0 && boldUnobservable.length === 0;
  return { ok, evidence: { ...rest.evidence, targetBold: { compared: op.cells.length, changed: boldChanged, unobservable: boldUnobservable, dependency: 'EO_INTERNAL_DEPENDENCY READ' } } };
}

function wholeTaskChecks(op, finalObservation) {
  const bad = op.cells.filter(({ a }) => isCleared(cellAt(finalObservation && finalObservation.primary, op.selector.sheet, a)) !== true).map(({ a }) => a);
  return [{ name: `every cell of ${op.selector.sheet}!${op.selector.range} is blank`, ok: bad.length === 0, detail: bad.length ? { notBlank: bad } : null }];
}

/** The target window, read WITH number format and fill (the proof needs both, before and after). */
const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.range, format: true }];

module.exports = { plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, isCleared, MAX_CELLS };
