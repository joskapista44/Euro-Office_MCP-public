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
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'range', 'scope']);

/**
 * L-14 / L-17 (standing release GO, internal tracker; LIVE-measured test/live/xlsx-l14-l17-clear-formats-probe.live.cjs, 1x3 + 2x2 merge,
 * session and reopen): `scope` picks what the clear removes.
 *   'contents' (the default, the code above UNCHANGED): the value only -- SetValue('') per cell, number format / fill / bold kept.
 *   'formats'  : range.ClearFormats() -- every format axis back to its base, the VALUE (and formula) KEPT.
 *   'all'      : range.Clear() -- value AND every format axis. (ClearContents() is NOT used: LIVE it equals Clear(), so it can never be the
 *                'contents' primitive.)
 * Both format scopes DISSOLVE a merge LIVE, so a range that touches any merged area (the anchor included) is refused by name before any
 * dispatch; re-merging is not part of this step. The base axes asserted after the write are the ones with an unambiguous reader.
 *
 * GH #141 P1 (clear-formats-l1-teljes-alap, 2026-10-17): the base axes asserted were ONLY the six
 * that existed when this file was written (nf/fill/bold/italic/wrap/rot) -- border, align.{h,v},
 * readingOrder/indent/shrink and font.{underline,strike} were readable (the family's own windowsOf
 * already asks for `style: true`, which the driver answers with ALL of these) but never compared
 * back to their own base after a 'formats'/'all' clear. Their "never touched" base values are the
 * SAME ones the family's own font/align/border readers already document and every existing test
 * fixture (test/fakes/fake-xlsx-driver.cjs styleOf()) already encodes, cited per-axis below; a LIVE
 * re-measurement of ClearFormats()/Clear() ITSELF against a cell carrying these axes is a SEPARATE,
 * follow-up LIVE round (the existing release-l14-l17-clear-scope-product.live.cjs harness only
 * seeds bold/italic) -- this file's own fail-closed shape (axesOffBase -> APPLY, never SATISFIED,
 * on a regression) means an unconfirmed axis fails SAFE: it would show as "still not at base"
 * rather than a silent false pass.
 * font.name / font.size stay OUT of BASE_AXES on purpose: the documented base ('Calibri', 11) is
 * THIS fixture's workbook default font, not a fixed engine constant (unlike wrap/rot/border/align/
 * readingOrder/indent/shrink, which are ECMA-376 cell-format defaults independent of any theme) --
 * asserting a hardcoded name/size after a clear would be wrong on a workbook with a different
 * default font. NEM-MERT, not silently assumed.
 */
const SCOPES = Object.freeze(['contents', 'formats', 'all']);
const BASE_NF = 'General';
const noFill = fill => fill === null || fill === undefined || /^S:No Fill$/i.test(String(fill));
/** border at base: all four sides + the diagonal read {s:0,c:0}, neither diagonal direction flag set
 *  (m6-w4c-internal-align-border-probe-20260928.json I2: a never-touched side reads exactly this). */
const BASE_BORDER_SIDE = side => !!side && typeof side === 'object' && side.s === 0 && side.c === 0;
const isBaseBorder = b => !!b && typeof b === 'object' && ['t', 'b', 'l', 'r', 'd'].every(k => BASE_BORDER_SIDE(b[k])) && b.dd === false && b.du === false;
/** align at base: horizontal unset (null, "General"), vertical the measured default 'bottom' (0) --
 *  m7-s31c-probe1-20261005.json / the m6-w4c I2 control cell, both MEASURED. */
const isBaseAlign = al => !!al && typeof al === 'object' && al.h === null && al.v === 0;
/** font.underline/strike at base: underline has no boolean off-state, its untouched/off reader is the
 *  enum string 'xlUnderlineStyleNone' (X-F05, m7-s21-probe1); strike is null (ambiguous with an
 *  explicit false, the same ambiguity bold/italic's own reader has) -- both MEASURED, not assumed. */
const isBaseFontEffects = f => !!f && typeof f === 'object' && f.underline === 'xlUnderlineStyleNone' && f.strike === null;
const BASE_AXES = Object.freeze([
  ['nf', c => c.nf, v => v === BASE_NF],
  ['fill', c => c.fill, v => noFill(v)],
  ['bold', c => c.bold, v => v === false],
  ['italic', c => c.italic, v => v === false],
  ['wrap', c => c.wrap, v => v === false],
  ['rot', c => c.rot, v => v === 0],
  // GH #141 P1 additions, all gated the SAME way as the original six: an 'UNOBSERVABLE:*' value on
  // the read makes the whole axis unread (axesOffBase below), never silently treated as "at base".
  ['align', c => c.align, v => isBaseAlign(v)],
  ['border', c => c.border, v => isBaseBorder(v)],
  ['readingOrder', c => c.readingOrder, v => v === 0],
  ['indent', c => c.indent, v => v === 0],
  ['shrink', c => c.shrink, v => v === false],
  // underline/strike share ONE reader (font.{underline,strike}) -- a gap on either makes BOTH
  // unread, the same "one capability, every axis it carries" rule align.h/v and border already use.
  ['underline', c => c.font, v => isBaseFontEffects(v)],
  ['strike', c => c.font, v => isBaseFontEffects(v)],
]);
const isFormatScope = op => op.scope === 'formats' || op.scope === 'all';

function plan(step, index) {
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `step ${index}: clear_range does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `step ${index}: clear_range requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `step ${index}: clear_range requires a plain A1 range on the named sheet (no Sheet!A1)` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `step ${index}: clear_range range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  if ('scope' in step && !SCOPES.includes(step.scope)) return { ok: false, detail: `step ${index}: clear_range scope must be one of ${SCOPES.join(' | ')}` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'clear_range',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      scope: step.scope || 'contents',
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
    if (cell.merge && cell.merge !== a) return { ok: false, reason: isFormatScope(op) ? 'xlsx-clear-scope-dissolves-merge' : 'xlsx-clear-in-merged-area-refused', cell: a, mergeArea: cell.merge };
    if (!('nf' in cell) || !('fill' in cell)) return { ok: false, reason: 'xlsx-target-format-unread', cell: a };
    if (typeof cell.bold !== 'boolean') return { ok: false, reason: 'xlsx-target-bold-unobservable', cell: a, bold: 'bold' in cell ? cell.bold : 'UNOBSERVABLE:not-read' };
    if (isFormatScope(op) && !('wrap' in cell && 'rot' in cell && 'italic' in cell && 'align' in cell && 'border' in cell && 'readingOrder' in cell && 'indent' in cell && 'shrink' in cell && 'font' in cell)) return { ok: false, reason: 'xlsx-target-style-unread', cell: a };
    found[a] = cell;
  }
  return { ok: true, cells: found };
}

/** The base axes that are NOT at base on a cell (an unreadable axis is not at base: null = unread). */
function axesOffBase(cell) {
  const off = [];
  for (const [name, read, atBase] of BASE_AXES) {
    const x = read(cell);
    if (typeof x === 'string' && x.startsWith('UNOBSERVABLE')) return null;
    if (!atBase(x)) off.push(name);
  }
  return off;
}

function classifyScoped(g, op) {
  const unread = [];
  const pending = [];
  for (const { a } of op.cells) {
    const cell = g.cells[a];
    const off = axesOffBase(cell);
    if (off === null) { unread.push(a); continue; }
    if (op.scope === 'all' && isCleared(cell) !== true) pending.push(a);
    else if (off.length) pending.push(a);
  }
  if (unread.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-target-style-unobservable', cells: unread.slice(0, 20) } };
  if (!pending.length) return { state: 'SATISFIED', evidence: { sheet: op.selector.sheet, range: op.selector.range, scope: op.scope, proof: 'fresh-live-cells-already-at-base' } };
  return { state: 'APPLY', evidence: { sheet: op.selector.sheet, range: op.selector.range, scope: op.scope, pending: pending.length } };
}

function classify(primary, op) {
  const g = gate(primary, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  if (isFormatScope(op)) return classifyScoped(g, op);
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
  if (isFormatScope(op)) {
    const call = op.scope === 'all' ? 'Clear' : 'ClearFormats';
    return {
      kind: 'callCommand',
      body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.GetRange(${JSON.stringify(op.selector.range)}).${call}();
  return { ok: true, cleared: ${addrs.length} };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
      meta: Object.freeze({ family: 'clear', action: op.scope === 'all' ? 'clear-all' : 'clear-formats', sheet, range: op.selector.range, cells: Object.freeze(addrs) }),
    };
  }
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

/** L1 of the format scopes: every target cell's base axes are at base; 'formats' keeps value AND formula, 'all' leaves the cell blank. */
function verifyScoped(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  const offBase = [];
  const valueChanged = [];
  const notBlank = [];
  for (const { a: addr } of op.cells) {
    const now = cellAt(a, op.selector.sheet, addr);
    const was = cellAt(b, op.selector.sheet, addr);
    const off = now ? axesOffBase(now) : null;
    if (off === null || off.length) offBase.push({ cell: addr, axes: off === null ? 'unobservable' : off });
    if (op.scope === 'formats') { if (!was || !now || was.f !== now.f || JSON.stringify(was.v) !== JSON.stringify(now.v)) valueChanged.push(addr); }
    else if (isCleared(now) !== true) notBlank.push(addr);
  }
  return {
    ok: offBase.length === 0 && valueChanged.length === 0 && notBlank.length === 0,
    evidence: { sheet: op.selector.sheet, range: op.selector.range, scope: op.scope, cells: op.cells.length, offBase: offBase.slice(0, 20), valueChanged, notBlank, assertedAxes: BASE_AXES.map(x => x[0]), proof: 'fresh-live-base-axes-and-value' },
  };
}

/** L1: every target cell blank AND its number format and fill equal to the read before the write. */
function verifyPrimitive(op, before, after) {
  if (isFormatScope(op)) return verifyScoped(op, before, after);
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
  if (isFormatScope(op)) return rest;
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
  if (isFormatScope(op)) {
    const v = verifyScoped(op, null, finalObservation);
    const bad = v.evidence.offBase.map(x => x.cell).concat(v.evidence.notBlank);
    return [{ name: `every cell of ${op.selector.sheet}!${op.selector.range} is at base format${op.scope === 'all' ? ' and blank' : ''}`, ok: bad.length === 0, detail: bad.length ? { cells: bad } : null }];
  }
  const bad = op.cells.filter(({ a }) => isCleared(cellAt(finalObservation && finalObservation.primary, op.selector.sheet, a)) !== true).map(({ a }) => a);
  return [{ name: `every cell of ${op.selector.sheet}!${op.selector.range} is blank`, ok: bad.length === 0, detail: bad.length ? { notBlank: bad } : null }];
}

/** The target window, read WITH number format and fill (the proof needs both, before and after). */
const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.range, ...(isFormatScope(op) ? { style: true } : { format: true }) }];

module.exports = { plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, isCleared, MAX_CELLS };
