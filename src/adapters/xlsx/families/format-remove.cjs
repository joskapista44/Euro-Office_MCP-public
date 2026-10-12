'use strict';

const { cellsOf, parseRange } = require('../a1.cjs');
const { cellAt } = require('../inventory.cjs');
const { STYLE_KEYS, unobservable, readKey } = require('../style-axis.cjs');
const { OPPOSITE, REMOVED, H_REV, V_REV, STYLE_REV, API_SIDE, API_STYLE, neighbour, sidesList } = require('./format-maps.cjs');

/**
 * X-B04 (border removal, mode 'remove') -- split out of format.cjs (review, point 2: qlty smells
 * flagged a complexity jump in gate()/lower()/verifyPreservation after the remove-mode helpers landed there).
 * Module-level functions, not a factory (review, point 1: a closure-based factory folds every
 * inner function's complexity into the factory itself, which defeats the split's own purpose) -- the shared
 * vocabulary (H/V/STYLE/SIDE/OPPOSITE/API_SIDE/API_STYLE/REMOVED, their reverse maps, `neighbour`, `sidesList`)
 * comes from `./format-maps.cjs`, required directly here and by format.cjs, so neither file requires the
 * other.
 *
 * format.cjs calls in through five entry points: `removeExpectation` (plan-time spec validation),
 * `removeCells` / `removeFacing` / `removeWant` (the same mode-branch points every other border mode already
 * has there), and `removeGate` (the one check gate() adds for `op.remove`). `removeMeta` bundles the three
 * remove-specific `meta` fields lower() needs into one call. `cellGap` and `unrestorable` are internal.
 *
 * See format.cjs's own top-of-file doc block for the FULL 'remove' mode semantics (the composite, the
 * both-cells decision #ref, the refusal classes) -- this file is the implementation, not a second
 * explanation of it.
 */

const windowCells = op => cellsOf(parseRange(op.ring));

/** sides -> { remove: { sides } }; or a refusal detail. No `style`/`color` -- nothing is being drawn. */
function removeExpectation(b) {
  const bad = Object.keys(b).filter(k => !['mode', 'sides'].includes(k));
  if (bad.length) return { error: `border key(s) not supported with mode 'remove': ${bad.join(', ')}` };
  const sides = sidesList(b.sides);
  if (!sides) return { error: 'border.sides must be a non-empty list of distinct top / bottom / left / right' };
  return { remove: { sides } };
}

/** cell -> the border sides that must read REMOVED there -- every target cell's directly-removed sides
 * (already in `op.want`, via `wantsAt`), union a neighbour's shared side (a maintainer decision #ref: both cells
 * of every removed edge, since the shared border's physical owner is not observable from here). */
function removeCells(op) {
  const m = new Map();
  if (!op.remove) return m;
  const add = (a, x) => { const s = m.get(a) || new Set(); s.add(x); m.set(a, s); };
  for (const { a } of op.cells) for (const x of op.remove.sides) { add(a, x); const n = neighbour(a, x); if (n) add(n, OPPOSITE[x]); }
  return m;
}

const removeWant = x => ({ key: `border.${x}`, expect: REMOVED });

/** the sides of cell `a` that must read REMOVED only because a NEIGHBOUR directly removed the facing side --
 * `a`'s own directly-removed sides are already in `op.want`/`wantsAt`; a ring cell (outside op.cells) reaches
 * REMOVED only through this. */
function removeFacing(op, a) {
  if (!op.remove) return [];
  const all = removeCells(op).get(a);
  if (!all) return [];
  const own = op.cells.some(c => c.a === a);
  return [...all].filter(x => !(own && op.remove.sides.includes(x)));
}

/** A window cell that is NOT itself cleared/restored (not in removeCells()) is only ever READ by this
 * family, never written -- it needs border + merge readability, nothing more. The same weaker contract
 * outline/inside/all's own window check already uses. */
function weakGap(cell, a) {
  if (!cell || !('wrap' in cell)) return { reason: 'xlsx-target-format-unread', cell: a };
  const v = readKey(cell, 'border.t');
  if (v === undefined || unobservable(v)) return { reason: 'xlsx-format-remove-unmeasured-axis', cell: a, key: 'border.t', why: v === undefined ? 'not-read' : v };
  return cell.merge && cell.merge !== a ? { reason: 'xlsx-format-remove-merged-ring-refused', cell: a, mergeArea: cell.merge } : null;
}

/** cell `a`'s own window-gap check for a removeCells() cell (review, point 4: split out of
 * removeWindowGap's loop body so each function carries its own, smaller complexity) -- fully style-readable
 * on EVERY axis (not just border -- the composite clears the whole cell), not merged, no diagonal. Diagonal
 * restore is NOT implemented here (a scope boundary, not a missing public path: P2.0, m7-b4-probe1-20261004,
 * measured that `SetBorders('DiagonalDown', ...)` after `ClearFormats()` DOES restore it byte-exact; zero
 * review #ref). */
function cellGap(cell, a) {
  const weak = weakGap(cell, a);
  if (weak) return weak;
  for (const k of STYLE_KEYS) {
    const v = readKey(cell, k);
    if (v === undefined || unobservable(v)) return { reason: 'xlsx-format-remove-unmeasured-axis', cell: a, key: k, why: v === undefined ? 'not-read' : v };
  }
  const d = readKey(cell, 'border.d');
  return d && d.s > 0 ? { reason: 'xlsx-format-remove-diagonal-unsupported', cell: a } : null;
}

/** every removeCells() cell passes the STRONG check (cellGap); every other window cell -- a ring corner, a
 * neighbour of a side that is NOT being removed -- only passes the WEAK one (weakGap), since the composite
 * never clears or restores it (review: the X-B04 usability gap -- an irrelevant empty/unreadable
 * neighbour used to refuse the whole removal even though it is never touched). */
function removeWindowGap(primary, op) {
  const targets = removeCells(op);
  for (const a of windowCells(op)) {
    const cell = cellAt(primary, op.selector.sheet, a);
    const bad = targets.has(a) ? cellGap(cell, a) : weakGap(cell, a);
    if (bad) return bad;
  }
  return null;
}

/** [key, isBad, why] rows -- table-driven (review, point 3), so `unrestorable`'s own body is a
 * `find` plus the one border-side loop, not eight separate `if`-returns. Each row's `isBad(cell)` answers
 * "is this KEPT axis outside our write vocabulary", `why(cell)` is what to report. */
const UNRESTORABLE_AXES = [
  ['align.h', cell => cell.align.h !== null && !(cell.align.h in H_REV), cell => cell.align.h],
  ['align.v', cell => cell.align.v !== null && !(cell.align.v in V_REV), cell => cell.align.v],
  ['font.name', cell => cell.font.name === null, () => null],
  ['font.size', cell => cell.font.size === null, () => null],
  ['font.color', cell => cell.font.color === null, () => null],
  ['nf', cell => typeof cell.nf !== 'string', cell => cell.nf],
];

/** does cell `a`'s KEPT axes (every one not being removed) carry only values our write vocabulary can
 * actually restore? Returns the first unrestorable {key, why}, or null. MUST run before the write (zero
 * review #ref): a value outside the vocabulary (an align/border code our H/V/STYLE maps don't cover, a
 * font axis read as `null` where P0.5/c forbids guessing a default) would otherwise flow, undetected, into
 * `removeRestoreData`'s output -- `ClearFormats()` runs, the restore call then either throws mid-loop (the
 * cell left PARTIALLY cleared, a real format loss nothing undoes) or silently writes a wrong/fabricated
 * value (`rgb(null)` reading as black). */
function unrestorable(cell, clear) {
  const axis = UNRESTORABLE_AXES.find(([, isBad]) => isBad(cell));
  if (axis) return { key: axis[0], why: axis[2](cell) };
  const side = ['t', 'b', 'l', 'r'].filter(x => !clear.has(x)).map(x => [x, cell.border[x]]).find(([, v]) => v && v.s > 0 && !(v.s in STYLE_REV));
  return side ? { key: `border.${side[0]}`, why: side[1].s } : null;
}

/** the ONE check format.cjs's gate() adds for `op.remove`: every window cell readable + none merged + none
 * diagonal (removeWindowGap); `found` extended to cover the whole window (lower() writes BOTH cells of
 * every removed edge, a maintainer decision #ref, so `target.before` must carry every one of them, not just
 * op.cells); every kept axis on every removeCells() cell restorable (unrestorable). Returns a refusal
 * object, or null meaning proceed. Mutates `found` in place, same as format.cjs's own gate() loop does. */
function removeGate(primary, op, found) {
  const window = removeWindowGap(primary, op);
  if (window) return { ok: false, ...window };
  for (const a of windowCells(op)) if (!(a in found)) found[a] = cellAt(primary, op.selector.sheet, a);
  for (const [a, clear] of removeCells(op)) {
    const bad = unrestorable(found[a], clear);
    if (bad) return { ok: false, reason: 'xlsx-format-remove-unrestorable-value', cell: a, key: bad.key, why: bad.why };
  }
  return null;
}

/** the restore-script data for every cell in removeCells(op) -- read from `target.before` (gate's window
 * read, which covers the range + ring for remove). Each entry carries only the KEPT axes' explicit values:
 * word strings for align/border (the same vocabulary the rest of the family writes with), colours left as
 * integers (they differ per cell and per side, so they convert to an explicit CreateColorFromRGB() call at
 * script-run time, not at plan time) -- skipping an axis that read null/absent so the restore never
 * fabricates a value the cell never had (P0.5/c: omitting a colour argument is NOT the same as writing one). */
function removeRestoreData(op, target) {
  const out = [];
  for (const [a, clear] of removeCells(op)) {
    const before = target.before[a];
    const border = [];
    for (const x of ['t', 'b', 'l', 'r']) {
      if (clear.has(x)) continue;
      const side = before.border[x];
      if (side && side.s > 0) border.push([API_SIDE[x], API_STYLE[STYLE_REV[side.s]], side.c]);
    }
    out.push([a, {
      wrap: before.wrap,
      nf: before.nf,
      fill: typeof before.fill === 'number' ? before.fill : null,
      bold: before.font.bold === true,
      italic: before.font.italic === true,
      name: before.font.name,
      size: before.font.size,
      color: before.font.color,
      h: before.align.h === null ? null : H_REV[before.align.h],
      v: before.align.v === null ? null : V_REV[before.align.v],
      border,
    }]);
  }
  return out;
}

/** ClearFormats() + restore every kept axis, on every cell removeCells(op) names (direct targets AND their
 * removed-edge neighbours, a maintainer decision #ref) -- one data-driven loop, one callCommand ('' without a
 * remove). `SetBorders(side, 'None')` never removes anything (MEASURED, P0.3/a-b); this composite is the
 * only working public path. */
function removeLine(op, target) {
  if (!op.remove) return '';
  const data = removeRestoreData(op, target);
  return `
  const rgb = n => Api.CreateColorFromRGB((n >> 16) & 255, (n >> 8) & 255, n & 255);
  for (const [a, rs] of ${JSON.stringify(data)}) {
    const c = ws.GetRange(a);
    c.ClearFormats();
    c.SetWrap(rs.wrap);
    c.SetNumberFormat(rs.nf);
    if (rs.fill !== null) c.SetFillColor(rgb(rs.fill));
    if (rs.bold) c.SetBold(true);
    if (rs.italic) c.SetItalic(true);
    c.SetFontName(rs.name);
    c.SetFontSize(rs.size);
    c.SetFontColor(rgb(rs.color));
    if (rs.h !== null) c.SetAlignHorizontal(rs.h);
    if (rs.v !== null) c.SetAlignVertical(rs.v);
    for (const [side, style, col] of rs.border) c.SetBorders(side, style, rgb(col));
  }`;
}

/** lower()'s three remove-specific `meta` fields, bundled into one call (review, point 2: lower()
 * itself should make ONE call for its op.remove branch, not three separate removeCells(op) calls inline). */
function removeMeta(op, target) {
  if (!op.remove) return {};
  const cellsMap = removeCells(op);
  return {
    removeCells: [...cellsMap.keys()],
    removeClear: Object.fromEntries([...cellsMap].map(([a, set]) => [a, [...set]])),
    restoreData: removeRestoreData(op, target),
  };
}

module.exports = { removeExpectation, removeCells, removeFacing, removeWant, removeGate, removeLine, removeMeta, unrestorable };
