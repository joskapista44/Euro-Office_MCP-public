'use strict';

const { cellsOf, parseRange } = require('../a1.cjs');
const { cellAt } = require('../inventory.cjs');
const { STYLE_KEYS, unobservable, readKey } = require('../style-axis.cjs');
const { diagonalRestoreOf, diagonalRestoreScript } = require('./format-diagonal.cjs');
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

/** A window cell is only ever READ by the readability part of this check; it needs border + merge readability, nothing more (the same weaker contract outline/inside/all's own window
 * check already uses). A merge never refuses on its own any more: a merge part the composite WRITES is released through C5 (UnMerge -> composite -> Merge, `c5Wrap`; L-16/3, LIVE-measured),
 * and a merge it does not write is read back unchanged (L-16/2, LIVE-measured) -- the merge text must only be re-issuable (parseable); what C5 was not measured for is refused in `c5Gap`. */
function weakGap(cell, a) {
  if (!cell || !('wrap' in cell)) return { reason: 'xlsx-target-format-unread', cell: a };
  const v = readKey(cell, 'border.t');
  if (v === undefined || unobservable(v)) return { reason: 'xlsx-format-remove-unmeasured-axis', cell: a, key: 'border.t', why: v === undefined ? 'not-read' : v };
  if (!cell.merge || cell.merge === a) return null;
  return parseRange(cell.merge) ? null : { reason: 'xlsx-format-remove-merged-ring-refused', cell: a, mergeArea: cell.merge };
}

/** cell `a`'s own window-gap check for a removeCells() cell (review, point 4: split out of
 * removeWindowGap's loop body so each function carries its own, smaller complexity) -- fully style-readable
 * on EVERY axis (not just border -- the composite clears the whole cell); a merge part is allowed (L-16/3, C5: `c5Areas`). A drawn diagonal is
 * RESTORED (L-16/1, MEASURED: DiagonalDown / DiagonalUp / both restore byte-exact after `ClearFormats()`); only a
 * diagonal that cannot be re-issued exactly refuses (diagonalRestoreOf's `bad`). */
function cellGap(cell, a) {
  const weak = weakGap(cell, a);
  if (weak) return weak;
  for (const k of STYLE_KEYS) {
    const v = readKey(cell, k);
    if (v === undefined || unobservable(v)) return { reason: 'xlsx-format-remove-unmeasured-axis', cell: a, key: k, why: v === undefined ? 'not-read' : v };
  }
  const dg = diagonalRestoreOf(cell);
  return dg && dg.bad !== undefined ? { reason: 'xlsx-format-remove-unrestorable-value', cell: a, key: 'border.d', why: dg.bad } : null;
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
  return c5Gap(c5Areas(op, a => cellAt(primary, op.selector.sheet, a)));
}

/** L-16/3 (C5): the distinct merge areas the composite would write into -- every removeCells() cell that is a merge part names its area. The composite's ClearFormats() dissolves
 * a merge (LIVE), so these areas are UnMerge()d first and Merge()d again after (C5); LIVE-measured on the realistic merged header and 2x2 / 1x3 / 3x1 x four sides, each alone. */
function c5Areas(op, readOf) {
  const out = new Set();
  for (const a of removeCells(op).keys()) { const cell = readOf(a); if (cell && cell.merge && cell.merge !== a) out.add(cell.merge); }
  return [...out];
}
const C5_MAX_CELLS = 400;
/** What C5 was NOT measured for stays refused, named: more than one merge area in one write (adjacent merges), and an area too large for the in-body member guard. */
function c5Gap(areas) {
  if (areas.length > 1) return { reason: 'xlsx-format-remove-multi-merge-unmeasured', areas: areas.slice(0, 5), count: areas.length };
  const m = areas.length ? parseRange(areas[0]) : null;
  const size = m ? cellsOf(m).length : 0;
  return size > C5_MAX_CELLS ? { reason: 'xlsx-format-remove-merge-too-large', mergeArea: areas[0], cells: size, max: C5_MAX_CELLS } : null;
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
      diagonal: diagonalRestoreOf(before),
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
  const composite = compositeLine(op, target);
  const areas = c5Areas(op, a => target.before[a]);
  return areas.length ? c5Wrap(areas, composite) : composite;
}

/** C5 (L-16/3): UnMerge the area, run the composite, Merge it again -- in ONE body. A guard BEFORE the UnMerge reads every non-anchor member through the raw GetValue() and aborts
 * (nothing touched) if one holds a value, because the re-Merge would discard it (a member the raw read cannot answer for does not abort: the guard is best-effort, and the LIVE
 * measurement could not create a member value on this DS). The re-Merge sits in a `finally`: a throw after the UnMerge still puts the area back (the outer
 * catch reports the original error, and the verification names a merge that did not come back). */
function c5Wrap(areas, composite) {
  const members = areas.flatMap(m => cellsOf(parseRange(m)).slice(1));
  return `
  const MERGES = ${JSON.stringify(areas)};
  for (const m of ${JSON.stringify(members)}) { let v; try { v = ws.GetRange(m).GetValue(); } catch (e) { v = undefined; } if (v !== undefined && v !== null && String(v) !== '') return { ok: false, stage: 'merge-member-value', cell: m }; }
  const unmerged = [];
  try {
    for (const m of MERGES) { ws.GetRange(m).UnMerge(); unmerged.push(m); }${composite}
  } finally {
    for (const m of unmerged) { try { ws.GetRange(m).Merge(false); } catch (e) { /* the verification names a merge that did not come back */ } }
  }`;
}

function compositeLine(op, target) {
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
    for (const [side, style, col] of rs.border) c.SetBorders(side, style, rgb(col));${diagonalRestoreScript('rgb')}
  }`;
}

/** lower()'s three remove-specific `meta` fields, bundled into one call (review, point 2: lower()
 * itself should make ONE call for its op.remove branch, not three separate removeCells(op) calls inline). */
function removeMeta(op, target) {
  if (!op.remove) return {};
  const cellsMap = removeCells(op);
  return {
    removeCells: [...cellsMap.keys()],
    ...(c5Areas(op, a => target.before[a]).length ? { mergeAreas: c5Areas(op, a => target.before[a]) } : {}),
    removeClear: Object.fromEntries([...cellsMap].map(([a, set]) => [a, [...set]])),
    restoreData: removeRestoreData(op, target),
  };
}

module.exports = { removeExpectation, removeCells, removeFacing, removeWant, removeGate, removeLine, removeMeta, unrestorable };
