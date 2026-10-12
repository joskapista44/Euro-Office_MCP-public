'use strict';

const { cellAt } = require('../inventory.cjs');
const { STYLE_KEYS, unobservable, readKey } = require('../style-axis.cjs');
const { H_REV, V_REV, STYLE_REV, API_SIDE, API_STYLE, neighbour } = require('./format-maps.cjs');
const { unrestorable } = require('./format-remove.cjs');

/**
 * M7 backlog X-F08 (fill removal, `format: { fill: 'none' }`) -- built on the X-B04 remove
 * machinery's own pattern (format-remove.cjs), reusing its `unrestorable()` directly: the ONE
 * working public path to clear a single axis is ClearFormats() + explicit restore of every KEPT
 * axis (there is no scoped SetFillColor(null)/() public setter -- both throw, m7-s15-probe2).
 * Unlike border-remove, fill-clear is SINGLE-CELL: clearing a cell's fill has no shared-edge
 * concept of its own, so there is no neighbour-pair write here (a maintainer decision #ref does not
 * apply) -- `unrestorable(cell, new Set())` is called with an EMPTY clear set, since no border
 * side is removed, every side is kept.
 *
 * Measured restore sequence (m7-s15-probe4, LIVE-accepted): ClearFormats() then
 * SetWrap / SetNumberFormat / SetBold(conditional) / SetItalic(conditional) / SetFontName /
 * SetFontSize / SetFontColor / SetAlignHorizontal(conditional) / SetAlignVertical(conditional) /
 * SetBorders(loop) -- the SAME sequence removeLine() uses, minus the fill-restore line (fill is
 * the one axis deliberately left cleared).
 *
 * Three refusal classes, all BEFORE the write, none measured by the probe (one cell, one side,
 * no neighbour interaction --):
 *   (1) any KEPT axis unmeasured (the same STYLE_KEYS observability loop format-remove.cjs's
 *       cellGap() runs, duplicated here rather than imported: cellGap also folds in the
 *       border-removal weak/strong gap split this feature does not need) or outside our H/V/
 *       STYLE write vocabulary (`unrestorable`, reused as-is);
 *   (2) a diagonal border on the target cell (no diagonal restore path wired here, the same
 *       scope boundary format-remove.cjs draws for its own diagonal case);
 *   (3) the target carries a non-zero border side facing a READABLE neighbour cell at all --
 *       the probe's own accepted run (m7-s15-probe4) never checked a neighbour's state before or
 *       after (it reads only the target's own axes), so whether the restore's SetBorders() call
 *       disturbs the physically-shared edge (format.cjs's own doc: "a cell's read shows its
 *       neighbour's SHARED edge") is unmeasured REGARDLESS of what the neighbour itself carries
 *       -- not just when the neighbour's facing side happens to match.
 *
 * Merge is already refused generically (format.cjs's own `cellRefusal`, before this runs).
 *
 * format.cjs calls in through three entry points: `fillClearGate` (the one check gate() adds for
 * `op.fillClear`), `fillClearLine` (lower()'s callCommand body), `fillClearMeta` (the one extra
 * `meta` field lower() needs).
 */

/** every KEPT style key of cell `a` must be observable before the composite runs -- the same
 * STYLE_KEYS loop format-remove.cjs's cellGap() uses. */
function kept(cell, a) {
  for (const k of STYLE_KEYS) {
    const v = readKey(cell, k);
    if (v === undefined || unobservable(v)) return { reason: 'xlsx-format-fill-clear-unmeasured-axis', cell: a, key: k, why: v === undefined ? 'not-read' : v };
  }
  return null;
}

/** does cell `a` carry a non-zero border side facing a READABLE neighbour cell at all -- NOT
 * measured by the probe (it never checked a neighbour's state at all). Returns the first such
 * side found, or null. A side with NO readable neighbour (the sheet's outer edge) is fine: there
 * is nothing on the other side of it to disturb. */
function sharedEdge(primary, sheet, a, cell) {
  for (const x of ['t', 'b', 'l', 'r']) {
    const side = cell.border[x];
    if (!side || side.s <= 0) continue;
    const n = neighbour(a, x);
    if (!n) continue;
    if (cellAt(primary, sheet, n)) return { side: x, neighbour: n };
  }
  return null;
}

/** the ONE check format.cjs's gate() adds for `op.fillClear`: every target cell's kept axes
 * observable + restorable, none diagonal, none sharing an edge with a readable neighbour.
 * Returns a refusal object, or null meaning proceed. */
function fillClearGate(primary, op, found) {
  for (const { a } of op.cells) {
    const cell = found[a];
    const k = kept(cell, a);
    if (k) return { ok: false, ...k };
    const bad = unrestorable(cell, new Set());
    if (bad) return { ok: false, reason: 'xlsx-format-fill-clear-unrestorable-value', cell: a, key: bad.key, why: bad.why };
    const d = readKey(cell, 'border.d');
    if (d && d.s > 0) return { ok: false, reason: 'xlsx-format-fill-clear-diagonal-unsupported', cell: a };
    const shared = sharedEdge(primary, op.selector.sheet, a, cell);
    if (shared) return { ok: false, reason: 'xlsx-format-fill-clear-shared-edge-unsupported', cell: a, side: shared.side, neighbour: shared.neighbour };
  }
  return null;
}

/** the restore-script data for every op.cells target -- read from `target.before` (gate's
 * window read). Mirrors format-remove.cjs's removeRestoreData shape exactly, minus a `clear`
 * parameter (no side is removed here) and minus the `fill` field (the one axis deliberately
 * left cleared, so there is nothing to restore there). */
function fillClearRestoreData(op, target) {
  const out = [];
  for (const { a } of op.cells) {
    const before = target.before[a];
    const border = [];
    for (const x of ['t', 'b', 'l', 'r']) {
      const side = before.border[x];
      if (side && side.s > 0) border.push([API_SIDE[x], API_STYLE[STYLE_REV[side.s]], side.c]);
    }
    out.push([a, {
      wrap: before.wrap,
      nf: before.nf,
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

/** ClearFormats() + restore every kept axis, on every op.cells target -- one data-driven loop.
 * Fill is left cleared; every other axis matches its pre-write value exactly (m7-s15-probe4,
 * LIVE-accepted). '' when `op.fillClear` is not set. */
function fillClearLine(op, target) {
  if (!op.fillClear) return '';
  const data = fillClearRestoreData(op, target);
  return `
  const rgbFC = n => Api.CreateColorFromRGB((n >> 16) & 255, (n >> 8) & 255, n & 255);
  for (const [a, rs] of ${JSON.stringify(data)}) {
    const c = ws.GetRange(a);
    c.ClearFormats();
    c.SetWrap(rs.wrap);
    c.SetNumberFormat(rs.nf);
    if (rs.bold) c.SetBold(true);
    if (rs.italic) c.SetItalic(true);
    c.SetFontName(rs.name);
    c.SetFontSize(rs.size);
    c.SetFontColor(rgbFC(rs.color));
    if (rs.h !== null) c.SetAlignHorizontal(rs.h);
    if (rs.v !== null) c.SetAlignVertical(rs.v);
    for (const [side, style, col] of rs.border) c.SetBorders(side, style, rgbFC(col));
  }`;
}

/** lower()'s one fillClear-specific `meta` field. */
function fillClearMeta(op, target) {
  if (!op.fillClear) return {};
  return { fillClearCells: op.cells.map(c => c.a), restoreData: fillClearRestoreData(op, target) };
}

module.exports = { fillClearGate, fillClearLine, fillClearMeta };
