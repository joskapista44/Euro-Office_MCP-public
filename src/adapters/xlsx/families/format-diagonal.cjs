'use strict';

const { readKey, unobservable } = require('../style-axis.cjs');
const { DIAGONAL, API_DIAGONAL, API_STYLE, STYLE_REV, diagonalSidesList } = require('./format-maps.cjs');

/**
 * X-B05 (diagonal borders) -- split out of format.cjs (review, point 2: qlty smells grew in
 * gate()/lower()/expectationsOf()/borderExpectation() past the baseline once the diagonal branch landed
 * there). Module-level functions, not a factory (same reasoning as format-remove.cjs, review
 * point 1) -- the shared vocabulary (DIAGONAL/API_DIAGONAL/diagonalSidesList) comes from
 * `./format-maps.cjs`, required directly here and by format.cjs, so neither file requires the other.
 *
 * format.cjs calls in through three entry points: `diagonalExpectation` (plan-time spec validation,
 * called from borderExpectation once the shared 'none'/style/colour checks already passed),
 * `diagonalWant` (the want-list entries expectationsOf needs), and `diagonalConflict` (the one check
 * gate() adds for `op.diagonal`). `diagonalLine` is lower()'s per-cell command-body fragment.
 *
 * See format.cjs's own top-of-file doc block for the FULL diagonal semantics (the shared `border.d`,
 * the conflicting-style refusal, 'None' not clearing a diagonal, diagonal removal out of scope) -- this
 * file is the implementation, not a second explanation of it.
 */

/** Is `sides` a non-empty list where EVERY name is a diagonal name? format.cjs's borderExpectation
 * calls this to decide which branch to take -- diagonal sides are never mixed with top/bottom/left/
 * right (a diagonal has no neighbour-facing edge, so it never shares the SIDE/OPPOSITE/neighbour
 * machinery outline/inside/all/remove depend on). Keeps the DIAGONAL map itself private to this file
 * and format-maps.cjs. */
const isDiagonalSides = sides => Array.isArray(sides) && sides.length > 0 && sides.every(x => x in DIAGONAL);

/** sides -> { diagonal: { codes, expect } }; or a refusal detail. Only called once format.cjs's
 * borderExpectation has confirmed (via isDiagonalSides) that every named side is a diagonal name. */
function diagonalExpectation(sides, expect) {
  const codes = diagonalSidesList(sides);
  return codes ? { diagonal: { codes, expect } } : { error: 'border.sides must be a non-empty list of distinct diagonalDown / diagonalUp' };
}

/** { codes, expect } -> the want-list entries: the shared border.d style+colour once, plus each
 * requested direction's own boolean flag (border.dd / border.du). */
function diagonalWant({ codes, expect }) {
  return [{ key: 'border.d', expect }, ...codes.map(c => ({ key: `border.${c}`, expect: true }))];
}

/** X-B05/S03 (MEASURED, m7-b05-probe1): the engine models both diagonal directions with ONE shared
 * `border.d` style+colour -- writing just one direction when the OTHER is already drawn with a
 * DIFFERENT style/colour would silently change that one too (the second SetBorders call overwrites
 * `d` for both). Only relevant when the op names exactly one direction (naming both together writes
 * the SAME style/colour to both, the measured-safe case). An unreadable flag of the other direction
 * refuses too (review: it cannot be told apart from "not drawn", and the shared border.d
 * is a requested key, so L3 could not see that direction's style being silently overwritten --
 * fail-closed). The requested direction's own readability is caught by gate()'s generic per-key check,
 * not here. */
function diagonalConflict(op, cell) {
  if (!op.diagonal || op.diagonal.codes.length !== 1) return null;
  const [only] = op.diagonal.codes;
  const other = only === 'dd' ? 'du' : 'dd';
  const drawn = readKey(cell, `border.${other}`);
  if (drawn === undefined || unobservable(drawn)) return { reason: 'xlsx-format-diagonal-unobservable', key: `border.${other}`, why: drawn === undefined ? 'not-read' : drawn };
  if (drawn !== true) return null;
  const existing = readKey(cell, 'border.d');
  if (!existing || typeof existing !== 'object') return null;
  const { expect } = op.diagonal;
  const same = existing.s === expect.s && (existing.c === expect.c || (expect.c === 0 && existing.c === null));
  return same ? null : { reason: 'xlsx-format-diagonal-conflicting-style', existing, requested: expect };
}

/** lower()'s per-cell SetBorders('DiagonalDown'|'DiagonalUp', ...) call(s) -- one per requested
 * direction, same style+colour for both if both are requested (the measured-safe case). Returned as one
 * joined string so format.cjs's lower() can push it into `onCell` exactly like the regular-sides branch
 * right next to it. */
function diagonalLine(op, colour) {
  const { style, color } = op.spec.border;
  return op.diagonal.codes.map(c => `c.SetBorders(${JSON.stringify(API_DIAGONAL[c])}, ${JSON.stringify(API_STYLE[style])}, ${colour(color)});`).join(' ');
}

/** L-16/1: the diagonal a ClearFormats()-based restore must bring back, read from the cell's own `before`.
 * null = no diagonal drawn. A diagonal that IS drawn but cannot be re-issued exactly (a style code outside our
 * write vocabulary, or the shared `border.d` set while neither direction flag is true) is `{ bad: <why> }` --
 * the caller refuses BEFORE the write (the restore must never fabricate or drop a diagonal). Direction flags
 * that are unreadable are already refused by the STYLE_KEYS observability loop of both callers. MEASURED
 * (l16-diagonal-probe-live-20261016, zero): DD / DU / both restore byte-exact with this body. */
function diagonalRestoreOf(cell) {
  const d = readKey(cell, 'border.d');
  if (!d || typeof d !== 'object' || !(d.s > 0)) return null;
  const dd = readKey(cell, 'border.dd') === true;
  const du = readKey(cell, 'border.du') === true;
  if (!(d.s in STYLE_REV)) return { bad: d.s };
  if (!dd && !du) return { bad: 'no-direction' };
  return { style: API_STYLE[STYLE_REV[d.s]], color: d.c, dd, du };
}

/** the restore-script lines for one cell's diagonal (`rs.diagonal`, `rgbFn` = the caller's colour helper). */
function diagonalRestoreScript(rgbFn) {
  return `
    if (rs.diagonal) {
      if (rs.diagonal.dd) c.SetBorders(${JSON.stringify(API_DIAGONAL.dd)}, rs.diagonal.style, ${rgbFn}(rs.diagonal.color));
      if (rs.diagonal.du) c.SetBorders(${JSON.stringify(API_DIAGONAL.du)}, rs.diagonal.style, ${rgbFn}(rs.diagonal.color));
    }`;
}

module.exports = { diagonalRestoreOf, diagonalRestoreScript, isDiagonalSides, diagonalExpectation, diagonalWant, diagonalConflict, diagonalLine };
