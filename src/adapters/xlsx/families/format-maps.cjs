'use strict';

const { addr } = require('../a1.cjs');

/**
 * The `format` family's shared, pure vocabulary -- the alignment/style/side maps and the handful of tiny
 * address helpers both `format.cjs` and `format-remove.cjs` need. Split out so neither file duplicates them
 * and neither has to require the other (review, point 1: no circular require between format.cjs
 * and format-remove.cjs).
 */

const H = Object.freeze({ left: 1, center: 2, right: 0, justify: 3 });
const V = Object.freeze({ bottom: 0, center: 1, distributed: 2, justify: 3, top: 4 });
// X-B06 (MEASURED, m7-b06-probe1): 9 additional styles, each a distinct numeric code, no
// collision with the original four or with each other, reopen-stable.
const STYLE = Object.freeze({
  thin: 7, medium: 12, thick: 13, dashed: 6,
  double: 1, hair: 2, dashDotDot: 3, dashDot: 4, dotted: 5, mediumDashDotDot: 8, slantDashDot: 9, mediumDashDot: 10, mediumDashed: 11,
});
const SIDE = Object.freeze({ top: 't', bottom: 'b', left: 'l', right: 'r' });
const OPPOSITE = Object.freeze({ t: 'b', b: 't', l: 'r', r: 'l' });
const API_SIDE = Object.freeze({ t: 'Top', b: 'Bottom', l: 'Left', r: 'Right' });
// X-B05: a deliberately SEPARATE map from SIDE -- a diagonal name must never reach OPPOSITE/neighbour
// (both t/b/l/r-only; a diagonal has no neighbour-facing side) through sidesList()/requestedSides().
const DIAGONAL = Object.freeze({ diagonalDown: 'dd', diagonalUp: 'du' });
const API_DIAGONAL = Object.freeze({ dd: 'DiagonalDown', du: 'DiagonalUp' });
const API_STYLE = Object.freeze({
  thin: 'Thin', medium: 'Medium', thick: 'Thick', dashed: 'Dashed',
  double: 'Double', hair: 'Hair', dashDotDot: 'DashDotDot', dashDot: 'DashDot', dotted: 'Dotted', mediumDashDotDot: 'MediumDashDotDot', slantDashDot: 'SlantDashDot', mediumDashDot: 'MediumDashDot', mediumDashed: 'MediumDashed',
});
const REMOVED = Object.freeze({ s: 0, c: 0 });

const reverseOf = o => Object.fromEntries(Object.entries(o).map(([k, v]) => [v, k]));
const H_REV = reverseOf(H);
const V_REV = reverseOf(V);
const STYLE_REV = reverseOf(STYLE);

const col = a => /^[A-Z]+/.exec(a)[0];
const row = a => Number(/\d+$/.exec(a)[0]);
const colNum = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

/** The address one side away from `a`, or null off the sheet's top/left edge. */
const neighbour = (a, side) => {
  const c = colNum(col(a));
  const r = row(a);
  const [dc, dr] = { t: [0, -1], b: [0, 1], l: [-1, 0], r: [1, 0] }[side];
  return c + dc < 1 || r + dr < 1 ? null : addr(c + dc, r + dr);
};

/** A `sides` option (word list) -> internal side codes, or null if it isn't a non-empty list of distinct,
 * recognized side words. */
const sidesList = sides => (!Array.isArray(sides) || !sides.length || new Set(sides).size !== sides.length || !sides.every(x => x in SIDE) ? null : sides.map(x => SIDE[x]));

/** Same shape as sidesList, over DIAGONAL instead of SIDE. */
const diagonalSidesList = sides => (!Array.isArray(sides) || !sides.length || new Set(sides).size !== sides.length || !sides.every(x => x in DIAGONAL) ? null : sides.map(x => DIAGONAL[x]));

module.exports = { H, V, STYLE, SIDE, OPPOSITE, API_SIDE, API_STYLE, REMOVED, DIAGONAL, API_DIAGONAL, H_REV, V_REV, STYLE_REV, col, row, colNum, neighbour, sidesList, diagonalSidesList };
