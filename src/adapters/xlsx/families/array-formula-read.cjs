'use strict';

const { parseRange, addr, rangeText, MAX_COL, MAX_ROW } = require('../a1.cjs');
const { memberOf } = require('./array-formula.cjs');
const formula = require('../formula.cjs');

/**
 * GH #343 (F01): the PRODUCT READ behind `read_array_formulas` -- the array formulas of one range: anchor (top-left cell),
 * ref (the whole area), formula, type 'array', mirroring write_array_formula's own `range` / `formula` fields.
 *
 * What the public API gives (LIVE probe, docs/evidence/gh301-d4-array-formula-probe-20261025.json and the -live- file,
 * the same facts array-formula.cjs's header states): GetFormulaArray is NOT an extent reader -- any cell of an array returns
 * the array's formula, a superset null, a plain formula null, a value the value, an empty cell ''. So a cell is a MEMBER
 * when GetFormulaArray starts with '=' AND its own GetFormula is not empty (array-formula.cjs `memberOf`, the SAME
 * function the write guard uses -- not a copy), and an extent is proven by the area's every cell being a member while no
 * cell of its one-cell RING is one of the same formula. That is why the read asks the requested range PLUS a one-cell ring.
 *
 * Per area (a 4-neighbour block of members with the same formula, formula.sameFormula):
 *   extent 'proven'                        the block lies inside the requested range, is a full rectangle, and its ring has no
 *                                          member of that formula -> anchor / ref / formula
 *   extent 'crosses-the-requested-range'   the block reaches a ring cell outside the requested range: the array extends past
 *                                          the boundary and its true ref is NOT observable here -> anchor / ref are null; only the
 *                                          part INSIDE the range is named (insideRange). No ref is invented.
 *   extent 'not-a-rectangle'               members of one formula that do not fill a rectangle: no single array area fits
 * Two ADJACENT arrays with the SAME formula cannot be told apart through the public API; they form ONE block. `note` says so
 * in every answer, and an 'proven' area is "the maximal block of members of that formula", never proof of exactly one array.
 * Any unreadable cell (range or ring) makes the whole answer a NAMED gap (`arrays: null`, reason), never a partial guess.
 * In session the computed values stay '' until a reopen; this read returns formulas only.
 */

const NOTE = 'two adjacent arrays with the same formula cannot be told apart through the public API (GetFormulaArray is not an extent reader): each area is the maximal block of members of one formula, not proof of exactly one array';

/** the requested range and its one-cell ring (clipped to the sheet) */
function ringed(range) {
  const g = parseRange(range);
  const ring = { c1: Math.max(1, g.c1 - 1), c2: Math.min(MAX_COL, g.c2 + 1), r1: Math.max(1, g.r1 - 1), r2: Math.min(MAX_ROW, g.r2 + 1) };
  ring.count = (ring.c2 - ring.c1 + 1) * (ring.r2 - ring.r1 + 1);
  return { g, ring };
}

/** the readLive request fields: a window over the ring (cellAt needs it) and every ring cell for the GetFormulaArray axis */
function readRequest(a) {
  const { ring } = ringed(a.range);
  const arrayCells = [];
  for (let r = ring.r1; r <= ring.r2; r += 1) for (let c = ring.c1; c <= ring.c2; c += 1) arrayCells.push({ sheet: a.sheet, cell: addr(c, r) });
  return { windows: [{ sheet: a.sheet, range: rangeText(ring) }], arrayCells };
}

const boundsOf = cells => ({ c1: Math.min(...cells.map(x => x.c)), c2: Math.max(...cells.map(x => x.c)), r1: Math.min(...cells.map(x => x.r)), r2: Math.max(...cells.map(x => x.r)) });
const withCount = b => ({ ...b, count: (b.c2 - b.c1 + 1) * (b.r2 - b.r1 + 1) });

function projectArrayFormulasRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, range: a.range, error: 'SHEET_NOT_FOUND' };
  const { g, ring } = ringed(a.range);
  const members = new Map();
  for (let r = ring.r1; r <= ring.r2; r += 1) {
    for (let c = ring.c1; c <= ring.c2; c += 1) {
      const m = memberOf(inv, a.sheet, addr(c, r));
      if (m.problem) return { sheet: a.sheet, range: a.range, arrays: null, reason: m.problem, why: m.why || null, cell: addr(c, r) };
      if (m.member) members.set(`${c},${r}`, { c, r, raw: m.raw });
    }
  }
  const seen = new Set();
  const arrays = [];
  const inside = x => x.c >= g.c1 && x.c <= g.c2 && x.r >= g.r1 && x.r <= g.r2;
  for (const start of members.values()) {
    if (seen.has(`${start.c},${start.r}`)) continue;
    const block = [];
    const queue = [start];
    seen.add(`${start.c},${start.r}`);
    while (queue.length) {
      const cur = queue.pop();
      block.push(cur);
      for (const [dc, dr] of [[1, 0], [-1, 0], [0, 1], [0, -1]]) {
        const k = `${cur.c + dc},${cur.r + dr}`;
        const nb = members.get(k);
        if (nb && !seen.has(k) && formula.sameFormula(start.raw, nb.raw)) { seen.add(k); queue.push(nb); }
      }
    }
    const ins = block.filter(inside);
    if (!ins.length) continue;
    const b = withCount(boundsOf(block));
    const order = { _r: Math.min(...ins.map(x => x.r)), _c: Math.min(...ins.filter(x => x.r === Math.min(...ins.map(y => y.r))).map(x => x.c)) };
    let entry;
    if (block.length !== b.count) entry = { anchor: null, ref: null, formula: start.raw, type: 'array', extent: 'not-a-rectangle', cells: block.map(x => addr(x.c, x.r)).sort().slice(0, 20) };
    else if (block.some(x => !inside(x))) entry = { anchor: null, ref: null, formula: start.raw, type: 'array', extent: 'crosses-the-requested-range', insideRange: rangeText(withCount(boundsOf(ins))) };
    else entry = { anchor: addr(b.c1, b.r1), ref: rangeText(b), formula: start.raw, type: 'array', extent: 'proven' };
    arrays.push({ ...order, entry });
  }
  arrays.sort((x, y) => x._r - y._r || x._c - y._c);
  return { sheet: a.sheet, range: a.range, arrays: arrays.map(x => x.entry), note: NOTE };
}

module.exports = { NOTE, readRequest, projectArrayFormulasRead };
