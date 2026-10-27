'use strict';

const { parseRange, colName } = require('../a1.cjs');

/**
 * L-09 unit 1: the PROJECTION behind the `read_layout` read tool -- a pure field-select over `inv.layout[sheet]` (the driver's own in-session read: per column
 * GetEntireColumn().GetColumnWidth() / GetHidden(), per row GetEntireRow().GetRowHeight() / GetHidden(); xlsx-driver.cjs, "read only when a layout step asks"),
 * shaped as the MIRROR of the layout_range write schema (layout.cjs): the same units (width in (0, 255], height in (0, 409]), hidden a boolean. No new bridge call.
 *
 * What the engine returns that the write schema has no word for is NEVER turned into a value: it is named.
 *   - a hidden column / row (the engine reads width 0 / height null, W4b LIVE)  -> { hidden: true, width|height: null, reason: 'HIDDEN: ...' }
 *   - a row whose height was never set (null, not hidden)                       -> { height: null, hidden: false, reason: 'UNOBSERVABLE: ...' }
 *   - a value outside the write range / not a number                            -> the entry stays, `unrepresentable: [{ field, engine, reason }]`
 *   - a whole sheet that could not be read                                      -> { columns: null, rows: null, reason: 'UNOBSERVABLE:...' }
 * The driver reads columns 1..maxCol and rows 1..maxRow (and refuses the whole sheet above MAX_LAYOUT_ITEMS), so the extent of a request is the range's far corner.
 */

const CEILING = 500; // the write side's MAX_ITEMS (layout.cjs): one request names at most this many columns and this many rows
const WIDTH_MAX = 255;
const HEIGHT_MAX = 409;
const CAVEATS = Object.freeze(['a never-set row reads the engine default height (15 after a reopen): a set 15 and an unset row are indistinguishable']);

/** The extent the driver must read for a range: its last column and last row (the driver reads from 1). null for an unparsable range. */
function layoutExtent(range) {
  const g = parseRange(String(range).replace(/\$/g, ''));
  return g ? { maxCol: g.c2, maxRow: g.r2 } : null;
}

const bad = (field, engine, reason) => ({ field, engine, reason });

/** One column / row entry -> the write-schema shape. `field` is 'width' (columns) or 'height' (rows). */
function entryOf(e, field, max) {
  if (!e || typeof e !== 'object') return { reason: 'UNOBSERVABLE:not-read' };
  if (typeof e.hidden !== 'boolean') return { reason: `UNOBSERVABLE:hidden-${e.hidden === undefined ? 'not-read' : 'not-a-boolean'}` };
  if (e.hidden) return { hidden: true, [field]: null, reason: `HIDDEN: the engine reads ${field} ${field === 'width' ? '0' : 'null'} on a hidden ${field === 'width' ? 'column' : 'row'}, the ${field} is not observable` };
  const v = e[field];
  if (v === null || v === undefined) return { [field]: null, hidden: false, reason: `UNOBSERVABLE: the ${field} is not set (the engine reads null before a reopen)` };
  if (typeof v !== 'number' || !Number.isFinite(v) || v <= 0) return { [field]: null, hidden: false, unrepresentable: [bad(field, v, 'not a positive number')] };
  if (v > max) return { [field]: null, hidden: false, unrepresentable: [bad(field, v, `outside the writable (0, ${max}]`)] };
  return { [field]: v, hidden: false };
}

/** inv.layout[sheet] -> the tool result. A string is a whole-sheet gap, never guessed as "default layout". */
function projectLayoutRead(inv, a) {
  const base = { sheet: a.sheet, range: a.range };
  if (!inv.sheets.some(s => s.name === a.sheet)) return { ...base, error: 'SHEET_NOT_FOUND' };
  const g = parseRange(String(a.range).replace(/\$/g, ''));
  if (!g) return { ...base, columns: null, rows: null, reason: 'INVALID_RANGE' };
  const nCols = g.c2 - g.c1 + 1;
  const nRows = g.r2 - g.r1 + 1;
  if (nCols > CEILING || nRows > CEILING) return { ...base, columns: null, rows: null, reason: `RANGE_OVER_CEILING: ${nCols} columns x ${nRows} rows, at most ${CEILING} of each` };
  const L = inv.layout ? inv.layout[a.sheet] : undefined;
  if (L === undefined || L === null) return { ...base, columns: null, rows: null, reason: 'UNOBSERVABLE:not-read' };
  if (typeof L === 'string') return { ...base, columns: null, rows: null, reason: L };
  if (!L.cols || typeof L.cols !== 'object' || !L.rows || typeof L.rows !== 'object') return { ...base, columns: null, rows: null, reason: 'UNOBSERVABLE:unexpected-shape' };
  const columns = {};
  for (let c = g.c1; c <= g.c2; c += 1) columns[colName(c)] = entryOf(L.cols[colName(c)], 'width', WIDTH_MAX);
  const rows = {};
  for (let r = g.r1; r <= g.r2; r += 1) rows[String(r)] = entryOf(L.rows[String(r)], 'height', HEIGHT_MAX);
  return { ...base, columns, rows, caveats: CAVEATS };
}

module.exports = { projectLayoutRead, layoutExtent, CEILING };
