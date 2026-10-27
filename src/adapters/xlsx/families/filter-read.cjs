'use strict';

/**
 * L-10 (c): the PROJECTION behind the `read_filter` read tool -- a pure field-select over `inv.filterReads[sheet]` (the driver's own in-session read: the PUBLIC AutoFilter getters
 * GetRange / GetFilters -> GetOn / GetOperator / GetCriteria1 / GetCriteria2, and each filter's field through the DS Bridge sheet.autoFilter.fieldIdentity.read/v1), shaped as the
 * MIRROR of the filter_range write schema (filter.cjs): `field` (1-based column inside the range) plus exactly one of `value` (one plain value), `values` (two different plain values,
 * xlOr, read back `=a` / `=b`) or `and` (two numeric comparisons `>n` / `<m`, xlAnd). No new bridge call.
 *
 * What the engine returns that the write schema has no word for is NEVER turned into a value: it is named. A value LIST (criteria1 an ARRAY, xlFilterValues) is read by its RAW SHAPE, never
 * stringified: 1 element -> value, 2 different plain elements -> values, 3+ or anything else -> unsupported (L-10b LIVE r1: String(['2','4']) had turned a two-value list into the one value '2,4').
 * GH #301 D2 (AF-02 LIVE): 3..10 different plain elements under xlFilterValues -> values; xlTop10Items with a whole-number criteria1 -> top; xlFilterDynamic with
 * xlFilterAboveAverage / xlFilterBelowAverage -> dynamic. The cell-colour filter (criteria1 reads back null) and any other dynamic type stay named gaps.
 * `engineShapes` (same order as `filters`) is a DIAGNOSTIC: the raw operator and the TYPE of criteria1 / criteria2, so a LIVE run sees the engine's shape without going through the projection.
 *   - a criteria combination the writer cannot create   -> { field, unsupported: true, engine: { operator, criteria1, criteria2 }, reason: 'NO_WRITE_PATH: ...' }
 *   - an inactive filter (a cleared one keeps its controls, B-D13)  -> { field, on: false }
 *   - a sheet whose filter could not be read             -> { range: null, filters: null, reason: 'UNOBSERVABLE:...' }  (never "no filter")
 * Which rows are hidden is NOT part of this read (filter.cjs: not measured).
 */

const SINGLE_OPERATORS = new Set(['', 'xlAnd', 'xlOr', 'xlFilterValues']);
const COMPARE = /^[<>]-?\d+(\.\d+)?$/; // filter.cjs COMPARE: only > and < were measured
const bare = x => String(x === null || x === undefined ? '' : x).replace(/^=/, '');
const empty = x => x === null || x === undefined || x === '';
const plain = v => v !== '' && !/^[=<>]/.test(v);
const engineOf = f => ({ operator: f.operator, criteria1: f.criteria1, criteria2: f.criteria2 });
const unsupported = (field, f, why) => ({ field, unsupported: true, engine: engineOf(f), reason: `NO_WRITE_PATH: ${why}` });

const DYNAMIC_OF = Object.freeze({ xlFilterAboveAverage: 'aboveAverage', xlFilterBelowAverage: 'belowAverage' }); // filter.cjs DYNAMIC, inverted
const LIST_OPERATORS = new Set(['', 'xlFilterValues', 'xlOr']); // xlFilterValues = how SetAutoFilter(field, [a, b]) is made; xlOr = how filter.cjs documents its read-back
/** A criteria1 ARRAY (a value list) by its raw shape; only the writer's two list shapes are values. */
function listOf(field, f) {
  const items = f.criteria1.map(x => (typeof x === 'string' || Number.isFinite(x) ? String(x) : null));
  if (!empty(f.criteria2) || !LIST_OPERATORS.has(f.operator ?? '') || items.some(x => x === null || !plain(x))) return unsupported(field, f, 'a value list is writable only as 1 or 2 different plain values (xlFilterValues, no second criterion)');
  if (items.length === 1) return { field, value: items[0] };
  if (items.length === 2) return items[0] !== items[1] ? { field, values: items } : unsupported(field, f, 'a value list is writable only as different plain values (xlFilterValues, no second criterion)');
  // GH #301 D2: 3..10 different plain values, only as xlFilterValues (how SetAutoFilter(field, [..]) reads back)
  return f.operator === 'xlFilterValues' && items.length >= 3 && items.length <= 10 && new Set(items).size === items.length
    ? { field, values: items }
    : unsupported(field, f, 'a value list is writable only as 2 to 10 different plain values (xlFilterValues, no second criterion)');
}

function criteriaOf(field, f) {
  if (f.on === false) return { field, on: false };
  if (Array.isArray(f.criteria1) || Array.isArray(f.criteria2)) return Array.isArray(f.criteria1) && !Array.isArray(f.criteria2) ? listOf(field, f) : unsupported(field, f, 'criteria2 is a list, not writable');
  const c1 = bare(f.criteria1);
  // GH #301 D2: the two non-value single criteria the writer can leave behind
  if (empty(f.criteria2) && f.operator === 'xlTop10Items') return Number.isInteger(f.criteria1) && f.criteria1 >= 1 && f.criteria1 <= 500 ? { field, top: f.criteria1 } : unsupported(field, f, 'a top-10 filter is writable only as a whole number of items, 1 to 500');
  if (empty(f.criteria2) && f.operator === 'xlFilterDynamic') return Object.hasOwn(DYNAMIC_OF, f.criteria1) ? { field, dynamic: DYNAMIC_OF[f.criteria1] } : unsupported(field, f, 'a dynamic filter is writable only as above / below average');
  if (empty(f.criteria2)) {
    // a single plain criterion is a one-value filter only for the operators the writer's SetAutoFilter(field, value) can leave behind; any other operator (top-10, colour, ...) is not a value
    if (plain(c1) && SINGLE_OPERATORS.has(f.operator ?? '')) return { field, value: c1 };
    return unsupported(field, f, COMPARE.test(c1) ? 'a single comparison is not writable (only value / values / and)' : 'criteria1 is not one plain value');
  }
  if (f.operator === 'xlOr') {
    const c2 = bare(f.criteria2);
    return plain(c1) && plain(c2) && c1 !== c2 && String(f.criteria1).startsWith('=') && String(f.criteria2).startsWith('=')
      ? { field, values: [c1, c2] }
      : unsupported(field, f, 'an xlOr pair is writable only as two different plain `=value` criteria');
  }
  if (f.operator === 'xlAnd') {
    return COMPARE.test(String(f.criteria1)) && COMPARE.test(String(f.criteria2))
      ? { field, and: [String(f.criteria1), String(f.criteria2)] }
      : unsupported(field, f, 'an xlAnd pair is writable only as two numeric `>n` / `<n` comparisons');
  }
  return unsupported(field, f, `operator ${JSON.stringify(f.operator)} is not writable`);
}

const typeOf = x => (x === null ? 'null' : Array.isArray(x) ? 'array' : typeof x);
const shapeOf = f => (f && typeof f === 'object'
  ? { field: f.field, operator: f.operator, criteria1: typeOf(f.criteria1), ...(Array.isArray(f.criteria1) ? { criteria1Length: f.criteria1.length } : {}), criteria2: typeOf(f.criteria2), ...(Array.isArray(f.criteria2) ? { criteria2Length: f.criteria2.length } : {}) }
  : { shape: typeOf(f) });

function filterOf(f) {
  if (!f || typeof f !== 'object') return { reason: 'UNOBSERVABLE:unexpected-shape' };
  if (!Number.isInteger(f.field) || f.field < 1) return { unsupported: true, engine: engineOf(f), reason: 'UNOBSERVABLE:field-identity-not-a-column-number' };
  return criteriaOf(f.field, f);
}

/** inv.filterReads[sheet] -> the tool result. A string is a whole-sheet gap, never guessed as "no filter". */
function projectFilterRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND' };
  const r = inv.filterReads ? inv.filterReads[a.sheet] : undefined;
  if (r === undefined || r === null) return { sheet: a.sheet, range: null, filters: null, reason: 'UNOBSERVABLE:not-read' };
  if (typeof r === 'string') return { sheet: a.sheet, range: null, filters: null, reason: r };
  if (typeof r !== 'object' || !Array.isArray(r.filters) || (r.range !== null && typeof r.range !== 'string')) return { sheet: a.sheet, range: null, filters: null, reason: 'UNOBSERVABLE:unexpected-shape' };
  return { sheet: a.sheet, range: r.range === null ? null : r.range.replace(/\$/g, ''), filters: r.filters.map(filterOf), engineShapes: r.filters.map(shapeOf) };
}

module.exports = { projectFilterRead };
