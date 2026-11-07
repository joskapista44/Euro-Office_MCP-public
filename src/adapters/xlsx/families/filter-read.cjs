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
 * GH #344 / #345 (LIVE-measured, docs/evidence/gh344-gh345-filter-probe-live-r2-20261031.json, in session AND after a save + fresh reopen):
 *   - a COLOUR filter reads GetOperator xlFilterCellColor | xlFilterFontColor and GetCriteria1 / GetCriteria2 null; the public API has no colour getter
 *     (engine research: docs/GH344-GH345-FILTER-COLOR-EMPTY-ENGINE-RESEARCH-20261031.md), the colour lives only in the saved <colorFilter dxfId> -> <dxf>.
 *     It is named: { field, unsupported: true, kind: 'cellColor' | 'fontColor', color: null, engine, reason: 'COLOR_NOT_PUBLIC: ...' } -- `color` is null, never a guess,
 *     and the old NO_WRITE_PATH reason (the colour filter IS writable through SetAutoFilter, it just cannot be read back) is gone for this measured shape.
 *   - an EMPTY value list reads, in session, on:true xlFilterValues with GetCriteria1 []; it is saved as an empty <filterColumn/> and after a reopen the same field reads
 *     on:false (persistence MEASURED: a persisted empty list is indistinguishable from an inactive filter). `[]` itself is AMBIGUOUS to the public getters (Object.keys of
 *     Filters.Values; Blank and Dates are not exposed): nothing visible | blanks only | date groups only. It is named: { field, unsupported: true, values: [],
 *     ambiguous: ['empty-list', 'blanks-only', 'dates-only'], engine, reason: 'AMBIGUOUS_EMPTY_LIST: ...' } -- `values: []` is the raw list, NOT a claim that it is the empty list.
 *   Both keep `unsupported: true` (not a filter_range shape) so existing consumers keep skipping them. Only the MEASURED shapes get the new names; any other criteria under
 *   a colour operator, and `[]` under any other operator / with a second criterion, stay the NO_WRITE_PATH gap.
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

const COLOUR_KIND = Object.freeze({ xlFilterCellColor: 'cellColor', xlFilterFontColor: 'fontColor' });
const colourGap = (field, f) => ({ field, unsupported: true, kind: COLOUR_KIND[f.operator], color: null, engine: engineOf(f),
  reason: 'COLOR_NOT_PUBLIC: the colour of a colour filter has no public getter (GetCriteria1 is null); it is only in the saved <colorFilter dxfId> -> <dxf>' });
/**
 * GH #344 (Bridge /21 sheet.autoFilter.colorFilter.read/v1): the colour, read in the reader's own session from the filter's dxf. `colours` is the
 * capability value ({ columns: [{ field, kind, colour, slot }] }) or a named string (a runtime without it -> the old COLOR_NOT_PUBLIC entry, plus
 * `colorReason`). The colour is the capability's own form ({ form: 'rgb', hex } | { form: 'theme', theme, tint } | { form: 'indexed', indexed }), never
 * converted. Still `unsupported: true`: a colour filter is not a filter_range shape. A dxf without a colour, a column the capability does not list, or a
 * kind that disagrees with the public operator is named, never guessed.
 */
function colourOf(field, f, colours) {
  const gap = colourGap(field, f);
  if (colours === undefined) return gap;
  if (typeof colours === 'string') return { ...gap, colorReason: colours };
  const col = colours && Array.isArray(colours.columns) ? colours.columns.find(c => c && c.field === field) : undefined;
  if (!col) return { ...gap, colorReason: 'UNOBSERVABLE:colour-filter-column-not-listed' };
  if (col.kind !== gap.kind) return { ...gap, colorReason: `UNOBSERVABLE:colour-filter-kind-mismatch:${col.kind}` };
  if (col.colour === null) return { ...gap, reason: 'COLOR_NONE: the colour filter dxf holds no colour' };
  return { field, unsupported: true, kind: gap.kind, color: col.colour, engine: engineOf(f), reason: 'COLOUR_FILTER: written by filter_range { colour: { kind, rgb } }; its colour is read through sheet.autoFilter.colorFilter.read/v1' };
}
const emptyListGap = (field, f) => ({ field, unsupported: true, values: [], ambiguous: ['empty-list', 'blanks-only', 'dates-only'], engine: engineOf(f),
  reason: 'AMBIGUOUS_EMPTY_LIST: GetCriteria1 [] is Object.keys(Filters.Values); an empty list, a blanks-only list and a date-groups-only list all read [] (Blank and Dates are not public)' });

function criteriaOf(field, f, colours) {
  if (f.on === false) return { field, on: false };
  // GH #344: the MEASURED colour shape only (both criteria empty); anything else under a colour operator falls through to the old named gap
  if (Object.hasOwn(COLOUR_KIND, f.operator) && empty(f.criteria1) && empty(f.criteria2)) return colourOf(field, f, colours);
  // GH #345: the MEASURED empty list only (xlFilterValues, no second criterion)
  if (f.operator === 'xlFilterValues' && Array.isArray(f.criteria1) && f.criteria1.length === 0 && empty(f.criteria2)) return emptyListGap(field, f);
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

function filterOf(f, colours) {
  if (!f || typeof f !== 'object') return { reason: 'UNOBSERVABLE:unexpected-shape' };
  if (!Number.isInteger(f.field) || f.field < 1) return { unsupported: true, engine: engineOf(f), reason: 'UNOBSERVABLE:field-identity-not-a-column-number' };
  return criteriaOf(f.field, f, colours);
}

/** inv.filterReads[sheet] -> the tool result. A string is a whole-sheet gap, never guessed as "no filter". */
function projectFilterRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND' };
  const r = inv.filterReads ? inv.filterReads[a.sheet] : undefined;
  if (r === undefined || r === null) return { sheet: a.sheet, range: null, filters: null, reason: 'UNOBSERVABLE:not-read' };
  if (typeof r === 'string') return { sheet: a.sheet, range: null, filters: null, reason: r };
  if (typeof r !== 'object' || !Array.isArray(r.filters) || (r.range !== null && typeof r.range !== 'string')) return { sheet: a.sheet, range: null, filters: null, reason: 'UNOBSERVABLE:unexpected-shape' };
  return { sheet: a.sheet, range: r.range === null ? null : r.range.replace(/\$/g, ''), filters: r.filters.map(f => filterOf(f, r.colours)), engineShapes: r.filters.map(shapeOf) };
}

// GH #344 colour filter WRITE (internal tracker): `colourOf` is exported so filter.cjs's verifyPrimitive can resolve a
// written colour filter's colour through the SAME Bridge /21 route the reader uses, instead of duplicating the
// dxf/colours-capability resolution logic a second time.
module.exports = { projectFilterRead, colourOf };
