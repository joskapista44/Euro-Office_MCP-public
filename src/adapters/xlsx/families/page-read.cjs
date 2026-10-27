'use strict';

const { colName } = require('../a1.cjs');
const { PAPER } = require('./print.cjs');

/**
 * L-10 (b): the PROJECTION behind the `read_page_setup` read tool -- a pure field-select over TWO of the driver's own in-session reads, `inv.pageReads[sheet]` (the PUBLIC page getters,
 * set_page_layout, page.cjs) and `inv.printReads[sheet]` (the DS Bridge sheet.printSettings.read/v1, the print family), shaped as the MIRROR of those write schemas: the same field names
 * and units (orientation 'portrait' | 'landscape', margins in millimetres, fitToWidth / fitToHeight, header / footer codes, differentFirst / differentOddEven, manual page breaks as the `at`
 * of set_page_break, first page number, centering, scale, a paper name from the editor's own table). No new bridge call. The two sources fail INDEPENDENTLY: a gap in one is a named
 * reason on that half only.
 *
 * What the engine returns that the write schema has no word for is NEVER turned into a value: it is named.
 *   - a paper size that is not one of the editor's sizes      -> paper: null + unrepresentable [{ field: 'paper', engine: { width, height }, ... }]
 *   - an AUTOMATIC page break (man false -- not measured, never writable) -> counted in `automaticBreaks`, never listed
 *   - an out-of-schema value (orientation, margin, scale ...)  -> the field is null + unrepresentable
 *   - a half that could not be read                            -> { reason: 'UNOBSERVABLE:...' }
 * NOT in this read: print area and print titles are defined names (set_print_area / set_print_titles, names.cjs), read through the defined-names reader, not these two -- `caveats` says so.
 */

const CAVEATS = Object.freeze(['print area and print titles are defined names (set_print_area / set_print_titles), not part of this read']);
const ORIENT_OF = Object.freeze({ xlPortrait: 'portrait', xlLandscape: 'landscape' });
const MARGINS = Object.freeze(['topMargin', 'bottomMargin', 'leftMargin', 'rightMargin']);
const MAX_MARGIN_MM = 200;      // page.cjs
const SCALE_MIN = 10; const SCALE_MAX = 400; // print.cjs set_print_scale
const HF_SLOTS = Object.freeze(['oddHeader', 'oddFooter', 'firstHeader', 'firstFooter', 'evenHeader', 'evenFooter']);
const PAPER_OF = Object.freeze(Object.fromEntries(Object.entries(PAPER).map(([name, [w, h]]) => [`${w}x${h}`, name])));

const bad = (field, engine, reason) => ({ field, engine, reason });
const isNull = x => x === null || x === undefined;
const textOrNull = x => x === null || typeof x === 'string';

function pageHalf(r) {
  if (isNull(r)) return { reason: 'UNOBSERVABLE:not-read' };
  if (typeof r === 'string') return { reason: r };
  if (typeof r !== 'object') return { reason: 'UNOBSERVABLE:unexpected-shape' };
  const list = [];
  const out = {};
  if (Object.hasOwn(ORIENT_OF, r.orientation)) out.orientation = ORIENT_OF[r.orientation]; else { out.orientation = null; list.push(bad('orientation', r.orientation, "not xlPortrait / xlLandscape")); }
  for (const k of ['printGridlines', 'printHeadings']) { if (typeof r[k] === 'boolean') out[k] = r[k]; else { out[k] = null; list.push(bad(k, r[k], 'not a boolean')); } }
  for (const k of MARGINS) { if (Number.isFinite(r[k]) && r[k] >= 0 && r[k] <= MAX_MARGIN_MM) out[k] = r[k]; else { out[k] = null; list.push(bad(k, r[k], `not millimetres in [0, ${MAX_MARGIN_MM}]`)); } }
  return list.length ? { ...out, unrepresentable: list } : out;
}

const manualAts = (breaks, axis) => breaks.filter(b => b && b.man === true && Number.isInteger(b.id) && b.id >= 1).map(b => (axis === 'row' ? b.id + 1 : colName(b.id + 1)));
const autoCount = breaks => breaks.filter(b => !(b && b.man === true)).length;

function paperOf(r, list) {
  const name = PAPER_OF[`${r.width}x${r.height}`];
  if (name) return name;
  list.push(bad('paper', { width: r.width, height: r.height }, "not one of the editor's paper sizes"));
  return null;
}

function printHalf(r) {
  if (isNull(r)) return { reason: 'UNOBSERVABLE:not-read' };
  if (typeof r === 'string') return { reason: r };
  if (typeof r !== 'object' || !Array.isArray(r.rowBreaks) || !Array.isArray(r.colBreaks) || !Number.isInteger(r.fitW) || !Number.isInteger(r.fitH)) return { reason: 'UNOBSERVABLE:unexpected-shape' };
  const list = [];
  const out = { fitToWidth: r.fitW, fitToHeight: r.fitH, fitToPage: r.fitFlag === true };
  if (Number.isInteger(r.scale) && r.scale >= SCALE_MIN && r.scale <= SCALE_MAX) out.scale = r.scale; else { out.scale = null; list.push(bad('scale', r.scale, `not an integer percentage in [${SCALE_MIN}, ${SCALE_MAX}]`)); }
  out.paper = paperOf(r, list);
  for (const k of HF_SLOTS) { if (textOrNull(r[k])) out[k] = r[k] ?? null; else { out[k] = null; list.push(bad(k, r[k], 'not a header / footer code or null')); } }
  for (const [k, from] of [['differentFirst', 'diffFirst'], ['differentOddEven', 'diffOddEven'], ['horizontal', 'hCenter'], ['vertical', 'vCenter']]) {
    if (typeof r[from] === 'boolean') out[k] = r[from]; else { out[k] = null; list.push(bad(k, r[from], 'not a boolean')); }
  }
  out.firstPageNumber = r.useFirstPage === true && Number.isInteger(r.firstPage) ? r.firstPage : null;
  out.rowBreaks = manualAts(r.rowBreaks, 'row');
  out.columnBreaks = manualAts(r.colBreaks, 'column');
  out.automaticBreaks = { rows: autoCount(r.rowBreaks), columns: autoCount(r.colBreaks) };
  return list.length ? { ...out, unrepresentable: list } : out;
}

/** inv.pageReads + inv.printReads -> the tool result. Each half is independent; a string is a whole-half gap, never guessed as "defaults". */
function projectPageSetupRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND' };
  return {
    sheet: a.sheet,
    page: pageHalf(inv.pageReads ? inv.pageReads[a.sheet] : undefined),
    print: printHalf(inv.printReads ? inv.printReads[a.sheet] : undefined),
    caveats: CAVEATS,
  };
}

module.exports = { projectPageSetupRead, PAPER_OF };
