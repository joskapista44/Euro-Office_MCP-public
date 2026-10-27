'use strict';

const { colNum, colName, MAX_COL } = require('../a1.cjs');

/**
 * GH #134 (P2): the PROJECTION behind the `read_defined_names` read tool -- a pure field-select over the driver's own in-session
 * defined-names read (`inv.definedNames`, the DS Bridge workbook.definedNames.read/v1 behind `dependents: true`, the same read
 * read_print_area and the names family use). No new bridge call.
 *
 * Every entry is shaped as the MIRROR of the set_defined_name step (names.cjs planName): `name`, `refersTo` with its '=', and
 * `scope: '<sheet>'` ONLY for a sheet-level name (omitted = workbook level, the write's own convention). Two built-in names
 * carry a marker, never a guess:
 *   Print_Area   -> `builtIn: 'Print_Area'` (its structured form is read_print_area's job)
 *   Print_Titles -> `builtIn: 'Print_Titles'` + `titles: { columns: 'A:B', rows: 'n:m' }` -- the set_print_titles shape (columns /
 *                   rows / both, each member qualified with the OWN sheet); any other stored shape -> titles 'UNOBSERVABLE:<why>'
 *                   with the raw refersTo kept next to it
 * The names keep the driver's own filter and order (hidden names and '_xlnm'-NAMED names are dropped there -- NOT read here).
 *   the names reader not read  -> names 'UNOBSERVABLE:<why>' (never an empty list)
 *   more than MAX_NAMES        -> names 'UNOBSERVABLE:more-than-<MAX_NAMES>-names' + count (never a partial list)
 */

const MAX_NAMES = 500;
const SHEET = "('(?:[^']|'')+'|[^'!,]+)";
const COLS = new RegExp(`^${SHEET}!\\$([A-Z]{1,3}):\\$([A-Z]{1,3})$`);
const ROWS = new RegExp(`^${SHEET}!\\$(\\d+):\\$(\\d+)$`);
const unquote = s => (s.startsWith("'") ? s.slice(1, -1).replace(/''/g, "'") : s);

function namesGap(inv) {
  // dependencies.cjs collector: { axis, why, host }; the names gap is noted on axis 'defined-names'
  const u = inv.dependencies && Array.isArray(inv.dependencies.unobservable) ? inv.dependencies.unobservable.find(x => x && x.axis === 'defined-names') : null;
  return u && typeof u.why === 'string' ? u.why : 'UNOBSERVABLE:defined-names-not-read';
}

/** A stored Print_Titles refersTo -> { columns, rows } (set_print_titles shape), or the reason it is not one. */
function titlesOf(ref, sheet) {
  const out = {};
  for (const part of ref.split(',')) {
    const c = COLS.exec(part.trim());
    const r = c ? null : ROWS.exec(part.trim());
    const m = c || r;
    if (!m) return 'UNOBSERVABLE:not-a-titles-shape';
    if (unquote(m[1]) !== sheet) return 'UNOBSERVABLE:part-on-another-sheet';
    const key = c ? 'columns' : 'rows';
    if (out[key]) return 'UNOBSERVABLE:not-a-titles-shape';
    if (c) {
      const a = colNum(m[2]);
      const b = colNum(m[3]);
      if (!(a <= b && b <= MAX_COL)) return 'UNOBSERVABLE:not-a-titles-shape';
      out.columns = `${colName(a)}:${colName(b)}`;
    } else {
      if (Number(m[2]) < 1 || Number(m[3]) < Number(m[2])) return 'UNOBSERVABLE:not-a-titles-shape';
      out.rows = `${Number(m[2])}:${Number(m[3])}`;
    }
  }
  // a fixed key order (columns, rows), whatever the stored member order: the same titles read the same
  return { ...(out.columns ? { columns: out.columns } : {}), ...(out.rows ? { rows: out.rows } : {}) };
}

function entryOf(d) {
  const raw = String(d.ref);
  const e = { name: d.name, refersTo: raw.startsWith('=') ? raw : `=${raw}`, ...(d.scopeSheet === null || d.scopeSheet === undefined ? {} : { scope: d.scopeSheet }) };
  const lower = String(d.name).toLowerCase();
  if (lower === 'print_area' && e.scope !== undefined) e.builtIn = 'Print_Area';
  if (lower === 'print_titles' && e.scope !== undefined) {
    e.builtIn = 'Print_Titles';
    e.titles = titlesOf(e.refersTo.slice(1), e.scope);
  }
  return e;
}

function projectDefinedNamesRead(inv) {
  if (!Array.isArray(inv.definedNames)) return { names: namesGap(inv) };
  if (inv.definedNames.length > MAX_NAMES) return { names: `UNOBSERVABLE:more-than-${MAX_NAMES}-names`, count: inv.definedNames.length };
  return { names: inv.definedNames.map(entryOf), count: inv.definedNames.length };
}

module.exports = { projectDefinedNamesRead, MAX_NAMES };
