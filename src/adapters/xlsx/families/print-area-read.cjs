'use strict';

const { parseRange, rangeText } = require('../a1.cjs');

/**
 * GH #134 (P1): the PROJECTION behind the `read_print_area` read tool -- a pure field-select over the driver's own in-session
 * defined-names read (`inv.definedNames`, the DS Bridge workbook.definedNames.read/v1 behind `dependents: true`), shaped as the
 * MIRROR of the set_print_area write schema (names.cjs planPrint): one rectangle -> `range`, two or more -> `areas`, in the
 * stored order (the engine keeps the given order, X-P02). No new bridge call.
 *
 * A print area is the SHEET-level name Print_Area (the engine stores it with the '_xlnm.' prefix stripped; the driver keeps it).
 *   no Print_Area on the sheet                          -> { set: false }   (absence, NOT unobservable)
 *   one, every part a plain rectangle on THIS sheet     -> { set: true, range } | { set: true, areas }
 *   a shape the writer has no word for (a whole column / row, a part on another sheet, two names) -> printArea 'UNOBSERVABLE:<why>'
 *     with the stored `ref` -- never turned into a value
 *   the names reader itself not read                    -> printArea 'UNOBSERVABLE:<why>'
 * NOT in this read: Print_Titles and every other defined name (#134 P2).
 */

const quote = s => (/^[A-Za-z_À-￿][\w.À-￿]*$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`);
const PART = /^('(?:[^']|'')+'|[^'!]+)!\$([A-Z]{1,3})\$(\d+)(?::\$([A-Z]{1,3})\$(\d+))?$/;

/** One stored part 'Sheet!$A$1:$B$2' -> the plain 'A1:B2' on `sheet`, or the reason it is not one. */
function partOf(text, sheet) {
  const m = PART.exec(text.trim());
  if (!m) return { why: 'not-a-plain-rectangle' };
  const owner = m[1].startsWith("'") ? m[1].slice(1, -1).replace(/''/g, "'") : m[1];
  if (owner !== sheet && m[1] !== quote(sheet)) return { why: 'part-on-another-sheet' };
  const g = parseRange(m[4] ? `${m[2]}${m[3]}:${m[4]}${m[5]}` : `${m[2]}${m[3]}`);
  return g ? { range: rangeText(g) } : { why: 'not-a-plain-rectangle' };
}

function namesGap(inv) {
  // dependencies.cjs collector: { axis, why, host }; the names gap is noted on axis 'defined-names'
  const u = inv.dependencies && Array.isArray(inv.dependencies.unobservable) ? inv.dependencies.unobservable.find(x => x && x.axis === 'defined-names') : null;
  return u && typeof u.why === 'string' ? u.why : 'UNOBSERVABLE:defined-names-not-read';
}

function projectPrintAreaRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND' };
  if (!Array.isArray(inv.definedNames)) return { sheet: a.sheet, printArea: namesGap(inv) };
  const hits = inv.definedNames.filter(d => d.scopeSheet === a.sheet && d.name.toLowerCase() === 'print_area');
  if (hits.length === 0) return { sheet: a.sheet, printArea: { set: false } };
  if (hits.length > 1) return { sheet: a.sheet, printArea: 'UNOBSERVABLE:more-than-one-print-area', refs: hits.map(d => d.ref) };
  const ref = String(hits[0].ref).replace(/^=/, '');
  const parts = ref.split(',').map(p => partOf(p, a.sheet));
  const bad = parts.find(p => p.why);
  if (bad) return { sheet: a.sheet, printArea: `UNOBSERVABLE:${bad.why}`, ref };
  return { sheet: a.sheet, printArea: parts.length === 1 ? { set: true, range: parts[0].range } : { set: true, areas: parts.map(p => p.range) } };
}

module.exports = { projectPrintAreaRead };
