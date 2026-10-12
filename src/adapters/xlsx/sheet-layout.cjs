'use strict';

const { parseRange } = require('./a1.cjs');

/**
 * M7 S2a (owner GO 2026-09-30, owner channel): a sheet's AutoFilter, manual page breaks and frozen pane around a
 * structural step. Pure. docs/M7-S2-SCOPE-PLAN.md section 6, docs/evidence/M7-S2A-READ-PROBE.md.
 *
 * Readers (the driver's; every sheet is read before AND after, so the L3 claims exactly the read scope):
 *   filterReads[sheet]  the W6 AutoFilter reader: { range: 'A1:B5' | null, filters: [...] }
 *   printReads[sheet]   the W8 print READ (sheet.printSettings.read/v1, B09): rowBreaks [{ id, man, min, max }]; colBreaks
 *                       the same shape through the DS Bridge (sheet.columnBreaks.read/v1, B12; owner channel, asked
 *                       by `colBreaks: true`); a reader gap is 'UNOBSERVABLE:<why>' and refuses (never an empty set)
 *   panes[sheet]        the W4d pane READ, tri-state (DS Bridge, B06): { state: 'none' } | { state: 'frozen',
 *                       row0, col0 }; anything else (a split / unknown pane, a reader gap) is unobservable
 * An AUTOMATIC break (man false) is not measured (Talk 30096): present on any sheet, every structural step refuses.
 * A break `id` is the 0-based index of the row / column it precedes (print.cjs: a row break at N is id N-1). A frozen
 * pane's row0 / col0 = the number of frozen rows / columns.
 *
 * Positions on the step's axis, N = the 1-based at: BEFORE = a range ends at N-2 or earlier / a break id <= N-2 / the
 * frozen span <= N-2. Everything else (BOUNDARY, INSIDE, AFTER) is NOT_MEASURED and refused. The other axis is never
 * classified: a row step does not move column breaks or col0 -- the L3 verifies them unchanged.
 */

const unobs = x => typeof x === 'string' && x.startsWith('UNOBSERVABLE');
const isBreak = b => Number.isInteger(b && b.id) && b.id >= 0 && typeof b.man === 'boolean';
const isBreaks = v => Array.isArray(v) && v.every(isBreak);
const breaksGap = (v, axis) => ({ axis, why: unobs(v) ? v : v === undefined ? 'not-read' : 'unexpected-shape' });

/** The AutoFilter of a sheet: { filter: bbox | null } or { gap }. */
function filterOf(fr) {
  if (!fr || unobs(fr) || typeof fr !== 'object') return { gap: { axis: 'autofilter', why: unobs(fr) ? fr : 'not-read' } };
  if (fr.range === null) return { filter: null };
  const g = parseRange(String(fr.range).replace(/^.*!/, '').replace(/\$/g, ''));
  return g ? { filter: g } : { gap: { axis: 'autofilter', why: 'range-unparsable' } };
}

/** The row / column breaks of a sheet from the print READ: { rowBreaks, colBreaks } or gaps. */
function breaksOf(pr) {
  if (!pr || unobs(pr) || typeof pr !== 'object') return { gaps: [{ axis: 'page-breaks', why: unobs(pr) ? pr : 'not-read' }] };
  const gaps = [];
  if (!isBreaks(pr.rowBreaks)) gaps.push(breaksGap(pr.rowBreaks, 'row-breaks'));
  if (!isBreaks(pr.colBreaks)) gaps.push(breaksGap(pr.colBreaks, 'column-breaks'));
  return gaps.length ? { gaps } : { rowBreaks: pr.rowBreaks, colBreaks: pr.colBreaks, gaps };
}

/** A readable pane: none, or frozen with integer row0 / col0 >= 0 -- anything else (split / unknown) is a gap. */
const paneReadable = x => !!x && typeof x === 'object' && (x.state === 'none' || (x.state === 'frozen' && [x.row0, x.col0].every(v => Number.isInteger(v) && v >= 0)));

/** One sheet's layout objects, or the axes that cannot be read. */
function sheetLayout(p, sheet) {
  const f = filterOf(p.filterReads ? p.filterReads[sheet] : undefined);
  const br = breaksOf(p.printReads ? p.printReads[sheet] : undefined);
  const pane = p.panes ? p.panes[sheet] : undefined;
  const paneGap = paneReadable(pane) ? [] : [{ axis: 'pane', why: unobs(pane) ? pane : (pane ? 'split-or-unknown' : 'not-read') }];
  const gaps = [...(f.gap ? [f.gap] : []), ...br.gaps, ...paneGap];
  return { sheet, gaps, filter: f.filter || null, rowBreaks: br.rowBreaks || null, colBreaks: br.colBreaks || null, pane: paneGap.length ? null : pane };
}

/** The shifted sheet's objects that are not BEFORE the cut on the step's axis (the other axis is never classified). */
function notBefore(l, shift) {
  const rows = shift.axis === 'rows';
  const before = last => last <= shift.at - 2;
  const hits = [];
  if (l.filter && !before(rows ? l.filter.r2 : l.filter.c2)) hits.push({ object: 'autofilter', range: l.filter, axis: shift.axis });
  for (const b of rows ? l.rowBreaks : l.colBreaks) if (!before(b.id)) hits.push({ object: rows ? 'row-break' : 'column-break', id: b.id, man: b.man });
  // a pane frozen on the OTHER axis only (0 here) still has its frozen cell at index 0: a step at 1 touches it
  if (l.pane.state === 'frozen' && !before(rows ? l.pane.row0 : l.pane.col0)) hits.push({ object: 'frozen-pane', row0: l.pane.row0, col0: l.pane.col0 });
  return hits;
}

/**
 * The pre-write guard of a structural step ({ sheet, axis: 'rows' | 'cols', at, count }): null when every object of
 * the shifted sheet lies BEFORE the cut on the step's axis and every axis of every sheet is readable; else why not.
 */
function layoutGuard(p, shift) {
  const all = (p.sheets || []).map(s => sheetLayout(p, s.name));
  const gaps = all.flatMap(l => l.gaps.map(g => ({ sheet: l.sheet, ...g })));
  const l = all.find(x => x.sheet === shift.sheet);
  if (!l) gaps.push({ sheet: shift.sheet, axis: 'sheet', why: 'not-read' });
  if (gaps.length) return { reason: 'xlsx-structure-layout-unobservable', unobservable: gaps.slice(0, 20), count: gaps.length };
  const auto = all.flatMap(x => [...x.rowBreaks.filter(b => !b.man).map(b => ({ sheet: x.sheet, object: 'row-break', id: b.id })),
    ...x.colBreaks.filter(b => !b.man).map(b => ({ sheet: x.sheet, object: 'column-break', id: b.id }))]);
  if (auto.length) return { reason: 'xlsx-structure-layout-automatic-break-not-measured', objects: auto.slice(0, 20), count: auto.length };
  const hits = notBefore(l, shift);
  return hits.length ? { reason: 'xlsx-structure-layout-object-not-before-the-cut', objects: hits.slice(0, 20), count: hits.length } : null;
}

/** L3: every sheet's AutoFilter, row / column breaks and pane equal before and after; an unreadable side differs. */
function layoutDiffs(b, a) {
  if (!b || !a) return [{ axis: 'observation-missing' }];
  const sheets = [...new Set([...(b.sheets || []), ...(a.sheets || [])].map(s => s.name))];
  const diffs = [];
  for (const n of sheets) {
    const x = sheetLayout(b, n);
    const y = sheetLayout(a, n);
    if (x.gaps.length || y.gaps.length) { diffs.push({ sheet: n, axis: 'unobservable', before: x.gaps, after: y.gaps }); continue; }
    const fr = s => JSON.stringify(s.filterReads[n]);
    if (fr(b) !== fr(a)) diffs.push({ sheet: n, axis: 'autofilter' });
    if (JSON.stringify(x.rowBreaks) !== JSON.stringify(y.rowBreaks)) diffs.push({ sheet: n, axis: 'row-breaks' });
    if (JSON.stringify(x.colBreaks) !== JSON.stringify(y.colBreaks)) diffs.push({ sheet: n, axis: 'column-breaks' });
    if (JSON.stringify(x.pane) !== JSON.stringify(y.pane)) diffs.push({ sheet: n, axis: 'pane' });
  }
  return diffs;
}

module.exports = { sheetLayout, layoutGuard, layoutDiffs };
