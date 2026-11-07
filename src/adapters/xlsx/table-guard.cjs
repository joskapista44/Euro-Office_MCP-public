'use strict';

const { parseRange, overlaps } = require('./a1.cjs');

/**
 * GH #331 (T06): a write over a range against the sheet's Excel Tables (read_tables' own `tableReads`, cell Bridge /18).
 * Measured LIVE on the installed product (test/live/xlsx-gh331-t06-table-aware-baseline-probe.live.cjs): a merge across a
 * table's edge and a sort of a range only partly over a table are both APPLIED by the engine (Excel refuses both); a sort
 * of exactly the table's range sorts the table and keeps it. So the families that use this refuse an overlap, named.
 *
 * When the tables cannot be read (a runtime without the Bridge /18 read answers CAPABILITY_MISSING) the result is
 * `unobservable` and the caller keeps its behaviour from before T06 -- fail-closed there would refuse every merge and sort
 * on such a runtime; the gap is recorded in the caller's evidence, never read as "no table".
 *
 *   -> { table: null }                         no table overlaps the range
 *   -> { table: { name, ref }, exact }         the first overlapping table; exact = the range IS the table's range
 *   -> { unobservable: '<why>' }               the tables of the sheet could not be read
 */
function tableOverlap(p, sheet, range) {
  const read = p && p.tableReads ? p.tableReads[sheet] : undefined;
  if (read === undefined) return { unobservable: 'not-read' };
  if (typeof read === 'string') return { unobservable: read };
  if (!read || !Array.isArray(read.tables)) return { unobservable: 'malformed-read' };
  const g = parseRange(range);
  for (const t of read.tables) {
    const tg = parseRange(String(t.ref));
    if (!tg) return { unobservable: `table-ref-unparsable:${t.name}` };
    // GH #361 2. egyseg: autoFilter rides along additively (merge.cjs/sort.cjs ignore it, filter.cjs's own
    // table-owned-AutoFilter guard below is the only consumer) -- t.autoFilter is read_tables' own
    // presence-only boolean (tables-read.cjs), never guessed when the field itself is missing.
    if (overlaps(tg, g)) return { table: { name: t.name, ref: t.ref, autoFilter: t.autoFilter === true }, exact: tg.r1 === g.r1 && tg.r2 === g.r2 && tg.c1 === g.c1 && tg.c2 === g.c2 };
  }
  return { table: null };
}

/**
 * GH #361 2. egyseg: remove_autofilter /
 * clear_filter are BLANKET, sheet-wide operations on the ONE sheet-level AutoFilter the bare SetAutoFilter()
 * toggle addresses (filter.cjs's own callOf()/classify() documentation) -- unlike filter_range /
 * set_autofilter_range, their risk is not limited to whichever range happens to be asked: the LIVE measurement
 * (docs/evidence/gh361-table-autofilter-live-20261104/, R_EFFECT) showed remove_autofilter reporting SATISFIED
 * ("nothing of its own to remove at the sheet level") while a Table's own AutoFilter stayed completely intact.
 * So the question here is sheet-wide, not range-scoped: does ANY table on this sheet carry its own AutoFilter
 * at all?
 *
 *   -> { table: null }                  no table on the sheet carries an AutoFilter
 *   -> { table: { name, ref } }        the first table (document order) whose own AutoFilter flag is true
 *   -> { unobservable: '<why>' }       the sheet's tables could not be read
 */
function anyAutoFilterTable(p, sheet) {
  const read = p && p.tableReads ? p.tableReads[sheet] : undefined;
  if (read === undefined) return { unobservable: 'not-read' };
  if (typeof read === 'string') return { unobservable: read };
  if (!read || !Array.isArray(read.tables)) return { unobservable: 'malformed-read' };
  const hit = read.tables.find(t => t && t.autoFilter === true);
  return { table: hit ? { name: hit.name, ref: hit.ref, autoFilter: true } : null };
}

module.exports = { tableOverlap, anyAutoFilterTable };
