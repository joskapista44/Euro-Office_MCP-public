'use strict';

/**
 * GH #298 D: the PROJECTION behind the `read_tables` read tool -- a pure field-select over `inv.tableReads[sheet]`,
 * the driver's own in-session read (xlsx-driver.cjs readQuery tableSheets: sheet.tables.read/v1 through the DS
 * Bridge, patchVersion /18). Every Excel Table of the sheet, document order:
 *   { name, ref, headerRow, totalsRow, columns, style, autoFilter }
 * `ref` is already $-stripped by the bridge handler itself; `style` is already shaped as { name, rowStripes,
 * columnStripes, firstColumn, lastColumn } or null (no style info); `autoFilter` is a boolean PRESENCE flag only
 * -- its own range/filter detail is out of this capability's scope (docs/DS-9.3.5-BRIDGE18-HYPERLINKS-TABLES-READ.md).
 * On a /16 or /17 runtime (or any other read gap) the WHOLE axis degrades to bridgeAxis's own named
 * 'UNOBSERVABLE:<reason>' string -- CAPABILITY_MISSING, never a guessed value:
 *   the sheet has no table                   -> { sheet, tables: [] }
 *   the read could not be done                -> { sheet, tables: null, reason: 'UNOBSERVABLE:...' }
 *
 * GH #361 2. egyseg: when a table DOES carry
 * its own AutoFilter (autoFilter: true), this read names what it does NOT see, rather than leaving the caller
 * to assume the presence flag is the whole story -- `autoFilterCriteria: 'UNOBSERVABLE:table-autofilter-
 * criteria-not-read'`, a NAMED hole, never a guessed value. The MEASURED gap: `read_filter` (filter-read.cjs,
 * the family's own criteria reader) is sheet-level only and never sees a table-owned AutoFilter at all
 * (docs/evidence/gh361-table-autofilter-live-20261104/, T_SCOPE_AFTER_CREATE: SHEET_BLIND_TO_TABLE_FILTER) --
 * so a table's own filter range/per-field criteria has NO reader anywhere in this product today. When
 * autoFilter is false (or not strictly true), the field is simply absent -- there is nothing to name a gap
 * for.
 *
 * GH #361 route probe, route (c): the driver ALSO tries a CANDIDATE capability (sheet.table.autoFilter.
 * criteria.read/v1, NOT installed anywhere -- documentserver-patches/9.3.5.0/euro-office-bridge-v1/
 * rev28-table-autofilter-criteria-candidate/) per table, attached as `t.autoFilterDetail`. On every runtime
 * accepted today this capability is CAPABILITY_MISSING, so `autoFilterCriteria` stays BYTE-IDENTICAL to the
 * string above (proven by this family's own pure test, not assumed). If the capability is ever installed and
 * the table's own range is readable, `autoFilterCriteria` becomes `{ range, fields }` -- `fields` is the
 * candidate handler's own MEASURED per-column shape (a LIVE run against a real, OOXML-preseeded fixture with
 * a table-owned AutoFilter): each entry `{ colId, values }` (a value filter, `values` the selected keys) OR
 * `{ colId, custom: { and, conditions: [{ operator, operatorName, value }] } }` (a custom filter); a column
 * of a kind never observed, or whose observed-but-unverified sub-fields are non-default, carries a named
 * `unverified` array instead of a guessed meaning. This projection picks each field KEY BY KEY (colId/values/
 * custom/unverified), never a blind spread of the handler's raw output.
 */

const CRITERIA_NOT_READ = 'UNOBSERVABLE:table-autofilter-criteria-not-read';

/** one FilterColumns entry, picked field-by-field (never a blind spread of the candidate handler's raw
 * output): colId plus EITHER values (an array of selected value-filter keys) OR custom ({and, conditions:
 * [{operator, operatorName, value}]}); unverified (an array of named internal-shape-unverified:<gap> strings)
 * carries any sub-field the handler could not measure, never a guessed value in its place. */
function fieldOf(f) {
  if (!f || typeof f !== 'object') return { colId: null };
  const out = { colId: f.colId };
  if (Array.isArray(f.values)) out.values = f.values;
  if (f.custom && typeof f.custom === 'object') {
    out.custom = { and: f.custom.and, conditions: Array.isArray(f.custom.conditions) ? f.custom.conditions.map(conditionOf) : [] };
  }
  if (Array.isArray(f.unverified)) out.unverified = f.unverified;
  return out;
}

function conditionOf(c) {
  return c && typeof c === 'object' ? { operator: c.operator, operatorName: c.operatorName, value: c.value } : { operator: null, operatorName: null, value: null };
}

/** the candidate capability's raw per-table result -> `autoFilterCriteria`'s value, never a guess: anything
 * other than a clean { present: true, range: <string> } falls back to the SAME named string the field has
 * always carried -- a missing/not-installed capability, a flag mismatch, or an unreadable range are all
 * equally "we do not have the criteria". A clean result with a `fields` array is projected field-by-field
 * (fieldOf); the rare shape of a string range with no fields array (should the handler ever produce one) still
 * falls back to its own named reason, never silently dropped. */
function autoFilterCriteriaOf(detail) {
  if (!detail || typeof detail !== 'object' || detail.present !== true || typeof detail.range !== 'string') return CRITERIA_NOT_READ;
  if (!Array.isArray(detail.fields)) return typeof detail.reason === 'string' ? `UNOBSERVABLE:${detail.reason}` : CRITERIA_NOT_READ;
  return { range: detail.range, fields: detail.fields.map(fieldOf) };
}

function projectTablesRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND' };
  const read = inv.tableReads ? inv.tableReads[a.sheet] : undefined;
  if (read === undefined) return { sheet: a.sheet, tables: null, reason: 'UNOBSERVABLE:not-read' };
  if (typeof read === 'string') return { sheet: a.sheet, tables: null, reason: read };
  if (!read || !Array.isArray(read.tables)) return { sheet: a.sheet, tables: null, reason: 'UNOBSERVABLE:malformed-read' };
  return {
    sheet: a.sheet,
    tables: read.tables.map(t => ({
      name: t.name, ref: t.ref, headerRow: t.headerRow, totalsRow: t.totalsRow, columns: t.columns, style: t.style, autoFilter: t.autoFilter,
      ...(t.autoFilter === true ? { autoFilterCriteria: autoFilterCriteriaOf(t.autoFilterDetail) } : {}),
    })),
  };
}

module.exports = { projectTablesRead, autoFilterCriteriaOf, CRITERIA_NOT_READ };
