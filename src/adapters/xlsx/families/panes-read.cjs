'use strict';

/**
 * L-10 (a): the PROJECTION behind the `read_freeze_panes` read tool -- a pure field-select over `inv.panes[sheet]` (the driver's own in-session read of every sheet's OWN pane through
 * the DS Bridge sheet.freezeState.read/v1, "read only when a freeze step asked", xlsx-driver.cjs), shaped as the MIRROR of the freeze_panes write schema (freeze.cjs):
 * `rows` = the number of frozen rows (the pane's row0), `columns` = the number of frozen columns (col0), each present only when > 0 (freeze_panes takes rows and / or columns).
 * No new bridge call. What the reader cannot say is NAMED (a split / unknown pane, a bridge gap), never "not frozen".
 */

const MAX_ROWS = 100000; // freeze.cjs MAX_ROWS (X-L09)
const MAX_COLS = 1000;   // freeze.cjs MAX_COLS (X-L09)

const gap = (a, reason) => ({ sheet: a.sheet, frozen: null, reason });
const intOk = n => Number.isInteger(n) && n >= 0;
const bad = (field, engine, reason) => ({ field, engine, reason });

function projectPanesRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND' };
  const p = inv.panes ? inv.panes[a.sheet] : undefined;
  if (p === undefined || p === null) return gap(a, 'UNOBSERVABLE:not-read');
  if (typeof p === 'string') return gap(a, p);
  if (typeof p !== 'object') return gap(a, 'UNOBSERVABLE:unexpected-shape');
  if (p.state === 'none') return { sheet: a.sheet, frozen: false };
  if (p.state !== 'frozen' || !intOk(p.row0) || !intOk(p.col0) || (p.row0 === 0 && p.col0 === 0)) return gap(a, 'UNOBSERVABLE:split-or-unknown-pane');
  const out = { sheet: a.sheet, frozen: true };
  const list = [];
  if (p.row0 > 0) { out.rows = p.row0; if (p.row0 > MAX_ROWS) list.push(bad('rows', p.row0, `outside the writable [1, ${MAX_ROWS}]`)); }
  if (p.col0 > 0) { out.columns = p.col0; if (p.col0 > MAX_COLS) list.push(bad('columns', p.col0, `outside the writable [1, ${MAX_COLS}]`)); }
  return list.length ? { ...out, unrepresentable: list } : out;
}

module.exports = { projectPanesRead };
