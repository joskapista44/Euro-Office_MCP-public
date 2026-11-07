'use strict';

/**
 * GH #338-#342: the PROJECTION behind the `read_protection` read tool -- the READ mirror of protect_sheet / unprotect_sheet /
 * set_cell_protection (protection-write.cjs). A pure field-select over the #326 guard's own axes, the driver's in-session read
 * (xlsx-driver.cjs readQuery protectionSheets / protectionRanges: sheet.protection.read/v1 through the DS Bridge, /17+):
 *   no range  -> { sheet, protected, forbids, workbook }                 forbids: the 15 flags (true = forbidden), null when unprotected
 *   a range   -> { sheet, range, protected, forbids, cells: { nonEmpty, unlocked, hidden }, workbook }   over the NON-EMPTY cells of the range
 *                (an unlisted cell is the Excel default: locked, not hidden)
 * The read never carries a credential (the handler reads only the named booleans, docs/DS-9.3.5-BRIDGE17-PROTECTION-READ.md): a
 * password-protected sheet reads as protected, and its password is neither shown nor hinted. A read gap -> the named reason, never a
 * guessed state.
 *
 * GH #338-#342 (/24 only, W1, docs/DS-9.3.5-BRIDGE24-WORKBOOK-PROTECTION-DESIGN.md): `workbook: { lockStructure }` is
 * ADDITIVE -- a workbook-wide axis, independent of the requested sheet's own existence or protection state, so it is
 * attached to every response including SHEET_NOT_FOUND. On a pre-/24 runtime the driver's workbookProtectionRead
 * degrades to the named 'UNOBSERVABLE:CAPABILITY_MISSING' string (the SAME convention the per-sheet axis above
 * already uses); this projection never guesses `lockStructure` in that case.
 */

function workbookProtectionField(inv) {
  const read = inv.workbookProtectionRead;
  if (read === undefined) return { lockStructure: null, reason: 'UNOBSERVABLE:not-read' };
  if (typeof read === 'string') return { lockStructure: null, reason: read };
  if (!read || typeof read.lockStructure !== 'boolean') return { lockStructure: null, reason: 'UNOBSERVABLE:malformed-read' };
  return { lockStructure: read.lockStructure };
}

function projectProtectionRead(inv, a) {
  const workbook = workbookProtectionField(inv);
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND', workbook };
  const ranged = typeof a.range === 'string';
  const read = ranged ? (inv.protectionRangeReads || {})[`${a.sheet}!${a.range}`] : (inv.protectionReads || {})[a.sheet];
  const where = ranged ? { sheet: a.sheet, range: a.range } : { sheet: a.sheet };
  if (read === undefined) return { ...where, protected: null, reason: 'UNOBSERVABLE:not-read', workbook };
  if (typeof read === 'string') return { ...where, protected: null, reason: read, workbook };
  if (!read || typeof read.protected !== 'boolean') return { ...where, protected: null, reason: 'UNOBSERVABLE:malformed-read', workbook };
  const out = { ...where, protected: read.protected, forbids: read.protected ? { ...read.forbids } : null, workbook };
  if (ranged) {
    const c = read.cells;
    if (!c || !Array.isArray(c.unlocked) || !Array.isArray(c.hidden)) return { ...out, cells: null, reason: 'UNOBSERVABLE:cells-not-read' };
    out.cells = { nonEmpty: c.nonEmpty, unlocked: c.unlocked.slice(), hidden: c.hidden.slice() };
  }
  return out;
}

module.exports = { projectProtectionRead, workbookProtectionField };
