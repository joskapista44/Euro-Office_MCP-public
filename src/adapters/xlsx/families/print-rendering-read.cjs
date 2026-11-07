'use strict';

/**
 * GH #356 / #357: the PROJECTION behind the `read_print_rendering` read tool (contract:
 * docs/GH357-READ-PRINT-RENDERING-CONTRACT-DESIGN-20261031.md) -- a pure field-select over `inv.printPageReads[sheet]`, the
 * driver's own in-session read (xlsx-driver.cjs readQuery printPageSheets: sheet.printPageCount.read/v1 through the DS Bridge,
 * patchVersion /23; docs/DS-9.3.5-BRIDGE23-PRINT-PAGE-COUNT-READ-DESIGN.md). The editor's own rendered page count of ONE sheet
 * under its stored print settings; `pages` and `reason` exclude each other -- never an estimate derived from the stored fit /
 * scale, never 0 for a hidden sheet, never a truncated count:
 *   counted                    -> { sheet, pages: n, authority: 'bridge:sheet.printPageCount.read/v1', capped: false, scope: 'sheet' }
 *   at the engine cap (1500)   -> { sheet, pages: null, authority: 'none', capped: true, scope: 'sheet', reason: 'UNOBSERVABLE:BRIDGE:page-count-capped' }
 *   any other gap              -> { sheet, pages: null, authority: 'none', capped: false, scope: 'sheet', reason: 'UNOBSERVABLE:<reason>' }
 *     (BRIDGE:sheet-hidden, CAPABILITY_MISSING on a /16../22 runtime, BRIDGE:print-calc-threw, not-read, ...) -- the driver's string verbatim
 */

const AUTHORITY = 'bridge:sheet.printPageCount.read/v1';
/** the bridge client prefixes the handler's own reason (UNOBSERVABLE:BRIDGE:<reason>); the cap is told by that reason alone */
const isCapped = reason => /(^|:)page-count-capped$/.test(reason);

function projectPrintRenderingRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND' };
  const gap = reason => ({ sheet: a.sheet, pages: null, authority: 'none', capped: isCapped(reason), scope: 'sheet', reason });
  const read = inv.printPageReads ? inv.printPageReads[a.sheet] : undefined;
  if (read === undefined) return gap('UNOBSERVABLE:not-read');
  if (typeof read === 'string') return gap(read.startsWith('UNOBSERVABLE:') ? read : `UNOBSERVABLE:${read}`);
  if (!read || !Number.isInteger(read.pages) || read.pages < 0) return gap('UNOBSERVABLE:malformed-read');
  return { sheet: a.sheet, pages: read.pages, authority: AUTHORITY, capped: false, scope: 'sheet' };
}

module.exports = { projectPrintRenderingRead, AUTHORITY };
