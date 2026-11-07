'use strict';

/**
 * GH #300 L1: the PROJECTION behind the `read_hyperlinks` read tool -- a pure field-select over
 * `inv.hyperlinkReads[sheet]`, the driver's own in-session read (xlsx-driver.cjs readQuery hyperlinkSheets:
 * sheet.hyperlinks.read/v1 through the DS Bridge, patchVersion /18). Every hyperlink of the sheet, row-then-
 * column order:
 *   { ref, address, location, tooltip, kind }
 * `ref` is already $-stripped by the bridge handler itself; `address`/`location`/`tooltip` are string or null;
 * `kind` is 'external' | 'internal' | 'unknown:<n>', from the engine's own hyperlink-type code. The DISPLAYED
 * text is NOT part of this output -- it is the cell's own value (read_range covers that already,
 * docs/DS-9.3.5-BRIDGE18-HYPERLINKS-TABLES-READ.md). On a /16 or /17 runtime (or any other read gap) the WHOLE
 * axis degrades to bridgeAxis's own named 'UNOBSERVABLE:<reason>' string -- CAPABILITY_MISSING, never a
 * guessed value:
 *   the sheet has no hyperlink                -> { sheet, links: [] }
 *   the read could not be done                 -> { sheet, links: null, reason: 'UNOBSERVABLE:...' }
 */

function projectHyperlinksRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND' };
  const read = inv.hyperlinkReads ? inv.hyperlinkReads[a.sheet] : undefined;
  if (read === undefined) return { sheet: a.sheet, links: null, reason: 'UNOBSERVABLE:not-read' };
  if (typeof read === 'string') return { sheet: a.sheet, links: null, reason: read };
  if (!read || !Array.isArray(read.links)) return { sheet: a.sheet, links: null, reason: 'UNOBSERVABLE:malformed-read' };
  return { sheet: a.sheet, links: read.links.map(l => ({ ref: l.ref, address: l.address, location: l.location, tooltip: l.tooltip, kind: l.kind })) };
}

module.exports = { projectHyperlinksRead };
