'use strict';

/**
 * GH #133: the PROJECTION behind the `read_sheet_state` read tool -- a pure field-select over the driver's own reads, no new call:
 * `inv.visibility` (the public ApiWorksheet.GetVisible(), a boolean per sheet) and `inv.active` (the public
 * Api.GetActiveSheet().GetName()), plus the sheet list (name, index) the inventory already carries.
 *
 * Visibility is the two states the public getter can tell apart: true -> 'visible', false -> 'hidden'. The getter is boolean, so a
 * very-hidden sheet is NOT told apart from a hidden one (whether GetVisible() reports it at all is not measured) -- no third value is invented.
 * Anything the driver could not read is NAMED, never a default:
 *   - a sheet whose visibility is a gap string  -> { name, index, visibility: null, reason: 'UNOBSERVABLE:<why>' }
 *   - no visibility axis / a sheet missing from it / an unexpected value -> reason 'UNOBSERVABLE:not-read' / 'UNOBSERVABLE:unexpected-shape'
 *   - the active sheet: a name that is one of the workbook's sheets, else { active: null, activeReason: 'UNOBSERVABLE:<why>' } --
 *     a gap string (inventory.cjs keeps any string) or a name that is not a sheet of this workbook is never passed off as the active sheet.
 * The active sheet may itself be hidden; it is reported as the editor reports it.
 */

const isGap = v => typeof v === 'string' && v.startsWith('UNOBSERVABLE:');

function visibilityOf(vis, name) {
  const v = vis ? vis[name] : undefined;
  if (v === true) return { visibility: 'visible' };
  if (v === false) return { visibility: 'hidden' };
  if (isGap(v)) return { visibility: null, reason: v };
  return { visibility: null, reason: v === undefined || v === null ? 'UNOBSERVABLE:not-read' : 'UNOBSERVABLE:unexpected-shape' };
}

function projectSheetStateRead(inv) {
  const list = Array.isArray(inv.sheets) ? inv.sheets : [];
  const vis = inv.visibility && typeof inv.visibility === 'object' ? inv.visibility : null;
  const sheets = list.map(s => ({ name: s.name, index: s.index, ...visibilityOf(vis, s.name) }));
  const a = inv.active;
  if (typeof a !== 'string' || a === '') return { sheets, active: null, activeReason: 'UNOBSERVABLE:not-read' };
  if (isGap(a)) return { sheets, active: null, activeReason: a };
  if (!list.some(s => s.name === a)) return { sheets, active: null, activeReason: 'UNOBSERVABLE:active-sheet-not-in-workbook' };
  return { sheets, active: a };
}

module.exports = { projectSheetStateRead };
