'use strict';

/**
 * GH #133: the PROJECTION behind the `read_sheet_state` read tool -- a pure field-select over the driver's own reads, no new call:
 * `inv.visibility` (the public ApiWorksheet.GetVisible(), a boolean per sheet) and `inv.active` (the public
 * Api.GetActiveSheet().GetName()), plus the sheet list (name, index) the inventory already carries.
 *
 * Visibility is the two states the public getter can tell apart: true -> 'visible', false -> 'hidden'. The getter is boolean, so a
 * very-hidden sheet is NOT told apart from a hidden one (whether GetVisible() reports it at all is not measured) -- no third value is invented.
 *
 * GH #365 2nd unit: every 'hidden' entry carries a NAMED limit, `visibilityDetail:
 * 'UNOBSERVABLE:hidden-vs-veryhidden-not-distinguished'` -- `visibility: 'hidden'` itself STAYS
 * unchanged, so the common, plain-hidden case is byte-identical to before this field existed. A
 * stored-OOXML read was considered and rejected: the engine itself loses the
 * veryHidden/hidden distinction on the FIRST save after loading a file (GetVisible() already reads
 * `false` for both, and that first save re-writes the sheet's stored `state` as plain `hidden`) --
 * so even reading the saved file's raw bytes would only be correct until the first save of the
 * session, not a durable fix. `visible` and already-`UNOBSERVABLE` entries get NO new field.
 *
 * GH #365 option C U2: the 2nd unit's own limit is now RESOLVED wherever the
 * driver's `veryHiddenReads` axis (xlsx-driver.cjs, sheet.veryHidden.read/v1 through the Bridge, /29
 * ONLY) answers a clean boolean for a sheet GetVisible() reported hidden -- `visibility` becomes
 * `'veryHidden'` (true) or STAYS `'hidden'` but WITHOUT `visibilityDetail` (false: the limit named
 * "not distinguished" no longer applies -- we now positively know it is plain hidden). Capability-
 * gated, never a fallback: on any runtime without /29, or when the read itself failed, `veryHiddenReads`
 * carries a named `'UNOBSERVABLE:<reason>'` string instead of a boolean, and the entry degrades to the
 * EXACT pre-U2 shape (`visibility: 'hidden'` + the same `visibilityDetail`) -- byte-identical to before
 * this unit, by construction (the string check below is the ONLY new branch; a visible sheet never has
 * a `veryHiddenReads` entry at all, matching the driver's own "never asked" convention).
 * Anything the driver could not read is NAMED, never a default:
 *   - a sheet whose visibility is a gap string  -> { name, index, visibility: null, reason: 'UNOBSERVABLE:<why>' }
 *   - no visibility axis / a sheet missing from it / an unexpected value -> reason 'UNOBSERVABLE:not-read' / 'UNOBSERVABLE:unexpected-shape'
 *   - the active sheet: a name that is one of the workbook's sheets, else { active: null, activeReason: 'UNOBSERVABLE:<why>' } --
 *     a gap string (inventory.cjs keeps any string) or a name that is not a sheet of this workbook is never passed off as the active sheet.
 * The active sheet may itself be hidden; it is reported as the editor reports it.
 */

const isGap = v => typeof v === 'string' && v.startsWith('UNOBSERVABLE:');

// GH #365 2nd unit: the named limit every 'hidden' entry now carries -- see the module header.
const HIDDEN_VS_VERYHIDDEN_DETAIL = 'UNOBSERVABLE:hidden-vs-veryhidden-not-distinguished';

function visibilityOf(vis, veryHidden, name) {
  const v = vis ? vis[name] : undefined;
  if (v === true) return { visibility: 'visible' };
  if (v === false) {
    // GH #365 option C U2: resolve the 2nd unit's own gap wherever the capability-gated veryHidden
    // axis answered a clean boolean for THIS sheet -- a string (CAPABILITY_MISSING or any other
    // named reason, including "never asked" when veryHidden itself is null on a pre-U2 inventory
    // shape) falls through to the untouched pre-U2 shape below.
    const vh = veryHidden ? veryHidden[name] : undefined;
    if (vh === true) return { visibility: 'veryHidden' };
    if (vh === false) return { visibility: 'hidden' };
    return { visibility: 'hidden', visibilityDetail: HIDDEN_VS_VERYHIDDEN_DETAIL };
  }
  if (isGap(v)) return { visibility: null, reason: v };
  return { visibility: null, reason: v === undefined || v === null ? 'UNOBSERVABLE:not-read' : 'UNOBSERVABLE:unexpected-shape' };
}

function projectSheetStateRead(inv) {
  const list = Array.isArray(inv.sheets) ? inv.sheets : [];
  const vis = inv.visibility && typeof inv.visibility === 'object' ? inv.visibility : null;
  const veryHidden = inv.veryHiddenReads && typeof inv.veryHiddenReads === 'object' ? inv.veryHiddenReads : null;
  const sheets = list.map(s => ({ name: s.name, index: s.index, ...visibilityOf(vis, veryHidden, s.name) }));
  const a = inv.active;
  if (typeof a !== 'string' || a === '') return { sheets, active: null, activeReason: 'UNOBSERVABLE:not-read' };
  if (isGap(a)) return { sheets, active: null, activeReason: a };
  if (!list.some(s => s.name === a)) return { sheets, active: null, activeReason: 'UNOBSERVABLE:active-sheet-not-in-workbook' };
  return { sheets, active: a };
}

module.exports = { projectSheetStateRead, HIDDEN_VS_VERYHIDDEN_DETAIL };
