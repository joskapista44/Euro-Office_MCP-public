'use strict';

const { parseRange, cellsOf } = require('./a1.cjs');

/**
 * GH #326 (U4b, owner decision): the sheet-protection WRITE guard -- task-level,
 * fail-closed, and re-checked right before the first mutation (U5's/U4a's sheet.protection.read/v1 reads are the
 * data; this module is the DECISION over them). Pure; no editor knowledge, no Bridge call of its own.
 *
 * The condition table (GH299 PR-B02 evidence): `write_range` / `clear_range` / `copy_range` (the TARGET only --
 * Copy leaves the source unchanged) / `move_range` (BOTH the source and the target -- Cut clears the source and
 * writes the target) need every cell of their target to be in `cells.unlocked` (an EMPTY cell is never listed --
 * the Excel default is LOCKED, the fail-closed direction for a write guard, per docs/DS-9.3.5-BRIDGE17-PROTECTION-
 * READ.md). `format_range` / `insert_rows` need only the sheet's own named `forbids` flag. Every OTHER intent on a
 * protected sheet is refused, named with the intent -- nothing about it is measured. An unprotected sheet: every
 * intent is unchanged, with no restriction from this module at all.
 *
 * `gate()` returns `{ refusal, raceChecks }`: `refusal` is `null` (proceed) or `{ reason, sheet, range }` (the
 * operation's own evidence, surfaced as `FAIL_CLOSED`); `raceChecks` is the list of `{ sheet, range?, expected }`
 * the DRIVER must re-verify, unchanged, in the SAME editor session, immediately before the first setter runs
 * (xlsx-driver.cjs's `protectionRacePrelude`) -- empty whenever this op's sheet(s) carry no restriction to begin
 * with (nothing protected can have "changed before the mutation").
 */

const CELL_LEVEL_INTENTS = Object.freeze(['write_range', 'clear_range']);
const FLAG_OF = Object.freeze({ format_range: 'formatCells', insert_rows: 'insertRows' });
// GH #338-#342: the protection writes themselves. protect / unprotect ARE the protection change -- allowed on a protected sheet,
// still watched (the driver re-verifies the read state before the write). A cell-protection change is refused on a protected
// sheet (Excel does not allow it either, and it could turn #327's redaction off); on an unprotected one it is unrestricted.
const SHEET_PROTECTION_INTENTS = Object.freeze(['protect_sheet', 'unprotect_sheet']);
const CELL_PROTECTION_INTENTS = Object.freeze(['set_cell_protection']);
// GH #338-#342 remainder (W2, docs/DS-9.3.5-BRIDGE24-WORKBOOK-PROTECTION-DESIGN.md "W2 -- the five structural
// guards", Bridge /24 only): intents that change the sheet GRAPH itself (add/rename/remove/move/copy a sheet),
// closing WB-03's measured workbook-structure-lock bypass. Orthogonal to the per-sheet axis above -- a workbook
// lock does not change a per-sheet protection decision, and vice versa; `move_sheet`/`copy_sheet` already route
// through the per-sheet `unmeasured()` catch-all today (additive, not a replacement of that check).
const WORKBOOK_STRUCTURE_INTENTS = Object.freeze(['create_sheet', 'rename_sheet', 'delete_sheet', 'move_sheet', 'copy_sheet']);

const unread = raw => (typeof raw === 'string' ? raw.slice('UNOBSERVABLE:'.length) : 'not-read');
const shortList = a => (a.length <= 10 ? a.join(',') : `${a.slice(0, 10).join(',')},...(+${a.length - 10})`);

/** write_range / clear_range / a copy|move target: every cell of `range` must be in `cells.unlocked`. */
function cellLevel(primary, sheet, range) {
  const raw = primary.protectionRangeReads && primary.protectionRangeReads[`${sheet}!${range}`];
  if (raw === undefined || typeof raw === 'string') return { refusal: { reason: `xlsx-protection-unobservable:${unread(raw)}`, sheet, range }, raceChecks: [] };
  if (raw.protected !== true) return { refusal: null, raceChecks: [{ sheet, range, expected: raw }] };
  const unlocked = new Set((raw.cells && raw.cells.unlocked) || []);
  const locked = cellsOf(parseRange(range)).filter(a => !unlocked.has(a));
  if (locked.length) return { refusal: { reason: `xlsx-protected-locked-cells:${shortList(locked)}`, sheet, range }, raceChecks: [] };
  return { refusal: null, raceChecks: [{ sheet, range, expected: raw }] };
}

/** format_range / insert_rows: the sheet's own named `forbids` flag decides; no per-cell check. */
function flagLevel(primary, sheet, flag) {
  const raw = primary.protectionReads && primary.protectionReads[sheet];
  if (raw === undefined || typeof raw === 'string') return { refusal: { reason: `xlsx-protection-unobservable:${unread(raw)}`, sheet, range: null }, raceChecks: [] };
  if (raw.protected !== true) return { refusal: null, raceChecks: [{ sheet, expected: raw }] };
  if (raw.forbids[flag] === true) return { refusal: { reason: `xlsx-protected-sheet-forbids:${flag}`, sheet, range: null }, raceChecks: [] };
  return { refusal: null, raceChecks: [{ sheet, expected: raw }] };
}

/** protect_sheet / unprotect_sheet: never refused here (the protection change itself), the read state watched. */
function sheetProtectionChange(primary, sheet) {
  const raw = primary.protectionReads && primary.protectionReads[sheet];
  if (raw === undefined || typeof raw === 'string') return { refusal: { reason: `xlsx-protection-unobservable:${unread(raw)}`, sheet, range: null }, raceChecks: [] };
  return { refusal: null, raceChecks: [{ sheet, expected: raw }] };
}

/** set_cell_protection: refused on a protected sheet; unprotected -> unchanged, the range read watched. */
function cellProtectionChange(primary, sheet, range) {
  const raw = primary.protectionRangeReads && primary.protectionRangeReads[`${sheet}!${range}`];
  if (raw === undefined || typeof raw === 'string') return { refusal: { reason: `xlsx-protection-unobservable:${unread(raw)}`, sheet, range }, raceChecks: [] };
  if (raw.protected === true) return { refusal: { reason: 'xlsx-protected-sheet-forbids:cellProtection', sheet, range }, raceChecks: [] };
  return { refusal: null, raceChecks: [{ sheet, range, expected: raw }] };
}

/** Every other intent: refused on a protected sheet (not measured); unprotected -> unchanged, still watched. */
function unmeasured(primary, sheet, intent) {
  const raw = primary.protectionReads && primary.protectionReads[sheet];
  if (raw === undefined || typeof raw === 'string') return { refusal: { reason: `xlsx-protection-unobservable:${unread(raw)}`, sheet, range: null }, raceChecks: [] };
  if (raw.protected === true) return { refusal: { reason: `xlsx-protected-sheet-intent-not-measured:${intent}`, sheet, range: null }, raceChecks: [] };
  return { refusal: null, raceChecks: [{ sheet, expected: raw }] };
}

/** The real sheet an op targets, or null -- `create_sheet` names no EXISTING sheet, and the `names` family's
 * `\u0000name:...` selector is a claim identifier, never a real sheet (sheetProtection does not apply to it). */
const realSheet = op => (op.selector && typeof op.selector.sheet === 'string' && !op.selector.sheet.startsWith('\u0000') ? op.selector.sheet : null);
const realDest = op => (op.dest && typeof op.dest.sheet === 'string' ? { sheet: op.dest.sheet, range: op.dest.range } : null);

/** A missing / case-mismatched sheet is NOT a protection question -- the family's own classify() already
 * refuses it, named (`xlsx-sheet-not-found` and friends); this gate only has an opinion once that sheet is
 * real. Exact match, same as the driver's own `ws.GetName() !== n` check (sheetProtection is case-sensitive
 * the same way the workbook itself is). */
const exists = (primary, sheet) => !!(primary.sheets && primary.sheets.some(s => s.name === sheet));

/**
 * GH #338-#342 remainder (W2, Bridge /24 only): the workbook-wide structure lock, decided on a SINGLE
 * `primary.workbookProtectionRead` snapshot -- a string is an unobserved axis, an object is the real read.
 * `CAPABILITY_MISSING` is a DELIBERATE pass-through (fix 1, release sequencing: this guard's own unit lands
 * only once `/24` is already accepted everywhere, so by construction it never runs against a pre-/24
 * runtime in production -- this branch is defense-in-depth for a mixed-fleet/rollback session that still
 * reaches one anyway); every OTHER unobservable reason fails closed, unchanged from that. `raceChecks`
 * carries a workbook-wide check (no `sheet` key -- there is no per-sheet target) for `protectionRacePrelude`'s
 * workbook branch to re-verify immediately before the first setter of a task this gate already passed.
 */
function workbookGate(primary) {
  const raw = primary.workbookProtectionRead;
  if (typeof raw === 'string') {
    if (unread(raw) === 'CAPABILITY_MISSING') return { refusal: null, raceChecks: [] }; // pre-/24 runtime: no lock possible, pass through
    return { refusal: { reason: `xlsx-protection-unobservable:workbook-${unread(raw)}`, sheet: null, range: null }, raceChecks: [] };
  }
  // Review-blocking defect, fixed here (internal tracker review of this unit's first patch): `raw === undefined` alone
  // missed the ALSO-live `null` shape (inventory.cjs:141 normalizes a missing read to `null`, not `undefined`)
  // -- `raw.lockStructure` on `null` THROWS, a TypeError instead of a named refusal. And a malformed object read
  // (`{}`, `{ lockStructure: null }`, a non-boolean) is NOT the same question as "is it locked" -- `!== true`
  // silently read a malformed shape as "not locked", FAIL-OPEN, the exact hazard W3's own `gate()`
  // (families/workbook-protection-write.cjs:62-66) already refuses as `malformed-read`. This mirrors it: `raw`
  // must be a real object with a BOOLEAN `lockStructure` before either outcome (locked / not-locked) is decided.
  if (raw === undefined || raw === null) return { refusal: { reason: 'xlsx-protection-unobservable:workbook-not-read', sheet: null, range: null }, raceChecks: [] };
  if (typeof raw !== 'object' || typeof raw.lockStructure !== 'boolean') return { refusal: { reason: 'xlsx-protection-unobservable:workbook-malformed-read', sheet: null, range: null }, raceChecks: [] };
  if (raw.lockStructure === true) return { refusal: { reason: 'xlsx-protected-workbook-locked-structure', sheet: null, range: null }, raceChecks: [] };
  return { refusal: null, raceChecks: [{ workbook: true, expected: raw }] };
}

// GH #326 (U4b)'s existing per-sheet logic, UNCHANGED, renamed (was `gate`) so the new wrapper below can
// compose with it without touching a single line of the shipped per-sheet checks.
function gateSheetLevel(primary, op) {
  const sheet = realSheet(op);
  if (!sheet || !exists(primary, sheet)) return { refusal: null, raceChecks: [] };
  if (CELL_LEVEL_INTENTS.includes(op.intent)) return cellLevel(primary, sheet, op.selector.range);
  if (op.intent === 'copy_range') {
    const dest = realDest(op);
    return dest && exists(primary, dest.sheet) ? cellLevel(primary, dest.sheet, dest.range) : { refusal: null, raceChecks: [] };
  }
  if (op.intent === 'move_range') {
    const src = cellLevel(primary, sheet, op.selector.range);
    if (src.refusal) return src;
    const dest = realDest(op);
    if (!dest || !exists(primary, dest.sheet)) return src;
    const dst = cellLevel(primary, dest.sheet, dest.range);
    if (dst.refusal) return dst;
    return { refusal: null, raceChecks: [...src.raceChecks, ...dst.raceChecks] };
  }
  if (SHEET_PROTECTION_INTENTS.includes(op.intent)) return sheetProtectionChange(primary, sheet);
  if (CELL_PROTECTION_INTENTS.includes(op.intent)) return cellProtectionChange(primary, sheet, op.selector.range);
  if (FLAG_OF[op.intent]) return flagLevel(primary, sheet, FLAG_OF[op.intent]);
  return unmeasured(primary, sheet, op.intent);
}

/**
 * GH #338-#342 remainder (W2): for the five sheet-graph intents, the workbook-wide check runs FIRST --
 * BEFORE `gateSheetLevel`'s own `realSheet`/`exists` early exit, which `create_sheet` would otherwise hit
 * immediately (it names no EXISTING sheet) and never reach any check at all. Composes as a MERGE, never an
 * overwrite: `gateSheetLevel`'s own early-exit paths return `raceChecks: []` unconditionally, which would
 * otherwise silently discard `workbookGate`'s own race check for exactly the one case (an unprotected-sheet
 * target under an unlocked workbook) this guard exists to re-verify.
 */
function gate(primary, op) {
  if (!WORKBOOK_STRUCTURE_INTENTS.includes(op.intent)) return gateSheetLevel(primary, op);
  const wb = workbookGate(primary);
  if (wb.refusal) return wb;
  const inner = gateSheetLevel(primary, op); // create_sheet still exits via the existing !sheet guard, with []
  if (inner.refusal) return inner;
  return { refusal: null, raceChecks: [...wb.raceChecks, ...inner.raceChecks] };
}

/** The (sheet[, range]) pairs this op's gate() will need read -- pure in `op` alone (no `primary` yet: this
 * runs at liveScope, before any observation), for the SAME per-request-range convention cfRanges/validationRanges/
 * protectionRanges already use (xlsx-driver.cjs TIERED_EXTRAS). `range` present -> `protectionRanges`; absent ->
 * `protectionSheets` (U4a's sheet-only axis -- cheaper, no per-cell walk, no 10000-cell cap to worry about). */
function targetsOf(op) {
  const sheet = realSheet(op);
  if (!sheet) return [];
  if (CELL_LEVEL_INTENTS.includes(op.intent) || CELL_PROTECTION_INTENTS.includes(op.intent)) return [{ sheet, range: op.selector.range }];
  if (op.intent === 'copy_range') { const d = realDest(op); return d ? [d] : []; }
  if (op.intent === 'move_range') { const d = realDest(op); return [{ sheet, range: op.selector.range }, ...(d ? [d] : [])]; }
  return [{ sheet }];
}

module.exports = { gate, targetsOf, workbookGate, gateSheetLevel, WORKBOOK_STRUCTURE_INTENTS };
