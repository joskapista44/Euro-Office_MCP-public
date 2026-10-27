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

function gate(primary, op) {
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
  if (FLAG_OF[op.intent]) return flagLevel(primary, sheet, FLAG_OF[op.intent]);
  return unmeasured(primary, sheet, op.intent);
}

/** The (sheet[, range]) pairs this op's gate() will need read -- pure in `op` alone (no `primary` yet: this
 * runs at liveScope, before any observation), for the SAME per-request-range convention cfRanges/validationRanges/
 * protectionRanges already use (xlsx-driver.cjs TIERED_EXTRAS). `range` present -> `protectionRanges`; absent ->
 * `protectionSheets` (U4a's sheet-only axis -- cheaper, no per-cell walk, no 10000-cell cap to worry about). */
function targetsOf(op) {
  const sheet = realSheet(op);
  if (!sheet) return [];
  if (CELL_LEVEL_INTENTS.includes(op.intent)) return [{ sheet, range: op.selector.range }];
  if (op.intent === 'copy_range') { const d = realDest(op); return d ? [d] : []; }
  if (op.intent === 'move_range') { const d = realDest(op); return [{ sheet, range: op.selector.range }, ...(d ? [d] : [])]; }
  return [{ sheet }];
}

module.exports = { gate, targetsOf };
