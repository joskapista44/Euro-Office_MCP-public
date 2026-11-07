'use strict';

/**
 * GH #363 2nd unit -- the PROJECTION behind the `read_effective_fill` read tool: per named cell
 * (<=200, enforced at request validation, read-tools.cjs checkEffectiveFillBound), the conditional-
 * format EVALUATED fill, read through the worksheet's own `getCompiledStyle(row, col)` getter via
 * the DS Bridge candidate `sheet.cells.compiledStyle.read/v1` (documentserver-patches/9.3.5.0/
 * euro-office-bridge-v1/rev28-table-autofilter-criteria-candidate/, NOT installed anywhere) -- a
 * question `read_conditional_formats` cannot answer at all (it mirrors the RULE, never the per-cell
 * EVALUATED result; its own evidence string says so: "rendered effect X-O02 NOT_MEASURED").
 *
 * LIVE MEASURED (docs/GH363-CF-EFFECTIVE-READBACK-MEASUREMENT-20261105.md, docs/evidence/gh363-cf-
 * effective-readback-live-20261105/internal-route-probe-v2-measured.json): a cell that satisfies its
 * rule reads a compiled style whose fill equals the rule's own fill (16711680 == `#FF0000`); a cell
 * the rule does not fire on reads a null compiled style. On a pre-/28 runtime the WHOLE axis
 * degrades to bridgeAxis's own named `'UNOBSERVABLE:CAPABILITY_MISSING'` string -- proven by this
 * family's own pure test, not assumed.
 *
 * SCOPE: ONLY the fill is read (the candidate handler's own measured path: `getFill()` ->
 * `getColor()` or `getSolidFill()`). Every other compiled-style field (font, border, number format,
 * ...) is NOT_MEASURED by this capability -- never guessed, never returned. One measured pair, one
 * rule type (`cellValue greater`, a solid fill); other CF rule types (colour scale, data bar, icon
 * set, font/text colour) are untested by this probe and this capability alike.
 */

/** A raw fill colour (an integer, possibly negative/out-of-range if the engine's own encoding ever
 * differs) -> a `'#RRGGBB'` hex string, or null when not a usable number. Masked to 24 bits so a
 * sign-extended or out-of-range integer never produces a malformed (too-long) hex string. */
function hexOfInt(n) {
  if (typeof n !== 'number' || !Number.isFinite(n)) return null;
  return `#${(Math.trunc(n) & 0xFFFFFF).toString(16).padStart(6, '0').toUpperCase()}`;
}

/** One raw per-cell entry from the capability -> `{ cell, fill }`. `fill` is a hex string, or null
 * when the compiled style itself was null (no CF-evaluated override -- e.g. the rule never fires on
 * this cell) or its own fill was not a readable number. A per-cell read failure (a bad address, a
 * throw) carries its own named reason, never silently dropped or folded into `fill: null` without
 * explanation. Fields are picked one by one -- never a blind spread of the handler's raw output. */
function cellOf(entry) {
  if (!entry || typeof entry !== 'object' || typeof entry.cell !== 'string') return { cell: undefined, fill: null, reason: 'UNOBSERVABLE:malformed-entry' };
  if (typeof entry.reason === 'string') return { cell: entry.cell, fill: null, reason: `UNOBSERVABLE:${entry.reason}` };
  if (!entry.compiled || typeof entry.compiled !== 'object') return { cell: entry.cell, fill: null };
  return { cell: entry.cell, fill: hexOfInt(entry.compiled.fill) };
}

/** `inv.compiledStyleReads[sheet!range]`: undefined (not read, the read_validations/read_conditional_
 * formats convention) | a degraded 'UNOBSERVABLE:<reason>' string (the WHOLE read failed, fail-
 * closed -- never a per-cell guess) | an array of raw per-cell entries. */
function projectEffectiveFillRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, range: a.range, error: 'SHEET_NOT_FOUND' };
  const read = inv.compiledStyleReads ? inv.compiledStyleReads[`${a.sheet}!${a.range}`] : undefined;
  if (read === undefined) return { sheet: a.sheet, range: a.range, fill: 'UNOBSERVABLE:not-read' };
  if (typeof read === 'string') return { sheet: a.sheet, range: a.range, fill: read };
  if (!Array.isArray(read)) return { sheet: a.sheet, range: a.range, fill: 'UNOBSERVABLE:malformed-read' };
  return { sheet: a.sheet, range: a.range, cells: read.map(cellOf) };
}

module.exports = { projectEffectiveFillRead, hexOfInt };
