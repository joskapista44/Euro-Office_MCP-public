'use strict';

const { parseRange, parseCell, rangeText, cellsOf, addr } = require('../a1.cjs');
const { translateFormula } = require('../shift.cjs');
const cells = require('./cells.cjs');
const formula = require('../formula.cjs');

/**
 * XLSX `validation` family: set_validation / clear_validation (M6 wave 6).
 *
 * Two readers, both already in the driver: the PUBLIC range-level GetValidation() getters (the source's measured
 * reader, B-D16) for the rule's parameters, and the W3-proven, build-gated INTERNAL dataValidations reader for the
 * rule's exact RANGES (the public reader answers for a range, not per cell -- a partial rule is invisible to it).
 * Audit: B-D16 / D25 MIGRATE (whole number, between, stop); B-D17..D22 NOT_MEASURED; B-D23 an existing DIFFERENT
 * rule was silently overwritten through Modify.
 *
 * Contract (the W6 probe measured the public getters; nothing wider is supported): set supports the measured shape only -- type 'wholeNumber',
 * operator 'between', alertStyle 'stop', integer bounds. Refused before the write: a validation that overlaps the
 * target PARTIALLY or covers it with another range, and an existing rule on exactly the target with other
 * parameters (no silent Modify); an unobservable reader. clear: SATISFIED when no rule touches the target; refused
 * when a rule reaches beyond it (Delete would clear cells outside the request -- unmeasured).
 * GH #96 (owner request, explicit opt-in): `replace: true` on set_validation lifts BOTH of the above pre-existing-
 * rule refusals (an exact-but-different rule, and any rule overlapping/fragmenting the target -- e.g. a prior
 * per-cell validation covering the same range) -- Delete then Add, instead of the refusal. A fragmented target
 * (per-cell validations) is measured (the P96 LIVE probe) to fail EARLIER than that -- the AGGREGATE public
 * reader itself refuses the whole gate with xlsx-validation-read-unobservable, before classify()'s own
 * overlap check ever runs -- so `replace` also tolerates THIS specific gate()-level gap, proceeding on the
 * SEPARATE internal rules reader alone. Without `replace` both refusals are UNCHANGED. The replace path's L1
 * additionally reads EVERY individual cell of the target (not only the aggregate range-level public reader,
 * whose own module comment already says it "answers for a range, not per cell") -- a fragmented pre-existing
 * state may not collapse into one rule the same way a fresh range does, LIVE-pending (the probe measuring
 * Delete+Add's exact write mechanics is a separate, parallel unit).
 * L1 the public read + exactly one internal rule on exactly the range (set) / none touching it (clear);
 * L3 every other rule of every sheet and every cell. Retry SATISFIED 0/0.
 *
 * M7 R17 (owner M7 GO 2026-10-03; tier D probe docs/evidence/M7-D-PROBE.md, every getter read back exactly in
 * session and after a reopen): type wholeNumber | decimal | textLength with operator between | notBetween (min, max)
 * or greater | lessEqual (value); type list with `items` (inline, no comma / quote in an item) or `listRange` (an
 * absolute range on the same sheet); alertStyle stop | warning | information. Still refused (measured or not
 * measured): date (formula1 reads back as locale text, M/D vs D/M ambiguous), custom (the formula is re-anchored
 * to A1 -- its meaning changes), the other operators, input / error messages and the ignore-blank / dropdown flags.
 *
 * L-22a (allowed EURO GO; the L-22 probe test/live/xlsx-l22-validation-types-probe.live.cjs, type time,
 * MEASURED_OK): type 'time' (xlValidateTime), same operators as the other bound-taking types. UNLIKE those,
 * min/max/value are 24-hour 'HH:MM' | 'HH:MM:SS' TEXT (timeRule()), converted to a day fraction (0..1) for
 * formula1/formula2 -- the SAME call shape, a different bounds encoding. The read-back formula1/formula2 is
 * LOCALE TEXT (measured: '6:00:00 AM' / '6:00:00 PM', AM/PM -- but a 24-hour locale is not excluded), parsed back
 * to a day fraction by timeReadFraction() for the L1/retry compare (sameRule()); a read-back text that cannot be
 * parsed unambiguously (AM/PM OR 24-hour, HH:MM(:SS) only) is a NAMED gap (xlsx-validation-time-unparseable),
 * never guessed as matching or not. (date: see L-22d below.)
 *
 * L-22b (allowed EURO GO; the L-22 probe test/live/xlsx-l22-validation-types-probe.live.cjs, custom, MEASURED):
 * RELATIVE formula (`=G2>0` on G2:G5) reads back RE-ANCHORED (`=A1>0` -- the MEANING changes), stays refused.
 * ABSOLUTE formula (`=$H$2>0`) reads back UNCHANGED (session = reopen, MEASURED_OK) -- type 'custom'
 * (xlValidateCustom) is released for this shape ONLY: NO operator, a single `formula` whose EVERY cell/range
 * reference in it is bare (same sheet) and FULLY absolute (`$COL$ROW` or `$COL$ROW:$COL$ROW`), via
 * absoluteOnlyProblem() -- a relative or mixed reference, a sheet qualifier (another sheet, or the measured
 * re-anchor defect's own class), an external/array `[` syntax, a defined name, INDIRECT/OFFSET or any other
 * function are ALL refused, named, never guessed (none of those is measured). The Add() call's own operator
 * argument is the SAME measured placeholder every probe and the M7-D tier used (`xlBetween`) -- the editor
 * ignores it for this type, but the call still takes one. L1/retry compare formula1 via formula.cjs's own
 * canonical() (the SAME helper range.cjs/structure.cjs/sheet-copy.cjs already use for formula identity),
 * never a bare string compare -- the editor's own '= ' + case/space canonicalization (formula.cjs's own
 * module comment) would otherwise make an unchanged formula look different.
 *
 * L-22d v2 (docs/L22D-VALIDATION-DATE-ANALYSIS.md + docs/L22D-VALIDATION-DATE-RELEASE.md; the L-22 probe measured
 * `GetFormula1 = '3/4/2024'` for serial 45355, AND the DS Bridge `range.validations.read/v1` row for the SAME cell:
 * `formula1: "45355"` -- the written serial, as a string, session AND after a reopen, for every tested case including the
 * genuinely ambiguous one): type 'date' (xlValidateDate), same operators as the other bound-taking types, min/max/value are
 * ISO 'YYYY-MM-DD' TEXT (dateInput()), any real calendar date 1900-03-01..9999-12-31, written as the 1900-system serial --
 * NO locale-readability restriction (v1's dateReadable() is gone: the bridge measurement showed EVERY date, ambiguous or
 * not, round-trips through the raw serial). L1/retry identity is NOT the public GetFormula1 locale text (month/day order
 * is undecidable from it alone) -- it is the bridge's OWN rules.bySheet row for the exact target range (dependencies.cjs
 * rulesOf(), already read for every set_validation step; `g.exact[0].formula`/`.formula2`), a plain string-serial compare
 * (sameRule()'s 3rd argument). The public text is used ONLY as a secondary cross-check: if it parses to a date (one or two
 * readings) and NEITHER reading equals the bridge serial, that is a NAMED conflict (xlsx-validation-date-read-conflict,
 * classify()) -- the two sources disagreeing is never silently resolved either way; unparseable text is ignored (not a
 * conflict -- it says nothing). The replace-path PER-CELL check (verifyPrimitive()) has no per-cell bridge row, so for
 * date it compares type/alertStyle/operator SHAPE only -- the formula identity is already covered once by the aggregate
 * bridge row. A pre-existing date rule whose bridge formula is not a serial, or a different one, is the existing
 * xlsx-validation-exists-different / -ambiguous path, unchanged.
 *
 * L-22e (MEASURED by three LIVE rounds of test/live/xlsx-l22e-custom-relative-ref-probe.live.cjs,
 * r1-r3, docs/L22E-VALIDATION-CUSTOM-RELATIVE-RELEASE.md): L-22b's own RELATIVE-formula refusal ('=G2>0' on G2:G5 reads back
 * '=A1>0') was a READ-SIDE artifact, not a write-side one. MEASURED FACTS: the STORED OOXML is
 * always correctly anchored to the range's own first cell (file-verifier.cjs's
 * readWorksheetDataValidations(), x14, session AND reopen); the WRITE itself is anchor-independent
 * (moving the active cell off the target before Add() does not change what gets stored); the
 * PUBLIC Formula1 getter reports the formula AS IF WRITTEN AT the rule's OWN SHEET's CURRENT
 * active cell (r3: S active, reading at A1/G2/C10 gives back '=A1>0'/'=G2>0'/'=C10>0' exactly,
 * confirmed via the SAME call's own Api.GetActiveSheet().GetName() + S's GetActiveCell(); r2's
 * apparent contradiction -- every anchor reading '=A1>0' -- was itself a probe artifact: the
 * Select() calls had not actually run with S active) -- switching the GLOBALLY active sheet away
 * still reflects S's own LAST active cell (r3's cross-sheet case, FOLLOWS_RULE_SHEET_LAST_CELL),
 * never the globally-active sheet's cell and never the target range's own anchor.
 *
 * bareSameSheetRefProblem() (renamed from the L-22b absoluteOnlyProblem(), same shape check) now
 * accepts a relative OR mixed ($COL/COL$ROW) reference, not only fully absolute -- still refused,
 * named, never guessed: another sheet (bare or quoted, `!`), an external/array `[` reference, a
 * defined name, INDIRECT/OFFSET or any other function (none of those is measured). Identity
 * (classify()/verifyPrimitive(), via sameRule() -> reAnchoredFormula()) RE-ANCHORS the read text
 * from the driver-captured activeCell (xlsx-driver.cjs's validationReads, read in the SAME call as
 * Formula1 -- the reference cell is never inferred after the fact) to the comparison target, using
 * the SAME per-cell shift a copy would apply (translateFormula(), shift.cjs's own relative/mixed-
 * part semantics -- absolute parts never move). A FULLY ABSOLUTE formula skips this entirely
 * (formulaHasOnlyAbsoluteRefs(): anchor-independent on read too, matching the already-released
 * L-22b shape and its own tests unchanged) -- only a relative/mixed formula depends on the
 * captured activeCell, and an unobservable/unparseable one is a NAMED non-match, NEVER guessed as
 * satisfied (FAIL_CLOSED falls through to the existing xlsx-validation-exists-different path).
 * `replace: true` with a relative/mixed custom formula is refused at plan time (customRule()): the
 * GH #96 per-cell verify has no per-cell bridge row and (unlike the aggregate check) was never
 * measured to capture a per-cell activeCell either -- doing the per-cell re-anchor math without
 * that measurement would be guessed, not verified. A fully-absolute formula under `replace` is
 * unaffected (unchanged since L-22b/GH#96).
 */

const INTENTS = Object.freeze(['set_validation', 'clear_validation']);
const MAX_CELLS = 1000;
// L-22c: type list's `items` -- the ceiling is the inline list-source LITERAL LENGTH (the comma-joined items,
// quotes excluded), not the item count. 255 is the documented Excel "List Source" field limit -- the editor
// motor itself accepted a longer inline list in the L-22 LIVE probe (100 short items, well over 255 chars), so
// this is an Excel-compatibility ceiling we choose to enforce, not a motor write/read limitation we measured.
// MAX_LIST_ITEMS is a loose, DoS-shaped upper bound only (128 one-character items join to exactly 255
// characters, so the literal-length check below is what fires first for any longer item); it exists so the
// schema/plan never need an unbounded array.
const MAX_LIST_LITERAL_CHARS = 255;
const MAX_LIST_ITEMS = 128;

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const set = step.intent === 'set_validation';
  const fields = set ? ['intent', 'id', 'sheet', 'range', 'type', 'operator', 'alertStyle', 'min', 'max', 'value', 'items', 'listRange', 'replace', 'formula'] : ['intent', 'id', 'sheet', 'range'];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `${at} requires a plain A1 range` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `${at} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  // GH #96: `replace` is set_validation-only, and only a boolean (not required -- absent == false, today's
  // fail-closed behaviour).
  if ('replace' in step && set && typeof step.replace !== 'boolean') return { ok: false, detail: `${at}: replace must be true or false` };
  let want = null;
  if (set) {
    want = ruleOf(step);
    if (want.problem) return { ok: false, detail: `${at}: ${want.problem}` };
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'validation',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      want: want ? Object.freeze(want) : null,
      replace: set && step.replace === true,
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const TYPES = Object.freeze({ wholeNumber: 'xlValidateWholeNumber', decimal: 'xlValidateDecimal', textLength: 'xlValidateTextLength', list: 'xlValidateList', time: 'xlValidateTime', date: 'xlValidateDate', custom: 'xlValidateCustom' });
// L-22b: the measured Add() call for custom ALWAYS carries this operator argument -- the editor ignores it for
// this type, but the API call still takes one (the probe's own shape: ['xlValidateCustom', alertStyle, 'xlBetween', formula]).
const CUSTOM_OPERATOR = 'xlBetween';
const ALERTS = Object.freeze({ stop: 'xlValidAlertStop', warning: 'xlValidAlertWarning', information: 'xlValidAlertInformation' });
const TWO = Object.freeze({ between: 'xlBetween', notBetween: 'xlNotBetween' });
// GH #301 D1 (public-API delta, DS-patch NELKUL; LIVE V-01 on 9.3.5: equal/greaterEqual PASS via Add()/Modify()):
// the remaining four one-value operators -- equal, notEqual, less, greaterEqual -- on the SAME ONE axis, same
// `value` field, same types (wholeNumber/decimal/textLength via boundsRule, time via timeRule, date via
// dateRule -- all three already dispatch through this one shared table). notEqual and less are NOT-MEASURED
// LIVE yet (a separate LIVE round covers them); they take the identical Add() call shape as every other ONE
// entry, so the same offline acceptance (xlsx-m7-d-rules.test.cjs) covers them until that round lands.
const ONE = Object.freeze({ greater: 'xlGreater', lessEqual: 'xlLessEqual', equal: 'xlEqual', notEqual: 'xlNotEqual', less: 'xlLess', greaterEqual: 'xlGreaterEqual' });
const plainNumber = n => typeof n === 'number' && Number.isFinite(n) && !/e/i.test(String(n));

/** L-22a: user input, STRICT 24-hour 'HH:MM' | 'HH:MM:SS' (zero-padded) -> a day fraction [0, 1), or null. */
const TIME_INPUT = /^([01]\d|2[0-3]):([0-5]\d)(?::([0-5]\d))?$/;
function timeInputFraction(text) {
  const m = typeof text === 'string' ? TIME_INPUT.exec(text) : null;
  if (!m) return null;
  return (Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3] || 0)) / 86400;
}

/** L-22a: the read-back LOCALE text (measured: '6:00:00 AM' AM/PM -- a 24-hour locale is not excluded, never
 * guessed either way) -> a day fraction, or null if NOT unambiguously parseable. */
const TIME_READBACK = /^(\d{1,2}):([0-5]\d)(?::([0-5]\d))?\s*(AM|PM)?$/i;
function timeReadFraction(text) {
  const m = TIME_READBACK.exec(String(text ?? '').trim());
  if (!m) return null;
  const min = Number(m[2]); const s = Number(m[3] || 0);
  let h = Number(m[1]);
  if (m[4]) {
    if (h < 1 || h > 12) return null;
    h = m[4].toUpperCase() === 'AM' ? (h === 12 ? 0 : h) : (h === 12 ? 12 : h + 12);
  } else if (h > 23) return null;
  return (h * 3600 + min * 60 + s) / 86400;
}

// L-22d: the 1900 date system the editor writes serials in: days since 1899-12-30, valid from 1900-03-01 (the 1900-leap-year compatibility
// quirk makes serials before it off by one, so they are not offered) to 9999-12-31.
const DATE_INPUT = /^(\d{4})-(\d{2})-(\d{2})$/;
const DATE_MIN_SERIAL = 61;      // 1900-03-01
const DATE_MAX_SERIAL = 2958465; // 9999-12-31
const calendarSerial = (y, mo, da) => {
  const t = Date.UTC(y, mo - 1, da);
  const d = new Date(t);
  if (d.getUTCFullYear() !== y || d.getUTCMonth() !== mo - 1 || d.getUTCDate() !== da) return null;
  return Math.round((t - Date.UTC(1899, 11, 30)) / 86400000);
};
/** L-22d v2: user input, STRICT ISO 'YYYY-MM-DD' (a real calendar date, 1900-03-01 .. 9999-12-31) -> { serial }, or null.
 * No locale-readability restriction (v1's dateReadable() is gone -- see the family header comment). */
function dateInput(text) {
  const m = typeof text === 'string' ? DATE_INPUT.exec(text) : null;
  if (!m) return null;
  const serial = calendarSerial(Number(m[1]), Number(m[2]), Number(m[3]));
  return serial === null || serial < DATE_MIN_SERIAL || serial > DATE_MAX_SERIAL ? null : { serial, month: Number(m[2]), day: Number(m[3]) };
}
/** L-22d v2: the SECONDARY cross-check only (identity is the bridge serial, see sameRule()) -- the read-back LOCALE
 * text N/N/YYYY -> the distinct serials of its valid readings (M/D/YYYY and D/M/YYYY), or null if it is not that
 * shape at all (an unparseable text is ignored by the cross-check, not a conflict). */
const DATE_READBACK = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
function dateReadSerials(text) {
  const m = DATE_READBACK.exec(String(text === undefined || text === null ? '' : text).trim());
  if (!m) return null;
  const a = Number(m[1]); const b = Number(m[2]); const y = Number(m[3]);
  const set = [...new Set([calendarSerial(y, a, b), calendarSerial(y, b, a)].filter(x => x !== null))];
  return set.length ? set : null;
}

/** The rule a set_validation step asks for, in the editor's terms -- or { problem }. */
function ruleOf(step) {
  const type = TYPES[step.type];
  const alertStyle = ALERTS[step.alertStyle];
  if (!type) return { problem: `type must be one of ${Object.keys(TYPES).join(' | ')}` };
  if (!alertStyle) return { problem: `alertStyle must be one of ${Object.keys(ALERTS).join(' | ')}` };
  if (step.type === 'list') return listRule(step, type, alertStyle);
  if (step.type === 'custom') return customRule(step, type, alertStyle);
  for (const k of ['items', 'listRange']) if (k in step) return { problem: `${k} belongs to type list` };
  if (step.type === 'time') return timeRule(step, type, alertStyle);
  if (step.type === 'date') return dateRule(step, type, alertStyle);
  return boundsRule(step, type, alertStyle);
}

/** L-22e: custom -- NO operator/min/max/value/items/listRange, a single `formula` whose every cell/range
 * reference must be bare (same sheet); relative, mixed or fully absolute are all accepted (L-22b's own
 * fully-absolute-only restriction is lifted -- the identity check re-anchors the read text instead, see
 * reAnchoredFormula()). `replace: true` with a non-absolute formula is refused here, named: the GH #96
 * per-cell verify has no per-cell activeCell capture to re-anchor against (never measured), so doing that
 * math would be guessed, not verified -- a fully-absolute formula under `replace` is unaffected. */
function customRule(step, type, alertStyle) {
  for (const k of ['operator', 'min', 'max', 'value', 'items', 'listRange']) if (k in step) return { problem: `type custom does not take ${k}` };
  if (typeof step.formula !== 'string' || !step.formula.startsWith('=') || step.formula.length < 2) return { problem: 'formula must be a non-empty string starting with =' };
  if (formula.canonical(step.formula) === null) return { problem: 'formula has an unterminated string literal or quoted name' };
  const problem = bareSameSheetRefProblem(step.formula);
  if (problem) return { problem };
  if (step.replace === true && !formulaHasOnlyAbsoluteRefs(step.formula)) return { problem: 'type custom with a relative or mixed reference does not support replace: true (not measured)' };
  return { type, alertStyle, operator: CUSTOM_OPERATOR, formula1: step.formula, formula2: '' };
}

// L-22b/e: a bare (no sheet qualifier) cell or range reference, letters then digits, optionally $-prefixed, an
// optional second cell for a range -- matched on the formula text OUTSIDE string literals / quoted names only
// (formula.cjs's own tokenize(), the SAME literal-skipping the rest of this codebase already relies on).
const REF_RE = /\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/g;
const REF_PARTS = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?$/;
/**
 * Every cell/range reference in `text` must be bare, or a NAMED problem -- never guessed. A sheet qualifier
 * (another sheet, bare or quoted -- `!` always lands in an 'other' token either way) or a `[` (external book /
 * structured-table syntax) refuses first; after every reference is removed, ANY letter still in the text is a
 * defined name, INDIRECT/OFFSET or any other function -- none of those is measured, so none is guessed to be
 * safe. L-22e: the $ combination (relative/mixed/absolute) is no longer checked here -- see
 * formulaHasOnlyAbsoluteRefs() for where that distinction still matters (the read-side identity check).
 */
function bareSameSheetRefProblem(text) {
  const toks = formula.tokenize(text);
  const other = (toks || []).filter(t => t.kind === 'other').map(t => t.text).join('');
  if (other.includes('!')) return 'formula must not reference another sheet';
  if (other.includes('[')) return 'formula must not use an external reference or a structured-table/array syntax';
  if (/[A-Za-z]/.test(other.replace(REF_RE, ' '))) return 'formula uses a name, INDIRECT/OFFSET or another unsupported function';
  return null;
}

/** L-22e: does `text` contain ONLY fully-absolute ($COL$ROW) cell/range references (possibly none at all)?
 * A pure shape check, no refusal semantics -- used to skip the active-cell-based re-anchoring entirely for a
 * formula that is anchor-independent on read too (the already-released, already-tested L-22b shape). */
function formulaHasOnlyAbsoluteRefs(text) {
  const toks = formula.tokenize(text);
  const other = (toks || []).filter(t => t.kind === 'other').map(t => t.text).join('');
  const refs = other.match(REF_RE) || [];
  return refs.every(r => { const m = REF_PARTS.exec(r); return !!(m && m[1] && m[3] && (m[6] === undefined || (m[5] && m[7]))); });
}

/** L-22a: time -- same operators as boundsRule, min/max/value are 24-hour text, converted to a day fraction. */
function timeRule(step, type, alertStyle) {
  const what = "24-hour 'HH:MM' or 'HH:MM:SS' time text";
  if (TWO[step.operator]) {
    if ('value' in step) return { problem: `operator ${step.operator} takes min and max, not value` };
    const min = timeInputFraction(step.min); const max = timeInputFraction(step.max);
    if (min === null || max === null) return { problem: `min and max must be ${what}` };
    if (min > max) return { problem: `min and max must be ${what}, min <= max` };
    return { type, alertStyle, operator: TWO[step.operator], formula1: String(min), formula2: String(max) };
  }
  if (ONE[step.operator]) {
    if ('min' in step || 'max' in step) return { problem: `operator ${step.operator} takes value, not min / max` };
    const value = timeInputFraction(step.value);
    if (value === null) return { problem: `value must be ${what}` };
    return { type, alertStyle, operator: ONE[step.operator], formula1: String(value), formula2: '' };
  }
  return { problem: `operator must be one of ${[...Object.keys(TWO), ...Object.keys(ONE)].join(' | ')} (the others are not measured)` };
}

/** L-22d v2: date -- same operators as boundsRule, min/max/value are ISO 'YYYY-MM-DD' text converted to a serial. Every
 * real calendar date in range is accepted (v1's locale-readability restriction is gone -- see the header comment:
 * identity is now the bridge's raw serial, not the locale-ambiguous public text). */
function dateRule(step, type, alertStyle) {
  const what = "ISO 'YYYY-MM-DD' date text (a real date, 1900-03-01 or later)";
  const check = text => { const d = dateInput(text); return d === null ? { problem: `bounds must be ${what}` } : d; };
  if (TWO[step.operator]) {
    if ('value' in step) return { problem: `operator ${step.operator} takes min and max, not value` };
    const min = check(step.min); if (min.problem) return min;
    const max = check(step.max); if (max.problem) return max;
    if (min.serial > max.serial) return { problem: `min and max must be ${what}, min <= max` };
    return { type, alertStyle, operator: TWO[step.operator], formula1: String(min.serial), formula2: String(max.serial) };
  }
  if (ONE[step.operator]) {
    if ('min' in step || 'max' in step) return { problem: `operator ${step.operator} takes value, not min / max` };
    const value = check(step.value); if (value.problem) return value;
    return { type, alertStyle, operator: ONE[step.operator], formula1: String(value.serial), formula2: '' };
  }
  return { problem: `operator must be one of ${[...Object.keys(TWO), ...Object.keys(ONE)].join(' | ')} (the others are not measured)` };
}

/** wholeNumber / decimal / textLength: two bounds (between, notBetween) or one value (greater, lessEqual). */
function boundsRule(step, type, alertStyle) {
  const isNum = n => (step.type === 'decimal' ? plainNumber(n) : Number.isInteger(n)) && (step.type !== 'textLength' || n >= 0);
  const what = step.type === 'decimal' ? 'plain finite numbers' : (step.type === 'textLength' ? 'non-negative integers' : 'integers');
  if (TWO[step.operator]) {
    if ('value' in step) return { problem: `operator ${step.operator} takes min and max, not value` };
    if (!isNum(step.min) || !isNum(step.max) || step.min > step.max) return { problem: `min and max must be ${what}, min <= max` };
    return { type, alertStyle, operator: TWO[step.operator], formula1: String(step.min), formula2: String(step.max) };
  }
  if (ONE[step.operator]) {
    if ('min' in step || 'max' in step) return { problem: `operator ${step.operator} takes value, not min / max` };
    if (!isNum(step.value)) return { problem: `value must be one of the ${what}` };
    return { type, alertStyle, operator: ONE[step.operator], formula1: String(step.value), formula2: '' };
  }
  return { problem: `operator must be one of ${[...Object.keys(TWO), ...Object.keys(ONE)].join(' | ')} (the others are not measured)` };
}

function listRule(step, type, alertStyle) {
  for (const k of ['operator', 'min', 'max', 'value']) if (k in step) return { problem: `type list does not take ${k}` };
  if (('items' in step) === ('listRange' in step)) return { problem: 'type list takes exactly one of items | listRange' };
  if ('items' in step) {
    const shape = Array.isArray(step.items) && step.items.length > 0 && step.items.length <= MAX_LIST_ITEMS && step.items.every(x => typeof x === 'string' && x !== '' && x === x.trim() && !/[,"]/.test(x));
    if (!shape) return { problem: `items must be 1..${MAX_LIST_ITEMS} non-empty strings without a comma, a double quote or outer spaces` };
    const literal = step.items.join(',');
    if (literal.length > MAX_LIST_LITERAL_CHARS) return { problem: `items' combined list-source literal is ${literal.length} characters (the Excel list-source limit is ${MAX_LIST_LITERAL_CHARS})` };
    return { type, alertStyle, operator: 'xlBetween', formula1: `"${literal}"`, formula2: '' };
  }
  const g = parseRange(step.listRange);
  if (!g) return { problem: 'listRange must be a plain A1 range on the same sheet' };
  const abs = `$${rangeText(g).replace(/^([A-Z]+)(\d+)/, '$1$$$2').replace(/:([A-Z]+)(\d+)$/, ':$$$1$$$2')}`;
  return { type, alertStyle, operator: 'xlBetween', formula1: `=${abs}`, formula2: '' };
}

const keyOf = op => `${op.selector.sheet}!${op.selector.range}`;
const overlaps = (a, b) => a.c1 <= b.c2 && b.c1 <= a.c2 && a.r1 <= b.r2 && b.r1 <= a.r2;
/** L-22e: the A1 address of `range`'s own first (top-left) cell, or null -- sameRule()'s `wantAt` for the
 * aggregate (non-per-cell) checks. */
const rangeFirstCellAddr = range => { const g = parseRange(range); return g ? addr(g.c1, g.r1) : null; };

/** The internal validation rules of the target sheet that touch the target: exact / other. */
function touching(p, op) {
  const t = parseRange(op.selector.range);
  const rules = (p.rules.bySheet[op.selector.sheet] || []).filter(r => r.kind === 'validation');
  const exact = [];
  const other = [];
  for (const r of rules) {
    const parts = String(r.range).split(',').map(x => parseRange(x.trim()));
    if (!parts.some(x => x && overlaps(x, t))) continue;
    if (parts.length === 1 && parts[0] && rangeText(parts[0]) === op.selector.range) exact.push(r); else other.push(r);
  }
  return { exact, other };
}

const bare = f => String(f === undefined || f === null ? '' : f).replace(/^=/, '');
// L-22a: a want.formula* of type 'time' is ALREADY a day fraction (ruleOf's own timeRule) -- the read-back is
// LOCALE TEXT, so it is parsed back first (timeReadFraction), never string-compared. An unparseable read text
// returns false here (not matching) -- classify() names the gap explicitly BEFORE this is ever called, so this
// branch is only reached once the text is known to parse.
const timeFormulaSame = (readText, wantFraction) => (wantFraction === '' ? bare(readText) === '' : timeReadFraction(readText) !== null && Math.abs(timeReadFraction(readText) - Number(wantFraction)) < 1e-9);
// L-22d v2: a want.formula* of type 'date' is ALREADY a serial; the identity source is the BRIDGE's raw formula
// (rules.bySheet row, a plain string-serial), NOT the locale-ambiguous public text -- a non-digit string, or a
// digit string naming a different serial, is never a match.
const dateFormulaSame = (bridgeFormula, wantSerial) => {
  if (wantSerial === '') return bare(bridgeFormula) === '';
  return /^\d+$/.test(String(bridgeFormula === undefined || bridgeFormula === null ? '' : bridgeFormula)) && Number(bridgeFormula) === Number(wantSerial);
};
// L-22b: formula identity compares through formula.cjs's own canonical() (the SAME helper range.cjs/
// structure.cjs/sheet-copy.cjs already use) -- the editor's '= ' + case/space canonicalization (formula.cjs's
// own module comment) would otherwise make an UNCHANGED absolute formula look different from a bare compare.
const sameFormula = (a, b) => { const ca = formula.canonical(a); const cb = formula.canonical(b); return ca !== null && cb !== null && ca === cb; };
/**
 * L-22e: the public Formula1 getter for type 'custom' reports the formula AS IF WRITTEN AT the rule's OWN
 * SHEET's CURRENT active cell (measured, r1-r3 of xlsx-l22e-custom-relative-ref-probe.live.cjs), never the target range's own anchor and never the globally-active sheet's cell. A
 * fully-absolute formula needs none of this (measured anchor-independent, formulaHasOnlyAbsoluteRefs()'s own
 * fast path -- unchanged from L-22b). Only a relative/mixed formula is re-anchored, by the SAME per-cell shift
 * a copy from `readActiveCell` to `wantAt` would apply (translateFormula(), shift.cjs -- the editor's own
 * relative-reference semantics). `readActiveCell` unobservable (the driver's own UNOBSERVABLE: sentinel) or
 * unparseable, `wantAt` unparseable, or the shift itself breaks (#REF!, out of the sheet) -> null, NEVER
 * guessed as matching.
 */
function reAnchoredFormula(text, readActiveCell, wantAt) {
  if (formulaHasOnlyAbsoluteRefs(text)) return text;
  if (typeof readActiveCell !== 'string' || readActiveCell.startsWith('UNOBSERVABLE:')) return null;
  const from = parseCell(readActiveCell);
  const to = parseCell(wantAt);
  if (!from || !to) return null;
  const { formula: shifted, broken } = translateFormula(text, to.r - from.r, to.c - from.c);
  return broken ? null : shifted;
}
// L-22d v2: `bridgeRule` is the family's own rules.bySheet row for the EXACT target range (gate()'s `g.exact[0]`),
// undefined when there is none to compare against (the verifyPrimitive() replace-path per-cell check: there is no
// per-cell bridge row, only the aggregate one, already checked once -- date there is SHAPE-only, formula untested).
// L-22e: `wantAt` is the A1 cell `want.formula1` is itself anchored to (the target range's own first cell for the
// aggregate check; the individual cell for the per-cell replace check) -- only used for type 'custom'.
const sameRule = (read, want, bridgeRule, wantAt) => {
  if (!read || read.absent || !['type', 'alertStyle', 'operator'].every(k => read[k] === want[k])) return false;
  if (want.type === TYPES.time) return timeFormulaSame(read.formula1, want.formula1) && timeFormulaSame(read.formula2, want.formula2);
  if (want.type === TYPES.date) {
    if (bridgeRule === undefined) return true;
    return dateFormulaSame(bridgeRule.formula, want.formula1) && dateFormulaSame(bridgeRule.formula2, want.formula2);
  }
  if (want.type === TYPES.custom) {
    const reAnchored = reAnchoredFormula(read.formula1, read.activeCell, wantAt);
    return reAnchored !== null && sameFormula(reAnchored, want.formula1) && bare(read.formula2) === bare(want.formula2);
  }
  return bare(read.formula1) === bare(want.formula1) && bare(read.formula2) === bare(want.formula2);
};

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  if (!p.rules || !p.rules.complete || !p.rules.bySheet) return { ok: false, reason: 'xlsx-validation-rules-unobservable', unobservable: p.rules ? p.rules.unobservable : 'not-read' };
  const read = p.validationReads ? p.validationReads[keyOf(op)] : undefined;
  if (read === undefined || typeof read === 'string') {
    // GH #96 (the P96 probe): the AGGREGATE public reader itself is what fails over a range spanning
    // MULTIPLE different validations (e.g. a prior per-cell validation per cell) -- 'xlsx-validation-read-
    // unobservable', not an 'other'-overlap classify() ever gets a chance to see, since gate() returns here
    // first. `replace: true` tolerates this specific gap and proceeds on the SEPARATE internal reader alone
    // (`touching`, `p.rules.bySheet` -- unaffected by the public reader's per-range throw): `read: null` makes
    // classify()'s sameRule/SATISFIED check correctly false, falling through to the replace-APPLY branch.
    // Without `replace` this is unchanged: a hard refusal, named.
    if (op.replace) return { ok: true, read: null, ...touching(p, op) };
    return { ok: false, reason: 'xlsx-validation-read-unobservable', why: read === undefined ? 'not-read' : read };
  }
  return { ok: true, read, ...touching(p, op) };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, range: op.selector.range };
  if (op.intent === 'clear_validation') {
    if (g.other.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-overlaps-another-range', rules: g.other.map(r => r.range).slice(0, 20) } };
    if (!g.exact.length && g.read.absent) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-no-validation' } };
    return { state: 'APPLY', evidence: where };
  }
  // set_validation: already the requested rule, exactly, with nothing else overlapping -- SATISFIED either way.
  if (g.exact.length === 1 && !g.other.length) {
    // L-22a: the exact rule is ALSO type time -- before guessing match/no-match, name an unparseable read-back
    // text explicitly (classify() never silently treats "cannot parse" as "different").
    if (op.want.type === TYPES.time && g.read && g.read.type === TYPES.time) {
      const unparseable = [g.read.formula1, g.read.formula2].filter(f => bare(f) !== '' && timeReadFraction(f) === null);
      if (unparseable.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-time-unparseable', ...where, texts: unparseable } };
    }
    // L-22d v2: the exact rule is ALSO type date -- a SECONDARY cross-check only (identity is the bridge serial,
    // sameRule() below): if the PUBLIC text parses to one or two readings and the bridge serial is NEITHER of
    // them, the two sources disagree -- named, never silently resolved either way. Unparseable text says
    // nothing and is ignored (not a conflict).
    if (op.want.type === TYPES.date && g.read && g.read.type === TYPES.date) {
      const pairs = [[g.read.formula1, g.exact[0].formula], [g.read.formula2, g.exact[0].formula2]];
      const conflicts = pairs.filter(([text, bridgeF]) => {
        if (bare(text) === '') return false;
        const readings = dateReadSerials(text);
        return readings !== null && !readings.includes(Number(bridgeF));
      }).map(([text]) => text);
      if (conflicts.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-date-read-conflict', ...where, texts: conflicts, bridge: [g.exact[0].formula, g.exact[0].formula2] } };
    }
    if (sameRule(g.read, op.want, g.exact[0], rangeFirstCellAddr(op.selector.range))) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-the-rule-on-exactly-the-range' } };
  }
  if (!op.replace) {
    if (g.other.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-overlaps-another-range', rules: g.other.map(r => r.range).slice(0, 20) } };
    if (g.exact.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-ambiguous', count: g.exact.length } };
    if (g.exact.length === 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-exists-different', now: g.read } };
    return { state: 'APPLY', evidence: where };
  }
  // GH #96: `replace: true` -- any pre-existing rule touching the target (an exact-but-different rule, an
  // ambiguous exact count, or a fragmented/other overlap, e.g. a prior per-cell validation) is replaced,
  // Delete then Add, instead of refused. Named in the evidence so a caller can tell a replace happened.
  // A touching rule that EXTENDS beyond the target would be deleted or shrunk outside it by the range Delete(), and the
  // replace L3 strips every touching rule -- that outside change would go unseen. Only rules fully inside the target
  // are replaced; anything wider refuses, named, before dispatch (the caller widens the range to the whole rule).
  const t = parseRange(op.selector.range);
  const inside = x => x && x.c1 >= t.c1 && x.c2 <= t.c2 && x.r1 >= t.r1 && x.r2 <= t.r2;
  const wider = [...g.exact, ...g.other].filter(r => !String(r.range).split(',').map(x => parseRange(x.trim())).every(inside));
  if (wider.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-replace-rule-extends-beyond-range', rules: wider.map(r => r.range).slice(0, 20) } };
  return { state: 'APPLY', evidence: { ...where, replacing: { exact: g.exact.length, other: g.other.map(r => r.range).slice(0, 20) } } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range }) };
}

function lower(op) {
  const { sheet, range } = op.selector;
  const w = op.want;
  // a single-formula rule is added with four arguments (the measured call); a two-formula rule with five (M6 body)
  const args = w ? [w.type, w.alertStyle, w.operator, w.formula1, ...(w.formula2 === '' ? [] : [w.formula2])].map(x => JSON.stringify(x)).join(', ') : '';
  // GH #96 (replace: true, LIVE-pending -- a separate, parallel unit measures Delete+Add's exact mechanics on
  // a fragmented/per-cell pre-existing state; this is the most-likely shape stated ahead of that measurement,
  // not yet LIVE-confirmed): Delete the pre-existing rule(s) first, then Add -- never a silent Modify.
  const call = op.intent === 'set_validation' ? `${op.replace ? 'r.GetValidation().Delete();\n  ' : ''}r.GetValidation().Add(${args});` : 'r.GetValidation().Delete();';
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const r = ws.GetRange(${JSON.stringify(range)});
  ${call}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'validation', action: op.intent, sheet, range, want: w, replace: !!op.replace }),
  };
}

function verifyPrimitive(op, _before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  if (op.intent !== 'set_validation') {
    const ok = g.exact.length === 0 && g.other.length === 0 && g.read.absent === true;
    return { ok, evidence: { sheet: op.selector.sheet, range: op.selector.range, read: g.read, exact: g.exact.length, other: g.other.length, proof: 'fresh-live-public-read-and-internal-ranges' } };
  }
  const aggregateOk = g.exact.length === 1 && g.other.length === 0 && sameRule(g.read, op.want, g.exact[0], rangeFirstCellAddr(op.selector.range));
  if (!op.replace) {
    return { ok: aggregateOk, evidence: { sheet: op.selector.sheet, range: op.selector.range, read: g.read, exact: g.exact.length, other: g.other.length, proof: 'fresh-live-public-read-and-internal-ranges' } };
  }
  // GH #96 replace: the aggregate range-level reader's own module comment already says it "answers for a
  // range, not per cell" -- a fragmented pre-existing state (e.g. a prior per-cell validation) may not collapse
  // into its single-rule shape the way a fresh range does, LIVE-pending. Read EVERY cell of the target
  // individually and require each one to show the SAME new rule, in addition to the aggregate check.
  // L-22d v2: no per-cell bridge row exists here (only the aggregate one, already checked above) -- sameRule()'s
  // 3rd argument is intentionally omitted, so for date this is SHAPE-only (type/alertStyle/operator).
  const p = after && after.primary;
  const cellMisses = [];
  for (const a of cellsOf(parseRange(op.selector.range))) {
    const cellRead = p && p.validationReads ? p.validationReads[`${op.selector.sheet}!${a}`] : undefined;
    if (cellRead === undefined || typeof cellRead === 'string' || !sameRule(cellRead, op.want, undefined, a)) cellMisses.push(a);
  }
  return {
    ok: aggregateOk && cellMisses.length === 0,
    evidence: { sheet: op.selector.sheet, range: op.selector.range, read: g.read, exact: g.exact.length, other: g.other.length, cellMisses: cellMisses.slice(0, 20), proof: 'fresh-live-public-read-per-cell-and-internal-ranges' },
  };
}

/** L3: every cell, and every rule other than the target's validation unchanged (all sheets). GH #96 replace:
 * the BEFORE snapshot may carry rule(s) that legitimately disappear (the ones Delete+Add replaced) -- those
 * are excluded from the comparison too (by OVERLAP with the target, not just an exact-range-text match),
 * otherwise their intended removal would read as an unwanted change. Without replace, unchanged: only the
 * exact-range-text match is excluded, same as before this change. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const t = parseRange(op.selector.range);
  const touchesTarget = r => String(r.range).split(',').map(x => parseRange(x.trim())).some(x => x && overlaps(x, t));
  const strip = p => Object.fromEntries(Object.entries(p.rules.bySheet || {}).map(([n, rs]) => [n, rs.filter(r => !(n === op.selector.sheet && r.kind === 'validation' && (op.replace ? touchesTarget(r) : r.range === op.selector.range)))]));
  const b = before.primary;
  const a = after.primary;
  const rulesKept = !!(b.rules && a.rules && b.rules.bySheet && a.rules.bySheet) && JSON.stringify(strip(b)) === JSON.stringify(strip(a));
  // X-L13/6: once tiered, report preservationScope FULL -- b.rules.bySheet is fed by the `dependents`
  // bridge (dependentsBridgeBody(), "every data validation of every sheet", already in TIERED_EXTRAS and
  // already proven under tiered by set_defined_name/print-area), workbook-wide and unbounded by
  // MAX_READ_CELLS -- the cross-sheet check above is not reduced. Below 45000 unchanged.
  const tiered = b.tier === 'tiered';
  return { ok: rest.ok && rulesKept, evidence: { ...rest.evidence, otherRulesKept: rulesKept, ...(tiered ? { preservationScope: 'FULL' } : {}) } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, finalObservation, finalObservation);
  const what = op.intent !== 'set_validation' ? 'has no validation' : (op.want.type === 'xlValidateWholeNumber' && op.want.operator === 'xlBetween' ? 'validates a whole number between the bounds' : `validates ${op.want.type} ${op.want.operator}`);
  return [{ name: `${op.selector.sheet}!${op.selector.range} ${what}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
// GH #96: a replace step also asks for a per-cell validation read of every cell of the target -- see
// verifyPrimitive's own comment. Every other step's request is unchanged (one range-level read).
// GH #96 follow-up: the dependency readers are asked for the RULES only (rules.bySheet, the L3) -- a validation step
// never reads the T3 formula index, and its 5000-formula ceiling would make the observation incomplete
const liveExtras = op => ({
  dependents: true,
  formulaIndex: false,
  validationRanges: [
    { sheet: op.selector.sheet, range: op.selector.range },
    ...(op.replace ? cellsOf(parseRange(op.selector.range)).map(a => ({ sheet: op.selector.sheet, range: a })) : []),
  ],
});

module.exports = { INTENTS, MAX_LIST_LITERAL_CHARS, MAX_LIST_ITEMS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, reAnchoredFormula, rangeFirstCellAddr };
