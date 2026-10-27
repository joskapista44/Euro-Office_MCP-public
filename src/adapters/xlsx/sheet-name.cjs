'use strict';

/**
 * L-21/2: the ONE sheet-name rule of the released name-only / rename intents. Pure.
 *
 * Before: three layers each carried their own copy (the write-tools schema SHEET_BARE, its CONSTRAINTS bareNameProblem,
 * the family gates sheets.cjs nameProblem / charts.cjs SHEET) and the narrowest one won. Design and evidence:
 * docs/L-21-SHEET-NAME-DESIGN-20261015.md (code reading) + the LIVE run of 2026-10-15 (raw callCommand, session AND reopen):
 *   - a name with a space ("Munkalap 1") or an accent ("Kimutatás"): freeze / insert rows / chart title, axis title, rename
 *     all work; a rename is followed by the other sheet's formula with the engine's OWN quoting and by the chart series;
 *   - a name with an APOSTROPHE ("O'Brien"): the DS engine does not quote it by itself (`=O'Brien!B4` reads #NAME? after the
 *     reopen), a rename is not kept after the reopen and a chart series is not followed -> refused by name, never written.
 *
 * The released class is exactly the MEASURED one: the first character a letter or underscore, then letters, digits and
 * underscore, words joined by SINGLE spaces, at most 31 characters, not shaped like a cell reference, not a reserved word.
 * Punctuation, a leading digit and padded / doubled spaces are NOT measured and stay refused with their own reason.
 */

const MAX_SHEET_NAME = 31;
const APOSTROPHE_REASON = 'sheet-name-apostrophe-unsupported: the Document Server engine mishandles an apostrophe in a sheet name (a formula naming it reads #NAME? after a reopen, a rename is not kept, a chart series is not followed) -- L-21 LIVE run, docs/L-21-SHEET-NAME-DESIGN-20261015.md';

const CELL_REF_SHAPED = n => /^[A-Za-z]{1,3}\d+$/.test(n) || /^([RrCc]\d*|[Rr]\d*[Cc]\d*)$/.test(n);
const RESERVED = n => /^(history|true|false)$/i.test(n);
const MEASURED_CLASS = /^[\p{L}_][\p{L}\p{N}_]*( [\p{L}\p{N}_]+)*$/u;

/** A name that is not a cell reference and not a reserved word (the class the SHEET_ADDRESS intents keep), else why not. */
const bareNameProblem = n => (CELL_REF_SHAPED(n) ? 'shaped like a cell reference' : RESERVED(n) ? 'a reserved word' : null);

/** null when the name is inside the released (measured) class, otherwise why not. */
function sheetNameProblem(n) {
  if (typeof n !== 'string' || n === '') return 'a sheet name must be a non-empty string';
  if (n.length > MAX_SHEET_NAME) return `longer than ${MAX_SHEET_NAME} characters (the editor silently renames it)`;
  if (n.includes("'")) return APOSTROPHE_REASON;
  if (!MEASURED_CLASS.test(n)) {
    return 'sheet-name-class-not-measured: letters (accents included), digits and underscore, starting with a letter or underscore, words joined by single spaces';
  }
  return bareNameProblem(n);
}

module.exports = { MAX_SHEET_NAME, APOSTROPHE_REASON, bareNameProblem, sheetNameProblem };
