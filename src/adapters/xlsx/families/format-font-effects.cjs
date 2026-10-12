'use strict';

const { readKey } = require('../style-axis.cjs');

/**
 * M7 backlog X-F05 (font.underline / font.strikeout / font.subscript / font.superscript) --
 *, built on the SAME GetCharacters(1, len).GetFont() public path bold/italic/name/
 * size/color already use (W4a2); no new internal dependency, LIVE-measured by the S21 probes
 * (m7-s21-probe1/2).
 *
 *   font.underline   'single' | 'singleAccounting' | 'double' | 'doubleAccounting' | 'none' ->
 *                     range.SetUnderline(word); reads 'xlUnderlineStyle<Type>'.
 *   font.strikeout    true | false -> range.SetStrikeout(bool); reads: true -> true, false ->
 *                     null -- the SAME null/false ambiguity bold/italic's public reader already
 *                     has (false and never-touched are the same reading, LIVE-measured).
 *   font.subscript /
 *   font.superscript  true | false -> GetCharacters(1, len).GetFont().SetSubscript/
 *                     SetSuperscript(bool); reads: true -> its own field true, the OTHER field
 *                     null (they exclude each other, LIVE-measured); false -> null.
 *
 * Refused before the write (none measured to work, LIVE m7-s21-probe1/2):
 *   - subscript:true AND superscript:true requested together (plan-time: they cannot both hold)
 *   - subscript/superscript:true requested on a cell that is not plain TEXT (number / formula) --
 *     the setter call itself succeeds, but the readback silently stays null (a no-op nothing
 *     else would catch)
 *   - (a maintainer review) subscript:false / superscript:false requested while the OTHER one is true on the cell and not
 *     itself requested -- whether clearing one clears the other is NOT measured, and the shadow tolerance below
 *     would hide it (the other key's drop to null is tolerated whenever this key is requested)
 * An EMPTY cell is already refused generically: its whole `font` axis reads
 * 'UNOBSERVABLE:empty-cell', and every 'font.*' key falls through that same string (readKey).
 */

const UNDERLINE = Object.freeze({ single: 'xlUnderlineStyleSingle', singleAccounting: 'xlUnderlineStyleSingleAccounting', double: 'xlUnderlineStyleDouble', doubleAccounting: 'xlUnderlineStyleDoubleAccounting', none: 'xlUnderlineStyleNone' });

/** true/false -> the want entry for a boolean font-effects key: true -> its own field true,
 * false -> null (the same null/false ambiguity bold/italic's public reader already has). */
const boolWant = key => new Map([[true, Object.freeze({ key, expect: true })], [false, Object.freeze({ key, expect: null })]]);
const STRIKE_WANT = boolWant('font.strike');
const SUB_WANT = boolWant('font.sub');
const SUP_WANT = boolWant('font.sup');

/** The font spec's underline/strikeout/subscript/superscript keys -> their `want` entries, or a
 * refusal detail. Also validates the one plan-time constraint: subscript and superscript cannot
 * both be requested true (they exclude each other -- LIVE-measured, not a guess). */
const BOOL_EFFECTS = Object.freeze([['strikeout', STRIKE_WANT], ['subscript', SUB_WANT], ['superscript', SUP_WANT]]);

function fontEffectsExpectation(f) {
  // a maintainer review: an OWN key only -- `in` would accept an inherited name ('toString') as an underline type
  if ('underline' in f && !Object.hasOwn(UNDERLINE, f.underline)) return { error: `font.underline must be one of ${Object.keys(UNDERLINE).join(', ')}` };
  const badBool = BOOL_EFFECTS.find(([k]) => k in f && typeof f[k] !== 'boolean');
  if (badBool) return { error: `font.${badBool[0]} must be true or false` };
  if (f.subscript === true && f.superscript === true) return { error: 'font.subscript and font.superscript cannot both be true (they exclude each other)' };
  const underline = 'underline' in f ? [{ key: 'font.underline', expect: UNDERLINE[f.underline] }] : [];
  return { want: [...underline, ...BOOL_EFFECTS.filter(([k]) => k in f).map(([k, m]) => m.get(f[k]))] };
}

/** is cell `a` a plain TEXT cell -- no formula, not numeric, non-empty. The MEASURED condition
 * under which subscript/superscript:true actually changes anything (LIVE m7-s21-probe1: on a
 * number/formula/empty cell the setter call succeeds and the readback silently stays null). */
const isTextCell = cell => cell.f === '' && cell.count === 0 && cell.v !== '' && cell.v !== null;

/** the refusal for target cell `a` requesting subscript/superscript:true where the cell is not
 * plain text -- a silent no-op nothing else would catch. `want` is the op's full want list (this
 * family has no per-cell variation of it, unlike the border families). Returns a refusal object,
 * or null meaning proceed. */
function fontEffectsRefusal(want, cell, a) {
  const needsText = want.some(w => (w.key === 'font.sub' || w.key === 'font.sup') && w.expect === true);
  if (needsText && !isTextCell(cell)) return { reason: 'xlsx-format-font-effects-not-text-cell', cell: a };
  return crossScript(want, cell) ? { reason: 'xlsx-format-font-effects-cross-script-unmeasured', cell: a } : null;
}

/** a maintainer review: one script cleared (false) while the OTHER is true on the cell and not requested -- not measured. */
function crossScript(want, cell) {
  const clears = (mine, other) => want.some(w => w.key === mine && w.expect === null) && !want.some(w => w.key === other) && readKey(cell, other) === true;
  return clears('font.sub', 'font.sup') || clears('font.sup', 'font.sub');
}

/** subscript / superscript setter calls for one cell, a `false` BEFORE a `true` (clearing one after setting the other is
 * not measured; setting one true clears the other -- measured). */
function scriptCalls(f, J) {
  const calls = [['subscript', 'SetSubscript'], ['superscript', 'SetSuperscript']].filter(([k]) => k in f);
  calls.sort((x, y) => Number(f[x[0]] === true) - Number(f[y[0]] === true));
  return calls.map(([k, m]) => `cf.${m}(${J(f[k])});`).join(' ');
}

module.exports = { UNDERLINE, fontEffectsExpectation, fontEffectsRefusal, isTextCell, scriptCalls };
