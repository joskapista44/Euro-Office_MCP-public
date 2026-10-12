'use strict';

const { MAX_COL, MAX_ROW } = require('../a1.cjs');

/**
 * M7 backlog X-F07b (reading order) / X-F07c (indent, shrink-to-fit) -- owner-approved internal
 * dependency (the owner, owner channel, 2026-10-05; docs/M7-INTERNAL-DEPENDENCY-PROPOSAL-20261005.md),
 * on the SAME internal Range model format.cjs's align.h/v and border axes already read
 * (`cell.range`, `inner.getAlign(false)`) -- via the MODEL path S31c measured (NOT the UI path
 * S31b measured, which moves the user's active cell/sheet: asc_findCell + asc_getCellInfo;
 * Workbook.js:19665/19599/19358 model setters, Workbook.js:20217 getAlign(), WBE:5286-5310
 * getters). format.cjs's own gate() refuses the write BEFORE dispatch if any requested key reads
 * UNOBSERVABLE on the current state -- the same generic mechanism align.h/v and border already
 * rely on, so this file does not add a second, write-specific build check.
 *
 * align.readingOrder 'context' | 'ltr' | 'rtl' -> the PUBLIC ApiRange.SetReadingOrder (read back 0 | 1 | 2, OOXML's own
 *   readingOrder enum); unset reads back JS `undefined` (S31c MEASURED) -- the driver maps that
 *   to the semantic default 0 ('context'), never UNOBSERVABLE for that reason alone.
 * align.indent  an integer in [0, 15] -> setIndent(n). ONLY valid together with
 *   align.horizontal 'left' or 'right' IN THE SAME STEP -- the source zeroes indent for every
 *   other horizontal alignment, silently (not a throw, not a clamp), so a request without one of
 *   those two is refused BEFORE the write rather than accepted and later found wrong. NEVER on a
 *   whole-sheet range: setIndent there calls setAngle instead (Workbook.js:19673, the SAME edge
 *   format-rotation.cjs's own doc already names for rotation) -- refused before the write.
 * align.shrink  boolean -> setShrinkToFit(bool).
 *
 * The READ back is the DS Bridge range.alignmentExt.read/v1 (B05 READ, Wave B READ /4; a capability
 * SEPARATE from align.h/v's range.appearance.read/v1): a missing bridge or another patch reads
 * UNOBSERVABLE for these three keys, refusing the write before dispatch. The WRITE goes through the DS
 * Bridge WRITE capability since M11.5a.4 (format.cjs lower(), /6 only) -- no internal setter in the product.
 */

const READING_ORDER = Object.freeze({ context: 0, ltr: 1, rtl: 2 });
const READING_ORDER_REV = Object.freeze(Object.fromEntries(Object.entries(READING_ORDER).map(([k, v]) => [v, k])));

/** align.readingOrder's value -> its want entry, or a refusal detail. */
function readingOrderExpectation(value) {
  if (!(value in READING_ORDER)) return { error: `align.readingOrder must be one of ${Object.keys(READING_ORDER).join(', ')}` };
  return { want: [{ key: 'readingOrder', expect: READING_ORDER[value] }] };
}

/** align.indent's value -> its want entry, or a refusal detail. `horizontal` is align.horizontal's
 * OWN value from the SAME step (undefined if that key was not given this step). */
function indentExpectation(value, horizontal) {
  if (!(Number.isInteger(value) && value >= 0 && value <= 15)) return { error: 'align.indent must be an integer in [0, 15]' };
  if (horizontal !== 'left' && horizontal !== 'right') return { error: "align.indent requires align.horizontal 'left' or 'right' in the SAME step (every other alignment silently zeroes it)" };
  return { want: [{ key: 'indent', expect: value }] };
}

/** align.shrink's value -> its want entry, or a refusal detail. */
function shrinkExpectation(value) {
  if (typeof value !== 'boolean') return { error: 'align.shrink must be true or false' };
  return { want: [{ key: 'shrink', expect: value }] };
}

/** Is `g` (a parsed range grid) the WHOLE sheet -- setIndent there calls setAngle instead. */
const isWholeSheetRange = g => g.c1 === 1 && g.r1 === 1 && g.c2 === MAX_COL && g.r2 === MAX_ROW;

module.exports = { READING_ORDER, READING_ORDER_REV, readingOrderExpectation, indentExpectation, shrinkExpectation, isWholeSheetRange };
