'use strict';

/**
 * M7 backlog X-F07a (text rotation, `align.rotation`) --, built on the SAME public
 * ApiRange.SetOrientation/GetOrientation path the S24 probes measured (m7-s24-probe1/2).
 *
 * `align.rotation`  an integer in [-90, 90], or exactly 255 (vertical text) -> range-level
 *                   SetOrientation(n); read back via the SAME number (STYLE_KEYS 'rot', a FLAT
 *                   key -- PUBLIC, not through the DS Bridge the way align.h/v are).
 *
 * Every other value is refused BEFORE the write -- the S24 probe measured these to SILENTLY
 * RESET to 0 (not a no-op, not a clamp, not a throw): 91, -91, 180, 254, 256 (out of the measured
 * domain), any non-integer, and any string. Accepting one would corrupt the cell's rotation with
 * no proof it holds, violating P0.5/c (never guess a default / never accept an unproven write).
 * The four named values ('xlDownward'/'xlHorizontal'/'xlUpward'/'xlVertical') are not accepted as
 * spec input either: their own numeric forms (-90/0/90/255) already cover the same four values.
 */

const VERTICAL = 255;

/** `align.rotation`'s value -> its `want` entry, or a refusal detail. */
function rotationExpectation(rotation) {
  const ok = Number.isInteger(rotation) && ((rotation >= -90 && rotation <= 90) || rotation === VERTICAL);
  if (!ok) return { error: `align.rotation must be an integer in [-90, 90] or exactly ${VERTICAL} (the measured values -- other values were measured to silently reset to 0)` };
  return { want: [{ key: 'rot', expect: rotation }] };
}

module.exports = { VERTICAL, rotationExpectation };
