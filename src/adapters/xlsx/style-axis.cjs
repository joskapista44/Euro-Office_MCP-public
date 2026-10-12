'use strict';

/**
 * The cell STYLE axis as one observation (M6 wave 4 reads; M7 S3 shares it). Pure.
 *
 * A cell read WITH style carries `wrap`, `font`, `align`, `border` (and with formatSheets also `nf`, `fill`, `bold`);
 * an axis the build cannot read is the string 'UNOBSERVABLE:<why>'. STYLE_KEYS is the flat list every style check
 * compares.
 */

const STYLE_KEYS = Object.freeze(['nf', 'fill', 'bold', 'italic', 'wrap', 'rot', 'readingOrder', 'indent', 'shrink', 'font.bold', 'font.italic', 'font.name', 'font.size', 'font.color', 'font.underline', 'font.strike', 'font.sub', 'font.sup', 'align.h', 'align.v', 'border.t', 'border.b', 'border.l', 'border.r', 'border.d', 'border.dd', 'border.du']);

const unobservable = x => typeof x === 'string' && x.startsWith('UNOBSERVABLE');

// Flat (non-dotted) keys: read directly off the cell, no sub-axis to descend into. `rot`
// (X-F07a) is PUBLIC (GetOrientation(), ungated) -- flat like wrap, not nested
// under `align` (which stays the internally-gated h/v pair, read through a different path).
// `readingOrder`/`indent`/`shrink` (X-F07b/c, owner GO 2026-10-05) are flat for the SAME
// structural reason, but the opposite direction: they come from their OWN DS Bridge capability
// (range.alignmentExt.read/v1, B05 READ) independent of align.h/v's -- nesting them under `align`
// would force one shared UNOBSERVABLE marker for two differently-governed capabilities.
const FLAT_KEYS = new Set(['fill', 'nf', 'wrap', 'bold', 'italic', 'rot', 'readingOrder', 'indent', 'shrink']);

/** The read value of one flat key on a normalized cell; `undefined` = not read at all. */
function readKey(cell, k) {
  if (!cell) return undefined;
  if (FLAT_KEYS.has(k)) return k in cell ? cell[k] : undefined;
  const [axis, sub] = k.split('.');
  if (!(axis in cell)) return undefined;
  return unobservable(cell[axis]) ? cell[axis] : cell[axis][sub];
}

const styleRead = cell => !!cell && 'wrap' in cell;
const comparable = v => v !== undefined && !unobservable(v);

/**
 * M7 S3: the style of every cell read WITH style in `before` (any sheet) is unchanged in `after`. Used by an operation
 * that must not touch style at all (a write, a column width), wherever a composed task reads style. Per key (owner
 * review of 93ff8899):
 *   observable before AND after        -> exact equality
 *   observable before, NOT after       -> a change: the preservation is not proven (a getter that throws, a key gone)
 *   NOT observable before              -> no claim for that key (the font of an empty cell becomes readable once a
 *                                         write gives it characters)
 * A cell whose style was read before and not after is a change.
 */
function observedStyleChanges(before, after) {
  const changed = [];
  let compared = 0;
  for (const [k, was] of Object.entries(before.cells || {})) {
    if (!styleRead(was)) continue;
    const now = after.cells && after.cells[k];
    const cell = k.replace('\u0000', '!');
    if (!styleRead(now)) { changed.push({ cell, change: 'style-unread-after' }); continue; }
    compared += 1;
    const keys = STYLE_KEYS.filter(x => {
      const b = readKey(was, x);
      const n = readKey(now, x);
      return comparable(b) && (!comparable(n) || JSON.stringify(b) !== JSON.stringify(n));
    });
    if (keys.length) changed.push({ cell, keys });
  }
  return { compared, changed };
}

module.exports = { STYLE_KEYS, unobservable, readKey, observedStyleChanges };
