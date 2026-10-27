'use strict';

const { parseRange, cellsOf, rangeText, addr, MAX_COL, MAX_ROW } = require('../a1.cjs');
const { cellAt, key } = require('../inventory.cjs');
const cells = require('./cells.cjs');
const { STYLE_KEYS, unobservable, readKey } = require('../style-axis.cjs');
const { H, V, STYLE, SIDE, OPPOSITE, API_SIDE, API_STYLE, REMOVED, col, row, colNum, neighbour, sidesList } = require('./format-maps.cjs');
const { removeExpectation, removeFacing, removeWant, removeGate, removeLine, removeMeta } = require('./format-remove.cjs');
const { isDiagonalSides, diagonalExpectation, diagonalWant, diagonalConflict, diagonalLine } = require('./format-diagonal.cjs');
const { fillClearGate, fillClearLine, fillClearMeta } = require('./format-fill-clear.cjs');
const { fontEffectsExpectation, fontEffectsRefusal, scriptCalls } = require('./format-font-effects.cjs');
const { rotationExpectation } = require('./format-rotation.cjs');
const { readingOrderExpectation, indentExpectation, shrinkExpectation, isWholeSheetRange } = require('./format-align-ext.cjs');
const { XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS } = require('../runtime.cjs');

/**
 * B05-WRITE (align.indent / align.shrink) goes through the DS Bridge WRITE capability (M11.5a.4, owner GO the owner, Talk
 * owner channel, internal tracker) -- the direct internal setters are gone. The capability exists only on /6, so this
 * write path alone is narrowed to /6 (DS5_POLICY=B05_ONLY_FAIL_CLOSED): on any other runtime classify() refuses BEFORE
 * the write, after the prewrite READ (the driver's B05_BRIDGE_WRITE_RUNTIMES is the second line, the in-editor client's
 * exact patchVersion the third). Every other format key keeps the dual runtime list.
 */
const B05_BRIDGE_WRITE_RUNTIMES = XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS;
const wantsAlignExtWrite = op => op.want.some(w => w.key === 'indent' || w.key === 'shrink');

/**
 * XLSX `format` family: format_range (M6 wave 4). Every sub-option is verified PER CELL on a fresh read
 * (W4a LIVE: a range-level getter reports the top-left for the whole range).
 *
 *   fill          PUBLIC SetFillColor / GetFillColor().GetRGB()                 (W4a). GH #90 (LIVE-measured):
 *                 a target range containing a MERGED area is not refused -- a plain range-level
 *                 write lands on every member (one call, the whole merge), verified on every member too. A merge
 *                 that CROSSES the requested range's boundary stays refused, named (xlsx-format-merge-crosses-
 *                 range-refused) -- partial colouring of a merged block is misleading.
 *   fill 'none'   (M7 backlog X-F08): clears the fill, keeps every other axis -- the
 *                 ONE working public path is the SAME ClearFormats() + explicit-restore composite
 *                 border.mode 'remove' uses (format-fill-clear.cjs, built on format-remove.cjs's
 *                 own `unrestorable()`), minus a shared-edge write: fill-clear is single-cell, so
 *                 there is no neighbour-pair rewrite (a maintainer decision #ref does not apply here).
 *                 Refused before the write (none measured by the probe -- one cell, one side, no
 *                 neighbour interaction): a kept axis unobservable or outside our H/V/STYLE write
 *                 vocabulary, a diagonal border present, the target's border sharing an edge with
 *                 a readable neighbour (both sides drawn, same style + colour). `fill: 'none'`
 *                 must be the format's ONLY key (like border.mode 'remove' restricts its own
 *                 object) -- a simultaneous explicit change to another axis has no defined outcome
 *                 against a restore that reads 'before' and reapplies every kept axis.
 *   wrap          PUBLIC SetWrap / GetWrapText()                                (W4a)
 *   numberFormat  PUBLIC SetNumberFormat / GetNumberFormat(), exact string       (W4a)
 *   font          PUBLIC setters / GetCharacters(1, len).GetFont()              (W4a2) -- bold / italic only TRUE
 *                 (a false readback is not measured: an untouched cell reads null); not on an EMPTY cell (no
 *                 characters), not on a FORMULA cell (not measured), not on a mixed-run cell (the driver compares
 *                 the first / last character with the whole). GH #90 (LIVE-measured): a target range
 *                 containing a MERGED area is not refused -- the write and the verification both stay on
 *                 the anchor; every other member's value (and so its font axis) is blanked by the merge itself,
 *                 the SAME empty-cell case above, not a new gap. A merge crossing the range's boundary stays
 *                 refused, named, same as fill's.
 *   font.underline / font.strikeout / font.subscript / font.superscript  (M7 backlog X-F05, zero
 *                 #ref, LIVE m7-s21-probe1/2): the SAME public GetCharacters(1, len).GetFont()
 *                 path, format-font-effects.cjs. underline/strikeout are RANGE-level setters
 *                 (ApiRange.SetUnderline/SetStrikeout); subscript/superscript have no range-level
 *                 setter, only ApiFont.SetSubscript/SetSuperscript, so those write per cell.
 *                 underline reads back as its enum string ('xlUnderlineStyle<Type>'); strikeout
 *                 true -> true, false -> null (the SAME null/false ambiguity bold/italic's public
 *                 reader already has); subscript/superscript exclude each other (requesting one
 *                 true drops the other to null, tolerated as a shadow effect, not drift) and
 *                 cannot both be requested true (refused at plan time). subscript/superscript:true
 *                 is refused on a non-TEXT cell (number/formula) -- the setter succeeds but the
 *                 readback silently stays null, a no-op nothing else would catch.
 *   align         setters PUBLIC, the read EO_INTERNAL_DEPENDENCY READ, build-gated (W4c): horizontal
 *                 left / center / right / justify, vertical top / center / bottom / distributed / justify
 *   border        per cell SetBorders(side, style, colour), read EO_INTERNAL_DEPENDENCY READ (W4c): sides top /
 *                 bottom / left / right, styles thin / medium / thick / dashed (measured); REMOVAL refused (W4c:
 *                 SetBorders(side, 'None') does not remove); black reads 0 in the session and null after a reopen
 *                 (verified as 0 | null); a cell's read shows its neighbour's SHARED edge, so a neighbour's facing
 *                 side may change -- and nothing else of it.
 *   border mode 'outline' (M7 S4 S01, owner GO 2026-10-03; docs/evidence/M7-S4-S01-OUTLINE.md): the requested sides
 *                 (default all four) on the RANGE's outer edges only -- top = the first row's top, etc. Written per
 *                 edge cell (the measured W4c path, one callCommand); the range-level SetBorders gives the identical
 *                 state (S4 probe). Measured shared edge: a ring neighbour's own facing side that differs is cleared
 *                 by the engine, so its read shows the requested edge -- L3 requires EXACTLY that value there, and
 *                 every other side of the range and its one-cell ring (inner edges, diagonals, ring corners) unchanged.
 *                 Refused: an unreadable border anywhere in the range + ring, a merge in the range + ring.
 *   border mode 'inside' / 'all' (M7 backlog X-B02/X-B03, owner mandate 2026-10-03, owner channel;
 *                 docs/evidence/M7-B-PROBE.md section 2): 'inside' writes the range's INNER shared edges only (no
 *                 `sides` option -- always both directions): Bottom of every cell except the last row, Right of
 *                 every cell except the last column (the MEASURED write direction -- the opposite cell's Top/Left
 *                 shows it through the same shared edge, as for outline). 'all' is outline (all four outer sides,
 *                 no `sides` option) plus the same inside write, in one command. The shared-edge rule generalizes:
 *                 a cell's side may read as changed either because IT was requested, or because its neighbour
 *                 across that side requested the opposite side -- this now applies to a grid-internal neighbour
 *                 (inside/all) exactly as it already did to a one-cell-ring neighbour (outline).
 *   border mode 'remove' (M7 backlog X-B04, owner mandate 2026-10-03, owner channel; docs/evidence/
 *                 M7-B-PROBE.md sections 3-5, a maintainer decision #ref): the named sides, gone, on every target cell.
 *                 `SetBorders(side, 'None')` never removes anything, self-owned or shared (MEASURED, P0.3/a-b) --
 *                 the only working public path is the composite: read EVERY observable style axis first, THEN
 *                 `ClearFormats()`, THEN re-apply every KEPT axis (the other border sides, fill, wrap, number
 *                 format, font, align) with an EXPLICIT colour/value, never the removed sides (MEASURED, P0.5).
 *                 Because a shared edge's physical storage belongs to ONE of its two cells and that is not
 *                 observable from here (no 'own' read, only the synthesized 'full' one), the composite runs on
 *                 BOTH cells of every removed side (a maintainer decision #ref, candidate (C) over an unmeasured
 *                 internal 'own' read (A) or a write-then-verify-fails attempt (B)) -- LIVE-CONFIRMED by P2.0
 *                 (m7-b4-probe1-20261004): the shared edge disappears from both cells, every kept side/
 *                 fill/bold/italic/numFmt round-trips byte-exact (the black 0->null tolerance aside).
 *                 Refused before the write, two classes: (1) a KEPT axis reads a value outside our H/V/STYLE
 *                 write vocabulary -- observable is not restorable, review (`unrestorable()`);
 *                 (2) a DIAGONAL border is present on a target cell -- NOT implemented here (out of this
 *                 feature's scope, kept conservative), even though P2.0 measured that `SetBorders('DiagonalDown',
 *                 ...)` after `ClearFormats()` DOES restore it byte-exact; this is a scope boundary, not a
 *                 missing public path. A black border colour reads `0` in-session and `null` after a reopen
 *                 (P0.5/c) -- `holds()` already tolerates this for the removed sides; the KEPT sides' drift
 *                 check tolerates it too.
 * Refused before the write: an unobservable reader for a requested key (unverified build, empty / formula /
 * mixed-run cell for the font), a merged area, null / unknown / empty specs.
 * L1: every requested key holds on every target cell, every OTHER style key of the target cells is unchanged
 * (the shared-edge exception aside), the content unchanged. L3: every other cell's content (all sheets) and
 * every other cell's style on the target sheet (its used range + a one-cell ring) unchanged.
 * Retry = the terminal state on a fresh read -> SATISFIED, 0 write / 0 command. Limit, stated: vertical bottom
 * equals the untouched default -- a bottom request on an untouched cell is SATISFIED at once.
 */

const INTENTS = Object.freeze(['format_range']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'range', 'format']);
const MAX_CELLS = 2000; // X-B11 (m7-b11-probe1-4 MEASURED): every format axis
// (bold/fill/numberFormat/outline/inside/all/remove) satisfies the 3-part bound (command <=5s,
// L1+L3 complete in-session AND after reopen, save confirmed) up to and including N=2000. The
// workbook-level ~45000-used-cell read ceiling (format.cjs's own live read, not yet migrated to
// tiered observation) is unchanged and independent of this per-task write ceiling.
const HEX = /^#[0-9A-Fa-f]{6}$/;

const rgbOf = hex => [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16));
const intOf = hex => { const [r, g, b] = rgbOf(hex); return r * 65536 + g * 256 + b; };

/** The edge-based modes ('outline' / 'inside' / 'all') -> { outline: { sides, expect }, inside }; or a refusal
 * detail. 'inside' and 'all' take no `sides` -- always both inner directions / all four outer sides. */
function edgeModeExpectation(mode, b, expect) {
  if (mode === 'inside' || mode === 'all') {
    if ('sides' in b) return { error: `border.sides is not supported with mode '${mode}' (always ${mode === 'inside' ? 'both inner directions' : 'all four outer sides plus both inner directions'})` };
    return { outline: { sides: mode === 'all' ? Object.values(SIDE) : [], expect }, inside: true };
  }
  const sides = sidesList('sides' in b ? b.sides : Object.keys(SIDE));
  if (!sides) return { error: 'border.sides must be a non-empty list of distinct top / bottom / left / right' };
  return { outline: { sides, expect }, inside: false };
}

/** The border spec -> per-cell keys ({ want }), an edge-based mode (`edgeModeExpectation`), or a removal
 * (`removeExpectation`); or a refusal detail. */
function borderExpectation(b) {
  if (!b || typeof b !== 'object' || Array.isArray(b)) return { error: 'border must be an object' };
  if ('mode' in b && !['cells', 'outline', 'inside', 'all', 'remove'].includes(b.mode)) return { error: "border.mode must be 'cells' (every cell), 'outline' (the range's outer edges), 'inside' (the inner shared edges), 'all' (outline + inside) or 'remove' (the named sides, gone)" };
  if (b.mode === 'remove') return removeExpectation(b);
  const bad = Object.keys(b).filter(k => !['mode', 'sides', 'style', 'color'].includes(k));
  if (bad.length) return { error: `border key(s) not supported: ${bad.join(', ')}` };
  if (b.style === 'none') return { error: "border.style 'none' is not a style -- use border.mode 'remove' with the sides to remove" };
  if (!(b.style in STYLE)) return { error: `border.style must be one of ${Object.keys(STYLE).join(', ')} (the measured styles)` };
  // X-B07 (LIVE-measured, test/live/xlsx-xb07-auto-border-colour-probe.live.cjs): 'auto' = the editor's "Automatic"
  // border colour -- the PUBLIC Api.AutoColor() writes <color auto="1"/>, and the Bridge border read reports it as
  // c: null (an explicit black reads 0 and saves as indexed 64 -- a DIFFERENT state). Not offered on the diagonals
  // (unmeasured there). Named limit, unchanged: a BLACK request still tolerates a null read (the W4c 0 | null rule),
  // so a black request on an 'auto' edge reads SATISFIED.
  if (typeof b.color !== 'string' || !(HEX.test(b.color) || b.color === 'auto')) return { error: "border.color must be '#RRGGBB' or 'auto'" };
  if (b.color === 'auto' && isDiagonalSides(b.sides)) return { error: "border.color 'auto' is not supported on the diagonal sides (unmeasured)" };
  const expect = { s: STYLE[b.style], c: b.color === 'auto' ? null : intOf(b.color) };
  if (b.mode && b.mode !== 'cells') return edgeModeExpectation(b.mode, b, expect);
  // X-B05/S03: diagonal sides are a SEPARATE, same-cell-only vocabulary -- never mixed with top/bottom/
  // left/right in one call, and never routed through the SIDE/OPPOSITE/neighbour machinery outline/
  // inside/all/remove share (a diagonal has no neighbour-facing edge, and 'remove' has no diagonal
  // restore path -- MEASURED, SetBorders(side,'None') does not clear a diagonal either, m7-b05-probe1).
  return isDiagonalSides(b.sides) ? diagonalExpectation(b.sides, expect) : sideExpectation(b.sides, expect);
}

/** The plain sides (top / bottom / left / right) of a 'cells' border -> the want list, or a refusal detail. */
function sideExpectation(list, expect) {
  const sides = sidesList(list);
  return sides ? { want: sides.map(x => ({ key: `border.${x}`, expect })) } : { error: 'border.sides must be a non-empty list of distinct top / bottom / left / right' };
}

/** The spec -> the flat list of requested keys with their expected read value; or a refusal detail. */
/** font.bold -> the key it is verified on: true on the public character reader, false on the internal cell reader (X-F01). */
const BOLD_WANT = new Map([[true, Object.freeze({ key: 'font.bold', expect: true })], [false, Object.freeze({ key: 'bold', expect: false })]]);
/** font.italic -> the same shape as BOLD_WANT, on the italic-parallel internal cell reader (X-F01, owner GO 2026-10-04). */
const ITALIC_WANT = new Map([[true, Object.freeze({ key: 'font.italic', expect: true })], [false, Object.freeze({ key: 'italic', expect: false })]]);

/** A border in a mode of its own (outline / inside / all, remove, diagonal) -> the whole expectation; plain sides -> null. */
function borderModeResult(want, b) {
  if (b.outline) return { want, outline: b.outline, inside: !!b.inside, remove: null, diagonal: null };
  if (b.remove) return { want: [...want, ...b.remove.sides.map(x => ({ key: `border.${x}`, expect: REMOVED }))], outline: null, inside: false, remove: b.remove, diagonal: null };
  return b.diagonal ? { want: [...want, ...diagonalWant(b.diagonal)], outline: null, inside: false, remove: null, diagonal: b.diagonal } : null;
}

const FONT_KEYS = Object.freeze(['bold', 'italic', 'name', 'size', 'color', 'underline', 'strikeout', 'subscript', 'superscript']);
/** [key, valid, error, expect] for the value-carrying font keys, in their want order. */
const FONT_VALUES = Object.freeze([
  // L-02/1 LIVE: the xlsx font-name limit is 31 characters -- a longer name reads whole in the session and is cut AFTER the reopen
  ['name', v => typeof v === 'string' && !!v && v.length <= 31, 'font.name must be a non-empty string of at most 31 characters (the xlsx font-name limit)', v => v],
  ['size', v => Number.isFinite(v) && v >= 1 && v <= 409, 'font.size must be a number in [1, 409]', v => v],
  ['color', v => typeof v === 'string' && HEX.test(v), "font.color must be '#RRGGBB'", v => intOf(v)],
]);

/** The `font` object -> its want entries (bold, italic, name, size, color, then the X-F05 effects), or an error. */
function fontWant(f) {
  if (!f || typeof f !== 'object' || Array.isArray(f) || !Object.keys(f).length) return { error: 'font must be a non-empty object' };
  const bad = Object.keys(f).filter(k => !FONT_KEYS.includes(k));
  if (bad.length) return { error: `font key(s) not supported: ${bad.join(', ')}` };
    // X-F01/S07: font.bold:false targets the EXISTING cell-level internal reader ('bold', already read on
    // every style-read cell for the true-case drift tolerance below) -- a NEW CONSUMER of that read, not a
    // new dependency. font.bold:true keeps targeting the public character reader (font.bold), unchanged.
    // Owner GO 2026-10-04 (owner channel) for both. On an unverified build the 'bold' read is
    // UNOBSERVABLE and the gate refuses before the write. font.italic (X-F01, owner GO 2026-10-04) is
    // read the same way as bold's -- since the B01 / B02 cutover both through the DS Bridge
    // (textStyle.bold/italic.read/v1), fail-closed/UNOBSERVABLE when the bridge does not serve them.
  const flagError = ('bold' in f && !BOLD_WANT.has(f.bold) && 'font.bold must be true or false')
    || ('italic' in f && !ITALIC_WANT.has(f.italic) && 'font.italic must be true or false');
  const valueBad = FONT_VALUES.find(([k, valid]) => k in f && !valid(f[k]));
  if (flagError || valueBad) return { error: flagError || valueBad[2] };
  // X-F05: underline/strikeout/subscript/superscript, the SAME public
  // GetCharacters(1, len).GetFont() path, LIVE-measured by the S21 probes.
  const fx = fontEffectsExpectation(f);
  if (fx.error) return fx;
  const flags = [['bold', BOLD_WANT], ['italic', ITALIC_WANT]].filter(([k]) => k in f).map(([k, m]) => m.get(f[k]));
  return { want: [...flags, ...FONT_VALUES.filter(([k]) => k in f).map(([k, , , x]) => ({ key: `font.${k}`, expect: x(f[k]) })), ...fx.want] };
}

function expectationsOf(spec) {
  if (!spec || typeof spec !== 'object' || Array.isArray(spec)) return { error: 'format must be an object' };
  const top = Object.keys(spec);
  if (!top.length) return { error: 'format is empty' };
  const unknown = top.filter(k => !['fill', 'wrap', 'numberFormat', 'font', 'align', 'border'].includes(k));
  if (unknown.length) return { error: `format key(s) not supported: ${unknown.join(', ')}` };
  const fill = fillExpectation(spec, top.length);
  if (fill.error) return fill;
  const want = [...fill.want];
  if ('wrap' in spec) { if (typeof spec.wrap !== 'boolean') return { error: 'wrap must be true or false' }; want.push({ key: 'wrap', expect: spec.wrap }); }
  if ('numberFormat' in spec) { if (typeof spec.numberFormat !== 'string' || !spec.numberFormat || spec.numberFormat.length > 255) return { error: 'numberFormat must be a non-empty string' }; want.push({ key: 'nf', expect: spec.numberFormat }); }
  if ('font' in spec) {
    const fw = fontWant(spec.font);
    if (fw.error) return fw;
    want.push(...fw.want);
  }
  if ('align' in spec) {
    const a = spec.align;
    if (!a || typeof a !== 'object' || Array.isArray(a) || !Object.keys(a).length) return { error: 'align must be a non-empty object' };
    const bad = Object.keys(a).filter(k => !['horizontal', 'vertical', 'rotation', 'readingOrder', 'indent', 'shrink'].includes(k));
    if (bad.length) return { error: `align key(s) not supported: ${bad.join(', ')}` };
    if ('horizontal' in a) { if (!(a.horizontal in H)) return { error: `align.horizontal must be one of ${Object.keys(H).join(', ')}` }; want.push({ key: 'align.h', expect: H[a.horizontal] }); }
    // GH #93: SetAlignHorizontal('justify') also turns the cell's wrap ON (LIVE-measured, session + reopen, text /
    // number / empty cells, test/live/justify-text-cell-probe.live.cjs) -- expect that coupled wrap unless the step
    // asks for wrap itself. Vertical justify / distributed: no such coupling measured, not assumed.
    if (a.horizontal === 'justify' && !('wrap' in spec)) want.push({ key: 'wrap', expect: true });
    if ('vertical' in a) { if (!(a.vertical in V)) return { error: `align.vertical must be one of ${Object.keys(V).join(', ')}` }; want.push({ key: 'align.v', expect: V[a.vertical] }); }
    // X-F07a: align.rotation, the SAME public SetOrientation/GetOrientation path
    // the S24 probes measured -- format-rotation.cjs.
    if ('rotation' in a) {
      const rot = rotationExpectation(a.rotation);
      if (rot.error) return rot;
      want.push(...rot.want);
    }
    // X-F07b/c (owner GO 2026-10-05, owner channel): readingOrder / indent / shrink, the
    // internal Range model path S31c measured -- format-align-ext.cjs.
    if ('readingOrder' in a) {
      const ro = readingOrderExpectation(a.readingOrder);
      if (ro.error) return ro;
      want.push(...ro.want);
    }
    if ('indent' in a) {
      const ind = indentExpectation(a.indent, a.horizontal);
      if (ind.error) return ind;
      want.push(...ind.want);
    }
    if ('shrink' in a) {
      const sh = shrinkExpectation(a.shrink);
      if (sh.error) return sh;
      want.push(...sh.want);
    }
  }
  if ('border' in spec) {
    const b = borderExpectation(spec.border);
    if (b.error) return b;
    const modal = borderModeResult(want, b);
    if (modal) return modal;
    want.push(...b.want);
  }
  return { want, outline: null, inside: false, remove: null, diagonal: null, fillClear: fill.clear };
}

/** the `fill` key's expectation: '#RRGGBB' sets it; X-F08 'none' clears it and is a standalone clear-mode, like
 * border.mode 'remove' -- combining it with any other key has no defined outcome against the ClearFormats()+restore
 * composite (MEASURED, m7-s15-probe4: the restore reads 'before' and reapplies every kept axis). */
function fillExpectation(spec, keyCount) {
  if (!('fill' in spec)) return { want: [], clear: false };
  if (spec.fill === 'none') {
    if (keyCount > 1) return { error: "fill 'none' must be the only format key (use a separate format_range step for other changes)" };
    return { want: [{ key: 'fill', expect: null }], clear: true };
  }
  if (typeof spec.fill !== 'string' || !HEX.test(spec.fill)) return { error: "fill must be '#RRGGBB' or 'none'" };
  return { want: [{ key: 'fill', expect: intOf(spec.fill) }], clear: false };
}

function plan(step, index) {
  const at = `step ${index}: format_range`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `${at} requires a plain A1 range on the named sheet (no Sheet!A1)` };
  const e = expectationsOf(step.format);
  if (e.error) return { ok: false, detail: `${at}: ${e.error}` };
  // X-F07c: align.indent never on a whole-sheet range (setIndent there calls setAngle instead) --
  // checked BEFORE the cell-count ceiling, since a whole-sheet range always exceeds it too, and
  // this is the more specific, more informative refusal for that combination.
  if (e.want.some(w => w.key === 'indent') && isWholeSheetRange(g)) return { ok: false, detail: `${at}: align.indent is not supported on a whole-sheet range (the source calls setAngle there instead)` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `${at} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  const ring = { c1: Math.max(1, g.c1 - 1), r1: Math.max(1, g.r1 - 1), c2: Math.min(MAX_COL, g.c2 + 1), r2: Math.min(MAX_ROW, g.r2 + 1) };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'format_range',
      family: 'format',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      cells: Object.freeze(cellsOf(g).map(a => Object.freeze({ a }))),
      grid: Object.freeze({ c1: g.c1, r1: g.r1, c2: g.c2, r2: g.r2 }),
      ring: `${addr(ring.c1, ring.r1)}:${addr(ring.c2, ring.r2)}`,
      spec: Object.freeze(JSON.parse(JSON.stringify(step.format))),
      want: Object.freeze(e.want.map(Object.freeze)),
      outline: e.outline ? Object.freeze({ sides: Object.freeze(e.outline.sides), expect: Object.freeze(e.outline.expect) }) : null,
      inside: !!e.inside,
      remove: e.remove ? Object.freeze({ sides: Object.freeze(e.remove.sides) }) : null,
      diagonal: e.diagonal ? Object.freeze({ codes: Object.freeze(e.diagonal.codes), expect: Object.freeze(e.diagonal.expect) }) : null,
      fillClear: !!e.fillClear,
      // M7 S3: what the step modifies, per axis -- the style of its cells; a wrap also re-fits the AUTO height of its
      // rows (W4e LIVE: a long wrapped text grew its row), so it claims those rows too. Remove's footprint is the
      // ring too: clearing a shared edge can clear a ring neighbour's facing side exactly like outline/inside do.
      // X-F07a: rotation locks rows the SAME way -- the S24 probe measured the row height move
      // from null to a concrete value around a rotation write; the footprint stays conservative either way.
      footprint: Object.freeze({ style: (e.outline || e.remove) ? `${addr(ring.c1, ring.r1)}:${addr(ring.c2, ring.r2)}` : rangeText(g), ...(('wrap' in step.format || (step.format.align && ('rotation' in step.format.align || step.format.align.horizontal === 'justify'))) ? { rows: Object.freeze([g.r1, g.r2]) } : {}) }),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

function holds(cell, w) {
  const v = readKey(cell, w.key);
  // X-B05: border.dd/border.du are the diagonal DIRECTION FLAGS -- booleans, not {s,c} style objects --
  // so they take the plain equality branch even though the key starts with 'border.'.
  if (w.key === 'border.dd' || w.key === 'border.du') return v === w.expect;
  if (w.key.startsWith('border.')) return !!v && typeof v === 'object' && v.s === w.expect.s && (v.c === w.expect.c || (w.expect.c === 0 && v.c === null));
  // X-F08: 'no fill' is the ONE measured read -- after ClearFormats() GetFillColor() returns the string
  // 'No Fill' (no GetRGB; m7-s15-probe4/5), which xlsx-driver.cjs's fallback stores as 'S:No Fill'.
  // LIVE m7-f08-live1: the earlier unmeasured null / 'S:none' sentinel never matched (FAIL, copy kept).
  if (w.key === 'fill' && w.expect === null) return v === 'S:No Fill';
  return v === w.expect;
}

/** Every style key of a cell, flattened (an unobservable axis is one key). */
/** `blackIsNull`: a border colour that the ClearFormats()-and-restore composite (op.remove / op.fillClear) re-issues explicitly reads 0 in the writing session,
 * while the same black read null in the fresh read before (P0.5/c: black reads 0 in a session and null after a reopen). Without this, a kept BLACK side or diagonal
 * on a restored cell reads as drift and the product's own verification answers VERIFY_CANNOT_PROVE_TERMINAL_STATE for a write that landed (L-16/1 LIVE). Scoped to
 * those two ops and to border {s,c} objects only: any other op, a different colour or a different style still counts as drift. */
const nullToBlack = v => (v && typeof v === 'object' && v.c === null ? { ...v, c: 0 } : v);
const styleKeysOf = (cell, blackIsNull = false) => Object.fromEntries(STYLE_KEYS.map(k => {
  const v = readKey(cell, k);
  return [k, JSON.stringify(blackIsNull && k.startsWith('border.') ? nullToBlack(v) : v)];
}));

const inGrid = (op, a) => { const c = colNum(col(a)); const r = row(a); return c >= op.grid.c1 && c <= op.grid.c2 && r >= op.grid.r1 && r <= op.grid.r2; };
const requestedSides = op => op.want.filter(w => w.key.startsWith('border.')).map(w => w.key.slice(7));

/** Inside: the sides cell `a` carries by the inner-shared-edge write -- Bottom unless it's the last row, Right
 * unless it's the last column (both directions, always -- no `sides` option on 'inside'/'all'). */
function insideSides(op, a) {
  if (!op.inside || !inGrid(op, a)) return [];
  const c = colNum(col(a));
  const r = row(a);
  const s = [];
  if (r < op.grid.r2) s.push('b');
  if (c < op.grid.c2) s.push('r');
  return s;
}

/** The requested sides cell `a` carries directly -- outline's outer-edge sides (if `a` is on that outer edge) plus
 * inside's inner-shared-edge sides (if `a` is inside the grid at all). */
function edgeSides(op, a) {
  if (!op.outline || !inGrid(op, a)) return [...insideSides(op, a)];
  const c = colNum(col(a));
  const r = row(a);
  const on = { t: r === op.grid.r1, b: r === op.grid.r2, l: c === op.grid.c1, r: c === op.grid.c2 };
  return [...op.outline.sides.filter(x => on[x]), ...insideSides(op, a)];
}

/** The sides of cell `a` that face a NEIGHBOUR's directly-requested side (the shared edge, measured) -- a ring
 * neighbour's facing side for outline, or a grid-internal neighbour's facing side for inside/all. */
const facingSides = (op, a) => (!op.outline ? [] : Object.keys(OPPOSITE).filter(x => { const n = neighbour(a, x); return !!n && edgeSides(op, n).includes(OPPOSITE[x]); }));
const borderWant = (op, x) => ({ key: `border.${x}`, expect: op.outline.expect });

/** Every expectation at cell `a`: its own requested keys plus any side it shows only because a neighbour's
 * request shares that physical edge. Needed by L1 (verifyPrimitive) for inside/all: the receiving cell is
 * itself a target cell (unlike outline's ring), so L3's ring-only facing check does not reach it -- and
 * `wantsAt` alone only sees the writing side, not the side a shared edge makes it show after the write.
 * classify()'s SATISFIED check and wholeTaskChecks do NOT need this: both already combine `wantsAt` with a
 * separate `ringMisses(op, ...)` call, and `ringMisses` iterates the WHOLE window (range + ring), a superset
 * of `op.cells` -- so it already covers every facing side `expectationsAt` would add there, outline's ring
 * cells included. Only L1 has no such window-wide check of its own. Remove's `removeFacing` plays the exact
 * same role, for the sides cleared only because a neighbour's edge was the one directly removed. */
const expectationsAt = (op, a, cell) => [...wantsAt(op, a, cell), ...facingSides(op, a).map(x => borderWant(op, x)), ...removeFacing(op, a).map(removeWant)];

/** Every expectation of a range cell: the spec's keys, plus its outline sides -- or, for 'cells'-mode border on
 * a merge MEMBER cell (GH #83), only the requested sides on that merge's own perimeter (mergeAwareSides). `cell`
 * is the cell's own pre-write read (optional -- omitted call sites that can never see a merge member keep the
 * unfiltered behaviour, same as before this change). */
function wantsAt(op, a, cell) {
  if (cell && cell.merge && cell.merge !== a) {
    if (isBorderCellsModeOnly(op)) {
      const byKey = new Map(op.want.map(w => [w.key, w.expect]));
      return mergeAwareSides(op, a, cell).map(x => ({ key: `border.${x}`, expect: byKey.get(`border.${x}`) }));
    }
    // GH #90: a non-anchor merge member carries no font requirement at all -- its value (and so its font axis)
    // is blanked by the merge itself, LIVE-measured UNOBSERVABLE there; the write and the verification both stay
    // on the anchor. `cell.merge !== a` is true for the ANCHOR too (the probe's own second finding: the merge
    // text never equals a single member's own address) -- `isMergeAnchor` is the real distinction. Fill falls
    // through to the unfiltered return below -- every member (anchor included) is a legitimate target.
    if (isFontOnly(op) && !isMergeAnchor(cell, a)) return [];
  }
  return [...op.want, ...edgeSides(op, a).map(x => borderWant(op, x))];
}

/** The range + its one-cell ring, as addresses. */
const windowCells = op => cellsOf(parseRange(op.ring));

/** May side `side` of cell `a` change? Requested on a target; or the SHARED edge of a target whose facing side is
 * requested. GH #83: for 'cells' mode on a merge MEMBER (`cell` is its pre/post-write read, when available),
 * narrowed to that merge's OWN perimeter (mergeAwareSides) -- a bug that leaked a write onto an INTERIOR edge
 * must NOT be silently tolerated here just because the side was globally requested somewhere in the range. */
function sideMayChange(op, a, side, cell) {
  if (op.outline) return edgeSides(op, a).includes(side) || facingSides(op, a).includes(side);
  if (op.remove) return removeFacing(op, a).includes(side);
  // GH #95: the merge-member narrowing is for a TARGET member only; a RING cell that belongs to a merge (e.g. the right
  // edge of a merged label next to the target) keeps the plain shared-edge rule below -- its facing side legitimately
  // changes when the target draws that edge.
  if (isBorderCellsModeOnly(op) && inGrid(op, a) && cell && cell.merge && cell.merge !== a) return mergeAwareSides(op, a, cell).includes(side);
  const sides = requestedSides(op);
  if (!sides.length || !(side in OPPOSITE)) return false;
  if (inGrid(op, a) && sides.includes(side)) return true;
  const n = neighbour(a, side);
  return !!n && inGrid(op, n) && sides.includes(OPPOSITE[side]);
}

/** GH #83: border.mode 'cells' sets each target cell's OWN sides independently, with no shared-edge/ring
 * coordination at all (unlike outline/inside/all) -- the same per-cell independence `boundedEligible` already
 * relies on. X-M01 (cells.cjs write_range) already carries an anchor exemption for VALUE writes; 'cells'
 * border goes further (LIVE-measured, docs/GH83-BORDER-MERGED-DESIGN-20261014.md): writing ONLY
 * the anchor leaves the merged block's own far edge unset, so EVERY member of the merge is a legitimate
 * target here, not just the anchor -- `mergeAwareSides()` below is what keeps each member's write to its OWN
 * perimeter side(s). Nothing broader: not fill/font/numberFormat/align, not outline/inside/all/remove/diagonal. */
const isBorderCellsModeOnly = op => op.want.length > 0 && op.want.every(w => w.key.startsWith('border.')) && !op.outline && !op.inside && !op.remove && !op.diagonal;

/** GH #90 (LIVE-measured test/live/xlsx-gh90-merge-fill-font-probe.live.cjs, R1=R2):
 * a plain range-level `fill` write over a merged area's own range address lands on EVERY member (the
 * MERGE_B shape -- one range call, all four members read the colour back), unlike border.mode 'cells'
 * (per-cell SetBorders). So fill needs no per-member filtering of WHAT is written -- only the membership
 * refusal lifted, below -- and is verified on every member accordingly (LIVE-measured, see the probe cited above). */
const isFillOnly = op => op.want.length > 0 && op.want.every(w => w.key === 'fill') && !op.fillClear && !op.outline && !op.inside && !op.remove && !op.diagonal;

/** GH #90: the SAME range-level write for font lands only where the merge keeps characters -- the anchor;
 * every other member's value is blanked by the merge itself (LIVE-measured, R1=R2), so its font axis is
 * UNOBSERVABLE there (not a gap, the engine's own behaviour) -- `wantsAt()` below drops the requirement for
 * those members entirely instead of letting the unobservable read refuse the whole task. */
const FONT_WANT_KEYS = new Set(['bold', 'font.bold', 'italic', 'font.italic', 'font.name', 'font.size', 'font.color', 'font.underline', 'font.strike', 'font.sub', 'font.sup']);
const isFontOnly = op => op.want.length > 0 && op.want.every(w => FONT_WANT_KEYS.has(w.key)) && !op.outline && !op.inside && !op.remove && !op.diagonal;

/** GH #90: fill and font are both merge-aware the same way border 'cells' is -- membership no longer refuses
 * on its own, only a merge that crosses the requested range's boundary does (`cellRefusal`'s own check). */
const isFillOrFontOnly = op => isFillOnly(op) || isFontOnly(op);

/** X-B10 / L-15 (owner decision "A", internal tracker; LIVE-measured test/live/xlsx-xb10-merged-format-axes-probe.live.cjs,
 * LIVE probe, 1x3 + 2x2 merge, session + reopen): align (h+v), wrap, numberFormat, fill 'none' written through
 * SetFillColor('No Fill') and the border OUTLINE (per-edge-cell SetBorders) each land on EVERY member of a merged area,
 * the merge and the anchor value survive. Exactly these five axes are merge-aware; nothing broader (rotation / reading
 * order / indent / shrink, border inside / all / remove / diagonal, a mix with other keys keep refusing merge membership).
 * ClearFormats() is NOT among them: LIVE it DISSOLVES the merge, so the X-F08 clear composite never runs on a merge
 * (lower() writes 'No Fill' instead, see `noFillPath`). */
const MERGE_SAFE_KEYS = new Set(['align.h', 'align.v', 'wrap', 'nf']);
const isMergeSafePlainAxes = op => op.want.length > 0 && op.want.every(w => MERGE_SAFE_KEYS.has(w.key)) && !op.fillClear && !op.outline && !op.inside && !op.remove && !op.diagonal;
const isMergeSafeOutline = op => !!op.outline && !op.inside && !op.remove && !op.diagonal && op.want.length === 0 && !op.fillClear;
const isMergeSafeFillClear = op => !!op.fillClear && !op.outline && !op.inside && !op.remove && !op.diagonal;
const isMergeSafeAxis = op => isMergeSafePlainAxes(op) || isMergeSafeOutline(op) || isMergeSafeFillClear(op);

/** X-B10: is `cell` a member of a merged area (the anchor included -- its `merge` text is the whole range, never its own address). */
const inMerge = (cell, a) => !!cell && !!cell.merge && cell.merge !== a;

/** X-B10: fill 'none' over a range that touches ANY merged area is written as a plain 'No Fill' on the range (no ClearFormats
 * composite, no restore: nothing else is touched). A range without a merge keeps the X-F08 composite, unchanged. */
const noFillPath = (op, cellsByAddr) => op.fillClear && op.cells.some(({ a }) => inMerge(cellsByAddr && cellsByAddr[a], a));

/** GH #83: for 'cells' mode, the SIDE keys (t/b/l/r) cell `a` should actually carry -- every requested side
 * for a non-merged cell; for a merge MEMBER, only the requested sides on THAT merge's own outer perimeter
 * (top: its own top row; bottom: its own bottom row; left/right likewise) -- an INTERIOR edge of the block is
 * never written (LIVE-measured: it stays at its prior {s:0} state, session+reopen). `cell` is the
 * cell's PRE-write read (its `.merge` field names the merge's own range text, e.g. 'BA201:BD201'). */
function mergeAwareSides(op, a, cell) {
  const want = requestedSides(op);
  if (!cell || !cell.merge || cell.merge === a) return want;
  const m = parseRange(cell.merge);
  if (!m) return want;
  const c = colNum(col(a));
  const r = row(a);
  return want.filter(x => (x === 't' && r === m.r1) || (x === 'b' && r === m.r2) || (x === 'l' && c === m.c1) || (x === 'r' && c === m.c2));
}

/** GH #90: is address `a` the merge's own top-left (anchor) cell -- the one member a merge keeps the value
 * and font on. `cell.merge` (the merge's own range text, e.g. 'G2:H3') is NEVER equal to a single member's own
 * address, anchor included -- the probe's own second finding -- so position is the only way to tell. */
function isMergeAnchor(cell, a) {
  const m = parseRange(cell.merge);
  return !!m && colNum(col(a)) === m.c1 && row(a) === m.r1;
}

/** A target cell's own refusal before the per-key checks: unread format, inside a merge, a diagonal conflict (X-B05),
 * subscript/superscript:true requested on a non-text cell (X-F05 -- a silent no-op, LIVE m7-s21-probe1). */
function cellRefusal(op, cell, a) {
  if (!cell || !('wrap' in cell) || !('nf' in cell) || !('fill' in cell)) return { reason: 'xlsx-target-format-unread', cell: a };
  // GH #83: 'cells' mode is merge-aware (mergeAwareSides) -- EVERY member of a merge is a legitimate target,
  // not just the anchor (LIVE-measured, see isBorderCellsModeOnly's own comment). GH #90: fill and font are
  // merge-aware the same way -- fill writes and verifies the whole merged area (isFillOnly), font writes and
  // verifies the anchor only (isFontOnly, see wantsAt). X-B10: align / wrap / numberFormat / fill 'none' /
  // border outline are merge-aware too (isMergeSafeAxis -- every member written and verified). Every other mode/key
  // (border inside/all/remove/diagonal, rotation, reading order, indent, shrink, a mix of keys) refuses any merge
  // membership unchanged.
  // L-16/3: border remove is released over a merge through C5 (format-remove.cjs c5Wrap) -- a merged target must lie wholly inside the range, like every merge-aware mode
  const mergeAware = isBorderCellsModeOnly(op) || isFillOrFontOnly(op) || isMergeSafeAxis(op) || !!op.remove;
  if (cell.merge && cell.merge !== a && !mergeAware) return { reason: 'xlsx-format-in-merged-area-refused', cell: a, mergeArea: cell.merge };
  // GH #83 review (and GH #90, same reasoning): a merge-aware write draws its result only from the members
  // INSIDE the range -- a merge that crosses the range boundary would get a partial frame/fill/font (or none)
  // while L1 still holds, so it is refused, named, before any write: the caller widens the range to the whole merge.
  if (cell.merge && cell.merge !== a && mergeAware) {
    const m = parseRange(cell.merge);
    if (!m || m.c1 < op.grid.c1 || m.c2 > op.grid.c2 || m.r1 < op.grid.r1 || m.r2 > op.grid.r2) return { reason: 'xlsx-format-merge-crosses-range-refused', cell: a, mergeArea: cell.merge };
  }
  const fx = fontEffectsRefusal(op.want, cell, a);
  if (fx) return fx;
  const conflict = diagonalConflict(op, cell);
  return conflict ? { cell: a, ...conflict } : null;
}

/** GH #86: a shared-edge border step whose observation still came back truncated above the read bound -- since GH #86
 * (b) these modes HAVE a bounded variant, so this is a task the bounded read could not cover (e.g. mixed with a step
 * that has none, or a window over the bound). The refusal names the bound and, for the drawing modes, the merge-aware
 * 'cells' path. null = no hint. */
function truncatedDetail(op) {
  if (!op || !(op.outline || op.inside || op.remove)) return null;
  const bound = "the workbook is over the 45000-cell read bound and this task's bounded read was not available (border modes 'outline' / 'inside' / 'all' / 'remove' need it)";
  return op.remove ? bound : `${bound}; use border.mode 'cells' with sides ['top','bottom','left','right'] (merge-aware) for every cell edge`;
}

function gate(primary, op) {
  if (!primary || primary.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!primary.complete) {
    const detail = primary.truncated ? truncatedDetail(op) : null;
    return { ok: false, reason: primary.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: primary.unsupported, ...(detail ? { detail } : {}) };
  }
  const s = primary.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: primary.sheets.map(x => x.name) };
  const found = {};
  for (const { a } of op.cells) {
    const cell = cellAt(primary, op.selector.sheet, a);
    const refused = cellRefusal(op, cell, a);
    if (refused) return { ok: false, ...refused };
    for (const w of wantsAt(op, a, cell)) {
      const axis = w.key.split('.')[0];
      // X-F03 (S09 LIVE m7-s09-probe1; spec): bold/name/size/color are MEASURED to
      // work on a formula cell (byte-stable through a value-length change and a reopen, same as
      // a non-formula control). font.italic:true, the last refused axis, is measured the same way
      // (test/live/xlsx-xf03-italic-formula.live.cjs: product path, formula result 10 -> 100000,
      // reopen SATISFIED 0/0, saved file italic + formula preserved, a plain control identical)
      // -- no font axis is refused on a formula cell any more.
      const v = readKey(cell, w.key);
      if (v === undefined || unobservable(v)) return { ok: false, reason: `xlsx-format-${axis === 'nf' ? 'number-format' : axis}-unobservable`, cell: a, why: v === undefined ? 'not-read' : v };
    }
    found[a] = cell;
  }
  // X-B04: op.remove's one check (format-remove.cjs, review point 2): the window gap, `found` extended to the
  // whole window (lower() writes both cells of every removed edge), the unrestorable-value refusal
  // X-F08: op.fillClear's one check (format-fill-clear.cjs): every target cell's kept axes observable +
  // restorable, none diagonal, none sharing an edge with a readable neighbour.
  const window = windowGap(primary, op, found);
  return window ? { ok: false, ...window } : { ok: true, cells: found };
}

/** the one mode-specific window check gate() adds: outline, remove or fill-clear (they are mutually exclusive). */
function windowGap(primary, op, found) {
  if (op.outline) return outlineWindowGap(primary, op);
  if (op.remove) return removeGate(primary, op, found);
  // X-B10: over a merge the clear is a plain 'No Fill' (no restore of kept axes), so the composite's kept-axis /
  // diagonal gate does not apply; the diagonal conflict is still refused by cellRefusal.
  if (op.fillClear) return noFillPath(op, found) ? null : fillClearGate(primary, op, found);
  return null;
}

/** Outline: every cell of the range + ring read with a border, none merged (a merged shared edge is not measured). */
function outlineWindowGap(primary, op) {
  for (const a of windowCells(op)) {
    const cell = cellAt(primary, op.selector.sheet, a);
    if (!cell || !('wrap' in cell)) return { reason: 'xlsx-target-format-unread', cell: a };
    const v = readKey(cell, 'border.t');
    if (v === undefined || unobservable(v)) return { reason: 'xlsx-format-border-unobservable', cell: a, why: v === undefined ? 'not-read' : v };
    // X-B10: a merged area INSIDE the range is a legitimate outline target (isMergeSafeOutline, LIVE-measured); a merged
    // cell in the RING (outside the range) shares an edge whose facing side is not measured and stays refused.
    if (cell.merge && cell.merge !== a && !(isMergeSafeOutline(op) && inGrid(op, a))) return { reason: 'xlsx-format-outline-merged-ring-refused', cell: a, mergeArea: cell.merge };
  }
  return null;
}

/** The window cells whose shared-edge side does not yet show the expected value -- outline's requested edge
 * (facingSides), or remove's REMOVED (removeFacing); the two modes are mutually exclusive. */
function ringMisses(op, p) {
  if (op.remove) return windowCells(op).filter(a => removeFacing(op, a).some(x => !holds(cellAt(p, op.selector.sheet, a), removeWant(x))));
  return windowCells(op).filter(a => facingSides(op, a).some(x => !holds(cellAt(p, op.selector.sheet, a), borderWant(op, x))));
}

/** L3: cell `at`'s shared-edge mismatch, if any -- outline's ring facing side must show EXACTLY the
 * requested edge, remove's facing side must show REMOVED too (a maintainer decision #ref: both cells of every
 * removed edge). One call site for verifyPreservation's loop instead of two separate filter+if blocks. */
function sharedEdgeMismatch(op, at, now) {
  if (op.outline) {
    const facing = facingSides(op, at).filter(x => !holds(now, borderWant(op, x)));
    return facing.length ? { cell: at, sharedEdge: facing.map(x => `border.${x}`), want: op.outline.expect } : null;
  }
  if (op.remove) {
    const facing = removeFacing(op, at).filter(x => !holds(now, removeWant(x)));
    return facing.length ? { cell: at, sharedEdge: facing.map(x => `border.${x}`), want: REMOVED } : null;
  }
  return null;
}

function classify(primary, op) {
  const g = gate(primary, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, range: op.selector.range, keys: op.want.map(w => w.key), ...(op.outline ? { outline: op.outline.sides, inside: op.inside } : {}), ...(op.remove ? { remove: op.remove.sides } : {}) };
  if (op.cells.every(({ a }) => wantsAt(op, a, g.cells[a]).every(w => holds(g.cells[a], w))) && (!(op.outline || op.remove) || !ringMisses(op, primary).length)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-every-cell-every-key' } };
  const runtime = primary.runtime ? primary.runtime.identity : null;
  if (wantsAlignExtWrite(op) && !B05_BRIDGE_WRITE_RUNTIMES.includes(runtime)) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-align-ext-bridge-write-runtime-not-proven', runtime, proven: B05_BRIDGE_WRITE_RUNTIMES } };
  return { state: 'APPLY', evidence: where };
}

function resolve(primary, selector, ctx = {}) {
  const g = gate(primary, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range, before: g.cells }) };
}

/** Outline: [cell, [API side names]] of every edge cell -- the per-cell SetBorders calls, batched in one callCommand. */
const outlineEdges = op => op.cells.map(({ a }) => [a, edgeSides(op, a).map(x => API_SIDE[x])]).filter(([, x]) => x.length);

/** Outline: the batched per-edge-cell SetBorders loop of the command body ('' without an outline). */
const outlineLine = (op, colour) => (op.outline ? `for (const [a, sides] of ${JSON.stringify(outlineEdges(op))}) { const c = ws.GetRange(a); for (const x of sides) c.SetBorders(x, ${JSON.stringify(API_STYLE[op.spec.border.style])}, ${colour(op.spec.border.color)}); }` : '');

/** GH #83: 'cells' mode's address -> [API side names] map -- merge-aware (mergeAwareSides): a non-merged cell
 * keeps every requested side; a merge member keeps only the sides on ITS OWN perimeter. A member whose result
 * is EMPTY (fully interior to a larger block) is simply omitted -- no call, nothing written there. `target.before`
 * is the cell's pre-write read (its `.merge` names the merge, if any -- `resolve()`'s `g.cells`); missing/absent
 * (a caller with no `target`, e.g. a unit test of `lower()` alone) falls back to the unfiltered sides for every
 * cell, same as this family's behaviour before GH #83. */
function cellsSidesByAddr(op, target) {
  const out = {};
  for (const { a } of op.cells) {
    const sides = mergeAwareSides(op, a, target && target.before && target.before[a]).map(x => API_SIDE[x]);
    if (sides.length) out[a] = sides;
  }
  return out;
}

/** One callCommand: exact-name sheet check; range-level public setters (W4a / W4a2), per-cell alignment / borders (W4c). */
/** lower()'s per-cell border calls: the diagonal line(s), or the plain sides of a 'cells' border -- UNLESS this
 * op is border-ONLY (`isBorderCellsModeOnly`), in which case the merge-aware per-cell lookup (`cellsSidesSnippet`
 * below) handles it instead: the family ALSO allows border combined with other format keys in one step (only
 * the PRODUCT surface restricts border to the sole key, write-tools.cjs CONSTRAINTS) -- GH #83's merge-awareness
 * was designed and LIVE-measured for that border-only shape, so a combined request keeps the UNCHANGED, uniform
 * per-cell write (same as before GH #83; a merge in that case is still refused by cellRefusal as before). */
function borderLines(op, s, colour) {
  if (op.diagonal) return [diagonalLine(op, colour)];
  if (!s.border || op.outline || op.remove || isBorderCellsModeOnly(op)) return [];
  return s.border.sides.map(x => `c.SetBorders(${JSON.stringify(API_SIDE[SIDE[x]])}, ${JSON.stringify(API_STYLE[s.border.style])}, ${colour(s.border.color)});`);
}

/** GH #83: the per-cell snippet appended INSIDE the main per-cell loop for border.mode 'cells' ('' for every
 * other mode/key) -- looks up `a` in the once-declared `CELLS_SIDES` map (see `cellsSidesLine`'s caller). */
const cellsSidesSnippet = (op, colour) => (isBorderCellsModeOnly(op) ? `for (const x of (CELLS_SIDES[a] || [])) c.SetBorders(x, ${JSON.stringify(API_STYLE[op.spec.border.style])}, ${colour(op.spec.border.color)});` : '');

/** B05: the Bridge WRITE args of one cell -- ONLY the requested keys, in bridge order (indent, then shrink); null if none. */
function b05Args(al) {
  const args = {}; const requested = [];
  if ('indent' in al) { args.indent = al.indent; requested.push('indent'); }
  if ('shrink' in al) { args.shrink = al.shrink; requested.push('shrink'); }
  return requested.length ? { args, requested } : null;
}

function lower(op, target) {
  const { sheet, range } = op.selector;
  const s = op.spec;
  const colour = hex => (hex === 'auto' ? 'Api.AutoColor()' : `Api.CreateColorFromRGB(${rgbOf(hex).join(', ')})`);
  const J = JSON.stringify;
  const onRange = [];
  const noFill = noFillPath(op, target && target.before);
  if ('fill' in s && s.fill !== 'none') onRange.push(`r.SetFillColor(${colour(s.fill)});`);
  if (noFill) onRange.push("r.SetFillColor('No Fill');");
  if ('wrap' in s) onRange.push(`r.SetWrap(${J(s.wrap)});`);
  if ('numberFormat' in s) onRange.push(`r.SetNumberFormat(${J(s.numberFormat)});`);
  const f = s.font || {};
  if ('bold' in f) onRange.push(`r.SetBold(${J(f.bold)});`);
  if ('italic' in f) onRange.push(`r.SetItalic(${J(f.italic)});`);
  if ('name' in f) onRange.push(`r.SetFontName(${J(f.name)});`);
  if ('size' in f) onRange.push(`r.SetFontSize(${J(f.size)});`);
  if ('color' in f) onRange.push(`r.SetFontColor(${colour(f.color)});`);
  // X-F05: underline/strikeout are RANGE-level setters (ApiRange.SetUnderline/SetStrikeout, same
  // shape as bold/italic/name/size/color above); subscript/superscript have NO range-level
  // setter (LIVE-measured) -- only ApiFont.SetSubscript/SetSuperscript, so those go
  // in onCell below, through GetCharacters(1, text.length) (the gate already refused an empty
  // cell reaching here).
  if ('underline' in f) onRange.push(`r.SetUnderline(${J(f.underline)});`);
  if ('strikeout' in f) onRange.push(`r.SetStrikeout(${J(f.strikeout)});`);
  const onCell = [];
  const al = s.align || {};
  if ('horizontal' in al) onCell.push(`c.SetAlignHorizontal(${J(al.horizontal)});`);
  if ('vertical' in al) onCell.push(`c.SetAlignVertical(${J(al.vertical)});`);
  if ('rotation' in al) onCell.push(`c.SetOrientation(${J(al.rotation)});`);
  // X-F07b/c: the FIRST internal-model WRITE in this family (c.range, not a public ApiRange
  // setter) -- each guarded by its own 'function exists' check, so an unverified build fails
  // closed with a named reason instead of either throwing deep in a try/catch or silently no-op.
  // reading order: the PUBLIC ApiRange.SetReadingOrder (apiBuilder.js:10616) -- OD-A allows an internal call only
  // where no public path serves the capability, and this one does; only its READ is internal (getReadingOrder).
  if ('readingOrder' in al) onCell.push(`c.SetReadingOrder(${J(al.readingOrder)});`);
  // B05-WRITE (M11.5a.4): ONE Bridge WRITE per cell, carrying ONLY the requested keys, at the position the internal
  // setters had. __eoB05Write / __eoB05Fail come from the driver's injected client (no client -> ReferenceError ->
  // caught below -> not ok). A failed write stops the command at once: no second write, no fallback, no rollback.
  const b05 = b05Args(al);
  if (b05) onCell.push(`{ const w = __eoB05Write(a, ${J(b05.args)}, ${J(b05.requested)}); if (!w.ok) return __eoB05Fail(a, w); }`);
  if ('subscript' in f || 'superscript' in f) onCell.push(`{ const text = String(c.GetValue()); const cf = c.GetCharacters(1, text.length).GetFont(); ${scriptCalls(f, J)} }`);
  onCell.push(...borderLines(op, s, colour));
  const cellsSides = isBorderCellsModeOnly(op) ? cellsSidesByAddr(op, target) : null;
  if (cellsSides) onCell.push(cellsSidesSnippet(op, colour));
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${J(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ${b05 ? 'const g05 = __eoB05Gate(name); if (!g05.ok) return g05;' : ''}
  const r = ws.GetRange(${J(range)});
  ${onRange.join('\n  ')}
  ${cellsSides ? `const CELLS_SIDES = ${J(cellsSides)};` : ''}
  ${onCell.length ? `for (const a of ${J(op.cells.map(c => c.a))}) { const c = ws.GetRange(a); ${onCell.join(' ')} }` : ''}
  ${outlineLine(op, colour)}
  ${removeLine(op, target)}
  ${noFill ? '' : fillClearLine(op, target)}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'format', sheet, range, cells: op.cells.map(c => c.a), spec: s, bridgeAlignExtWrite: wantsAlignExtWrite(op), ...(op.outline ? { edges: outlineEdges(op) } : {}), ...(cellsSides ? { cellsSides } : {}), ...removeMeta(op, target), ...(noFill ? { fillClearCells: op.cells.map(c => c.a), fillClearMode: 'no-fill' } : fillClearMeta(op, target)) }),
  };
}

/**
 * Shadow-key co-transitions: writing one axis also changes an UNREQUESTED key, measured (not
 * assumed) on LIVE. Each row [key, otherRequestedKey, newValueOfKey]: `key` may drift to the JSON
 * text `newValueOfKey` without being flagged, as long as `otherRequestedKey` is in `requested`.
 *   bold/font.bold:       font.bold:true also flips internal 'bold' true (W4 LIVE 2026-09-29).
 *   font.bold/bold:       bold:false drops the public font.bold read to 'null' (S07 point 3 --
 *                         the public reader never reports a hard false, only true | null).
 *   font.italic/italic:   italic:false drops the public font.italic read to 'null' (S07 point 3,
 *                         same rule, other axis).
 *   italic/font.italic:   font.italic:true also flips internal 'italic' true -- measured on LIVE
 *                         (m7-f01-live1, cell D9), mirroring bold's own W4 finding.
 *   font.sub/font.sup:    requesting superscript:true drops a pre-existing subscript to 'null'
 *                         (X-F05, LIVE m7-s21-probe1: the two exclude each other).
 *   font.sup/font.sub:    the symmetric case -- requesting subscript:true drops a pre-existing
 *                         superscript to 'null'.
 */
const SHADOW_TOLERANCE = Object.freeze([
  ['bold', 'font.bold', 'true'],
  ['font.sub', 'font.sup', 'null'],
  ['font.sup', 'font.sub', 'null'],
  ['font.bold', 'bold', 'null'],
  ['font.italic', 'italic', 'null'],
  ['italic', 'font.italic', 'true'],
]);
const tolerated = (k, n, requested) => SHADOW_TOLERANCE.some(([key, other, val]) => key === k && requested.has(other) && n[k] === val);

/**
 * The style keys of `cell` that changed between two reads and are not allowed to. A requested font.bold IS the cell's
 * bold (W4 LIVE 2026-09-29: SetBold through the font turned the cell-level `bold` read true as well) -- so that key
 * may change, but only to true.
 */
function driftOf(op, a, was, now, requested) {
  const blackIsNull = !!(op.remove || op.fillClear);
  const b = styleKeysOf(was, blackIsNull);
  const n = styleKeysOf(now, blackIsNull);
  // GH #90: a non-anchor merge member of a font-only op carries no requirement (wantsAt returns []), but the
  // range-level font write still reaches its underlying storage the same way it reaches the anchor's -- only
  // the READ side is unobservable there (blanked value). A font-axis key changing on such a member (inside
  // op's OWN requested grid, never a ring cell of an unrelated merge elsewhere) is this, not a leak.
  const nonAnchorFontMember = isFontOnly(op) && inGrid(op, a) && now && now.merge && now.merge !== a && !isMergeAnchor(now, a);
  return STYLE_KEYS.filter(k => b[k] !== n[k] && !requested.has(k) && !tolerated(k, n, requested)
    && !(k.startsWith('border.') && sideMayChange(op, a, k.slice(7), now))
    && !(nonAnchorFontMember && FONT_WANT_KEYS.has(k)));
}

/** L1: every requested key holds on every target cell; every other style key and the content unchanged. */
function verifyPrimitive(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  const wrong = [];
  const drift = [];
  const content = [];
  const mergeChanged = [];
  for (const { a: at } of op.cells) {
    const now = cellAt(a, op.selector.sheet, at);
    const requested = new Set(wantsAt(op, at, now).map(w => w.key));
    const miss = expectationsAt(op, at, now).filter(w => !holds(now, w)).map(w => w.key);
    if (miss.length) wrong.push({ cell: at, keys: miss });
    const was = cellAt(b, op.selector.sheet, at);
    if (!was || !now) continue;
    const d = driftOf(op, at, was, now, requested);
    if (d.length) drift.push({ cell: at, keys: d });
    if (was.f !== now.f || (was.f === '' && JSON.stringify(was.v) !== JSON.stringify(now.v))) content.push(at);
    // X-B10: a format write never changes a merge. LIVE, ClearFormats() dissolved one (members read as separate cells).
    if (inMerge(was, at) && now.merge !== was.merge) mergeChanged.push({ cell: at, was: was.merge, now: now.merge });
  }
  return { ok: wrong.length === 0 && drift.length === 0 && content.length === 0 && mergeChanged.length === 0, evidence: { sheet: op.selector.sheet, range: op.selector.range, wrong: wrong.slice(0, 20), otherKeysChanged: drift.slice(0, 20), contentChanged: content.slice(0, 20), mergeChanged: mergeChanged.slice(0, 20), proof: 'fresh-live-per-cell-per-key' } };
}

/** L3: every other cell's content (all sheets) and, on the target sheet, every other read cell's style -- and its MERGE (L-16/3 (a): a format write never changes a merge, and L1 only
 * compares it on op.cells; a merge dissolving on a ring / neighbour cell was invisible to both layers). Only a definite change counts: both reads carry a merge text and they differ. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary;
  const a = after.primary;
  const target = new Set(op.cells.map(c => key(op.selector.sheet, c.a)));
  const prefix = key(op.selector.sheet, '');
  const changed = [];
  const mergeChanged = [];
  let compared = 0;
  for (const [k, was] of Object.entries(b.cells)) {
    if (!target.has(k) && k.startsWith(prefix) && typeof was.merge === 'string') {
      const nowMerge = a.cells[k] && a.cells[k].merge;
      if (typeof nowMerge === 'string' && nowMerge !== was.merge) mergeChanged.push({ cell: k.slice(prefix.length), was: was.merge, now: nowMerge });
    }
    if (target.has(k) || !k.startsWith(prefix) || !('wrap' in was)) continue;
    const now = a.cells[k];
    const at = k.slice(prefix.length);
    if (!now || !('wrap' in now)) { changed.push({ cell: at, change: 'style-unread-after' }); continue; }
    compared += 1;
    const d = driftOf(op, at, was, now, new Set());
    if (d.length) changed.push({ cell: at, keys: d });
    const mismatch = sharedEdgeMismatch(op, at, now);
    if (mismatch) changed.push(mismatch);
  }
  return { ok: rest.ok && changed.length === 0 && mergeChanged.length === 0, evidence: { ...rest.evidence, styleCompared: compared, styleChanged: changed.slice(0, 20), styleChangedCount: changed.length, mergeChanged: mergeChanged.slice(0, 20), mergeChangedCount: mergeChanged.length, ...scopeOf(b) } };
}

/**
 * X-N01b: what this L3 proves, machine-readable. FULL: the target sheet's every read cell's style (the sheet-wide
 * style read). BOUNDED (the reduced observation above the read bound, owner decision A): the style of the target and
 * its ring only; values and number formats of the rest from the digest. A BOUNDED pass is never a full style proof.
 */
function scopeOf(b) {
  if (!(b && b.bounded === true)) return { preservationScope: 'FULL' };
  return {
    preservationScope: 'BOUNDED',
    proven: ['target-value', 'target-requested-number-format', 'target-other-style', 'ring-style-and-value', 'sheet-values-outside-windows-digest', 'sheet-number-formats-outside-windows-digest'],
    notProven: ['style-outside-target-and-ring'],
  };
}

function wholeTaskChecks(op, finalObservation) {
  const p = finalObservation && finalObservation.primary;
  const wrong = op.cells.filter(({ a }) => { const now = cellAt(p, op.selector.sheet, a); return !wantsAt(op, a, now).every(w => holds(now, w)); }).map(({ a }) => a);
  const ring = (op.outline || op.remove) ? ringMisses(op, p) : [];
  const what = [...op.want.map(w => w.key), ...(op.outline ? [`the outline ${op.outline.sides.join('')}${op.inside ? ' + inside' : ''} with its shared ring edges`] : []), ...(op.remove ? [`the removed ${op.remove.sides.join('')} gone on both cells of every shared edge`] : [])].join(', ');
  return [{ name: `every cell of ${op.selector.sheet}!${op.selector.range} holds ${what}`, ok: wrong.length === 0 && ring.length === 0, detail: wrong.length || ring.length ? { wrong: wrong.slice(0, 20), ring: ring.slice(0, 20) } : null }];
}

/** The target plus a one-cell ring, read WITH style; the target sheet's used range with style too (L3). */
const windowsOf = op => [{ sheet: op.selector.sheet, range: op.ring, style: true }];
const liveExtras = op => ({ styleSheets: [op.selector.sheet] });

/**
 * X-N01b (owner decision A) + X-L13/8 (the (b)-group BOUNDED classification, owner decision --
 * see docs/X-L13-LARGE-BOOK-MATRIX-20261013.md): a step whose ONLY cell-level effect is one or
 * more PLAIN style fields (any `want` keys, including number format) may be observed in the BOUNDED
 * tier when the full observation does not fit the documented read bound. The extras that replace
 * the sheet-wide style read: the sheet's number formats in the digest (formatSheets). The target +
 * ring window stays read WITH style -- tier-independent, so the target + its ring remain fully
 * proven either way. Below the bound the request above is used unchanged (full L3).
 *
 * GH #86 (b) (owner GO, "igen"): the shared-edge border modes (`outline` / `inside` /
 * `all` / `remove`), kept out by X-L13/8 ("not proven unsafe", not measured LIVE), are ELIGIBLE now. Every cell
 * their checks reach -- the target, and the ring cell across each shared edge that `sharedEdgeMismatch()` /
 * `ringMisses()` / the outline+remove window gates read -- is inside the target + one-cell ring window, which the
 * bounded request still reads WITH style (`windowsOf`, tier-independent). What BOUNDED gives up is the same as for
 * every other format key: the style of cells outside the target and ring (scopeOf: notProven). An outline-only op
 * carries no `want` key of its own (its expectation lives on op.outline), hence the explicit `op.outline` term.
 *
 * null = this step has no bounded variant.
 */
const boundedEligible = op => op.want.length > 0 || !!op.outline;
const boundedExtras = op => (boundedEligible(op) ? { formatSheets: [op.selector.sheet] } : null);

module.exports = { INTENTS, B05_BRIDGE_WRITE_RUNTIMES, b05Args, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, boundedExtras, expectationsOf, holds, H, V, STYLE, MAX_CELLS };
