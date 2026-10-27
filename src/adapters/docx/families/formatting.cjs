'use strict';

const { norm, blockText, resolveTextUnit } = require('../inventory.cjs');
const { resolveStableTarget } = require('../resolver.cjs');

/**
 * DOCX `formatting` family: format_text (M4 Wave 1).
 *
 * EXTRACT_SEMANTICS from the source oracle (joskapista44/euro-mcp f8c6d6cf, read-only):
 *   wanted      docx-session-formatting.cjs normalizeWanted / semanticMatches (same keys, same
 *               comparison rules: case-insensitive names, 0.01 numeric tolerance)
 *   lower       coedit-docx-formatting.cjs:45 (exact unique range; run properties on the range,
 *               paragraph properties on its paragraph, style only if the document has it)
 *   L1          docx-session-formatting.cjs:98-102 the wanted formatting is observed
 *   integral    docx-session-formatting.cjs:103-108 PRIMITIVE_INTEGRAL (preservation audit): the
 *               target text is still there exactly once AND the live round-trip is safe
 *               (docx-round-trip-safety.cjs verifyLiveRoundTrip: every OTHER paragraph/heading
 *               survives, table shapes and images survive). NOT the rewrite's seven run properties.
 *
 * W1.1 -- a substring target (`matchText`, source docx-agent-w012-live-lowering.cjs:36-43):
 *   target      the PARENT paragraph/heading by the semantic selector (or ref), then the LITERAL
 *               matchText exactly once inside it. More than once in the parent -> FAIL_CLOSED (the
 *               source has no occurrence parameter). The owner's contract (W1.1 point 9) replaces
 *               the source's global-uniqueness rule: the paragraph context discriminates, so the
 *               same text in ANOTHER paragraph is not an ambiguity -- but the parent paragraph's own
 *               text must be addressable once in the whole document (the write finds it by Search).
 *   observe     the needle-free run list of the fresh LIVE read (inventory block.runs): the range's
 *               formatting is the composition of the runs it overlaps. Runs that disagree = MIXED
 *               (a definite, not-terminal state); a run whose property is unread = UNKNOWN
 *               (not provable -> FAIL_CLOSED). No run list (a non-run child) -> FAIL_CLOSED.
 *   write       one callCommand through the common dispatch: the parent by its exact text, then
 *               ApiParagraph.Search(matchText) exactly once; text properties on that range,
 *               paragraph properties on the paragraph (source coedit-docx-formatting.cjs:45 parity).
 *   L1 / L2     the range itself, from the runs -- never a whole-paragraph property.
 *   L3          the whole-paragraph integral set above, PLUS per character: outside the range all
 *               seven text properties unchanged, inside it every NON-requested one unchanged, and
 *               every other paragraph's run formatting unchanged. (Stronger than the source, which
 *               checks text only -- W1.1 point 11 / 13.)
 *
 * W-08 (zero internal tracker, this session): font colour / highlight / vertical alignment (sub-
 * /superscript) added to the SAME pipeline as the other TEXT_KEYS -- measured PUBLIC r/w on both
 * ApiRun and run.GetTextPr() (ApiTextPr): SetColor/GetColor, SetHighlight/GetHighlight,
 * SetVertAlign/GetVertAlign. The call is direct on `range`, the SAME way SetBold/SetFontFamily
 * already are (not through GetTextPr()) -- both paths exist per the measurement, this one matches
 * the existing style.
 *
 * CORRECTED after the FIRST LIVE acceptance: the first attempt's SetColor argument
 * used `Api.CreateColorFromRGB`, cross-referenced from an XLSX probe (xlsx-m7-s22-font-values-
 * probe.live.cjs) -- that constructor DOES NOT EXIST on the DOCX Builder (measured: TypeError).
 * LIVE-measured DOCX colour constructors: Api.RGB, Api.RGBA, Api.HexColor, Api.CreateRGBColor,
 * Api.AutoColor, Api.ThemeColor, Api.CreatePresetColor; `range.SetColor(Api.HexColor('RRGGBB'))`
 * is the one actually measured working (docx-w08-getcolor-shape-probe, w08-colorshape2). Whether
 * ApiRange.SetColor accepts anything beyond this is NOT-MEASURED and not assumed.
 *
 * GetColor's return shape is ALSO now measured, not hypothesized: on a run with no explicit colour
 * it returns `null` (a definite, observable state -- 'auto', not "unobservable"); after a write it
 * returns an ApiColor object whose GetHex() gives '#RRGGBB' and IsAutoColor() a boolean. GetRGB()
 * returns an {r,g,b} OBJECT, not a number or string -- the first hypothesis's speculative branches
 * (GetRGB as number/string, GetR/GetG/GetB, bare r/g/b properties) were all unmeasured guesses and
 * are removed; colorHex now reads exactly the measured shape (null -> 'auto', IsAutoColor() -> same,
 * else GetHex()). Highlight and vertAlign's plain-string hypothesis was CONFIRMED correct by the
 * same LIVE run (GetHighlight()/GetVertAlign() echo back what was set; the fixture's own default is
 * 'none' / 'baseline').
 * W-09 (indentation + shading, implementation phase + the matrix probe it asked for):
 *   indentLeft/indentRight/indentFirstLine, like spaceBefore/spaceAfter, are POINTS externally and
 *   twips (`* 20` / `/ 20`) at the Builder boundary -- not a new unit family, the existing one
 *   (measured: `SetIndLeft(720)` round-trips as 720 in the saved `<w:ind>`, i.e. the Builder's own
 *   unit is twips already). A negative indentFirstLine is accepted as-is (hanging indent); the
 *   Builder itself stores it as `w:hanging` in the file -- the adapter does not special-case it.
 *   shading is a 6-digit hex string (e.g. "FFFF00"), rejected otherwise -- measured: the ONLY
 *   call shape that does not silently write the wrong colour is `SetShd('clear', r, g, b)` with
 *   r/g/b as separate integers; every hex-string variant of SetShd wrote black (or white) WITHOUT
 *   throwing, so the hex-to-RGB split happens here, never passed through as a string.
 */

const BOOLEAN_KEYS = Object.freeze(['bold', 'italic', 'underline', 'strike']);
const NUMBER_KEYS = Object.freeze(['fontSize', 'spaceBefore', 'spaceAfter', 'lineSpacing', 'indentLeft', 'indentRight', 'indentFirstLine']);
const NAME_KEYS = Object.freeze(['fontFamily', 'style', 'language']);
const ALIGNMENTS = Object.freeze(['left', 'right', 'center', 'justify']);
const HEX_RE = /^#[0-9a-fA-F]{6}$/;
const HIGHLIGHT_VALUES = Object.freeze(['none', 'black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'red', 'white', 'yellow']);
const VERT_ALIGN_VALUES = Object.freeze(['baseline', 'subscript', 'superscript']);
const SHADING_KEY = 'shading';
const SHADING_HEX_RE = /^[0-9a-fA-F]{6}$/;
// W-09 (tabs / per-side border / pagination flags, W09-1a..1e MEASURED):
// tab align and border style are each restricted to the ONE value this session's own LIVE probes
// measured -- 'center'/'decimal' tab alignment and any border style other than 'single' are
// explicitly NOT MEASURED (the design doc's own point 4) and refused, not guessed.
// GH #290 (docs/evidence/gh290-paragraph-format-public-api-20261021.json): 'center' is stored and survives a reopen ->
// released; 'decimal' (and an unknown value) is silently stored as 'left', and SetTabs ignores a leader -> still refused.
// Border 'none' writes an explicit w:val="none" (the other sides unchanged) -> released; double/dotted/dashed/thick are
// silently stored as 'none' (the border disappears) -> still refused.
const TAB_ALIGNS = Object.freeze(['left', 'right', 'center']);
const BORDER_STYLES = Object.freeze(['single', 'none']);
const BORDER_SIDES = Object.freeze(['Top', 'Bottom', 'Left', 'Right']);
const BORDER_KEYS = Object.freeze(BORDER_SIDES.map(s => `border${s}`));
// PAGINATION_KEYS do NOT join BOOLEAN_KEYS: that loop blindly coerces any value with Boolean(),
// which would silently accept `false` as a request -- clearing (false) is explicitly NOT
// MEASURED (design doc point 4) and must be refused with a named reason, not silently no-op'd.
const PAGINATION_KEYS = Object.freeze(['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl']);
const KNOWN_KEYS = Object.freeze([...BOOLEAN_KEYS, ...NUMBER_KEYS, ...NAME_KEYS, 'alignment', 'color', 'highlight', 'vertAlign', SHADING_KEY, 'tabs', ...BORDER_KEYS, ...PAGINATION_KEYS]);
// shading compares like a NAME_KEY (case-insensitive string) without joining NAME_KEYS itself --
// NAME_KEYS also drives normalizeWanted's blind trim-and-accept loop, and shading must instead be
// format-validated (SHADING_HEX_RE) and rejected outright when malformed (measured: a malformed
// SetShd argument silently writes the wrong colour instead of throwing -- the adapter must not
// repeat that at its own boundary).
const CASE_INSENSITIVE_KEYS = Object.freeze([...NAME_KEYS, SHADING_KEY]);
const FORMATTABLE_KINDS = Object.freeze(['paragraph', 'heading']);
const TEXT_KEYS = Object.freeze(['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'language', 'color', 'highlight', 'vertAlign']);
// W-09's new fields, for L3's own extension below -- every OTHER paragraph's state on these
// keys must be unchanged (the whole-paragraph L3 previously only checked TEXT/EXISTENCE
// survival of other paragraphs, never their format -- this is new coverage, scoped to exactly
// these new keys, not a retroactive fix of the pre-existing fields).
const NEW_PARA_LEVEL_KEYS = Object.freeze(['tabs', ...BORDER_KEYS, ...PAGINATION_KEYS]);
// GH #290: L3 on the OTHER paragraphs also compares the explicit (direct) pagination state -- a stray explicit `false`
// reads exactly like an absent flag through the boolean keys
const L3_PARA_LEVEL_KEYS = Object.freeze([...NEW_PARA_LEVEL_KEYS, 'paginationDirect']);

/** W-09 shading: absent -> nothing; else exactly six hex digits (an optional leading '#'), upper-cased. */
function normalizeShading(value) {
  if (value == null) return { ok: true, value: undefined };
  const hex = String(value).trim().replace(/^#/, '');
  if (!SHADING_HEX_RE.test(hex)) return { ok: false, detail: 'shading must be a 6-digit hex colour (e.g. "FFFF00")' };
  return { ok: true, value: hex.toUpperCase() };
}

/** W-09 tabs: an array of {pos (points, >= 0), align}. align restricted to TAB_ALIGNS -- a tab
 * leader is NOT accepted at all (the MEASURED write never passes one; NOT MEASURED for v1).
 * GH #290 F-02 (docs/evidence/gh290-paragraph-format-public-api-20261021.json P3): an EMPTY list removes the paragraph's own
 * tabs (SetTabs([], []) -- no w:tabs stored, the read gives []); a tab the paragraph's style defines is not touched. */
function normalizeTabs(value) {
  if (value == null) return { ok: true, value: undefined };
  if (!Array.isArray(value)) return { ok: false, detail: 'tabs must be an array of {pos, align} ([] removes the paragraph\'s own tabs)' };
  const out = [];
  for (const t of value) {
    if (!t || typeof t !== 'object') return { ok: false, detail: 'each tabs entry must be an object {pos, align}' };
    const pos = Number(t.pos);
    if (!Number.isFinite(pos) || pos < 0) return { ok: false, detail: 'tabs[].pos must be a non-negative number (points)' };
    const align = String(t.align == null ? '' : t.align).trim();
    if (!TAB_ALIGNS.includes(align)) return { ok: false, detail: `tabs[].align must be one of ${TAB_ALIGNS.join(', ')} (decimal is silently stored as left by the editor -- refused)` };
    out.push({ pos, align });
  }
  return { ok: true, value: out };
}

/** W-09 border side: {style (restricted to BORDER_STYLES), size (points, > 0), color (#RRGGBB)}. */
function normalizeBorderSide(key, value) {
  if (value == null) return { ok: true, value: undefined };
  if (!value || typeof value !== 'object' || Array.isArray(value)) return { ok: false, detail: `${key} must be an object {style, size, color}` };
  const style = String(value.style == null ? '' : value.style).trim();
  if (!BORDER_STYLES.includes(style)) return { ok: false, detail: `${key}.style must be one of ${BORDER_STYLES.join(', ')} (double/dotted/dashed/thick are silently stored as none by the editor -- refused)` };
  // GH #290: 'none' removes the side (an explicit w:val="none", which also overrides a border the style defines)
  if (style === 'none') {
    if (Object.keys(value).some(k => k !== 'style')) return { ok: false, detail: `${key} {style: 'none'} takes no size or color` };
    return { ok: true, value: { style: 'none', size: 0, color: '000000' } };
  }
  const size = Number(value.size);
  if (!Number.isFinite(size) || size <= 0) return { ok: false, detail: `${key}.size must be a positive number (points)` };
  // stored WITHOUT the leading '#' -- matches both the driver read's own shape (borderSideFormat
  // builds a bare 6-hex-digit string from r/g/b) and the write site's slice(0,2)/slice(2,4)/
  // slice(4,6) RGB split (SetTopBorder etc. take separate r/g/b ints, like SetShd already does)
  const color = String(value.color == null ? '' : value.color);
  if (!HEX_RE.test(color)) return { ok: false, detail: `${key}.color must be #RRGGBB` };
  return { ok: true, value: { style, size, color: color.replace(/^#/, '').toUpperCase() } };
}

/** W-09 tabs deep-equality: same length, each entry's pos within 0.01pt and align exact. */
function tabsEqual(actual, wanted) {
  if (!Array.isArray(actual) || actual.length !== wanted.length) return false;
  return wanted.every((w, i) => actual[i] && Number.isFinite(Number(actual[i].pos)) && Math.abs(Number(actual[i].pos) - w.pos) < 0.01 && actual[i].align === w.align);
}

/** W-09 border side deep-equality: style exact, size within 0.01pt, color case-insensitive. */
function borderSideEqual(actual, wanted) {
  if (!actual || typeof actual !== 'object') return false;
  // GH #290: a removed side is an explicit 'none' (size / colour carry nothing); an ABSENT side is not proof -- the
  // paragraph's style may define a border there
  if (wanted.style === 'none') return actual.style === 'none';
  return actual.style === wanted.style && Number.isFinite(Number(actual.size)) && Math.abs(Number(actual.size) - wanted.size) < 0.01 && String(actual.color || '').toUpperCase() === wanted.color.toUpperCase();
}

/** Source normalizeWanted, plus an explicit refusal of anything it would silently drop. */
/** W-08's three new keys, split out to keep normalizeWanted's own complexity down (
 * the same remedy X-P02 used). Mutates `out`; returns an error detail string, or null on success. */
function normalizeW08(format, out) {
  if (format.color != null) {
    const c = String(format.color);
    if (!HEX_RE.test(c)) return 'color must be #RRGGBB';
    out.color = c.toUpperCase();
  }
  if (format.highlight != null) {
    const h = String(format.highlight);
    if (!HIGHLIGHT_VALUES.includes(h)) return `highlight must be one of ${HIGHLIGHT_VALUES.join(', ')}`;
    out.highlight = h;
  }
  if (format.vertAlign != null) {
    const v = String(format.vertAlign);
    if (!VERT_ALIGN_VALUES.includes(v)) return `vertAlign must be one of ${VERT_ALIGN_VALUES.join(', ')}`;
    out.vertAlign = v;
  }
  return null;
}

function normalizeWanted(format) {
  if (!format || typeof format !== 'object' || Array.isArray(format)) return { ok: false, detail: 'format must be an object' };
  const unknown = Object.keys(format).filter(k => !KNOWN_KEYS.includes(k));
  if (unknown.length) return { ok: false, detail: `unknown format properties: ${unknown.join(', ')}` };
  const out = {};
  for (const k of BOOLEAN_KEYS) if (k in format) out[k] = Boolean(format[k]);
  for (const k of NUMBER_KEYS) {
    if (format[k] == null) continue;
    const n = Number(format[k]);
    if (!Number.isFinite(n)) return { ok: false, detail: `${k} must be a number` };
    out[k] = n;
  }
  for (const k of NAME_KEYS) if (format[k] != null) out[k] = String(format[k]).trim();
  if (format.alignment != null) {
    const a = String(format.alignment).toLowerCase();
    if (!ALIGNMENTS.includes(a)) return { ok: false, detail: `alignment must be one of ${ALIGNMENTS.join(', ')}` };
    out.alignment = a;
  }
  const w08err = normalizeW08(format, out);
  if (w08err) return { ok: false, detail: w08err };
  const shading = normalizeShading(format.shading);
  if (!shading.ok) return shading;
  if (shading.value !== undefined) out.shading = shading.value;
  const tabsR = normalizeTabs(format.tabs);
  if (!tabsR.ok) return tabsR;
  if (tabsR.value !== undefined) out.tabs = tabsR.value;
  for (const k of BORDER_KEYS) {
    const r = normalizeBorderSide(k, format[k]);
    if (!r.ok) return r;
    if (r.value !== undefined) out[k] = r.value;
  }
  // W-09 pagination flags: a real boolean only (never coerced). GH #290: `false` is released -- SetKeepNext(false) etc.
  // write an explicit w:val="false", which overrides a style that sets the flag (Heading 1 keepNext, measured). GH #290 F-07
  // (LIVE, docs/evidence/gh289-290-remaining-probe-20261023.json): SetWidowControl(false) the same (stored, kept after a reopen).
  for (const k of PAGINATION_KEYS) {
    if (format[k] == null) continue;
    if (typeof format[k] !== 'boolean') return { ok: false, detail: `${k} must be a boolean` };
    out[k] = format[k];
  }
  if (Object.keys(out).length === 0) return { ok: false, detail: 'empty format' };
  return { ok: true, wanted: Object.freeze(out) };
}

/**
 * Source semanticMatches. A wanted property the live read did not return AT ALL (undefined) is
 * UNOBSERVABLE, not a mismatch. W-09 (tabs / border sides): `null` here is a DEFINITE "not set"
 * state (no w:tabs / no that pBdr side), not unobservable -- UNLIKE every other field here,
 * where `null` still means unreadable. Checked BEFORE the generic null-is-unobservable branch.
 */
function compare(actual, wanted) {
  const missing = [];
  const unobservable = [];
  for (const [k, v] of Object.entries(wanted)) {
    if (!actual || actual[k] === undefined) { unobservable.push(k); continue; }
    if (k === 'tabs') {
      if (!tabsEqual(actual[k], v)) missing.push({ property: k, expected: v, actual: actual[k] });
      continue;
    }
    if (BORDER_KEYS.includes(k)) {
      if (!borderSideEqual(actual[k], v)) missing.push({ property: k, expected: v, actual: actual[k] });
      continue;
    }
    // GH #290: a `false` flag is proven only by an EXPLICIT direct false (the driver's paginationDirect) -- an absent
    // flag reads false but the paragraph's style may set it (Heading 1 keepNext), so it is not the requested state
    if (PAGINATION_KEYS.includes(k) && v === false) {
      const dv = actual.paginationDirect ? actual.paginationDirect[k] : undefined;
      if (dv === undefined) unobservable.push(k);
      else if (dv !== false) missing.push({ property: k, expected: false, actual: dv === null ? 'absent (the style may set it)' : dv });
      continue;
    }
    if (actual[k] === null) { unobservable.push(k); continue; }
    if (CASE_INSENSITIVE_KEYS.includes(k)) {
      if (String(actual[k]).trim().toLowerCase() !== String(v).trim().toLowerCase()) missing.push({ property: k, expected: v, actual: actual[k] });
    } else if (NUMBER_KEYS.includes(k)) {
      if (!Number.isFinite(Number(actual[k])) || Math.abs(Number(actual[k]) - Number(v)) > 0.01) missing.push({ property: k, expected: v, actual: actual[k] });
    } else if (actual[k] !== v) {
      missing.push({ property: k, expected: v, actual: actual[k] });
    }
  }
  return { ok: missing.length === 0 && unobservable.length === 0, missing, unobservable };
}

function plan(step, index) {
  let matchText;
  if (step.matchText !== undefined) {
    if (typeof step.matchText !== 'string' || step.matchText === '' || !step.matchText.trim()) return { ok: false, detail: `step ${index}: matchText must be a non-empty string` };
    matchText = step.matchText;
  }
  if (!step.selector || typeof step.selector !== 'object') return { ok: false, detail: `step ${index}: format_text requires a selector${matchText !== undefined ? ' (the parent paragraph of matchText)' : ''}` };
  if (step.selector.contains === true) return { ok: false, detail: `step ${index}: format_text on a contained substring is not migrated in M4 Wave 1` };
  const w = normalizeWanted(step.format);
  if (!w.ok) return { ok: false, detail: `step ${index}: ${w.detail}` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'format_text',
      selector: step.selector,
      ...(matchText !== undefined ? { matchText } : {}),
      format: w.wanted,
      preservationMode: 'PRIMITIVE_INTEGRAL',
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

/**
 * The selector this operation resolves by. A ref names an earlier operation's semantic result; the
 * adapter planner records that operation's planned semantic text as `op.refText` (adapter.cjs), and
 * resolve() cross-checks it against the ref's actual evidence.
 */
function effectiveSelector(op) {
  if (op.selector && op.selector.ref !== undefined) return typeof op.refText === 'string' ? { text: op.refText } : null;
  return op.selector;
}

function resolveFormattable(primary, selector) {
  if (!selector) return { ok: false, target: null, selector, ambiguity: 'REF_WITHOUT_TEXT' };
  const r = resolveStableTarget(primary, selector);
  if (r.ok && !FORMATTABLE_KINDS.includes(r.target.kind)) {
    return { ok: false, target: null, selector, ambiguity: 'UNSUPPORTED_TARGET_KIND', kind: r.target.kind };
  }
  return r;
}

const isRange = op => typeof op.matchText === 'string';
const splitWanted = wanted => ({
  text: Object.fromEntries(Object.entries(wanted).filter(([k]) => TEXT_KEYS.includes(k))),
  para: Object.fromEntries(Object.entries(wanted).filter(([k]) => !TEXT_KEYS.includes(k))),
});
const occurrences = (hay, needle) => (needle ? String(hay).split(needle).length - 1 : 0);

/** The effective text formatting of every character of a covered block (from its runs). */
function charFormats(block) {
  if (!block || !Array.isArray(block.runs) || typeof block.liveText !== 'string') return null;
  const out = [];
  for (const r of block.runs) for (let i = r.start; i < r.end; i += 1) out[i] = r.format;
  return out.length === block.liveText.length ? out : null;
}

/**
 * The formatting of the exact range `matchText` inside `block`, composed from the runs it overlaps.
 * Per text key: `format[k]` when every character agrees; `mixed` when they disagree; `unknown` when
 * any character's run did not return the property.
 */
function rangeState(block, matchText) {
  const chars = charFormats(block);
  if (!chars) return { ok: false, reason: 'range-format-unobservable' };
  const n = occurrences(block.liveText, matchText);
  if (n === 0) return { ok: false, reason: 'range-absent' };
  if (n > 1) return { ok: false, reason: 'range-ambiguous-in-paragraph', occurrences: n };
  const start = block.liveText.indexOf(matchText);
  const end = start + matchText.length;
  const format = {};
  const mixed = [];
  const unknown = [];
  for (const k of TEXT_KEYS) {
    const vals = chars.slice(start, end).map(f => f[k]);
    if (vals.some(v => v === undefined || v === null)) unknown.push(k);
    else if (vals.every(v => JSON.stringify(v) === JSON.stringify(vals[0]))) format[k] = vals[0];
    else mixed.push(k);
  }
  return { ok: true, start, end, format, mixed, unknown, chars };
}

/** The parent must be addressable by its text exactly once in the whole document (the write's Search). */
function parentAddressable(primary, block) {
  const t = norm(block.text);
  let n = 0;
  for (const b of (primary && primary.blocks) || []) {
    if (b.kind === 'table') for (const row of b.rows || []) for (const c of row) n += occurrences(norm(c.text), t);
    else if (b.kind === 'list') for (const it of b.items || []) n += occurrences(norm(it.text), t);
    else n += occurrences(norm(b.text), t);
  }
  return n === 1;
}

/** Range classification / verification core: {state, evidence} with state HOLDS | NOT_TERMINAL | FAIL_CLOSED. */
function rangeTerminal(primary, block, op) {
  if (!parentAddressable(primary, block)) return { state: 'FAIL_CLOSED', evidence: { reason: 'format-parent-not-addressable', text: block.text } };
  const r = rangeState(block, op.matchText);
  if (!r.ok) return { state: 'FAIL_CLOSED', evidence: { reason: r.reason, text: block.text, matchText: op.matchText, occurrences: r.occurrences } };
  const w = splitWanted(op.format);
  const unknown = Object.keys(w.text).filter(k => r.unknown.includes(k));
  if (unknown.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'range-terminal-property-unobservable', properties: unknown } };
  const para = compare(block.format, w.para);
  if (para.unobservable.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'format-terminal-property-unobservable', properties: para.unobservable } };
  const mixed = Object.keys(w.text).filter(k => r.mixed.includes(k));
  const text = compare(r.format, Object.fromEntries(Object.entries(w.text).filter(([k]) => !mixed.includes(k))));
  const holds = mixed.length === 0 && text.ok && para.ok;
  return {
    state: holds ? 'HOLDS' : 'NOT_TERMINAL',
    evidence: { text: block.text, matchText: op.matchText, range: [r.start, r.end], observed: r.format, mixed, missing: [...text.missing, ...para.missing] },
    range: r,
  };
}

/**
 * Fresh LIVE classification (owner brief Wave 1 point 4):
 *   target ambiguous / absent / not formattable            -> FAIL_CLOSED
 *   formatting state unreadable                            -> FAIL_CLOSED
 *   a wanted property the live read cannot observe         -> FAIL_CLOSED (terminal not provable)
 *   every wanted property already holds                    -> SATISFIED (0 command, 0 write)
 *   otherwise (none or part of it holds)                   -> APPLY
 */
/**
 * W2.1: under a declared later list step, the formatted paragraph may already be a list item. The
 * postcondition is then read from that TEXT UNIT (pure text / ref selectors only). HOLDS -> the
 * terminal state is proven; anything else fails closed -- a format is never re-applied blind onto a
 * representation a successor already transformed.
 */
function successorUnit(primary, op) {
  const sel = effectiveSelector(op);
  if (!Array.isArray(op.successorTransforms) || !op.successorTransforms.length || !sel || Object.keys(sel).some(k => k !== 'text')) return null;
  const u = resolveTextUnit(primary, sel.text);
  return u.ok && ['paragraph', 'list_item'].includes(u.unit.kind) ? u.unit : { missing: true, reason: u.reason || 'not-a-text-unit' };
}

function successorHolds(primary, node, op) {
  if (isRange(op)) return rangeTerminal(primary, node, op).state === 'HOLDS';
  const m = node.format ? compare(node.format, op.format) : { ok: false };
  return m.ok;
}

/**
 * W-10-9: `primary.document.styles` lists every style the observation could read --
 * an unknown `wanted.style` name is refused PRE-DISPATCH, with a NAMED reason, the SAME way every
 * other unobservable/ambiguous classification in this family already fails closed.
 *
 * CORRECTED (LIVE regression): the W10-9 comment here claimed "format_text never
 * creates a style, so this family needs none of style-definition.cjs's GetAllStyles-blindness
 * workaround" -- MEASURED FALSE. `doc.GetAllStyles()` does not see ANY custom style, not only a
 * freshly-created one in the same task (the reviewer's own LIVE diagnostic: a `define_style`'d custom
 * style, APPLIED, was then invisible to a `format_text` step requesting that SAME, genuinely
 * existing style -- a case that WORKED before this family's W10-9 fix, because the old bare
 * `doc.GetStyle(name)` call found it directly). `styleNamesOf` below (the SAME shape as
 * style-definition.cjs's own export) lets `docxAdapter.observationRequestFor` (adapter.cjs)
 * include this family's own `format.style` names in the SAME targeted-read request W10-8 already
 * wires through -- the docx-driver merges a targeted hit into `document.styles`, winning over any
 * same-named (or absent) `GetAllStyles()` entry.
 *
 * CORRECTED AGAIN (a SECOND LIVE regression, from W10-10): `buildTargetedStyleQuery`
 * writes `out[name] = null` for a name it was asked for but could not find (docx-driver.cjs) --
 * `readLive`'s union merge then adds that name as a KEY of `document.styles`, with a `null`
 * VALUE. Checking key-presence alone (the ORIGINAL W10-9 shape) therefore now also matches a
 * genuinely MISSING name, once W10-10 started requesting it -- the exact inversion of the pre-
 * dispatch refusal this function exists for. `style-definition.cjs`'s `styleAt` already treats a
 * `null` value the SAME as "not a key at all" (`s.st !== null` is its own existence test) -- this
 * function now does the SAME: a key counts only with a non-null value.
 */
function styleExists(primary, name) {
  const styles = primary && primary.document && primary.document.styles;
  if (!styles || typeof styles !== 'object') return false;
  const want = String(name).trim().toLowerCase();
  return Object.entries(styles).some(([k, v]) => k.trim().toLowerCase() === want && v !== null);
}

/** W-10-10: the style names this task's own `format_text` steps request -- exported
 * so `docxAdapter.observationRequestFor` can merge them with style-definition.cjs's own
 * `styleNamesOf(task)` into one targeted-read request. */
function styleNamesOf(task) {
  const steps = (task && Array.isArray(task.steps)) ? task.steps : [];
  const names = steps.filter(s => s && s.intent === 'format_text' && s.format && typeof s.format.style === 'string').map(s => String(s.format.style).trim());
  return [...new Set(names)];
}

function classify(primary, op) {
  if (op.format && op.format.style !== undefined && !styleExists(primary, op.format.style)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `format-style-does-not-exist:${op.format.style}` } };
  }
  const su = successorUnit(primary, op);
  if (su && !su.missing && su.kind === 'list_item') {
    if (successorHolds(primary, su.node, op)) return { state: 'SATISFIED', evidence: { text: su.text, format: op.format, proof: 'fresh-live-format-terminal-on-text-unit', successors: op.successorTransforms } };
    return { state: 'FAIL_CLOSED', evidence: { reason: 'format-target-transformed-by-successor-not-terminal', text: su.text } };
  }
  const r = resolveFormattable(primary, effectiveSelector(op));
  if (!r.ok) return { state: 'FAIL_CLOSED', evidence: { reason: `format-target-${String(r.ambiguity).toLowerCase()}` } };
  if (!r.target.format) return { state: 'FAIL_CLOSED', evidence: { reason: 'format-state-unreadable', text: r.target.text } };
  if (isRange(op)) {
    const t = rangeTerminal(primary, r.target, op);
    if (t.state === 'FAIL_CLOSED') return t;
    if (t.state === 'HOLDS') return { state: 'SATISFIED', evidence: { ...t.evidence, format: op.format, proof: 'fresh-live-range-format-already-terminal' } };
    return { state: 'APPLY', evidence: t.evidence };
  }
  const m = compare(r.target.format, op.format);
  if (m.unobservable.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'format-terminal-property-unobservable', properties: m.unobservable } };
  if (m.ok) return { state: 'SATISFIED', evidence: { text: r.target.text, format: op.format, proof: 'fresh-live-format-already-terminal' } };
  return { state: 'APPLY', evidence: { text: r.target.text, missing: m.missing } };
}

function resolve(primary, selector, ctx = {}) {
  if (selector && selector.ref !== undefined) {
    const ref = ctx.refs && ctx.refs[selector.ref];
    const text = ref && ref.evidence && ref.evidence.text;
    if (typeof text !== 'string') return { ok: false, target: null, selector, ambiguity: 'REF_WITHOUT_TEXT' };
    const op = ctx.operation;
    if (op && typeof op.refText === 'string' && norm(op.refText) !== norm(text)) {
      return { ok: false, target: null, selector, ambiguity: 'REF_RESULT_MISMATCH' };
    }
    return resolveFormattable(primary, { text });
  }
  return resolveFormattable(primary, selector);
}

function preDispatchCheck(op, target) {
  if (!target.format) return { ok: false, dispatched: false, detail: 'format-state-unreadable' };
  if (isRange(op) && !target.runs) return { ok: false, dispatched: false, detail: 'range-format-unobservable' };
  return { ok: true };
}

function lower(op, target) {
  if (isRange(op)) {
    const parentText = String(target.liveText);
    return {
      kind: 'callCommand',
      body: rangeCommandBody(parentText, op.matchText, op.format),
      meta: Object.freeze({ family: 'formatting', action: 'format', targetText: parentText, matchText: op.matchText, format: op.format }),
    };
  }
  const targetText = String(target.text || '');
  return {
    kind: 'callCommand',
    body: commandBody(targetText, op.format),
    meta: Object.freeze({ family: 'formatting', action: 'format', targetText, format: op.format }),
  };
}

/** Source coedit-docx-formatting.cjs:45 re-expressed: the same API calls, in the same order. */
function commandBody(targetText, wanted) {
  return `
try {
  const targetText = ${JSON.stringify(targetText)};
  const wanted = ${JSON.stringify(wanted)};
  const doc = Api.GetDocument();
  if (!doc || !targetText) return;
  const ranges = doc.Search(targetText, { MatchCase: true }) || [];
  if (ranges.length !== 1) return;
  const range = ranges[0];
  if (String(range.GetText({ Numbering: false }) || '') !== targetText) return;
  const paragraph = range.GetParagraph(0);
  if (!paragraph) return;
  let requestedStyle = null;
  // W-10-9 (the SAME bug class as style-definition.cjs's commandBody, W10-7):
  // doc.GetStyle(<unknown name>) falls back to the 'Normal' style object instead of null -- a
  // match is only valid when the returned style's own GetName() equals the REQUESTED name.
  if (Object.prototype.hasOwnProperty.call(wanted, 'style')) {
    const rawStyle = doc.GetStyle(wanted.style);
    requestedStyle = rawStyle && typeof rawStyle.GetName === 'function' && rawStyle.GetName() === wanted.style ? rawStyle : null;
    if (!requestedStyle) return;
  }
  if ('bold' in wanted) range.SetBold(wanted.bold);
  if ('italic' in wanted) range.SetItalic(wanted.italic);
  if ('underline' in wanted) range.SetUnderline(wanted.underline);
  if ('strike' in wanted) range.SetStrikeout(wanted.strike);
  if ('fontFamily' in wanted) range.SetFontFamily(wanted.fontFamily);
  if ('fontSize' in wanted) range.SetFontSize(wanted.fontSize);
  if ('language' in wanted) { const textPr = range.GetTextPr(); if (!textPr || typeof textPr.SetLanguage !== 'function') return; textPr.SetLanguage(wanted.language); }
  if ('color' in wanted) range.SetColor(Api.HexColor(wanted.color.replace('#', '')));
  if ('highlight' in wanted) range.SetHighlight(wanted.highlight);
  if ('vertAlign' in wanted) range.SetVertAlign(wanted.vertAlign);
  const paraPr = paragraph.GetParaPr();
  if (!paraPr) return;
  if ('alignment' in wanted) paraPr.SetJc(wanted.alignment === 'justify' ? 'both' : wanted.alignment);
  if ('spaceBefore' in wanted) paraPr.SetSpacingBefore(wanted.spaceBefore * 20, false);
  if ('spaceAfter' in wanted) paraPr.SetSpacingAfter(wanted.spaceAfter * 20, false);
  if ('lineSpacing' in wanted) paraPr.SetSpacingLine(wanted.lineSpacing * 240, 'auto');
  if ('indentLeft' in wanted) paraPr.SetIndLeft(wanted.indentLeft * 20);
  if ('indentRight' in wanted) paraPr.SetIndRight(wanted.indentRight * 20);
  if ('indentFirstLine' in wanted) paraPr.SetIndFirstLine(wanted.indentFirstLine * 20);
  if ('shading' in wanted) {
    const hex = wanted.shading;
    paraPr.SetShd('clear', parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16));
  }
  if ('tabs' in wanted) {
    const t = wanted.tabs;
    paraPr.SetTabs(t.map(x => Math.round(x.pos * 20)), t.map(x => x.align));
  }
  if ('borderTop' in wanted) {
    const b = wanted.borderTop;
    paraPr.SetTopBorder(b.style, Math.round(b.size * 8), 0, parseInt(b.color.slice(0, 2), 16), parseInt(b.color.slice(2, 4), 16), parseInt(b.color.slice(4, 6), 16));
  }
  if ('borderBottom' in wanted) {
    const b = wanted.borderBottom;
    paraPr.SetBottomBorder(b.style, Math.round(b.size * 8), 0, parseInt(b.color.slice(0, 2), 16), parseInt(b.color.slice(2, 4), 16), parseInt(b.color.slice(4, 6), 16));
  }
  if ('borderLeft' in wanted) {
    const b = wanted.borderLeft;
    paraPr.SetLeftBorder(b.style, Math.round(b.size * 8), 0, parseInt(b.color.slice(0, 2), 16), parseInt(b.color.slice(2, 4), 16), parseInt(b.color.slice(4, 6), 16));
  }
  if ('borderRight' in wanted) {
    const b = wanted.borderRight;
    paraPr.SetRightBorder(b.style, Math.round(b.size * 8), 0, parseInt(b.color.slice(0, 2), 16), parseInt(b.color.slice(2, 4), 16), parseInt(b.color.slice(4, 6), 16));
  }
  if ('keepNext' in wanted) paraPr.SetKeepNext(wanted.keepNext);
  if ('keepLines' in wanted) paraPr.SetKeepLines(wanted.keepLines);
  if ('pageBreakBefore' in wanted) paraPr.SetPageBreakBefore(wanted.pageBreakBefore);
  if ('widowControl' in wanted) paraPr.SetWidowControl(wanted.widowControl);
  if (requestedStyle) paragraph.SetStyle(requestedStyle);
} catch (_) {}
`;
}

/**
 * The substring write: the parent by its exact text (one hit), then the literal matchText exactly
 * once INSIDE that paragraph (ApiParagraph.Search, measured live: paragraph-scoped counts 2 vs 1).
 * Text properties on the range, paragraph properties on the paragraph -- the source's split.
 */
function rangeCommandBody(parentText, matchText, wanted) {
  return `
try {
  const parentText = ${JSON.stringify(parentText)};
  const matchText = ${JSON.stringify(matchText)};
  const wanted = ${JSON.stringify(wanted)};
  const textOpts = { Numbering: false };
  const doc = Api.GetDocument();
  if (!doc || !parentText || !matchText) return;
  const parents = (doc.Search(parentText, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || '') === parentText);
  if (parents.length !== 1) return;
  const paragraph = parents[0].GetParagraph(0);
  if (!paragraph || typeof paragraph.Search !== 'function') return;
  const hits = (paragraph.Search(matchText, true) || []).filter(r => String(r.GetText(textOpts) || '') === matchText);
  if (hits.length !== 1) return;
  const range = hits[0];
  let requestedStyle = null;
  // W-10-9: the SAME GetName()===name guard as commandBody above -- see its comment.
  if (Object.prototype.hasOwnProperty.call(wanted, 'style')) {
    const rawStyle = doc.GetStyle(wanted.style);
    requestedStyle = rawStyle && typeof rawStyle.GetName === 'function' && rawStyle.GetName() === wanted.style ? rawStyle : null;
    if (!requestedStyle) return;
  }
  if ('bold' in wanted) range.SetBold(wanted.bold);
  if ('italic' in wanted) range.SetItalic(wanted.italic);
  if ('underline' in wanted) range.SetUnderline(wanted.underline);
  if ('strike' in wanted) range.SetStrikeout(wanted.strike);
  if ('fontFamily' in wanted) range.SetFontFamily(wanted.fontFamily);
  if ('fontSize' in wanted) range.SetFontSize(wanted.fontSize);
  if ('language' in wanted) { const textPr = range.GetTextPr(); if (!textPr || typeof textPr.SetLanguage !== 'function') return; textPr.SetLanguage(wanted.language); }
  if ('color' in wanted) range.SetColor(Api.HexColor(wanted.color.replace('#', '')));
  if ('highlight' in wanted) range.SetHighlight(wanted.highlight);
  if ('vertAlign' in wanted) range.SetVertAlign(wanted.vertAlign);
  const paraPr = paragraph.GetParaPr();
  if (!paraPr) return;
  if ('alignment' in wanted) paraPr.SetJc(wanted.alignment === 'justify' ? 'both' : wanted.alignment);
  if ('spaceBefore' in wanted) paraPr.SetSpacingBefore(wanted.spaceBefore * 20, false);
  if ('spaceAfter' in wanted) paraPr.SetSpacingAfter(wanted.spaceAfter * 20, false);
  if ('lineSpacing' in wanted) paraPr.SetSpacingLine(wanted.lineSpacing * 240, 'auto');
  if ('indentLeft' in wanted) paraPr.SetIndLeft(wanted.indentLeft * 20);
  if ('indentRight' in wanted) paraPr.SetIndRight(wanted.indentRight * 20);
  if ('indentFirstLine' in wanted) paraPr.SetIndFirstLine(wanted.indentFirstLine * 20);
  if ('shading' in wanted) {
    const hex = wanted.shading;
    paraPr.SetShd('clear', parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16));
  }
  if ('tabs' in wanted) {
    const t = wanted.tabs;
    paraPr.SetTabs(t.map(x => Math.round(x.pos * 20)), t.map(x => x.align));
  }
  if ('borderTop' in wanted) {
    const b = wanted.borderTop;
    paraPr.SetTopBorder(b.style, Math.round(b.size * 8), 0, parseInt(b.color.slice(0, 2), 16), parseInt(b.color.slice(2, 4), 16), parseInt(b.color.slice(4, 6), 16));
  }
  if ('borderBottom' in wanted) {
    const b = wanted.borderBottom;
    paraPr.SetBottomBorder(b.style, Math.round(b.size * 8), 0, parseInt(b.color.slice(0, 2), 16), parseInt(b.color.slice(2, 4), 16), parseInt(b.color.slice(4, 6), 16));
  }
  if ('borderLeft' in wanted) {
    const b = wanted.borderLeft;
    paraPr.SetLeftBorder(b.style, Math.round(b.size * 8), 0, parseInt(b.color.slice(0, 2), 16), parseInt(b.color.slice(2, 4), 16), parseInt(b.color.slice(4, 6), 16));
  }
  if ('borderRight' in wanted) {
    const b = wanted.borderRight;
    paraPr.SetRightBorder(b.style, Math.round(b.size * 8), 0, parseInt(b.color.slice(0, 2), 16), parseInt(b.color.slice(2, 4), 16), parseInt(b.color.slice(4, 6), 16));
  }
  if ('keepNext' in wanted) paraPr.SetKeepNext(wanted.keepNext);
  if ('keepLines' in wanted) paraPr.SetKeepLines(wanted.keepLines);
  if ('pageBreakBefore' in wanted) paraPr.SetPageBreakBefore(wanted.pageBreakBefore);
  if ('widowControl' in wanted) paraPr.SetWidowControl(wanted.widowControl);
  if (requestedStyle) paragraph.SetStyle(requestedStyle);
} catch (_) {}
`;
}

function targetTextOf(before, op) {
  const t = before ? resolveFormattable(before.primary, effectiveSelector(op)).target : null;
  return t ? t.text : null;
}

/** L1: the target, found again by its text in the FRESH after-state, carries the wanted formatting. */
function verifyPrimitive(op, before, after) {
  const text = targetTextOf(before, op);
  if (text === null) return { ok: false, evidence: { reason: 'target not resolvable in the before-state' } };
  const now = resolveFormattable(after.primary, { text });
  if (!now.ok) return { ok: false, evidence: { text, reason: `target-after-${String(now.ambiguity).toLowerCase()}` } };
  if (isRange(op)) {
    const t = rangeTerminal(after.primary, now.target, op);
    return { ok: t.state === 'HOLDS', evidence: { ...t.evidence, format: op.format, level: 'RANGE' } };
  }
  const m = compare(now.target.format, op.format);
  return { ok: m.ok, evidence: { text, format: op.format, observed: now.target.format, missing: m.missing, unobservable: m.unobservable } };
}

const paragraphKey = b => (FORMATTABLE_KINDS.includes(b.kind) ? `${b.kind}|${b.level || ''}|${norm(b.text)}` : null);
const tableKey = b => JSON.stringify((b.rows || []).map(r => r.map(c => norm(c.text))));
const imageKey = i => JSON.stringify({ src: i.src || null, alt: i.alt || null, title: i.title || null });

function containsMultiset(haystack, needles) {
  const h = new Map();
  for (const v of haystack) h.set(v, (h.get(v) || 0) + 1);
  const missing = [];
  const need = new Map();
  for (const v of needles) need.set(v, (need.get(v) || 0) + 1);
  for (const [k, n] of need) if ((h.get(k) || 0) < n) missing.push({ key: k, expected: n, actual: h.get(k) || 0 });
  return { ok: missing.length === 0, missing };
}

/**
 * L3, PRIMITIVE_INTEGRAL (source docx-session-formatting.cjs:103-108 + verifyLiveRoundTrip):
 * the target text is still there exactly once, every OTHER paragraph/heading of the fresh
 * pre-mutation state survives (kind|level|text multiset), table shapes and images survive.
 */
/**
 * W-09 (design doc point 3): "every OTHER paragraph's ParaPr unchanged" for the
 * NEW fields -- the pre-existing whole-paragraph L3 above only ever checked OTHER paragraphs'
 * TEXT/EXISTENCE survival (`paragraphKey`/`containsMultiset`), never their FORMAT. This is NEW
 * coverage, scoped to exactly `NEW_PARA_LEVEL_KEYS` (tabs / the four border sides / the four
 * pagination flags) -- not a retroactive fix of the pre-existing fields, which this unit was not
 * asked to touch. A paragraph matched by text more than once (ambiguous) is skipped here, same
 * as `rangePreservation`'s own "text survival is the multiset check elsewhere" rule.
 */
function otherParagraphsNewFieldsUnchanged(target, before, after) {
  const changed = [];
  for (const blk of (before.primary.blocks || [])) {
    if (blk === target || !FORMATTABLE_KINDS.includes(blk.kind)) continue;
    const hits = (after.primary.blocks || []).filter(x => FORMATTABLE_KINDS.includes(x.kind) && norm(x.text) === norm(blk.text));
    if (hits.length !== 1) continue;
    for (const k of L3_PARA_LEVEL_KEYS) {
      const bv = blk.format ? blk.format[k] : undefined;
      const av = hits[0].format ? hits[0].format[k] : undefined;
      if (JSON.stringify(bv) !== JSON.stringify(av)) changed.push({ text: norm(blk.text).slice(0, 60), property: k, before: bv, after: av });
    }
  }
  return { ok: changed.length === 0, changed: changed.slice(0, 20) };
}

function verifyPreservation(op, before, after) {
  const target = before ? resolveFormattable(before.primary, effectiveSelector(op)).target : null;
  if (!target) return { ok: false, evidence: { reason: 'no before-target' } };
  const afterBlocks = after.primary.blocks || [];
  const targetCount = afterBlocks.filter(b => FORMATTABLE_KINDS.includes(b.kind) && norm(b.text) === norm(target.text)).length;
  const preserve = (before.primary.blocks || []).map(paragraphKey).filter(Boolean);
  const own = preserve.indexOf(paragraphKey(target));
  if (own >= 0) preserve.splice(own, 1);
  const text = containsMultiset(afterBlocks.map(paragraphKey).filter(Boolean), preserve);
  const tables = containsMultiset(afterBlocks.filter(b => b.kind === 'table').map(tableKey), (before.primary.blocks || []).filter(b => b.kind === 'table').map(tableKey));
  const images = containsMultiset((after.primary.images || []).map(imageKey), (before.primary.images || []).map(imageKey));
  const range = isRange(op) ? rangePreservation(op, target, before, after) : null;
  const otherNewFields = otherParagraphsNewFieldsUnchanged(target, before, after);
  const ok = targetCount === 1 && text.ok && tables.ok && images.ok && (!range || range.ok) && otherNewFields.ok;
  return { ok, evidence: { targetText: target.text, targetCount, untouchedText: text, tables, images, otherNewFields, ...(range ? { range } : {}), expectationSource: 'derived-from-fresh-pre-mutation-observation' } };
}

/**
 * The substring part of L3, per character from the runs of the fresh before- and after-states:
 * outside the range all seven text properties unchanged; inside it every property the operation
 * did NOT request unchanged; every other covered paragraph's run formatting unchanged.
 */
function rangePreservation(op, target, before, after) {
  const now = resolveFormattable(after.primary, { text: target.text });
  const b = rangeState(target, op.matchText);
  const a = now.ok ? rangeState(now.target, op.matchText) : { ok: false, reason: 'target-after-unresolved' };
  if (!b.ok || !a.ok) return { ok: false, reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` };
  if (b.start !== a.start || b.chars.length !== a.chars.length) return { ok: false, reason: 'range-position-changed' };
  const requested = Object.keys(op.format).filter(k => TEXT_KEYS.includes(k));
  const changed = [];
  for (let i = 0; i < b.chars.length; i += 1) {
    const inside = i >= b.start && i < b.end;
    for (const k of TEXT_KEYS) {
      if (inside && requested.includes(k)) continue;
      if (JSON.stringify(b.chars[i][k]) !== JSON.stringify(a.chars[i][k])) changed.push({ at: i, inside, property: k, before: b.chars[i][k], after: a.chars[i][k] });
    }
  }
  const others = [];
  for (const blk of (before.primary.blocks || [])) {
    if (blk === target || !FORMATTABLE_KINDS.includes(blk.kind) || !blk.runs) continue;
    const hits = (after.primary.blocks || []).filter(x => FORMATTABLE_KINDS.includes(x.kind) && norm(x.text) === norm(blk.text));
    if (hits.length !== 1) continue; // text survival is the multiset check above
    const cb = charFormats(blk);
    const ca = charFormats(hits[0]);
    if (!ca || JSON.stringify(cb) !== JSON.stringify(ca)) others.push(norm(blk.text).slice(0, 60));
  }
  return { ok: changed.length === 0 && others.length === 0, changedInTarget: changed.slice(0, 20), changedCount: changed.length, otherParagraphsChanged: others };
}

/** L2: in the final state the target is found once by its text and carries the wanted formatting. */
function wholeTaskChecks(op, finalObservation) {
  const sel = effectiveSelector(op);
  let text = sel && typeof sel.text === 'string' ? sel.text : null;
  // M5 sectionHeading + paragraphOrdinal: a formatting write keeps the paragraph count and text, so
  // the ordinal names the same paragraph in the final state -- its text is read from there
  if (text === null && sel && sel.sectionHeading !== undefined) {
    const r = resolveStableTarget(finalObservation.primary, sel);
    text = r.ok ? r.target.text : null;
  }
  if (text === null) return [{ name: `${op.id}: format target text known`, ok: false }];
  // W2.1: a later list step of this task declaredly turned the formatted paragraph into a list item
  // (planner metadata). The formatting postcondition is checked on that TEXT UNIT -- exactly one,
  // with its own format/run read -- not on a stale block kind.
  if (Array.isArray(op.successorTransforms) && op.successorTransforms.length) {
    const u = resolveTextUnit(finalObservation.primary, text);
    const node = u.ok && ['paragraph', 'list_item'].includes(u.unit.kind) ? u.unit.node : null;
    const holds = !!node && successorHolds(finalObservation.primary, node, op);
    return [
      { name: `${op.id}: format target present once (as a text unit; declared successor ${op.successorTransforms.join(', ')})`, ok: !!node },
      { name: `${op.id}: final ${isRange(op) ? 'RANGE ' : ''}formatting holds`, ok: holds },
    ];
  }
  const r = resolveFormattable(finalObservation.primary, { text });
  if (isRange(op)) {
    const t = r.ok ? rangeTerminal(finalObservation.primary, r.target, op) : { state: 'FAIL_CLOSED' };
    return [
      { name: `${op.id}: format parent present once`, ok: r.ok },
      { name: `${op.id}: final RANGE formatting holds`, ok: r.ok && t.state === 'HOLDS' },
    ];
  }
  const m = r.ok ? compare(r.target.format, op.format) : { ok: false };
  return [
    { name: `${op.id}: format target present once`, ok: r.ok },
    { name: `${op.id}: final formatting holds`, ok: r.ok && m.ok },
  ];
}

/** The semantic text this operation leaves behind (for a later operation's ref). */
function producedText(op) {
  const sel = effectiveSelector(op);
  return sel && typeof sel.text === 'string' ? sel.text : null;
}

module.exports = {
  plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation,
  wholeTaskChecks, producedText, normalizeWanted, compare, blockText, rangeState, rangeCommandBody,
  styleNamesOf,
};
