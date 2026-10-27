'use strict';

const { textUnits, norm, tablesContent } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');

/**
 * DOCX `define_style`/`modify_style` family (post-M7, W-10 -- a NEW capability, not a
 * source-migrated intent; registered via src/adapters/docx/capabilities.cjs's NEW_INTENTS,
 * disjoint from the 28/28 migration inventory, the `page-setup`/`numbering` precedent). See
 * docs/W10-STYLE-DEFINITION-DESIGN-20261007.md (v2) for the full design and the measured facts
 * this file is built on.
 *
 * TWO intents, ONE family module (page-setup's shape, section-global target by `name` instead of
 * section index) -- `define_style` and `modify_style` share the SAME observation/plan/write
 * machinery but have OPPOSITE fail-closed directions on the style's existence (design doc point 1,
 * "ket kulon intent, nem egy mode-mezos", the reason spelled out there is not repeated here).
 *
 * Measured live (W10-1a..1c, W10-2;, all accepted):
 *   - `doc.CreateStyle(name, 'paragraph')` (arity 2) creates a paragraph style; `doc.GetStyle(name)`
 *     is the session-side identity -- NEVER the saved file's `w:styleId` (an internal, auto-
 *     generated token, measured `'1_1'`, that does NOT match `name`). `basedOn` as a 3rd argument
 *     stores nothing (GH #292 S3, LIVE) -- the base is set with ApiStyle.SetBasedOn after the create (see STEP_KEYS).
 *   - `text.size` is passed to `SetFontSize` in HALF-POINTS, but the public field is PT (the
 *     `format_run` convention) -- `lower()` multiplies by 2 (W10-1c: `SetFontSize(31)` -> saved
 *     `<w:sz w:val="31"/>`, i.e. the Builder argument and the OOXML half-point unit are identical).
 *   - `text.color`: ONLY `Api.RGB(r,g,b)` (or `Api.HexColor(hex)`) produce a CORRECT saved
 *     `w:color` -- `Api.CreateRGBColor(r,g,b)` measured WRONG (`000000` regardless of input, the
 *     SAME trap as the table family's `SetShd`, W04-1c). v1 uses `Api.RGB` exclusively.
 *     `'auto'`: GH #292 S8 (LIVE) -- SetColor(Api.AutoColor()) persists (saved w:color auto, ToJSON color.auto); before that it was refused (FAIL_CLOSED,
 *     named); only a 6-hex-digit string is accepted.
 *   - `paragraph.indentLeft`/`spacingAfter` are twips at the Builder boundary, externalized as PT
 *     (the SAME /20 convention `docx-driver.cjs`'s paragraph-format read already uses for the
 *     run-level `indentLeft`/`spaceAfter` fields -- `lower()` multiplies by 20 the other way).
 *   - `SetSpacingAfter(n, isAuto)`: v1 ALWAYS passes `false` for the 2nd argument (W10-2, the review's
 *     exact words: "a termek MINDIG false-t adjon at, es az L1 az afterAutoSpacing
 *     hianyat/0-jat is nezze") -- auto-spacing can make Word ignore the explicit value.
 *   - `paragraph.align` -> `SetJc('left'|'center'|'right'|'both')` (justify is written 'both': GH #292 S6 LIVE,
 *     docs/evidence/gh292-style-definition-probe-20261021.json -- SetJc('justify') on a style stores nothing). The
 *     style read normalizes the ToJSON names 'start' / 'end' / 'both' back to left / right / justify (the saved w:jc
 *     is left / right / both). Before this, left / right / justify were accepted but never provable
 *     (docs/evidence/gh292-style-align-installed-release14-BUG-20261021.json).
 *
 * NAMED, UNMEASURED ASSUMPTION carried from the driver change (docx-driver.cjs's own comment,
 * repeated here because it is THIS family's L1/L3 foundation): the style's `GetTextPr()`/
 * `GetParaPr()` getter SURFACE (`GetBold`/`GetFontSize`/`GetColor`/`GetJc`/`GetIndLeft`/
 * `GetSpacingAfter`) was never independently LIVE-measured -- only the matching SETTERS were
 * (write-only probes). A wrong assumption degrades to `'UNOBSERVABLE:...'` (never a guessed
 * value), so classify()/verifyPrimitive() below NEVER grant a false SATISFIED from it.
 */

// GH #292 S1 / S2 (LIVE, docs/evidence/gh292-style-definition-probe-20261021.json): italic / underline / strike / fontFamily and
// spacingBefore / indentRight / indentFirstLine / keepNext / keepLines / pageBreakBefore persist on a style; the booleans only as
// true (false was not measured on a style -- a named refusal, not field-unsupported).
const TEXT_FIELDS = Object.freeze(['bold', 'size', 'color', 'italic', 'underline', 'strike', 'fontFamily']);
const PARAGRAPH_FIELDS = Object.freeze(['align', 'indentLeft', 'spacingAfter', 'spacingBefore', 'indentRight', 'indentFirstLine', 'keepNext', 'keepLines', 'pageBreakBefore']);
// GH #292 S8 (LIVE, docs/evidence/gh292s8-style-unset-probe-20261022.json): false persists too -- an EXPLICIT off (saved w:i w:val="0",
// w:u w:val="none", w:keepNext w:val="false"; ToJSON false), on a fresh style and after a true. It is NOT a removal: the style's
// TextPr / ParaPr has no Clear / Remove / Reset method, so a field cannot go back to 'inherited' through the public API.
const FLAGS = Object.freeze({ text: ['italic', 'underline', 'strike'], paragraph: ['keepNext', 'keepLines', 'pageBreakBefore'] });
const flagsOf = (group, obj, index, fields) => { for (const k of FLAGS[group]) if (obj[k] !== undefined) { if (typeof obj[k] !== 'boolean') return { ok: false, detail: `step ${index}: ${group}.${k} must be a boolean` }; fields[k] = obj[k]; } return null; };
const ALIGNMENTS = Object.freeze(['left', 'right', 'center', 'justify']);
// Shared by both intents. GH #292 S3 (LIVE, docs/evidence/gh292-style-definition-probe-20261021.json): `basedOn` on define_style is the
// NAME of an existing paragraph style -- ApiStyle.SetBasedOn(doc.GetStyle(base)) persists (saved w:basedOn, the reopened ToJSON basedOn =
// the base's styleId); CreateStyle's 3rd argument stores nothing. On modify_style it stays a named refusal (re-basing an existing style
// changes every style and paragraph built on it -- not measured).
// GH #292 S5 (LIVE, docs/evidence/gh292s5-character-style-probe-20261022.json): define_style `type: 'character'` creates a character
// style -- CreateStyle(name, 'run') (saved w:type="character"; 'character' as the argument makes a PARAGRAPH style). Every text setter
// persists on it; its ParaPr also takes SetJc (stored, meaningless on a character style) -- so the paragraph group is refused on one.
// SetBasedOn only takes a base of the SAME type (a paragraph base returns false and stores nothing). ToJSON styleType names the type.
const STYLE_TYPES = Object.freeze({ paragraph: { create: 'paragraph', getType: 'paragraph', styleType: 'paragraphStyle' }, character: { create: 'run', getType: 'run', styleType: 'characterStyle' } });
const STEP_KEYS = Object.freeze(['intent', 'id', 'dependsOn', 'name', 'type', 'basedOn', 'next', 'link', 'text', 'paragraph']);
// GH #292 S4 (owner GO 2026-10-22; LIVE docs/evidence/gh292s4-next-link-internal-probe-20261022.json + gh292s4b-...-run2-20261022.json):
// modify_style `next` / `link` = the NAME of another style. No public setter exists -- the write is ONE word /4 Bridge WRITE
// docx.style.nextLink.write/v1 per field (the editor's Set_Next / Set_Link; the handler resolves both names, checks the types and sets
// BOTH sides of a link). next: paragraph -> paragraph; link: a paragraph and a character style; a style already linked to a third style
// is refused, never re-pointed. Clearing either is not offered (not measured).
const NEXT_LINK = Object.freeze(['next', 'link']);

const isPosNum = v => typeof v === 'number' && Number.isFinite(v) && v > 0;
const isNonNegNum = v => typeof v === 'number' && Number.isFinite(v) && v >= 0;
const isUnobservable = v => typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0;

function validateText(text, index) {
  if (!text || typeof text !== 'object') return { ok: false, detail: `step ${index}: text requires an object` };
  for (const k of Object.keys(text)) if (!TEXT_FIELDS.includes(k)) return { ok: false, detail: `field-unsupported: step ${index} (style).text.${k}` };
  if (Object.keys(text).length === 0) return { ok: false, detail: `step ${index}: text requires at least one of ${TEXT_FIELDS.join(', ')}` };
  const fields = {};
  if (text.bold !== undefined) {
    if (typeof text.bold !== 'boolean') return { ok: false, detail: `step ${index}: text.bold must be a boolean` };
    fields.bold = text.bold;
  }
  if (text.size !== undefined) {
    if (!isPosNum(text.size)) return { ok: false, detail: `step ${index}: text.size must be a positive number (points)` };
    fields.size = text.size;
  }
  if (text.color !== undefined) {
    // GH #292 S8: 'auto' -> SetColor(Api.AutoColor()) (LIVE); the driver reads color.auto back as 'auto'
    if (text.color !== 'auto' && (typeof text.color !== 'string' || !/^[0-9A-Fa-f]{6}$/.test(text.color))) return { ok: false, detail: `step ${index}: text.color must be a 6-hex-digit string or 'auto'` };
    fields.color = text.color === 'auto' ? 'auto' : text.color.toUpperCase();
  }
  const t = flagsOf('text', text, index, fields); if (t) return t;
  if (text.fontFamily !== undefined) {
    if (typeof text.fontFamily !== 'string' || !text.fontFamily.trim() || text.fontFamily.length > 64) return { ok: false, detail: `step ${index}: text.fontFamily must be a non-empty font name (at most 64 characters)` };
    fields.fontFamily = text.fontFamily;
  }
  return { ok: true, fields };
}

function validateParagraph(paragraph, index) {
  if (!paragraph || typeof paragraph !== 'object') return { ok: false, detail: `step ${index}: paragraph requires an object` };
  for (const k of Object.keys(paragraph)) if (!PARAGRAPH_FIELDS.includes(k)) return { ok: false, detail: `field-unsupported: step ${index} (style).paragraph.${k}` };
  if (Object.keys(paragraph).length === 0) return { ok: false, detail: `step ${index}: paragraph requires at least one of ${PARAGRAPH_FIELDS.join(', ')}` };
  const fields = {};
  if (paragraph.align !== undefined) {
    if (!ALIGNMENTS.includes(paragraph.align)) return { ok: false, detail: `step ${index}: paragraph.align must be one of ${ALIGNMENTS.join(', ')}` };
    fields.align = paragraph.align;
  }
  if (paragraph.indentLeft !== undefined) {
    if (!isNonNegNum(paragraph.indentLeft)) return { ok: false, detail: `step ${index}: paragraph.indentLeft must be a non-negative number (points)` };
    fields.indentLeft = paragraph.indentLeft;
  }
  if (paragraph.spacingAfter !== undefined) {
    if (!isNonNegNum(paragraph.spacingAfter)) return { ok: false, detail: `step ${index}: paragraph.spacingAfter must be a non-negative number (points)` };
    fields.spacingAfter = paragraph.spacingAfter;
  }
  for (const k of ['spacingBefore', 'indentRight']) {
    if (paragraph[k] === undefined) continue;
    if (!isNonNegNum(paragraph[k])) return { ok: false, detail: `step ${index}: paragraph.${k} must be a non-negative number (points)` };
    fields[k] = paragraph[k];
  }
  if (paragraph.indentFirstLine !== undefined) {
    if (typeof paragraph.indentFirstLine !== 'number' || !Number.isFinite(paragraph.indentFirstLine)) return { ok: false, detail: `step ${index}: paragraph.indentFirstLine must be a number (points; negative = a hanging indent)` };
    fields.indentFirstLine = paragraph.indentFirstLine;
  }
  const t = flagsOf('paragraph', paragraph, index, fields); if (t) return t;
  return { ok: true, fields };
}

function plan(step, index) {
  const action = step.intent === 'define_style' ? 'define' : (step.intent === 'modify_style' ? 'modify' : null);
  if (!action) return { ok: false, detail: `step ${index}: ${step.intent} is not a style-definition intent` };
  for (const key of Object.keys(step)) {
    if (!STEP_KEYS.includes(key)) return { ok: false, detail: `field-unsupported: step ${index} (${step.intent}).${key}` };
  }
  if (typeof step.name !== 'string' || step.name.trim() === '') {
    return { ok: false, detail: `step ${index}: ${step.intent} requires a non-empty 'name'` };
  }
  if (step.type !== undefined) {
    if (action === 'modify') return { ok: false, detail: `field-unsupported: step ${index} (modify_style).type` };
    if (!Object.prototype.hasOwnProperty.call(STYLE_TYPES, step.type)) return { ok: false, detail: `step ${index}: define_style.type must be one of ${Object.keys(STYLE_TYPES).join(', ')}` };
    if (step.type === 'character' && step.paragraph !== undefined) return { ok: false, detail: `step ${index}: define_style.paragraph is not allowed on a character style` };
  }
  if (step.basedOn !== undefined) {
    if (action === 'modify') return { ok: false, detail: `step ${index}: modify_style.basedOn is not yet measured -- refused` };
    if (typeof step.basedOn !== 'string' || step.basedOn.trim() === '') return { ok: false, detail: `step ${index}: define_style.basedOn must be a non-empty style name` };
    if (step.basedOn === step.name) return { ok: false, detail: `step ${index}: define_style.basedOn names the style itself` };
  }
  for (const f of NEXT_LINK) {
    if (step[f] === undefined) continue;
    if (action === 'define') return { ok: false, detail: `step ${index}: define_style.${f} is not offered -- define the style, then modify_style.${f}` };
    if (typeof step[f] !== 'string' || step[f].trim() === '') return { ok: false, detail: `step ${index}: modify_style.${f} must be a non-empty style name` };
    if (f === 'link' && step[f] === step.name) return { ok: false, detail: `step ${index}: modify_style.link names the style itself` };
  }
  const present = ['text', 'paragraph', ...(action === 'define' ? ['basedOn'] : NEXT_LINK)].filter(f => step[f] !== undefined);
  if (present.length === 0) return { ok: false, detail: `step ${index}: ${step.intent} requires at least one of ${action === 'define' ? 'text, paragraph, basedOn' : 'text, paragraph, next, link'}` };
  const fields = {};
  if (action === 'define') fields.type = step.type || 'paragraph';
  if (step.basedOn !== undefined) fields.basedOn = step.basedOn;
  for (const f of NEXT_LINK) if (step[f] !== undefined) fields[f] = step[f];
  if (step.text !== undefined) {
    const v = validateText(step.text, index); if (!v.ok) return v;
    fields.text = v.fields;
  }
  if (step.paragraph !== undefined) {
    const v = validateParagraph(step.paragraph, index); if (!v.ok) return v;
    fields.paragraph = v.fields;
  }
  const selector = Object.freeze({ name: step.name });
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: step.intent, action, selector, name: step.name,
      fields: Object.freeze(fields), preservationMode: 'PRIMITIVE_INTEGRAL',
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

// ---- the observation ------------------------------------------------------------------------

function styleAt(primary, name) {
  const d = primary && primary.document;
  const styles = d ? d.styles : null;
  if (!styles || typeof styles !== 'object') return { ok: false, reason: 'styles-unreadable' };
  return { ok: true, styles, st: Object.prototype.hasOwnProperty.call(styles, name) ? styles[name] : null };
}

/** Which of the requested fields already match, reading each via the driver's own getter-sourced
 * value -- never the write args. UNOBSERVABLE never counts as a match (same rule as page-setup). */
function fieldMatches(st, group, key, want) {
  if (!st || !st[group]) return false;
  const have = st[group][key];
  if (isUnobservable(have)) return false;
  return have === want;
}

function groupMatches(st, group, wantFields) {
  return Object.keys(wantFields).every(k => fieldMatches(st, group, k, wantFields[k]));
}

/** GH #292 S3: the style's base, by NAME (the driver's identity.basedOnName); an unresolved base never matches. */
function basedOnMatches(st, want) {
  const have = st && st.identity ? st.identity.basedOnName : undefined;
  return typeof have === 'string' && !isUnobservable(have) && have === want;
}

/** GH #292 S5: the style's type, from the driver's identity.styleType (ToJSON styleType, measured 'paragraphStyle' / 'characterStyle'). */
function typeMatches(st, want) {
  return !!st && !!st.identity && st.identity.styleType === STYLE_TYPES[want].styleType;
}

/** GH #292 S4: the style's next / linked style, by NAME (identity.nextName / linkName); an unresolved one never matches. */
function nameFieldMatches(st, key, want) {
  const have = st && st.identity ? st.identity[key] : undefined;
  return typeof have === 'string' && !isUnobservable(have) && have === want;
}

function allRequestedMatch(st, fields) {
  if (fields.type !== undefined && !typeMatches(st, fields.type)) return false;
  if (fields.basedOn !== undefined && !basedOnMatches(st, fields.basedOn)) return false;
  if (fields.next !== undefined && !nameFieldMatches(st, 'nextName', fields.next)) return false;
  if (fields.link !== undefined && !nameFieldMatches(st, 'linkName', fields.link)) return false;
  if (fields.text && !groupMatches(st, 'text', fields.text)) return false;
  if (fields.paragraph && !groupMatches(st, 'paragraph', fields.paragraph)) return false;
  return true;
}

/** `group.key` of every requested field that does NOT read back as requested (a different value OR an unobservable one) -- the named part of the define_style name-exists refusal. */
function differingRequestedFields(st, fields) {
  const out = [];
  if (fields.type !== undefined && !typeMatches(st, fields.type)) out.push('type');
  if (fields.basedOn !== undefined && !basedOnMatches(st, fields.basedOn)) out.push('basedOn');
  for (const f of NEXT_LINK) if (fields[f] !== undefined && !nameFieldMatches(st, `${f}Name`, fields[f])) out.push(f);
  for (const group of ['text', 'paragraph']) if (fields[group]) for (const k of Object.keys(fields[group])) if (!fieldMatches(st, group, k, fields[group][k])) out.push(`${group}.${k}`);
  return out;
}

function unobservableRequestedFields(st, fields) {
  const out = [];
  const checkGroup = (group, keys) => { for (const k of Object.keys(keys)) if (!st || !st[group] || isUnobservable(st[group][k])) out.push(`${group}.${k}`); };
  if (fields.basedOn !== undefined && (!st || !st.identity || typeof st.identity.basedOnName !== 'string' || isUnobservable(st.identity.basedOnName))) out.push('basedOn');
  for (const f of NEXT_LINK) if (fields[f] !== undefined && (!st || !st.identity || isUnobservable(st.identity[`${f}Name`]) || st.identity[`${f}Name`] === undefined)) out.push(f);
  if (fields.text) checkGroup('text', fields.text);
  if (fields.paragraph) checkGroup('paragraph', fields.paragraph);
  return out;
}

// ---- classify / resolve ----------------------------------------------------------------------

function classify(primary, op) {
  const s = styleAt(primary, op.name);
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason } };
  if (op.action === 'define') {
    // GH #264 (supersedes the earlier "ALWAYS FAIL_CLOSED on a taken name" rule): a retry must be distinguishable from a name taken by someone else, like every other intent's
    // retry. The name exists AND every requested field reads back as requested (a getter-sourced read, UNOBSERVABLE never counts as a match) -> SATISFIED, nothing written. The name exists but
    // the style differs (or a requested field is unobservable) -> FAIL_CLOSED with a NAMED reason that lists the differing fields. The name is free -> APPLY, as before.
    if (s.st !== null) {
      if (allRequestedMatch(s.st, op.fields)) {
        if (op.recovery === UNRESOLVED) return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } };
        return { state: 'SATISFIED', evidence: { name: op.name, fields: op.fields, proof: 'define_style-name-exists-and-matches-every-requested-field' } };
      }
      return { state: 'FAIL_CLOSED', evidence: { reason: `define_style-name-already-exists-with-different-fields:${op.name}:${differingRequestedFields(s.st, op.fields).join(',')}` } };
    }
    // GH #292 S3: the base must be a style this read sees by name (the targeted read covers it: styleNamesOf adds the base)
    if (op.fields.basedOn !== undefined && !(s.styles[op.fields.basedOn])) return { state: 'FAIL_CLOSED', evidence: { reason: `define_style-base-style-does-not-exist:${op.fields.basedOn}` } };
    // GH #292 S5: SetBasedOn takes only a base of the same type (LIVE: a paragraph base on a character style returns false, stores nothing)
    if (op.fields.basedOn !== undefined && !typeMatches(s.styles[op.fields.basedOn], op.fields.type)) return { state: 'FAIL_CLOSED', evidence: { reason: `define_style-base-style-type-mismatch:${op.fields.basedOn}:${op.fields.type}` } };
    if (op.recovery === UNRESOLVED) return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } };
    return { state: 'APPLY', evidence: { name: op.name, fields: op.fields } };
  }
  // modify_style
  if (s.st === null) return { state: 'FAIL_CLOSED', evidence: { reason: `modify_style-name-does-not-exist:${op.name}` } };
  // GH #292 S5: a character style's ParaPr takes values that mean nothing on it -- the paragraph group is refused there
  if (op.fields.paragraph && s.st.identity && s.st.identity.styleType === STYLE_TYPES.character.styleType) return { state: 'FAIL_CLOSED', evidence: { reason: `modify_style-paragraph-fields-on-a-character-style:${op.name}` } };
  const nl = nextLinkRefusal(s.styles, s.st, op);
  if (nl) return { state: 'FAIL_CLOSED', evidence: { reason: nl } };
  // a link is a PAIR: matched only when the target points back too
  if (allRequestedMatch(s.st, op.fields) && (op.fields.link === undefined || nameFieldMatches(s.styles[op.fields.link], 'linkName', op.name))) {
    if (op.recovery === UNRESOLVED) return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } };
    return { state: 'SATISFIED', evidence: { name: op.name, fields: op.fields, proof: 'fresh-live-style-matches-every-requested-field' } };
  }
  if ((op.fields.next !== undefined || op.fields.link !== undefined) && !runtimeVerified(primary, 'styleNextLinkBridge')) return { state: 'FAIL_CLOSED', evidence: { reason: 'docx-style-next-link-bridge-write-runtime-not-proven' } };
  return { state: 'APPLY', evidence: { name: op.name, fields: op.fields, unobservable: unobservableRequestedFields(s.st, op.fields) } };
}

const runtimeVerified = (primary, dependency) => !!(primary && primary.internalGate && primary.internalGate[dependency] === true);
const styleTypeOf = st => (st && st.identity ? st.identity.styleType : null);

/** GH #292 S4: a next / link the Bridge handler would refuse is refused here first, named (the handler stays the second line). */
function nextLinkRefusal(styles, st, op) {
  const P = STYLE_TYPES.paragraph.styleType, C = STYLE_TYPES.character.styleType;
  if (op.fields.next !== undefined) {
    const t = styles[op.fields.next];
    if (!t) return `modify_style-next-style-does-not-exist:${op.fields.next}`;
    if (styleTypeOf(st) !== P || styleTypeOf(t) !== P) return `modify_style-next-type-mismatch:${op.name}:${op.fields.next}`;
  }
  if (op.fields.link !== undefined) {
    const t = styles[op.fields.link];
    if (!t) return `modify_style-link-style-does-not-exist:${op.fields.link}`;
    const pair = [styleTypeOf(st), styleTypeOf(t)].sort().join(',');
    if (pair !== [P, C].sort().join(',')) return `modify_style-link-type-mismatch:${op.name}:${op.fields.link}`;
    const mine = st.identity ? st.identity.linkName : undefined; const theirs = t.identity ? t.identity.linkName : undefined;
    if (isUnobservable(mine) || mine === undefined || isUnobservable(theirs) || theirs === undefined) return `modify_style-link-unread:${op.name}:${op.fields.link}`;
    if ((mine !== null && mine !== op.fields.link) || (theirs !== null && theirs !== op.name)) return `modify_style-link-already-linked-elsewhere:${op.name}:${op.fields.link}`;
  }
  return null;
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  if (!op || (op.intent !== 'define_style' && op.intent !== 'modify_style')) return { ok: false, target: null, selector, ambiguity: 'INVALID_SELECTOR' };
  const s = styleAt(primary, op.name);
  if (!s.ok) return { ok: false, target: null, selector, ambiguity: s.reason.toUpperCase().replace(/-/g, '_') };
  // GH #264: a define on an EXISTING name is no longer an ambiguity here -- classify() decides SATISFIED (every requested field matches) vs FAIL_CLOSED (named, with the differing fields)
  if (op.action === 'modify' && s.st === null) return { ok: false, target: null, selector, ambiguity: 'MODIFY_STYLE_NAME_DOES_NOT_EXIST' };
  return { ok: true, target: { kind: 'style_definition', name: op.name, action: op.action, fields: op.fields, styleNextLinkBridge: runtimeVerified(primary, 'styleNextLinkBridge') }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  const same = target && target.kind === 'style_definition' && target.name === op.name && target.action === op.action;
  if (!same) return { ok: false, dispatched: false, detail: 'style-definition-target-unresolved' };
  if ((op.fields.next !== undefined || op.fields.link !== undefined) && target.styleNextLinkBridge !== true) return { ok: false, dispatched: false, detail: 'docx-style-next-link-bridge-write-runtime-not-proven' };
  return { ok: true };
}

// ---- lower ------------------------------------------------------------------------------------

function lower(op, target) {
  // `action` is the FAMILY-level dispatch discriminator (the same role 'pageSetup'/'list'/etc
  // play in every sibling family's meta) -- `mode` carries define/modify, which is this family's
  // OWN sub-choice, not the dispatcher's.
  const meta = Object.freeze({ family: 'styleDefinition', action: 'styleDefinition', mode: target.action, name: target.name, fields: target.fields,
    ...(target.fields.type ? { styleType: STYLE_TYPES[target.fields.type] } : {}),
    ...(target.fields.next !== undefined || target.fields.link !== undefined ? { bridgeStyleNextLink: true } : {}) });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand, public API only. `define` creates via `doc.CreateStyle(name, 'paragraph')`;
 * `modify` requires the style to already exist -- BOTH branches are defensive no-ops if the
 * precondition plan()/classify() already checked turns out false at dispatch time (never a
 * silent overwrite, never a silent create-on-modify). `size`/`indentLeft`/`spacingAfter` are
 * converted here (pt -> half-point / twips), the SAME direction `page-setup.cjs`'s `margins`
 * reorder comment describes for its own unit boundary.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  // (LIVE run4): doc.GetStyle(<name>) was MEASURED to NOT return null for an unknown
  // name -- it falls back to the Normal style object. A bare truthiness check on the result is
  // ALWAYS true, so (before this fix) 'define' never created a style: style was always truthy
  // (the Normal fallback), 'if (style) return;' always fired, and CreateStyle was never reached.
  // A match is only valid when the returned style's own GetName() equals the REQUESTED name.
  const rawStyle = doc.GetStyle(m.name);
  let style = rawStyle && typeof rawStyle.GetName === 'function' && rawStyle.GetName() === m.name ? rawStyle : null;
  const kind = m.styleType || { create: 'paragraph', getType: 'paragraph' }; // GH #292 S5: a meta without a type is a paragraph style
  if (m.mode === 'define') {
    if (style) return;
    // GH #292 S3: the base is resolved BEFORE the create -- a missing or non-paragraph base writes nothing (the same GetName guard as above)
    let base = null;
    if (m.fields.basedOn !== undefined) {
      const rawBase = doc.GetStyle(m.fields.basedOn);
      base = rawBase && typeof rawBase.GetName === 'function' && rawBase.GetName() === m.fields.basedOn ? rawBase : null;
      if (!base || typeof base.GetType !== 'function' || base.GetType() !== kind.getType) return;
    }
    style = doc.CreateStyle(m.name, kind.create);
    if (!style) return;
    if (base && typeof style.SetBasedOn === 'function') style.SetBasedOn(base);
  } else {
    if (!style) return;
  }
  // GH #292 S4: the Bridge gate BEFORE any public setter -- a gate refusal writes nothing at all (the driver prepends __eoS4Gate / __eoS4Write)
  let s4 = null;
  if (m.bridgeStyleNextLink === true) { s4 = __eoS4Gate(); if (!s4.ok) return { ok: false, stage: 'bridge-write-gate', reason: s4.reason, mutation: 'NONE' }; }
  if (m.fields.text) {
    const tp = typeof style.GetTextPr === 'function' ? style.GetTextPr() : null;
    if (tp) {
      if (m.fields.text.bold !== undefined && typeof tp.SetBold === 'function') tp.SetBold(m.fields.text.bold);
      if (m.fields.text.size !== undefined && typeof tp.SetFontSize === 'function') tp.SetFontSize(m.fields.text.size * 2);
      if (m.fields.text.color === 'auto' && typeof tp.SetColor === 'function' && typeof Api.AutoColor === 'function') tp.SetColor(Api.AutoColor());
      else if (m.fields.text.color !== undefined && typeof tp.SetColor === 'function' && typeof Api.RGB === 'function') {
        const hex = m.fields.text.color;
        const r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
        tp.SetColor(Api.RGB(r, g, b));
      }
      // GH #292 S1 (LIVE): the further text setters, each stored on the style's rPr
      if (m.fields.text.italic !== undefined && typeof tp.SetItalic === 'function') tp.SetItalic(m.fields.text.italic);
      if (m.fields.text.underline !== undefined && typeof tp.SetUnderline === 'function') tp.SetUnderline(m.fields.text.underline);
      if (m.fields.text.strike !== undefined && typeof tp.SetStrikeout === 'function') tp.SetStrikeout(m.fields.text.strike);
      if (m.fields.text.fontFamily !== undefined && typeof tp.SetFontFamily === 'function') tp.SetFontFamily(m.fields.text.fontFamily);
    }
  }
  if (m.fields.paragraph) {
    const pp = typeof style.GetParaPr === 'function' ? style.GetParaPr() : null;
    if (pp) {
      // GH #292 S6 (LIVE): SetJc('justify') stores nothing on a style -- the justify value is 'both'
      if (m.fields.paragraph.align !== undefined && typeof pp.SetJc === 'function') pp.SetJc(m.fields.paragraph.align === 'justify' ? 'both' : m.fields.paragraph.align);
      if (m.fields.paragraph.indentLeft !== undefined && typeof pp.SetIndLeft === 'function') pp.SetIndLeft(m.fields.paragraph.indentLeft * 20);
      if (m.fields.paragraph.spacingAfter !== undefined && typeof pp.SetSpacingAfter === 'function') pp.SetSpacingAfter(m.fields.paragraph.spacingAfter * 20, false);
      // GH #292 S2 (LIVE): twips at the Builder boundary, auto-spacing always false (the spacingAfter rule); a negative first line is a hanging indent
      if (m.fields.paragraph.spacingBefore !== undefined && typeof pp.SetSpacingBefore === 'function') pp.SetSpacingBefore(m.fields.paragraph.spacingBefore * 20, false);
      if (m.fields.paragraph.indentRight !== undefined && typeof pp.SetIndRight === 'function') pp.SetIndRight(m.fields.paragraph.indentRight * 20);
      if (m.fields.paragraph.indentFirstLine !== undefined && typeof pp.SetIndFirstLine === 'function') pp.SetIndFirstLine(m.fields.paragraph.indentFirstLine * 20);
      if (m.fields.paragraph.keepNext !== undefined && typeof pp.SetKeepNext === 'function') pp.SetKeepNext(m.fields.paragraph.keepNext);
      if (m.fields.paragraph.keepLines !== undefined && typeof pp.SetKeepLines === 'function') pp.SetKeepLines(m.fields.paragraph.keepLines);
      if (m.fields.paragraph.pageBreakBefore !== undefined && typeof pp.SetPageBreakBefore === 'function') pp.SetPageBreakBefore(m.fields.paragraph.pageBreakBefore);
    }
  }
  if (s4) {
    for (const f of ['next', 'link']) {
      if (m.fields[f] === undefined) continue;
      const w = __eoS4Write(s4.bridge, { style: m.name, field: f, target: m.fields[f] });
      if (!w.ok) return Object.assign({ stage: 'bridge-write', field: f }, w);
    }
    return { ok: true };
  }
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** Same shape as page-setup.cjs's bodyKey -- the document body, unaffected by a style edit. */
function bodyKey(primary) {
  const d = (primary && primary.document) || {};
  return JSON.stringify({ units: textUnits(primary).map(u => [u.kind, norm(u.text)]), tables: tablesContent(d.tables), footnotes: d.footnotes || null, endnotes: d.endnotes || null, bookmarks: d.bookmarks || null, images: ((primary && primary.images) || []).length, inMainBody: d.inMainBody === undefined ? null : d.inMainBody });
}

function otherStylesKey(styles, exceptName, linkTarget) {
  if (!styles || typeof styles !== 'object') return null;
  const names = Object.keys(styles).filter(n => n !== exceptName).sort();
  // GH #292 S4: a link sets BOTH sides -- the target's own link / linkName is the requested change, every other field of it must stay
  const view = n => (n === linkTarget && styles[n] && styles[n].identity ? { ...styles[n], identity: { ...styles[n].identity, link: null, linkName: null } } : styles[n]);
  return JSON.stringify(names.map(n => [n, view(n)]));
}

/** L1: every requested field now matches, via the driver's own getter-sourced read; never UNOBSERVABLE. */
function verifyPrimitive(op, before, after) {
  const b = styleAt(before && before.primary, op.name);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const a = styleAt(after && after.primary, op.name);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  if (a.st === null) return { ok: false, evidence: { reason: 'style-not-found-after-write' } };
  const unobs = unobservableRequestedFields(a.st, op.fields);
  const pairOk = op.fields.link === undefined || nameFieldMatches(a.styles[op.fields.link], 'linkName', op.name);
  const ok = (op.action === 'define' ? b.st === null : true) && allRequestedMatch(a.st, op.fields) && unobs.length === 0 && pairOk;
  return { ok, evidence: { name: op.name, fields: op.fields, unobservableAfter: unobs, styleReadAfter: a.st } };
}

/**
 * L3, PRIMITIVE_INTEGRAL (the page-setup/numbering owner-placement rule: the body and every
 * OTHER style must not change, as part of success). `modify_style`'s effect on paragraphs
 * ALREADY using this style name is intentionally NOT separately verified here -- design doc
 * point 4, PENDING-closed by W10-2: the cascade is NOT independently observable on the
 * paragraph (no effective-size getter found live), so this L3 states what IS measured --
 * every OTHER style's own fields, and the body's own pStyle/direct-formatting, unchanged --
 * never a guessed claim about the rendered APPEARANCE of paragraphs using the modified style.
 */
function verifyPreservation(op, before, after) {
  const b = styleAt(before && before.primary, op.name);
  const a = styleAt(after && after.primary, op.name);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` } };
  const sameOtherStyles = otherStylesKey(b.styles, op.name, op.fields.link) === otherStylesKey(a.styles, op.name, op.fields.link);
  const sameBody = bodyKey(before.primary) === bodyKey(after.primary);
  const ok = sameOtherStyles && sameBody;
  return { ok, evidence: { otherStylesUnchanged: sameOtherStyles, bodyUnchanged: sameBody } };
}

/** L2: the final style holds exactly the requested fields. */
function wholeTaskChecks(op, finalObservation) {
  const s = styleAt(finalObservation.primary, op.name);
  const ok = s.ok && s.st !== null && allRequestedMatch(s.st, op.fields) && (op.fields.link === undefined || nameFieldMatches(s.styles[op.fields.link], 'linkName', op.name));
  return [{ name: `${op.id}: style '${op.name}' (${op.action}) matches ${Object.keys(op.fields).join('+')}`, ok }];
}

// refs: none -- a style definition is not a body text unit a later step can name by reference.
function producedText() {
  return null;
}

/**
 * W-10: `doc.GetAllStyles()` cannot see a freshly created style (measured live,
 * the docx-driver.cjs `buildTargetedStyleQuery` comment explains why) -- a caller MUST pass the
 * names this task's own `define_style`/`modify_style` steps target via
 * `createExecutor(task, { observation: { live: { styleNames: styleNamesOf(task) } } })`, or
 * every such step risks reading back as verify-uncertain. Exported so ANY caller (this family's
 * own LIVE harness, and -- once wired -- a real MCP entry point) can derive the right value
 * without knowing the driver's internal targeted-read mechanism.
 */
function styleNamesOf(task) {
  const steps = (task && Array.isArray(task.steps)) ? task.steps : [];
  const names = steps.filter(s => s && (s.intent === 'define_style' || s.intent === 'modify_style') && typeof s.name === 'string').map(s => s.name);
  // GH #292 S3: a define_style base is read by name too (a session-created base is invisible to GetAllStyles)
  for (const s of steps) if (s && s.intent === 'define_style' && typeof s.basedOn === 'string' && s.basedOn) names.push(s.basedOn);
  // GH #292 S4: a next / link target is read by name too (its type, and a link target's own link)
  for (const s of steps) if (s && s.intent === 'modify_style') for (const f of NEXT_LINK) if (typeof s[f] === 'string' && s[f]) names.push(s[f]);
  return [...new Set(names)];
}

module.exports = {
  TEXT_FIELDS, PARAGRAPH_FIELDS, ALIGNMENTS, STEP_KEYS, STYLE_TYPES,
  plan, classify, resolve, preDispatchCheck, lower, commandBody,
  verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  styleAt, fieldMatches, groupMatches, allRequestedMatch, basedOnMatches, typeMatches, unobservableRequestedFields, otherStylesKey, bodyKey,
  styleNamesOf,
};
