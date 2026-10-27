'use strict';

const { norm, textUnits, tablesContent } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');

/**
 * DOCX `header-footer` family (M4 Wave 5): edit_header, edit_footer (P1, PRIMITIVE_INTEGRAL).
 *
 * Source oracle (joskapista44/euro-mcp f8c6d6cf, read-only):
 *   registry    docx-agent-capabilities.cjs:33-34 (rich_structure set_header / set_footer;
 *               requires fresh_live_preflight + single_section)
 *   planner     docx-agent-planner.cjs:30 (section-global, NO selector), :66 (text required;
 *               variant default | even | title, default when absent)
 *   lower       docx-session-rich-structures.cjs:10-58 -- frame.evaluate on the INTERNAL logic
 *               document (the backdoor the owner brief forbids): more than one section ->
 *               'header-footer-multisection-not-live-accepted'; Get_/Set_<Kind>_<Variant>, a missing
 *               part created (Create_SectionHdrFtr); ClearContent() + AddText(text) = the WHOLE
 *               content replaced by one plain paragraph; even -> the DOCUMENT-wide even/odd flag;
 *               title -> the section's title-page flag; Set_CurHdrFtr(hf) leaves the cursor in the part
 *   retry/L2    docx-agent-idempotency.cjs:175-179, docx-agent-whole-verify.cjs:76,
 *               docx-rich-structures.cjs:94-100 (the part's normalized text equals the text; nothing else)
 *   source LIVE default header and footer on a ONE-section document (DOCX-WORD-POWER-USER-CAPABILITY-
 *               AUDIT.md:362); even/title session-level only
 *
 * Measured live before migrating (test/live/docx-header-footer-probe*.live.cjs, 2026-09-25):
 *   - the public API reads and writes every part: ApiDocument.GetSections / CreateSection,
 *     ApiSection.GetHeader(v, create) / GetFooter(v, create); with create=false a missing part is
 *     null. The HTML projection shows NO header/footer text (probe 1, 0 of 11), so the driver reads
 *     the parts itself. The section object is a new wrapper per call: the identity is the INDEX;
 *   - a section without a part of its own reads null; creating one writes ONLY that section (probe 1
 *     W5/W6: the others unchanged by the read);
 *   - RemoveAllElements leaves ONE (reset) paragraph; AddText into it gives exactly one plain
 *     paragraph: the text, default run properties, left alignment -- the old formatting is NOT
 *     kept (probe 2 Q1/Q4, a formatted and a two-paragraph part) = the source's semantics;
 *   - the public writes leave the cursor in the main body (internal docPosType 0 after every step,
 *     read-only control; a body write right after a header write landed in the body, probe 2 Q2);
 *   - the title-page and even/odd flags are WRITE-only in the public API (no getter); the title
 *     flag is readable only internally, the even/odd flag not even there (probe 2 Q5/Q6).
 *
 * Where the target differs from the source, stated:
 *   identity    section + kind + variant, explicit: `section: { index, of }` -- `of` is the
 *               section count the caller saw; a different count -> FAIL_CLOSED (never "the first").
 *               The source took the final section of a one-section document implicitly.
 *   variants    default, title, even: SUPPORTED (M5, owner decision after 28799 / 28808, option A).
 *               The flags that make a variant visible (title page: per section; even/odd: document-
 *               wide, target probe 2) have no public getter: they are read READ ONLY from the internal
 *               section (EO_INTERNAL_DEPENDENCY READ; unread -> FAIL_CLOSED ...-flag-unobservable).
 *               The write uses the public API only: GetHeader/GetFooter(variant, true) + the content,
 *               then SetTitlePage(true) / SetEvenAndOddHdrFtr(true) when the flag was off. Turning a
 *               flag on switches UNREQUESTED slots too (title: the other kind of that section's first
 *               page; even: both kinds on the even pages of every section) -- allowed only when each
 *               of them is PROVABLY unchanged (both provably empty, or both own parts with identical
 *               content); unknown / inherited -> FAIL_CLOSED before dispatch. No slot is created or
 *               copied. The terminal state is the part AND its flag on; L3 also requires every other
 *               flag unchanged. Re-measure the internal readers after every Euro-Office release
 *               (test/live/docx-hf-variant-target-probe.live.cjs).
 *   sections    more than one section is SUPPORTED when it is safe (the source refused it):
 *               a part is inherited by the following sections that have none of their own, so a
 *               write where a LATER section has no own part of the same kind+variant ->
 *               FAIL_CLOSED_UNSUPPORTED_SHAPE:later-section-inherits-this-part:<sections> (GH #289 H-07a: written when
 *               the caller names exactly those sections in `inheritedBy`); a missing part on a
 *               section that has a previous one would break the link ->
 *               FAIL_CLOSED_UNSUPPORTED_SHAPE:linked-to-previous. A one-section document keeps the
 *               source's create-if-missing.
 *   content     a part holding anything but paragraphs of runs (a field, a table, an image) is
 *               refused (the replacement would destroy it silently).
 *   L1          the part holds exactly one paragraph with the text (the source compared the text).
 *   L3 (P1)     the body (text units in order, tables, notes, bookmarks, images) and EVERY other part
 *               of every section unchanged, the section count unchanged -- owner W5 points 11/21/22.
 *               The source audit placed these intents P0; the target is P1 by the owner brief.
 *   retry       SATISFIED from the fresh part (one paragraph, the text), never from the old text.
 *   text        non-empty, no tab or line break (one paragraph).
 *
 * GH #289 H-04b: `pageNumber: true` (full mode only) puts a PAGE field after the text in the one paragraph (public
 * AddPageNumber, LIVE probe 1 W4); L1/retry: the paragraph holds the text AND exactly one PAGE field.
 *
 * GH #289 paragraph mode (docs/GH289-HF-DOCUMENTED-LIMITS-DECISION.md, H-01): `paragraph: { index, of }` changes the
 * text of ONE paragraph of an EXISTING part instead of replacing the part. `of` is the part's paragraph count the
 * caller saw (a different count -> FAIL_CLOSED, never "the nearest"). Measured LIVE (probe 1 W1/W2/W3/W6, kept after a
 * fresh reopen): the text of the paragraph's ONE text run is changed in place (ClearContent + AddText on that run), so
 * the paragraph style (e.g. the built-in Header), the alignment, the run's bold / font / size, and every run with no
 * text -- a PAGE field, an image -- are kept. The part must hold only paragraphs of runs (a table in the part: refused);
 * the paragraph must have exactly ONE run with text (two differently formatted runs: refused -- which format the new
 * text takes would be a guess). No part is created and no variant is activated in this mode. L1: that paragraph holds
 * the text, the count unchanged. L3: that paragraph's style, alignment, run properties and inline content unchanged,
 * every other paragraph of the part unchanged, plus everything the full mode compares (every other part, the body,
 * the flags). Retry: SATISFIED from the fresh paragraph text.
 */

const KIND = Object.freeze({ edit_header: 'header', edit_footer: 'footer' });
const VARIANTS = Object.freeze(['default', 'even', 'title']);
// M5 (owner decision after 28799 / 28808, option A): title and even are supported. Their visibility
// flags (title page: per section; even/odd: document-wide) are read internally, READ ONLY
// (EO_INTERNAL_DEPENDENCY READ -- the driver's section read); an unread flag refuses the write.
const SUPPORTED_VARIANTS = Object.freeze(['default', 'title', 'even']);

/** GH #289 H-07a: the optional `inheritedBy` acknowledgement -> null (absent) | a frozen ascending list | false (malformed). */
function inheritedByDescriptor(v, section) {
  if (v === undefined) return null;
  if (!Array.isArray(v) || !v.every(Number.isInteger)) return false;
  const ok = v.every((x, i) => x > section.index && x < section.of && (i === 0 || x > v[i - 1]));
  return ok ? Object.freeze([...v]) : false;
}

/**
 * GH #289 H-07a: the sections that SHOW this part without having one of their own -- the later sections, in order, up
 * to the first one with its own part of the same kind + variant (ECMA-376: a section without a header/footer reference
 * inherits the previous section's). A write to the part changes what these sections show.
 */
function inheritingSections(sections, op) {
  const out = [];
  for (let j = op.sectionIndex + 1; j < sections.length && !partAt(sections, j, op.part, op.variant); j += 1) out.push(j);
  return out;
}

/** GH #289 H-07a: an inheriting set is written only when the caller named it exactly (never implicit). */
function inheritanceGuard(s, op) {
  const aff = inheritingSections(s.sections, op);
  const named = op.inheritedBy || null;
  // the reason NAMES the set (the receipt carries the reason, not the evidence): the caller acknowledges it by inheritedBy
  if (named === null) return aff.length ? `FAIL_CLOSED_UNSUPPORTED_SHAPE:later-section-inherits-this-part:${aff.join(',')}` : null;
  return JSON.stringify(aff) === JSON.stringify(named) ? null : `FAIL_CLOSED:inherited-by-mismatch:${aff.join(',') || 'none'}`;
}

/** GH #289: the optional `paragraph` descriptor -> null (absent) | { index, of } | false (malformed). */
function paragraphDescriptor(q) {
  if (q === undefined) return null;
  const valid = q && typeof q === 'object' && Object.keys(q).sort().join(',') === 'index,of' && Number.isInteger(q.index) && Number.isInteger(q.of);
  return valid && q.of >= 1 && q.index >= 0 && q.index < q.of ? Object.freeze({ index: q.index, of: q.of }) : false;
}

function plan(step, index) {
  const kind = KIND[step.intent];
  if (!kind) return { ok: false, detail: `step ${index}: ${step.intent} is not a header-footer intent` };
  const variant = step.variant === undefined || step.variant === null ? 'default' : String(step.variant);
  if (!VARIANTS.includes(variant)) return { ok: false, detail: `step ${index}: ${step.intent} variant must be one of ${VARIANTS.join(', ')}` };
  const s = step.section;
  if (!s || typeof s !== 'object' || Object.keys(s).sort().join(',') !== 'index,of' || !Number.isInteger(s.index) || !Number.isInteger(s.of) || s.of < 1 || s.index < 0 || s.index >= s.of) {
    return { ok: false, detail: `step ${index}: ${step.intent} requires section: { index, of } (integers, 0 <= index < of = the section count) -- no implicit section` };
  }
  if (step.selector !== undefined) return { ok: false, detail: `step ${index}: ${step.intent} takes no text selector (the target is section + kind + variant)` };
  if (typeof step.text !== 'string' || !norm(step.text)) return { ok: false, detail: `step ${index}: ${step.intent} requires a non-empty text` };
  if (/[\t\r\n]/.test(step.text)) return { ok: false, detail: `step ${index}: ${step.intent} text must not contain a tab or line break (one paragraph)` };
  const para = paragraphDescriptor(step.paragraph);
  const inheritedBy = inheritedByDescriptor(step.inheritedBy, s);
  if (step.pageNumber !== undefined && typeof step.pageNumber !== 'boolean') return { ok: false, detail: `step ${index}: ${step.intent} pageNumber must be a boolean` };
  if (step.pageNumber === true && para) return { ok: false, detail: `step ${index}: ${step.intent} pageNumber is a full-part write (it cannot be combined with paragraph)` };
  if (inheritedBy === false) return { ok: false, detail: `step ${index}: ${step.intent} inheritedBy must be an ascending list of the later section indexes (> section.index, < section.of)` };
  if (para === false) return { ok: false, detail: `step ${index}: ${step.intent} paragraph must be { index, of } (integers, 0 <= index < of = the part's paragraph count)` };
  const selector = Object.freeze({ section: Object.freeze({ index: s.index, of: s.of }), part: kind, variant, ...(para ? { paragraph: para } : {}) });
  return { ok: true, operation: { id: step.id || `s${index}`, intent: step.intent, selector, part: kind, variant, sectionIndex: s.index, sectionCount: s.of, text: step.text,
    ...(para ? { paragraphIndex: para.index, paragraphCount: para.of } : {}), ...(inheritedBy ? { inheritedBy } : {}), ...(step.pageNumber === true ? { pageNumber: true } : {}), preservationMode: 'PRIMITIVE_INTEGRAL', ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) } };
}

// ---- the observation ----------------------------------------------------------------------------

const partAt = (sections, i, kind, variant) => (sections[i] && sections[i][kind] ? sections[i][kind][variant] : null);
const terminal = (part, text) => !!part && part.paragraphs === 1 && part.paras.length === 1 && part.paras[0].text === norm(text);
// GH #289 H-04b: with pageNumber the one paragraph also holds exactly one PAGE field (the public ToJSON names it pgNum) and
// nothing else; without it the paragraph holds no field (a stale field from an earlier write is not the requested state)
const pageNumberOk = (part, op) => { const inl = part && part.paras[0] && part.paras[0].inline; return Array.isArray(inl) && JSON.stringify(inl) === JSON.stringify(op.pageNumber ? ['pgNum'] : []); };
const paragraphMode = op => Number.isInteger(op.paragraphIndex);
// GH #289 H-05 (LIVE, docs/evidence/gh289-290-remaining-probe-20261023.json): in a part that also holds tables the paragraph mode
// counts the part's OWN paragraphs (a table and the paragraphs inside its cells are not counted) and leaves every table as it is
const paragraphCountOf = part => (part.withTables ? part.paras.length : part.paragraphs);
// GH #289 paragraph mode: the part still holds the caller's paragraph count, and the named paragraph holds the text
const paragraphTerminal = (part, op) => !!part && paragraphCountOf(part) === op.paragraphCount && part.paras.length === op.paragraphCount && part.paras[op.paragraphIndex].text === norm(op.text);
const OTHER_KIND = Object.freeze({ header: 'footer', footer: 'header' });

/** The even/odd flag is document-wide: every section must report the same readable value, else null (unknown). */
function evenFlag(sections) {
  const values = sections.map(s => s.evenAndOdd);
  return values.length && values.every(v => v === values[0] && typeof v === 'boolean') ? values[0] : null;
}

/** Is the variant's visibility flag on? true | false | null (unread / inconsistent). default: always visible. */
function flagOn(sections, op) {
  if (op.variant === 'title') { const v = sections[op.sectionIndex] && sections[op.sectionIndex].titlePage; return typeof v === 'boolean' ? v : null; }
  if (op.variant === 'even') return evenFlag(sections);
  return true;
}

/** The state that decides classification, resolution and verification, from ONE observation. */
function partState(primary, op) {
  const d = primary && primary.document;
  const sections = d ? d.sections : null;
  if (!Array.isArray(sections)) return { ok: false, reason: 'sections-unreadable' };
  if (sections.length !== op.sectionCount) return { ok: false, reason: 'FAIL_CLOSED:section-descriptor-mismatch', sectionCount: sections.length };
  const part = partAt(sections, op.sectionIndex, op.part, op.variant);
  const flag = flagOn(sections, op);
  return { ok: true, sections, part, flag, defaultText: d.defaultText || null, done: (paragraphMode(op) ? paragraphTerminal(part, op) : terminal(part, op.text) && (!op.pageNumber || pageNumberOk(part, op))) && flag === true };
}

/**
 * What a slot SHOWS, provably: OWN (the section's own part; key = its content), NONE (no own part and
 * no earlier section has one of this kind + variant: provably empty) or INHERITED (no own part but an
 * earlier section has one -- what it shows is not observable through the API: fail closed).
 */
function slotState(sections, i, kind, variant) {
  const own = partAt(sections, i, kind, variant);
  if (own) return { state: 'OWN', key: JSON.stringify([own.paragraphs, own.simple, own.paras.map(x => [x.text, x.format]), own.tables]) };
  for (let j = 0; j < i; j += 1) if (partAt(sections, j, kind, variant)) return { state: 'INHERITED' };
  return { state: 'NONE' };
}

/**
 * Turning a variant's flag ON makes the variant slot shown instead of the default one -- on the first
 * page of THAT section (title) or on the even pages of EVERY section (even/odd is document-wide), for
 * BOTH kinds. Owner decision (option A): activate only if every unrequested slot that switches is
 * PROVABLY unchanged -- both provably empty, or both own parts with identical content. The target
 * slot itself is the requested change. Nothing is created or copied.
 */
function activationGuard(s, op) {
  const pairs = op.variant === 'title'
    ? [[op.sectionIndex, OTHER_KIND[op.part]]]
    : s.sections.flatMap((_, i) => ['header', 'footer'].map(k => [i, k])).filter(([i, k]) => !(i === op.sectionIndex && k === op.part));
  for (const [i, k] of pairs) {
    const shown = slotState(s.sections, i, k, 'default');
    const after = slotState(s.sections, i, k, op.variant);
    const where = `section-${i}-${k}`;
    if (shown.state === 'INHERITED' || after.state === 'INHERITED') return `FAIL_CLOSED_UNSUPPORTED_SHAPE:${op.variant}-activation-inherited-${where}`;
    const same = (shown.state === 'NONE' && after.state === 'NONE') || (shown.state === 'OWN' && after.state === 'OWN' && shown.key === after.key);
    if (!same) return `FAIL_CLOSED_UNSUPPORTED_SHAPE:${op.variant}-activation-changes-unrequested-${where}`;
  }
  return null;
}

/** Why a write here is unsafe, or null (owner W5 points 12-14, 29-30). */
function flagGuard(s, op) {
  if (s.flag !== null) return null;
  return op.variant === 'even' && s.sections.some(x => typeof x.evenAndOdd === 'boolean')
    ? 'FAIL_CLOSED_UNSUPPORTED_SHAPE:even-flag-inconsistent-across-sections'
    : `FAIL_CLOSED_UNSUPPORTED_SHAPE:variant-${op.variant}-flag-unobservable`;
}

/** B21: every section's variant state must be KNOWN (true/false) -- the preservation check compares it before/after, and
 * an unavailable state (unverified runtime, reader missing/threw) is never taken as a real one. */
function variantStateGuard(s) {
  const unknown = s.sections.findIndex(x => typeof x.titlePage !== 'boolean' || typeof x.evenAndOdd !== 'boolean');
  return unknown >= 0 ? `FAIL_CLOSED_UNSUPPORTED_SHAPE:section-variant-state-unavailable-section-${unknown}` : null;
}

const variantGuard = op => (SUPPORTED_VARIANTS.includes(op.variant) ? null : `FAIL_CLOSED_UNSUPPORTED_SHAPE:variant-${op.variant}-unsupported`);

function shapeGuard(s, op) {
  const early = variantGuard(op) || flagGuard(s, op) || variantStateGuard(s);
  if (early) return early;
  if (paragraphMode(op)) return paragraphGuard(s, op);
  const inherited = inheritanceGuard(s, op);
  if (inherited) return inherited;
  if (!s.part && op.sectionIndex > 0) return 'FAIL_CLOSED_UNSUPPORTED_SHAPE:linked-to-previous';
  if (s.part && !s.part.simple) return 'header-footer-not-plain-paragraphs-of-runs';
  if (s.flag === false) { const a = activationGuard(s, op); if (a) return a; }
  return s.part ? replacementLoss(s.part, s.defaultText) : null;
}

/**
 * GH #289 paragraph mode: an existing, shown part of paragraphs of runs (no later section inheriting it), the caller's
 * paragraph count, and a target paragraph with exactly ONE run that holds text. Returns the refusal or null.
 */
function paragraphGuard(s, op) {
  const P = 'FAIL_CLOSED_UNSUPPORTED_SHAPE:header-footer-paragraph-';
  const inherited = inheritanceGuard(s, op);
  const part = s.part;
  const p = part && part.paras.length === op.paragraphCount ? part.paras[op.paragraphIndex] : null;
  const textRuns = p && Array.isArray(p.runs) ? p.runs.filter(r => r.text !== '') : [];
  // in order: the first one that holds is the refusal
  const checks = [
    [!!inherited, inherited],
    [!part, `${P}mode-needs-an-existing-part`],
    [() => s.flag !== true, `${P}mode-variant-not-shown`],
    [() => !part.simple && !part.withTables, 'header-footer-not-plain-paragraphs-of-runs'],
    [() => paragraphCountOf(part) !== op.paragraphCount || !p, 'FAIL_CLOSED:paragraph-descriptor-mismatch'],
    [() => !Array.isArray(p.runs), 'header-footer-run-format-unreadable'],
    [() => !Array.isArray(p.inline), 'header-footer-inline-content-unreadable'],
    [() => textRuns.length !== 1, () => `${P}text-runs:${textRuns.length}`],
    [() => norm(textRuns[0].text) !== p.text, `${P}text-not-in-one-run`],
  ];
  const hit = checks.find(([c]) => (typeof c === 'function' ? c() : c));
  if (!hit) return null;
  return typeof hit[1] === 'function' ? hit[1]() : hit[1];
}

/**
 * M5 gate 6 (owner decision 1 after Talk 28947, internal tracker): the full replacement yields ONE
 * paragraph with NO paragraph style, left alignment, and runs in the document's DEFAULT text
 * properties without bold / italic / underline / strike (header/footer probe 2 Q1, 2026-09-25). An
 * EXISTING part is replaced only when that result loses nothing the part holds -- the built-in
 * 'Header' / 'Footer' paragraph style included (the replacement drops it, so a part carrying it is NOT
 * "unformatted"). Anything else -> refused before dispatch, the lost property named. A MISSING part
 * (create) is a separate scope: there is nothing to lose. This is a SAFETY BOUNDARY narrower than the
 * source (which replaces any plain-runs part and drops its formatting) -- owner-accepted, not PARITY
 * and not an EO_LIMIT. The in-command re-check covers the STRUCTURAL guards only; this one is
 * evaluated on the fresh observation right before the dispatch.
 */
const LOSS = 'FAIL_CLOSED_UNSUPPORTED_SHAPE:header-footer-replacement-would-drop-';
function replacementLoss(part, defaults) {
  if (part.paragraphs !== 1 || part.paras.length !== 1) return `${LOSS}paragraphs`;
  const p = part.paras[0];
  // GH #289: a field (PAGE ...) or an image reads as runs with no text -- the replacement would drop it silently
  if (!Array.isArray(p.inline)) return 'header-footer-inline-content-unreadable';
  if (p.inline.length) return `${LOSS}inline:${p.inline.join(',')}`;
  const f = p.format || {};
  if (f.style) return `${LOSS}style:${f.style}`;
  if (f.alignment !== undefined && f.alignment !== 'left') return `${LOSS}alignment:${f.alignment}`;
  if (!Array.isArray(p.runs)) return 'header-footer-run-format-unreadable';
  if (!defaults) return 'header-footer-default-text-format-unreadable';
  for (const r of p.runs.filter(x => x.text !== '')) {
    for (const k of ['bold', 'italic', 'underline', 'strike']) if (r[k] !== false) return `${LOSS}run-${k}`;
    if (r.fontFamily !== defaults.fontFamily) return `${LOSS}run-font-family`;
    if (r.fontSize !== defaults.fontSize) return `${LOSS}run-font-size`;
  }
  return null;
}

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  const s = partState(primary, op);
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason, sectionCount: s.sectionCount } };
  if (s.done) {
    if (op.recovery === UNRESOLVED) return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } };
    return { state: 'SATISFIED', evidence: { section: op.sectionIndex, part: op.part, variant: op.variant, text: norm(op.text), proof: paragraphMode(op) ? 'fresh-live-paragraph-holds-the-text' : 'fresh-live-part-holds-exactly-one-paragraph-with-the-text' } };
  }
  const guard = shapeGuard(s, op);
  const aff = inheritingSections(s.sections, op);
  // H-07a: the preflight NAMES the sections a write would change, refused or not -- the caller acknowledges them by inheritedBy
  if (guard) return { state: 'FAIL_CLOSED', evidence: { reason: guard, ...(aff.length || op.inheritedBy ? { inheritingSections: aff } : {}) } };
  const how = paragraphMode(op) ? { paragraph: op.paragraphIndex, from: s.part.paras[op.paragraphIndex].text } : { create: !s.part, activate: s.flag === false };
  return { state: 'APPLY', evidence: { section: op.sectionIndex, part: op.part, variant: op.variant, ...how, ...(aff.length ? { inheritingSections: aff } : {}) } };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  if (!op || !KIND[op.intent]) return { ok: false, target: null, selector, ambiguity: 'INVALID_SELECTOR' };
  const s = partState(primary, op);
  const up = r => r.toUpperCase().replace(/-/g, '_');
  if (!s.ok) return { ok: false, target: null, selector, ambiguity: up(s.reason) };
  const guard = s.done ? null : shapeGuard(s, op);
  if (guard) return { ok: false, target: null, selector, ambiguity: up(guard) };
  const para = paragraphMode(op) && s.part ? { paragraphIndex: op.paragraphIndex, paragraphCount: op.paragraphCount, oldText: s.part.paras[op.paragraphIndex].text } : {};
  return { ok: true, target: { kind: 'section_part', sectionIndex: op.sectionIndex, sectionCount: op.sectionCount, part: op.part, variant: op.variant, create: !s.part, activate: s.flag === false, inheritedBy: inheritingSections(s.sections, op), ...para }, selector, ambiguity: null };
}

/** GH #289: a paragraph-mode target names the same paragraph, its old text, and neither creates nor activates; a full-mode target names no paragraph. */
function sameParagraphTarget(op, target) {
  if (!paragraphMode(op)) return target.paragraphIndex === undefined;
  return target.paragraphIndex === op.paragraphIndex && target.paragraphCount === op.paragraphCount && typeof target.oldText === 'string' && target.create === false && target.activate === false;
}

function preDispatchCheck(op, target) {
  const same = target && target.kind === 'section_part' && target.sectionIndex === op.sectionIndex && target.sectionCount === op.sectionCount
    && target.part === op.part && target.variant === op.variant && SUPPORTED_VARIANTS.includes(op.variant) && sameParagraphTarget(op, target)
    && Array.isArray(target.inheritedBy) && JSON.stringify(target.inheritedBy) === JSON.stringify(op.inheritedBy || []);
  return same ? { ok: true } : { ok: false, dispatched: false, detail: 'header-footer-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({ family: 'headerFooter', action: 'headerFooter', part: target.part, variant: target.variant, sectionIndex: target.sectionIndex, sectionCount: target.sectionCount, create: target.create, activate: target.activate === true, text: op.text,
    ...(paragraphMode(op) ? { mode: 'paragraph', paragraphIndex: target.paragraphIndex, paragraphCount: target.paragraphCount, oldText: target.oldText } : {}),
    inheritedBy: target.inheritedBy,
    ...(op.pageNumber ? { pageNumber: true } : {}),
    // M11.5e.4: a title/even command re-reads the variant state through the word /3 Bridge READ -- the driver prepends
    // its prelude and the second line refuses it off /3, before anything is sent
    ...(target.variant !== 'default' ? { bridgeSectionVariantRead: true } : {}) });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand, public API only (header/footer probes 1-2). It re-checks what the plan decided:
 * the section count; that the part exists exactly when it is not being created; with more than one
 * section, that every LATER section has its own part of the same kind + variant, and that a
 * created part is on the first section; that the part holds only paragraphs of runs. Then the whole
 * content is replaced: RemoveAllElements (one reset paragraph remains, probe 2) + AddText on it.
 * Any guard miss returns unchanged. No cursor call: the public write leaves it in the body.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  const secs = doc.GetSections() || [];
  if (secs.length !== m.sectionCount) return;
  const get = (sec, create) => (m.part === 'header' ? sec.GetHeader(m.variant, create) : sec.GetFooter(m.variant, create));
  const sec = secs[m.sectionIndex];
  if (!sec) return;
  // M11.5e.4: the variant's flag re-checked through the word /3 Bridge READ (__eoB21Variant, the driver's prelude);
  // a changed or unavailable state -> unchanged
  if (m.variant !== 'default') {
    const v = __eoB21Variant(m.sectionIndex);
    if (!v) return;
    const on = m.variant === 'title' ? v.titlePage : v.evenAndOdd;
    if (on === m.activate) return;
  }
  const existing = get(sec, false);
  if (!!existing === m.create) return;
  if (secs.length > 1) {
    // H-07a: the sections inheriting this part, re-read -- they must be exactly the ones the plan named
    const aff = [];
    for (let j = m.sectionIndex + 1; j < secs.length && !get(secs[j], false); j += 1) aff.push(j);
    if (JSON.stringify(aff) !== JSON.stringify(m.inheritedBy)) return;
    if (m.create && m.sectionIndex > 0) return;
  }
  // GH #289 paragraph mode: the text of the named paragraph's ONE text run, changed in place -- the paragraph, its
  // style and alignment, the run's properties and every run without text (a field, an image) stay as they are
  if (m.mode === 'paragraph') {
    if (!existing) return;
    const norm = t => String(t || '').replace(/\\s+/g, ' ').trim();
    const paras = [];
    for (let i = 0; i < existing.GetElementsCount(); i += 1) {
      const e = existing.GetElement(i);
      // GH #289 H-05: a table is left as it is (not counted, not entered)
      if (e && e.GetClassType() === 'table') continue;
      if (!e || e.GetClassType() !== 'paragraph') return;
      for (let j = 0; j < e.GetElementsCount(); j += 1) { const ch = e.GetElement(j); if (!ch || ch.GetClassType() !== 'run') return; }
      paras.push(e);
    }
    if (paras.length !== m.paragraphCount) return;
    const p = paras[m.paragraphIndex];
    if (norm(p.GetText({ Numbering: false })) !== m.oldText) return;
    const textRuns = [];
    for (let j = 0; j < p.GetElementsCount(); j += 1) { const ch = p.GetElement(j); if (String(ch.GetText({ Numbering: false }) || '') !== '') textRuns.push(ch); }
    if (textRuns.length !== 1) return;
    textRuns[0].ClearContent();
    textRuns[0].AddText(m.text);
    return;
  }
  if (existing) {
    for (let i = 0; i < existing.GetElementsCount(); i += 1) {
      const e = existing.GetElement(i);
      if (!e || e.GetClassType() !== 'paragraph') return;
      for (let j = 0; j < e.GetElementsCount(); j += 1) { const ch = e.GetElement(j); if (!ch || ch.GetClassType() !== 'run') return; }
    }
  }
  const hf = existing || get(sec, true);
  if (!hf) return;
  hf.RemoveAllElements();
  const p = hf.GetElement(0);
  if (!p || p.GetClassType() !== 'paragraph') return;
  p.AddText(m.text);
  // GH #289 H-04b: a PAGE field after the text (public ApiParagraph.AddPageNumber, LIVE probe 1 W4)
  if (m.pageNumber) p.AddPageNumber();
  // activation AFTER the part holds its text (public API): title page per section, even/odd document-wide
  if (m.activate && m.variant === 'title') sec.SetTitlePage(true);
  if (m.activate && m.variant === 'even') doc.SetEvenAndOddHdrFtr(true);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** Every part of every section except the target, as a comparable key (null parts included). */
function otherParts(sections, op) {
  const out = {};
  sections.forEach((s, i) => ['header', 'footer'].forEach(k => VARIANTS.forEach(v => {
    if (i === op.sectionIndex && k === op.part && v === op.variant) return;
    const p = s[k][v];
    out[`${i}:${k}:${v}`] = p ? JSON.stringify([p.paragraphs, p.simple, p.paras.map(x => [x.text, x.format]), p.tables]) : null;
  })));
  return out;
}

/**
 * M5: every visibility flag except the requested one must be unchanged; the requested one may only go
 * off -> on (the activation), never on -> off. title: the target section's title page; even: the
 * document-wide even/odd flag (reported per section).
 */
const unknownStates = (b, a, i) => ['titlePage', 'evenAndOdd'].filter(k => typeof b[k] !== 'boolean' || typeof a[k] !== 'boolean').map(k => `${i}:${k}-unknown`);

function flagChanges(before, after, op) {
  const out = [];
  before.forEach((b, i) => {
    const a = after[i] || {};
    const title = op.variant === 'title' && i === op.sectionIndex;
    // B21: an unknown state on either side is not a preserved one
    out.push(...unknownStates(b, a, i));
    if (b.titlePage !== a.titlePage && !(title && b.titlePage === false && a.titlePage === true)) out.push(`${i}:titlePage`);
    if (b.evenAndOdd !== a.evenAndOdd && !(op.variant === 'even' && b.evenAndOdd === false && a.evenAndOdd === true)) out.push(`${i}:evenAndOdd`);
  });
  return out;
}

/** The main body as the reads see it: text units in order, tables, notes, bookmarks, images, and whether the cursor is in it. */
function bodyKey(primary) {
  const d = (primary && primary.document) || {};
  return JSON.stringify({ units: textUnits(primary).map(u => [u.kind, norm(u.text)]), tables: tablesContent(d.tables), footnotes: d.footnotes || null, endnotes: d.endnotes || null, bookmarks: d.bookmarks || null, images: ((primary && primary.images) || []).length, inMainBody: d.inMainBody === undefined ? null : d.inMainBody });
}

/** L1: the target part holds exactly one paragraph with the text; the section count unchanged. */
function verifyPrimitive(op, before, after) {
  const b = partState(before && before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const a = partState(after && after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const part = a.part;
  const ok = !b.done && a.done;
  return { ok, evidence: { section: op.sectionIndex, part: op.part, variant: op.variant, created: !b.part, paragraphs: part ? part.paragraphs : null, text: part ? part.paras.map(x => x.text) : null } };
}

/**
 * L3, PRIMITIVE_INTEGRAL (owner W5 points 11, 21, 22): the body unchanged; every other part of every
 * section unchanged (text, paragraphs, format; a missing part stays missing); the section count
 * unchanged. The replaced part's own old formatting is NOT claimed (full replacement, probe 2).
 */
function verifyPreservation(op, before, after) {
  const b = partState(before && before.primary, op);
  const a = partState(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` } };
  const pb = otherParts(b.sections, op);
  const pa = otherParts(a.sections, op);
  const changedParts = Object.keys(pb).filter(k => pb[k] !== pa[k]);
  const sameBody = bodyKey(before.primary) === bodyKey(after.primary);
  const flags = flagChanges(b.sections, a.sections, op);
  const inPart = paragraphMode(op) ? targetPartChanges(b.part, a.part, op) : [];
  const ok = changedParts.length === 0 && sameBody && flags.length === 0 && inPart.length === 0;
  return { ok, evidence: { changedParts, bodyUnchanged: sameBody, unrequestedFlagChanges: flags, ...(paragraphMode(op) ? { targetPartChanges: inPart } : {}), sectionCount: a.sections.length, expectationSource: 'derived-from-fresh-pre-mutation-observation' } };
}

/**
 * GH #289 paragraph mode, L3 inside the target part: the paragraph count; every OTHER paragraph (text, format, runs,
 * inline content); the target paragraph's format, its run properties (every run, in order, text aside) and its inline
 * content. Only the target paragraph's text may change.
 */
function targetPartChanges(b, a, op) {
  if (!b || !a) return ['target-part-missing'];
  if (b.paragraphs !== a.paragraphs || b.paras.length !== a.paras.length) return ['paragraph-count'];
  // GH #289 H-05: every table of the part (where it sits, every cell's text) unchanged
  if (JSON.stringify(b.tables) !== JSON.stringify(a.tables)) return ['tables'];
  const runProps = runs => (Array.isArray(runs) ? runs.map(r => Object.fromEntries(Object.entries(r).filter(([k]) => k !== 'text'))) : null);
  const out = [];
  b.paras.forEach((x, i) => {
    const y = a.paras[i];
    if (i !== op.paragraphIndex) { if (JSON.stringify([x.text, x.format, x.runs, x.inline]) !== JSON.stringify([y.text, y.format, y.runs, y.inline])) out.push(`paragraph-${i}`); return; }
    if (JSON.stringify(x.format) !== JSON.stringify(y.format)) out.push(`paragraph-${i}-format`);
    if (runProps(x.runs) === null || JSON.stringify(runProps(x.runs)) !== JSON.stringify(runProps(y.runs))) out.push(`paragraph-${i}-run-properties`);
    if (!Array.isArray(x.inline) || JSON.stringify(x.inline) !== JSON.stringify(y.inline)) out.push(`paragraph-${i}-inline-content`);
  });
  return out;
}

/** L2: the final part (section + kind + variant) holds exactly one paragraph with the text; the section count as declared. */
function wholeTaskChecks(op, finalObservation) {
  const s = partState(finalObservation.primary, op);
  if (paragraphMode(op)) return [{ name: `${op.id}: section ${op.sectionIndex}/${op.sectionCount} ${op.part} (${op.variant}) paragraph ${op.paragraphIndex}/${op.paragraphCount} = "${norm(op.text)}"`, ok: s.ok && s.done }];
  return [{ name: `${op.id}: section ${op.sectionIndex}/${op.sectionCount} ${op.part} (${op.variant}) = exactly one paragraph "${norm(op.text)}"${op.pageNumber ? ' + a PAGE field' : ''}`, ok: s.ok && s.done }];
}

// refs: none -- the header/footer text is not a body text unit a later step can name
function producedText() {
  return null;
}

module.exports = { replacementLoss, paragraphGuard, targetPartChanges, plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, partState, shapeGuard, otherParts, bodyKey, slotState, activationGuard, flagChanges, SUPPORTED_VARIANTS };
