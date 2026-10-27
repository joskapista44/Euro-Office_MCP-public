'use strict';

const { textUnits, norm, tablesContent } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');

/**
 * DOCX `page-setup` family (post-M7, W11 -- a NEW capability, not a source-migrated intent;
 * registered via src/adapters/docx/capabilities.cjs's NEW_INTENTS, disjoint from the 28/28
 * migration inventory, per the owner decision, option b). See docs/W11-PAGE-SETUP-DESIGN-20261007.md for
 * the full design.
 *
 * Measured live (test/live/docx-w11-1a-section-reflection-probe.live.cjs,
 * docx-w11-1b-section-write-probe.live.cjs, both accepted; two follow-up probes closing the
 * design doc's open questions):
 *   - GetFinalSection()/GetSections() expose the section object; GetPageWidth/GetPageHeight/
 *     SetPageSize, GetPageMargin-and-SetPageMargins, GetColumnsCount(+derived widths)/
 *     SetEqualColumns/SetNotEqualColumns, GetStartPageNumber/SetStartPageNumber, SetTitlePage are
 *     confirmed public and persistent across two reopens -- 16/16 cross-axis attribution checks
 *     (W11-1b), each axis changes only at its own call, every other axis stays put.
 *   - `columns`: SetEqualColumns(count, space) with count=2, space=720 (distinct values) read back
 *     via GetColumnsCount()===2 -- the first-position-is-count order is CONFIRMED. `space` is ALSO
 *     independently re-readable via GetColumnsSpaces() (an array, count-1 gap widths) -- confirmed
 *     live, session + two reopens. An earlier version of this file believed `space` was
 *     not independently observable (confusing it with the DERIVED GetColumnsWidths) and verified
 *     count only; this was wrong and is fixed here -- both count and space are now part of L1/L3.
 *   - `startPageNumber`: single argument, no order question.
 *   - `size`: SetPageSize(width, height) -- the 2-arg call (3rd argument OMITTED) is MEASURED
 *     EQUIVALENT to passing true/false there for the one axis tried (portrait, width < height):
 *     three disposable copies, 3rd arg true/false/omitted, all three identical -- size applies,
 *     w:orient stays "portrait", persists across reopen (docs/evidence/
 *     docx-w11-1c-pagesize-orient-20261007.json). NOT MEASURED: the landscape case (width > height)
 *     -- v1 therefore REFUSES a `size` with width > height at plan time (never silently pass an
 *     unmeasured shape through).
 *   - `margins`: SetPageMargins's measured ARGUMENT ORDER is (left, top, right, bottom) -- NOT the
 *     task field order (top, right, bottom, left). Measured with four DISTINCT values
 *     (1001/1102/1203/1304) specifically to discriminate the order (the first W11-1b evidence used
 *     an identical value at all four positions and could not): SetPageMargins(1001, 1102, 1203,
 *     1304) -> left=1001, top=1102, right=1203, bottom=1304, confirmed via the section getters
 *     (session + reopen) AND the saved <w:pgMar> attributes (docs/evidence/
 *     docx-w11-1d-margins-order-20261007.json). lower() below reorders the task's
 *     {top,right,bottom,left} fields into the call's (left,top,right,bottom) positions.
 *
 * Shape follows families/header-footer.cjs (section-global target, no paragraph selector):
 * section: { index, of } required and explicit (header-footer.cjs:86-89's own validation, copied).
 */

const FIELDS = Object.freeze(['size', 'margins', 'columns', 'startPageNumber']);
const MARGIN_KEYS = Object.freeze(['top', 'right', 'bottom', 'left']);
// set_page_setup is a NEW_INTENTS capability (capabilities.cjs), not one of the 28 migrated
// intents -- it is deliberately NOT added to field-contract.cjs's STEP_KEYS, because that table's
// own acceptance test (test/acceptance/docx-m5-field-contract.test.cjs: "every registry intent has
// an accepted field set") asserts Object.keys(STEP_KEYS) equals the 28-intent registry 1:1; a 29th
// key there would break that gate, the same class of problem capabilities.cjs's own NEW_INTENTS
// split already solved for intents.cjs. field-contract.cjs's own fail-closed-unknown-field
// protection is therefore re-implemented HERE, locally, so it is not silently absent for this
// family's steps (field-contract.cjs's own comment: an intent outside STEP_KEYS is passed through
// with NO field check at all).
const STEP_KEYS = Object.freeze(['intent', 'id', 'dependsOn', 'section', ...FIELDS]);

const isPosInt = v => Number.isInteger(v) && v > 0;
const isNonNegInt = v => Number.isInteger(v) && v >= 0;

function validateSize(size, index) {
  if (!size || typeof size !== 'object' || Object.keys(size).sort().join(',') !== 'height,width') {
    return { ok: false, detail: `step ${index}: set_page_setup size requires exactly { width, height }` };
  }
  if (!isPosInt(size.width) || !isPosInt(size.height)) {
    return { ok: false, detail: `step ${index}: set_page_setup size.width/height must be a positive integer (twips)` };
  }
  // landscape (width > height) is NOT measured -- refused rather than silently passed through.
  if (size.width > size.height) {
    return { ok: false, detail: `step ${index}: set_page_setup size with width > height (landscape) is not yet measured -- portrait (width <= height) only` };
  }
  return { ok: true };
}

function validateMargins(margins, index) {
  if (!margins || typeof margins !== 'object' || Object.keys(margins).sort().join(',') !== MARGIN_KEYS.slice().sort().join(',')) {
    return { ok: false, detail: `step ${index}: set_page_setup margins requires all four of ${MARGIN_KEYS.join(', ')}` };
  }
  for (const k of MARGIN_KEYS) {
    if (!isNonNegInt(margins[k])) return { ok: false, detail: `step ${index}: set_page_setup margins.${k} must be a non-negative integer (twips)` };
  }
  return { ok: true };
}

function validateColumns(columns, index) {
  if (!columns || typeof columns !== 'object' || Object.keys(columns).sort().join(',') !== 'count,space') {
    return { ok: false, detail: `step ${index}: set_page_setup columns requires exactly { count, space }` };
  }
  if (!isPosInt(columns.count)) return { ok: false, detail: `step ${index}: set_page_setup columns.count must be a positive integer` };
  if (!isNonNegInt(columns.space)) return { ok: false, detail: `step ${index}: set_page_setup columns.space must be a non-negative integer (twips)` };
  return { ok: true };
}

function validateStartPageNumber(n, index) {
  if (!isPosInt(n)) return { ok: false, detail: `step ${index}: set_page_setup startPageNumber must be a positive integer` };
  return { ok: true };
}

function plan(step, index) {
  if (step.intent !== 'set_page_setup') return { ok: false, detail: `step ${index}: ${step.intent} is not a page-setup intent` };
  // no `selector` in STEP_KEYS at all -- a caller sending one is caught here, as an unknown field
  // (unlike header-footer.cjs, which accepts-then-refuses selector with its own reason, since
  // set_page_setup never had a source intent that took one).
  for (const key of Object.keys(step)) {
    if (!STEP_KEYS.includes(key)) return { ok: false, detail: `field-unsupported: step ${index} (set_page_setup).${key}` };
  }
  const s = step.section;
  if (!s || typeof s !== 'object' || Object.keys(s).sort().join(',') !== 'index,of' || !Number.isInteger(s.index) || !Number.isInteger(s.of) || s.of < 1 || s.index < 0 || s.index >= s.of) {
    return { ok: false, detail: `step ${index}: set_page_setup requires section: { index, of } (integers, 0 <= index < of = the section count) -- no implicit section` };
  }
  const present = FIELDS.filter(f => step[f] !== undefined);
  if (present.length === 0) return { ok: false, detail: `step ${index}: set_page_setup requires at least one of ${FIELDS.join(', ')}` };
  const fields = {};
  if (step.size !== undefined) {
    const v = validateSize(step.size, index); if (!v.ok) return v;
    fields.size = { width: step.size.width, height: step.size.height };
  }
  if (step.margins !== undefined) {
    const v = validateMargins(step.margins, index); if (!v.ok) return v;
    fields.margins = MARGIN_KEYS.reduce((o, k) => ({ ...o, [k]: step.margins[k] }), {});
  }
  if (step.columns !== undefined) {
    const v = validateColumns(step.columns, index); if (!v.ok) return v;
    fields.columns = { count: step.columns.count, space: step.columns.space };
  }
  if (step.startPageNumber !== undefined) {
    const v = validateStartPageNumber(step.startPageNumber, index); if (!v.ok) return v;
    fields.startPageNumber = step.startPageNumber;
  }
  const selector = Object.freeze({ section: Object.freeze({ index: s.index, of: s.of }) });
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: step.intent, selector,
      sectionIndex: s.index, sectionCount: s.of, fields: Object.freeze(fields),
      preservationMode: 'PRIMITIVE_INTEGRAL',
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

// ---- the observation ----------------------------------------------------------------------------

const sectionAt = (sections, i) => sections[i] || null;

/**
 * Equal-column spacing: GetColumnsSpaces() returns count-1 gap widths (count equal columns have
 * count-1 gaps between them). W11-4: this IS independently re-readable -- confirmed live,
 * session + two reopens (docs/evidence/docx-w11-1b-section-write-20261007.json R2/R3). The earlier
 * belief that only GetColumnsWidths (a DERIVED value, recomputed from size/margins) existed was
 * wrong -- GetColumnsCount's own sibling array getter was simply never checked.
 */
function spacesMatch(sec, wantSpace, wantCount) {
  const arr = sec && sec.columnsSpaces;
  if (!Array.isArray(arr)) return false;
  return arr.length === Math.max(0, wantCount - 1) && arr.every(v => v === wantSpace);
}

/** Which of the requested fields already match, reading each via its OWN getter -- never the write args. */
function fieldMatches(sec, field, want) {
  if (!sec) return false;
  if (field === 'size') return sec.pageWidth === want.width && sec.pageHeight === want.height;
  if (field === 'margins') return MARGIN_KEYS.every(k => sec.margins && sec.margins[k] === want[k]);
  if (field === 'columns') return sec.columnsCount === want.count && spacesMatch(sec, want.space, want.count);
  if (field === 'startPageNumber') return sec.startPageNumber === want;
  return false;
}

const isUnobservable = v => typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0;

/** Which requested fields currently read as UNOBSERVABLE (a missing/throwing getter, never a guessed number). */
function unobservableFields(sec, fields) {
  const out = [];
  for (const f of fields) {
    if (f === 'size' && (isUnobservable(sec.pageWidth) || isUnobservable(sec.pageHeight))) out.push(f);
    if (f === 'margins' && MARGIN_KEYS.some(k => isUnobservable(sec.margins && sec.margins[k]))) out.push(f);
    if (f === 'columns' && (isUnobservable(sec.columnsCount) || isUnobservable(sec.columnsSpaces))) out.push(f);
    if (f === 'startPageNumber' && isUnobservable(sec.startPageNumber)) out.push(f);
  }
  return out;
}

function sectionState(primary, op) {
  const d = primary && primary.document;
  const sections = d ? d.sections : null;
  if (!Array.isArray(sections)) return { ok: false, reason: 'sections-unreadable' };
  if (sections.length !== op.sectionCount) return { ok: false, reason: 'FAIL_CLOSED:section-descriptor-mismatch', sectionCount: sections.length };
  const sec = sectionAt(sections, op.sectionIndex);
  const fields = Object.keys(op.fields);
  const matches = fields.map(f => fieldMatches(sec, f, op.fields[f]));
  const done = sec !== null && matches.length > 0 && matches.every(Boolean);
  const unobservable = sec ? unobservableFields(sec, fields) : [];
  return { ok: true, sections, sec, done, fields, unobservable };
}

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  const s = sectionState(primary, op);
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason, sectionCount: s.sectionCount } };
  if (s.done) {
    if (op.recovery === UNRESOLVED) return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } };
    return { state: 'SATISFIED', evidence: { section: op.sectionIndex, fields: op.fields, proof: 'fresh-live-section-matches-every-requested-field' } };
  }
  if (!s.sec) return { state: 'FAIL_CLOSED', evidence: { reason: 'section-unreadable' } };
  if (s.unobservable.length) return { state: 'FAIL_CLOSED', evidence: { reason: `FAIL_CLOSED_UNSUPPORTED_SHAPE:page-setup-field-unobservable:${s.unobservable.join(',')}` } };
  return { state: 'APPLY', evidence: { section: op.sectionIndex, fields: op.fields } };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  if (!op || op.intent !== 'set_page_setup') return { ok: false, target: null, selector, ambiguity: 'INVALID_SELECTOR' };
  const s = sectionState(primary, op);
  const up = r => r.toUpperCase().replace(/-/g, '_');
  if (!s.ok) return { ok: false, target: null, selector, ambiguity: up(s.reason) };
  if (!s.sec) return { ok: false, target: null, selector, ambiguity: 'SECTION_UNREADABLE' };
  if (!s.done && s.unobservable.length) return { ok: false, target: null, selector, ambiguity: `PAGE_SETUP_FIELD_UNOBSERVABLE:${s.unobservable.join(',').toUpperCase()}` };
  return { ok: true, target: { kind: 'section_page_setup', sectionIndex: op.sectionIndex, sectionCount: op.sectionCount, fields: op.fields }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  const same = target && target.kind === 'section_page_setup' && target.sectionIndex === op.sectionIndex && target.sectionCount === op.sectionCount;
  return same ? { ok: true } : { ok: false, dispatched: false, detail: 'page-setup-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({ family: 'pageSetup', action: 'pageSetup', sectionIndex: target.sectionIndex, sectionCount: target.sectionCount, fields: target.fields });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand, public API only. Re-checks the section count and index (header-footer.cjs's own
 * re-check shape). `margins` is REORDERED here from the task's {top,right,bottom,left} into the
 * measured call order (left, top, right, bottom); the task-facing field names never
 * change, only the positional call this function builds.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  const secs = doc.GetSections() || [];
  if (secs.length !== m.sectionCount) return;
  const sec = secs[m.sectionIndex];
  if (!sec) return;
  if (m.fields.size) { const g = m.fields.size; sec.SetPageSize(g.width, g.height); }
  if (m.fields.margins) { const g = m.fields.margins; sec.SetPageMargins(g.left, g.top, g.right, g.bottom); }
  if (m.fields.columns) { const g = m.fields.columns; sec.SetEqualColumns(g.count, g.space); }
  if (m.fields.startPageNumber !== undefined) sec.SetStartPageNumber(m.fields.startPageNumber);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** The main body as the reads see it -- same shape as header-footer.cjs's bodyKey, reused. */
function bodyKey(primary) {
  const d = (primary && primary.document) || {};
  return JSON.stringify({ units: textUnits(primary).map(u => [u.kind, norm(u.text)]), tables: tablesContent(d.tables), footnotes: d.footnotes || null, endnotes: d.endnotes || null, bookmarks: d.bookmarks || null, images: ((primary && primary.images) || []).length, inMainBody: d.inMainBody === undefined ? null : d.inMainBody });
}

/** Every OTHER section's own page setup, unaffected by this op. */
function otherFieldsOfSection(sec) {
  if (!sec) return null;
  return JSON.stringify({ pageWidth: sec.pageWidth, pageHeight: sec.pageHeight, margins: sec.margins, columnsCount: sec.columnsCount, startPageNumber: sec.startPageNumber });
}

function otherSectionsKey(sections, exceptIndex) {
  return JSON.stringify(sections.map((s, i) => (i === exceptIndex ? null : otherFieldsOfSection(s))));
}

/** L1: every requested field now matches, via its own getter; the section count unchanged. */
function verifyPrimitive(op, before, after) {
  const b = sectionState(before && before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const a = sectionState(after && after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const ok = !b.done && a.done;
  return { ok, evidence: { section: op.sectionIndex, fields: op.fields, sectionReadAfter: a.sec } };
}

/**
 * L3, PRIMITIVE_INTEGRAL (same owner placement rule as edit_header/edit_footer -- intents.cjs's own
 * comment on a section-global, no-selector intent: the body and every other part/section must not
 * change, as part of success). Preserved: the body; every OTHER field of the SAME section not
 * requested by this step (the W11-1b write-probe's own 16/16 cross-axis matrix, promoted from a
 * measurement technique into a verification requirement, design doc point 3); every OTHER section's
 * own page setup.
 */
function verifyPreservation(op, before, after) {
  const b = sectionState(before && before.primary, op);
  const a = sectionState(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` } };
  const requested = new Set(Object.keys(op.fields));
  const axisKeys = { size: ['pageWidth', 'pageHeight'], margins: ['margins'], columns: ['columnsCount', 'columnsSpaces'], startPageNumber: ['startPageNumber'] };
  const changedOwnSectionFields = Object.keys(axisKeys).filter(f => !requested.has(f)).filter(f => {
    const keyOf = sec => JSON.stringify(axisKeys[f].map(k => sec && sec[k]));
    return keyOf(b.sec) !== keyOf(a.sec);
  });
  const sameOtherSections = otherSectionsKey(b.sections, op.sectionIndex) === otherSectionsKey(a.sections, op.sectionIndex);
  const sameBody = bodyKey(before.primary) === bodyKey(after.primary);
  const ok = changedOwnSectionFields.length === 0 && sameOtherSections && sameBody;
  return { ok, evidence: { changedOwnSectionFields, otherSectionsUnchanged: sameOtherSections, bodyUnchanged: sameBody, sectionCount: a.sections.length } };
}

/** L2: the final section holds exactly the requested fields; the section count as declared. */
function wholeTaskChecks(op, finalObservation) {
  const s = sectionState(finalObservation.primary, op);
  return [{ name: `${op.id}: section ${op.sectionIndex}/${op.sectionCount} page setup matches ${Object.keys(op.fields).join('+')}`, ok: s.ok && s.done }];
}

// refs: none -- a page-setup value is not a body text unit a later step can name by reference.
function producedText() {
  return null;
}

module.exports = {
  FIELDS, MARGIN_KEYS,
  plan, classify, resolve, preDispatchCheck, lower, commandBody,
  verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  sectionState, fieldMatches, otherFieldsOfSection, otherSectionsKey, bodyKey,
};
