'use strict';

/**
 * GH #306b 2. resz (write side, offline): the `insert_image` family's HEADER/FOOTER target
 * (src/adapters/docx/families/image-insert.cjs). plan/classify/resolve/verify against fabricated
 * `normalizeInventory` observations -- no editor, no VM. The commandBody itself is covered by the
 * dedicated LIVE product harness (zero runs it), not here.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const ii = require('../../src/adapters/docx/families/image-insert.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { check } = require('../../src/server/write-tools.cjs');

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const W = 914400;
const H = 457200;

// GH #306b's own part-join label format (docx-images-read.cjs: 'header:' + si + ':' + v).
const label = (part, i, v = 'default') => `${part}:${i}:${v}`;

const para = (text, inline = []) => ({ text, format: {}, runs: [{ text, bold: false, hasDrawing: false }], inline });
/** GH #317 step 4: a paragraph whose FIRST run carries the drawing (text '') and SECOND the text --
 * the SHAPE `position: 'start'` produces, read through read_headers_footers' own per-run
 * `hasDrawing` BOOLEAN (docx-driver.cjs, added alongside this candidate -- a PRIMITIVE; the
 * inline-type array this started as wiped the whole callCommand result LIVE). */
const paraStart = (text, paraInline = ['blipFill', 'image', 'paraDrawing']) => ({
  text, format: {}, runs: [{ text: '', bold: false, hasDrawing: true }, { text, bold: false, hasDrawing: false }], inline: paraInline,
});
const hfPart = (paras, over = {}) => ({ paragraphs: paras.length, simple: true, paras, tables: [], ...over });

/** One section's header/footer slots: hdr/ftr are null (missing) or the result of hfPart(). */
const sectionOf = (hdr, ftr) => ({ header: { default: hdr, even: null, title: null }, footer: { default: ftr, even: null, title: null } });

// GH #317 follow-up: `paragraphIndex` (the part-walk index) is the header/footer retry identity; default 0 = the first/only paragraph.
const imgRow = (part, paragraph, cx = W, cy = H, drawingType = 'inline', paragraphIndex = 0) => ({ index: 0, part, drawingType, wrapType: drawingType === 'inline' ? null : 'square', cx, cy, paragraph, paragraphIndex });

function obsOf(sections, images = []) {
  return { primary: normalizeInventory({ html: '<p>Body</p>', document: { images, sections } }) };
}

const op = (overrides = {}) => ii.plan({ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header' }, source: PNG, widthEmu: W, heightEmu: H, ...overrides }, 0).operation;

// ---- plan -----------------------------------------------------------------------------------------

test('plan (header/footer target): accepts { section, part }, variant defaults to \'default\'', () => {
  const p = ii.plan({ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header' }, source: PNG, widthEmu: W, heightEmu: H }, 0);
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operation.preservationMode, 'PRIMITIVE_INTEGRAL');
  assert.deepEqual(p.operation.selector, { section: { index: 0, of: 1 }, part: 'header', variant: 'default', position: 'end' });
  assert.equal(p.operation.sectionIndex, 0);
  assert.equal(p.operation.sectionCount, 1);
  assert.equal(p.operation.part, 'header');
  assert.equal(p.operation.variant, 'default');
  const explicit = ii.plan({ intent: 'insert_image', selector: { section: { index: 1, of: 2 }, part: 'footer', variant: 'default' }, source: PNG, widthEmu: W, heightEmu: H }, 0);
  assert.equal(explicit.ok, true, explicit.detail);
});

test('plan (header/footer target): refuses a bad section/part/variant, a position field, and an out-of-range index', () => {
  const base = { intent: 'insert_image', source: PNG, widthEmu: W, heightEmu: H };
  assert.equal(ii.plan({ ...base, selector: { section: { index: 0, of: 1 }, part: 'body' } }, 0).ok, false, 'part must be header/footer');
  assert.equal(ii.plan({ ...base, selector: { section: { index: 0, of: 1 }, part: 'header', variant: 'odd' } }, 0).ok, false, 'variant must be default/title/even (GH #296 part 3 accepts title/even)');
  assert.equal(ii.plan({ ...base, selector: { section: { index: 1, of: 1 }, part: 'header' } }, 0).ok, false, 'index out of range');
  assert.equal(ii.plan({ ...base, selector: { section: { index: 0, of: 0 }, part: 'header' } }, 0).ok, false, 'of must be >= 1');
  assert.equal(ii.plan({ ...base, selector: { part: 'header' } }, 0).ok, false, 'section is required');
  assert.equal(ii.plan({ ...base, selector: { section: { index: 0, of: 1 }, part: 'header' }, position: 'end' }, 0).ok, false, 'no position field on this target');
});

// ---- classify / resolve -----------------------------------------------------------------------------

test('classify (header/footer): FAIL_CLOSED when the section count mismatches or the part is missing', () => {
  assert.equal(ii.classify(obsOf([]).primary, op()).state, 'FAIL_CLOSED', 'section count mismatch');
  assert.equal(ii.classify(obsOf([sectionOf(null, null)]).primary, op()).state, 'FAIL_CLOSED', 'header missing');
});

test('classify (header/footer): FAIL_CLOSED when the part is not plain paragraphs of runs', () => {
  const sections = [sectionOf(hfPart([para('HA')], { simple: false }), null)];
  assert.equal(ii.classify(obsOf(sections).primary, op()).state, 'FAIL_CLOSED');
});

test('classify (header/footer): APPLY when the part exists but carries no matching image', () => {
  const sections = [sectionOf(hfPart([para('HA')]), null)];
  assert.equal(ii.classify(obsOf(sections).primary, op()).state, 'APPLY');
  const sections2 = [sectionOf(hfPart([para('HA')]), null)];
  assert.equal(ii.classify(obsOf(sections2, [imgRow(label('header', 0), 'HA', 1, 1)]).primary, op()).state, 'APPLY', 'wrong size is still APPLY');
  assert.equal(ii.classify(obsOf(sections2, [imgRow(label('footer', 0), 'HA')]).primary, op()).state, 'APPLY', 'right size+paragraph but WRONG part label is still APPLY');
});

test('classify (header/footer): SATISFIED when the part\'s own last paragraph carries a part-labeled matching inline image', () => {
  const sections = [sectionOf(hfPart([para('HA')]), null)];
  const c = ii.classify(obsOf(sections, [imgRow(label('header', 0), 'HA')]).primary, op());
  assert.equal(c.state, 'SATISFIED');
  // the footer target, a DIFFERENT section index
  const sections2 = [sectionOf(null, hfPart([para('FA')])), sectionOf(null, hfPart([para('FB')]))];
  const footerOp = op({ selector: { section: { index: 1, of: 2 }, part: 'footer' } });
  assert.equal(ii.classify(obsOf(sections2, [imgRow(label('footer', 1), 'FB')]).primary, footerOp).state, 'SATISFIED');
});

test('resolve / preDispatchCheck (header/footer): round-trip through the real planOperations pipeline', () => {
  const planned = planOperations({ steps: [{ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header' }, source: PNG, widthEmu: W, heightEmu: H }] });
  assert.equal(planned.ok, true, planned.detail);
  const o = planned.operations[0];
  const sections = [sectionOf(hfPart([para('HA')]), null)];
  const r = ii.resolve(obsOf(sections).primary, o.selector);
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.target.kind, 'insert_image_header_footer_target');
  assert.equal(ii.preDispatchCheck(o, r.target).ok, true);
  assert.equal(ii.preDispatchCheck(o, { kind: 'insert_image_header_footer_target', sectionIndex: 0, sectionCount: 1, part: 'footer', variant: 'default' }).ok, false);
});

test('resolve (header/footer): refuses when the part is missing or not plain paragraphs', () => {
  const missing = ii.resolve(obsOf([sectionOf(null, null)]).primary, { section: { index: 0, of: 1 }, part: 'header', variant: 'default' });
  assert.equal(missing.ok, false);
  assert.equal(missing.ambiguity, 'INSERT_IMAGE_HEADER_FOOTER_PART_MISSING');
  const notSimple = ii.resolve(obsOf([sectionOf(hfPart([para('HA')], { simple: false }), null)]).primary, { section: { index: 0, of: 1 }, part: 'header', variant: 'default' });
  assert.equal(notSimple.ok, false);
  assert.equal(notSimple.ambiguity, 'HEADER_FOOTER_NOT_PLAIN_PARAGRAPHS_OF_RUNS');
});

// ---- verify ---------------------------------------------------------------------------------------

test('verifyPrimitive (header/footer): ok only when the part-labeled matching image AND the read_headers_footers inline marker are both present', () => {
  const sections = [sectionOf(hfPart([para('HA')]), null)];
  const before = obsOf(sections);
  const good = ii.verifyPrimitive(op(), before, obsOf([sectionOf(hfPart([para('HA', ['blipFill', 'image', 'paraDrawing'])]), null)], [imgRow(label('header', 0), 'HA')]));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  const noImage = ii.verifyPrimitive(op(), before, before);
  assert.equal(noImage.ok, false, 'no image row at all');
  const noInlineMarker = ii.verifyPrimitive(op(), before, obsOf([sectionOf(hfPart([para('HA')]), null)], [imgRow(label('header', 0), 'HA')]));
  assert.equal(noInlineMarker.ok, false, 'image row matches but read_headers_footers inline marker is absent');
});

test('verifyPreservation (header/footer): OK when every prior image (by part+shape) survives, the body is unchanged, the OTHER part is unchanged, and the target part\'s own paragraphs/text are unchanged', () => {
  const beforeSections = [sectionOf(hfPart([para('HA')]), hfPart([para('FA')]))];
  const before = obsOf(beforeSections, [imgRow(label('footer', 0), 'FA', 100, 100)]);
  const afterSections = [sectionOf(hfPart([para('HA', ['blipFill', 'image', 'paraDrawing'])]), hfPart([para('FA')]))];
  const after = obsOf(afterSections, [imgRow(label('footer', 0), 'FA', 100, 100), imgRow(label('header', 0), 'HA')]);
  const good = ii.verifyPreservation(op(), before, after);
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

test('verifyPreservation (header/footer): FAILS if an existing image (any part) is lost, if the count does not grow by exactly one, if the OTHER part changes, or if the target part\'s text changes', () => {
  const beforeSections = [sectionOf(hfPart([para('HA')]), hfPart([para('FA')]))];
  const before = obsOf(beforeSections, [imgRow(label('footer', 0), 'FA', 100, 100)]);

  const lostFooterImage = ii.verifyPreservation(op(), before, obsOf([sectionOf(hfPart([para('HA')]), hfPart([para('FA')]))], [imgRow(label('header', 0), 'HA')]));
  assert.equal(lostFooterImage.ok, false, 'the footer image vanished');

  const noNew = ii.verifyPreservation(op(), before, obsOf([sectionOf(hfPart([para('HA')]), hfPart([para('FA')]))], [imgRow(label('footer', 0), 'FA', 100, 100)]));
  assert.equal(noNew.ok, false, 'count did not grow');

  const footerChanged = ii.verifyPreservation(op(), before, obsOf([sectionOf(hfPart([para('HA')]), hfPart([para('FA CHANGED')]))], [imgRow(label('footer', 0), 'FA', 100, 100), imgRow(label('header', 0), 'HA')]));
  assert.equal(footerChanged.ok, false, 'the OTHER (footer) part text changed');

  const headerTextChanged = ii.verifyPreservation(op(), before, obsOf([sectionOf(hfPart([para('HA CHANGED')]), hfPart([para('FA')]))], [imgRow(label('footer', 0), 'FA', 100, 100), imgRow(label('header', 0), 'HA CHANGED')]));
  assert.equal(headerTextChanged.ok, false, 'the target part\'s own paragraph text changed -- only the drawing may grow');
});

test('verifyPreservation (header/footer): a same-shape BODY image is never confused with the new header image (part is in the key)', () => {
  const beforeSections = [sectionOf(hfPart([para('HA')]), null)];
  const before = obsOf(beforeSections, [imgRow('body', 'HA', W, H)]); // a body image, coincidentally same paragraph text + size
  const after = obsOf([sectionOf(hfPart([para('HA')]), null)], [imgRow('body', 'HA', W, H), imgRow(label('header', 0), 'HA')]);
  const good = ii.verifyPreservation(op(), before, after);
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

test('wholeTaskChecks (header/footer): one named check, true only on a matching final image', () => {
  const sections = [sectionOf(hfPart([para('HA')]), null)];
  const ok = ii.wholeTaskChecks(op(), obsOf(sections, [imgRow(label('header', 0), 'HA')]));
  assert.equal(ok.length, 1);
  assert.equal(ok[0].ok, true);
  const fail = ii.wholeTaskChecks(op(), obsOf(sections));
  assert.equal(fail[0].ok, false);
});

// ---- GH #317 (internal tracker, step 3): selector.paragraph {index, of} addressing --------------------
// Targets a SPECIFIC paragraph instead of always the part's own last one. `of` is the caller's OWN
// SEEN paragraph count; a mismatch at dispatch time is FAIL_CLOSED, never a silent fallback.

test('plan (header/footer target): accepts selector.paragraph {index, of}; refuses a malformed one', () => {
  const ok = ii.plan({ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 1, of: 3 } }, source: PNG, widthEmu: W, heightEmu: H }, 0);
  assert.equal(ok.ok, true, ok.detail);
  assert.deepEqual(ok.operation.selector.paragraph, { index: 1, of: 3 });
  assert.equal(ok.operation.paragraphIndex, 1);
  assert.equal(ok.operation.paragraphCount, 3);
  const base = { intent: 'insert_image', source: PNG, widthEmu: W, heightEmu: H };
  assert.equal(ii.plan({ ...base, selector: { section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 3, of: 3 } } }, 0).ok, false, 'index must be < of');
  assert.equal(ii.plan({ ...base, selector: { section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 0, of: 0 } } }, 0).ok, false, 'of must be >= 1');
  assert.equal(ii.plan({ ...base, selector: { section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 0 } } }, 0).ok, false, 'of is required');
});

test('classify (header/footer, named paragraph): FAIL_CLOSED on a paragraph-descriptor mismatch; APPLY/SATISFIED target the NAMED paragraph, not the last one', () => {
  const sections = [sectionOf(hfPart([para('P0'), para('P1'), para('P2')]), null)];
  const named = op({ selector: { section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 0, of: 3 } } });
  // the descriptor's `of` no longer matches the live paragraph count
  const wrongCount = op({ selector: { section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 0, of: 5 } } });
  assert.equal(ii.classify(obsOf(sections).primary, wrongCount).state, 'FAIL_CLOSED');
  assert.equal(ii.classify(obsOf(sections).primary, wrongCount).evidence.reason, 'FAIL_CLOSED:paragraph-descriptor-mismatch');
  // no image anywhere -> APPLY, naming paragraph 0 (NOT "the last paragraph", which would be P2)
  const applyState = ii.classify(obsOf(sections).primary, named);
  assert.equal(applyState.state, 'APPLY');
  assert.equal(applyState.evidence.paragraph, 0);
  // an image sitting on the LAST paragraph (P2) does NOT satisfy a request naming paragraph 0
  const onLast = ii.classify(obsOf(sections, [imgRow(label('header', 0), 'P2', W, H, 'inline', 2)]).primary, named);
  assert.equal(onLast.state, 'APPLY', 'the image is on P2, not the named P0 -- still APPLY');
  // an image sitting on the NAMED paragraph (P0) satisfies it, even though it is not the last one
  const onNamed = ii.classify(obsOf(sections, [imgRow(label('header', 0), 'P0')]).primary, named);
  assert.equal(onNamed.state, 'SATISFIED');
  assert.equal(onNamed.evidence.proof, 'fresh-live-image-at-named-paragraph');
});

test('resolve / preDispatchCheck (header/footer, named paragraph): round-trip through the real planOperations pipeline; a target naming a DIFFERENT paragraph is unresolved', () => {
  const planned = planOperations({ steps: [{ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 1, of: 2 } }, source: PNG, widthEmu: W, heightEmu: H }] });
  assert.equal(planned.ok, true, planned.detail);
  const o = planned.operations[0];
  const sections = [sectionOf(hfPart([para('P0'), para('P1')]), null)];
  const r = ii.resolve(obsOf(sections).primary, o.selector);
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.target.paragraphIndex, 1);
  assert.equal(r.target.paragraphCount, 2);
  assert.equal(ii.preDispatchCheck(o, r.target).ok, true);
  const wrongTarget = { kind: 'insert_image_header_footer_target', sectionIndex: 0, sectionCount: 1, part: 'header', variant: 'default', paragraphIndex: 0, paragraphCount: 2 };
  assert.equal(ii.preDispatchCheck(o, wrongTarget).ok, false, 'resolved a DIFFERENT paragraph than the plan named');
  const unaddressedTarget = { kind: 'insert_image_header_footer_target', sectionIndex: 0, sectionCount: 1, part: 'header', variant: 'default' };
  assert.equal(ii.preDispatchCheck(o, unaddressedTarget).ok, false, 'an un-addressed target does not satisfy a named-paragraph plan');
});

test('resolve (header/footer, named paragraph): refuses when the descriptor\'s `of` no longer matches the live count', () => {
  const sections = [sectionOf(hfPart([para('P0'), para('P1'), para('P2')]), null)];
  const r = ii.resolve(obsOf(sections).primary, { section: { index: 0, of: 1 }, part: 'header', variant: 'default', paragraph: { index: 0, of: 5 } });
  assert.equal(r.ok, false);
  assert.equal(r.ambiguity, 'PARAGRAPH_DESCRIPTOR_MISMATCH');
});

test('verifyPrimitive (header/footer, named paragraph): checks the NAMED paragraph\'s own inline marker, not the last paragraph\'s', () => {
  const named = op({ selector: { section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 0, of: 2 } } });
  const before = obsOf([sectionOf(hfPart([para('P0'), para('P1')]), null)]);
  // the image landed on P0 (the named paragraph) -- ok
  const good = ii.verifyPrimitive(named, before, obsOf([sectionOf(hfPart([para('P0', ['blipFill', 'image', 'paraDrawing']), para('P1')]), null)], [imgRow(label('header', 0), 'P0')]));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  // an inline marker on P1 (the OTHER paragraph, which happens to be the part's last) must NOT count
  const wrongParagraph = ii.verifyPrimitive(named, before, obsOf([sectionOf(hfPart([para('P0'), para('P1', ['blipFill', 'image', 'paraDrawing'])]), null)], [imgRow(label('header', 0), 'P1', W, H, 'inline', 1)]));
  assert.equal(wrongParagraph.ok, false, 'the marker and the matching image are both on the WRONG (last, not named) paragraph');
});

test('wholeTaskChecks (header/footer, named paragraph): names the paragraph index in its own check name', () => {
  const named = op({ selector: { section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 0, of: 2 } } });
  const sections = [sectionOf(hfPart([para('P0'), para('P1')]), null)];
  const ok = ii.wholeTaskChecks(named, obsOf(sections, [imgRow(label('header', 0), 'P0')]));
  assert.match(ok[0].name, /paragraph 0\/2/);
  assert.equal(ok[0].ok, true);
});

// ---- GH #317 (internal tracker, step 4): selector.position 'start'|'end' ------------------------------
// 'end' (default) is the pre-existing, unchanged behaviour; 'start' inserts a fresh run with the
// image BEFORE the paragraph's existing runs. LIVE-measured, docx-gh317-image-start-of-paragraph-
// probe.live.cjs (comment #ref): AddElement(run, 0) works, AddElement(0, run) is a silent no-op.

test('plan (header/footer target): accepts selector.position \'start\'|\'end\'; defaults to \'end\'; refuses a bad value and a step-level position', () => {
  const okStart = ii.plan({ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header', position: 'start' }, source: PNG, widthEmu: W, heightEmu: H }, 0);
  assert.equal(okStart.ok, true, okStart.detail);
  assert.equal(okStart.operation.selector.position, 'start');
  assert.equal(okStart.operation.position, 'start');
  const defaulted = ii.plan({ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header' }, source: PNG, widthEmu: W, heightEmu: H }, 0);
  assert.equal(defaulted.operation.selector.position, 'end');
  const base = { intent: 'insert_image', source: PNG, widthEmu: W, heightEmu: H };
  assert.equal(ii.plan({ ...base, selector: { section: { index: 0, of: 1 }, part: 'header', position: 'after' } }, 0).ok, false, '\'after\' is the BODY target\'s own value, not this one');
  assert.equal(ii.plan({ ...base, selector: { section: { index: 0, of: 1 }, part: 'header' }, position: 'start' }, 0).ok, false, 'step-level position is still the body target\'s field, not this one');
});

test('classify (header/footer, position start): SATISFIED only when the FIRST run carries the drawing, not merely when one exists anywhere in the paragraph', () => {
  const startOp = op({ selector: { section: { index: 0, of: 1 }, part: 'header', position: 'start' } });
  const sections = [sectionOf(hfPart([paraStart('HA')]), null)];
  const satisfied = ii.classify(obsOf(sections, [imgRow(label('header', 0), 'HA')]).primary, startOp);
  assert.equal(satisfied.state, 'SATISFIED');
  // the SAME part+paragraph+size match, but the drawing sits on the LAST run (an 'end'-shaped
  // paragraph) -- must NOT satisfy a 'start' request (the gap zero named, #ref: the pre-existing
  // part+paragraph-text+size check alone cannot tell the two apart)
  const endShaped = [sectionOf(hfPart([para('HA', ['blipFill', 'image', 'paraDrawing'])]), null)];
  const notSatisfied = ii.classify(obsOf(endShaped, [imgRow(label('header', 0), 'HA')]).primary, startOp);
  assert.equal(notSatisfied.state, 'APPLY', 'the image exists, but not at the start -- still APPLY for a start request');
});

test('verifyPrimitive (header/footer, position start): ok only when the FIRST run\'s own inline marker shows the drawing', () => {
  const startOp = op({ selector: { section: { index: 0, of: 1 }, part: 'header', position: 'start' } });
  const before = obsOf([sectionOf(hfPart([para('HA')]), null)]);
  const good = ii.verifyPrimitive(startOp, before, obsOf([sectionOf(hfPart([paraStart('HA')]), null)], [imgRow(label('header', 0), 'HA')]));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  const wrongRun = ii.verifyPrimitive(startOp, before, obsOf([sectionOf(hfPart([para('HA', ['blipFill', 'image', 'paraDrawing'])]), null)], [imgRow(label('header', 0), 'HA')]));
  assert.equal(wrongRun.ok, false, 'the paragraph-level inline marker is present, but the FIRST run is not the drawing one');
});

test('runHasDrawing: reads the run\'s own hasDrawing boolean (a primitive), never its text emptiness -- a prior probe\'s own measured flaw', () => {
  assert.equal(ii.runHasDrawing({ text: '', hasDrawing: true }), true);
  assert.equal(ii.runHasDrawing({ text: '', hasDrawing: false }), false, 'an empty run with hasDrawing=false is not a drawing');
  assert.equal(ii.runHasDrawing({ text: 'x', hasDrawing: false }), false);
  assert.equal(ii.runHasDrawing(null), false);
  assert.equal(ii.runHasDrawing(undefined), false);
});

// ---- helpers exported for this test file's own use ------------------------------------------------

test('isHF / hfPartLabel: the discriminator and the label format match docx-images-read.cjs\'s own join labels', () => {
  assert.equal(ii.isHF(op()), true);
  assert.equal(ii.isHF(ii.plan({ intent: 'insert_image', selector: { afterParagraph: 'x' }, source: PNG, widthEmu: W, heightEmu: H }, 0).operation), false);
  assert.equal(ii.hfPartLabel(op()), 'header:0:default');
});

// ---- schema (docx-write-tools.cjs's IMAGE_INSERT_TARGET anyOf) -------------------------------------

const step = (selector, over = {}) => ({ intent: 'insert_image', selector, source: PNG, widthEmu: W, heightEmu: H, ...over });
const schemaProblem = s => check(D.DOCX_STEP_SCHEMAS_ALL.insert_image, s, 'step');

test('schema: both selector shapes validate; a mixed/malformed selector is refused by the anyOf', () => {
  assert.equal(schemaProblem(step({ afterParagraph: 'x' })), null, 'body shape');
  assert.equal(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header' })), null, 'header/footer shape, variant omitted');
  assert.equal(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header', variant: 'default' })), null, 'header/footer shape, variant explicit');
  assert.equal(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header', variant: 'title' })), null, 'GH #296 part 3: title is accepted');
  assert.ok(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header', variant: 'odd' })), 'a variant outside default/title/even is refused');
  assert.ok(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'body' })), 'part must be header/footer');
  assert.ok(schemaProblem(step({ afterParagraph: 'x', section: { index: 0, of: 1 }, part: 'header' })), 'mixing both shapes\' keys matches neither');
  assert.ok(schemaProblem(step({})), 'an empty selector matches neither shape');
});

test('schema (GH #317): selector.paragraph {index, of} validates; a malformed one is refused', () => {
  assert.equal(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 0, of: 2 } })), null);
  assert.ok(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header', paragraph: { index: 0 } })), 'of is required');
  assert.ok(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header', paragraph: { index: -1, of: 2 } })), 'index must be >= 0');
});

test('schema (GH #317 step 4): selector.position \'start\'|\'end\' validates; a bad value is refused', () => {
  assert.equal(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header', position: 'start' })), null);
  assert.equal(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header', position: 'end' })), null);
  assert.ok(schemaProblem(step({ section: { index: 0, of: 1 }, part: 'header', position: 'after' })), '\'after\' is not valid for this shape');
});
