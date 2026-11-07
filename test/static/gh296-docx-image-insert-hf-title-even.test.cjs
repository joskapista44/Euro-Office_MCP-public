'use strict';

/**
 * GH #296 part 3 (offline): insert_image's header/footer target now accepts variant title/even --
 * ONLY on an already-existing, already-active part (image-insert.cjs header comment; LIVE-measured
 * via test/live/docx-gh296-hf-title-even-image-insert-probe.live.cjs). Fabricated
 * normalizeInventory observations, no editor.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const ii = require('../../src/adapters/docx/families/image-insert.cjs');

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const W = 914400;
const H = 457200;

const para = text => ({ text, format: {}, runs: [{ text, bold: false, hasDrawing: false }], inline: [] });
const hfPart = paras => ({ paragraphs: paras.length, simple: true, paras, tables: [] });
/** One section: `parts` maps 'header.title' etc. to an hfPart; flags are the section's variant state. */
function sectionOf(parts, titlePage, evenAndOdd) {
  const slot = kind => ({ default: parts[`${kind}.default`] || null, title: parts[`${kind}.title`] || null, even: parts[`${kind}.even`] || null });
  return { header: slot('header'), footer: slot('footer'), titlePage, evenAndOdd };
}
const imgRow = (part, paragraph) => ({ index: 0, part, drawingType: 'inline', wrapType: null, cx: W, cy: H, paragraph, paragraphIndex: 0 });
const obsOf = (sections, images = []) => normalizeInventory({ html: '<p>Body</p>', document: { images, sections } });

const opOf = (part, variant, index = 0, of = 1) => {
  const p = ii.plan({ intent: 'insert_image', selector: { section: { index, of }, part, variant }, source: PNG, widthEmu: W, heightEmu: H }, 0);
  assert.equal(p.ok, true, p.detail);
  return p.operation;
};

test('plan: title and even are accepted, an unknown variant is refused', () => {
  assert.equal(opOf('header', 'title').variant, 'title');
  assert.equal(opOf('footer', 'even').variant, 'even');
  const odd = ii.plan({ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header', variant: 'odd' }, source: PNG, widthEmu: W, heightEmu: H }, 0);
  assert.equal(odd.ok, false);
});

test('classify: an existing, active title header -> APPLY; the same image already there -> SATISFIED', () => {
  const sections = [sectionOf({ 'header.title': hfPart([para('T')]) }, true, false)];
  assert.equal(ii.classify(obsOf(sections), opOf('header', 'title')).state, 'APPLY');
  const done = ii.classify(obsOf(sections, [imgRow('header:0:title', 'T')]), opOf('header', 'title'));
  assert.equal(done.state, 'SATISFIED');
  // a same-size image in the DEFAULT header is not a match for the title target
  assert.equal(ii.classify(obsOf(sections, [imgRow('header:0:default', 'T')]), opOf('header', 'title')).state, 'APPLY');
});

test('classify: an inactive variant is FAIL_CLOSED, never activated', () => {
  const titleOff = [sectionOf({ 'header.title': hfPart([para('T')]) }, false, false)];
  const c = ii.classify(obsOf(titleOff), opOf('header', 'title'));
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'insert-image-title-variant-inactive');
  const evenOff = [sectionOf({ 'footer.even': hfPart([para('E')]) }, false, false)];
  assert.equal(ii.classify(obsOf(evenOff), opOf('footer', 'even')).evidence.reason, 'insert-image-even-variant-inactive');
});

test('classify: even is document-wide -- one section with evenAndOdd off makes it inactive', () => {
  const sections = [sectionOf({ 'footer.even': hfPart([para('E0')]) }, false, true), sectionOf({}, false, false)];
  assert.equal(ii.classify(obsOf(sections), opOf('footer', 'even', 0, 2)).evidence.reason, 'insert-image-even-variant-inactive');
  const allOn = [sectionOf({ 'footer.even': hfPart([para('E0')]) }, false, true), sectionOf({}, false, true)];
  assert.equal(ii.classify(obsOf(allOn), opOf('footer', 'even', 0, 2)).state, 'APPLY');
});

test('classify: an unknown variant state on ANY section is FAIL_CLOSED; default is unaffected', () => {
  const sections = [sectionOf({ 'header.title': hfPart([para('T')]), 'header.default': hfPart([para('D')]) }, true, null)];
  assert.equal(ii.classify(obsOf(sections), opOf('header', 'title')).evidence.reason, 'section-variant-state-unavailable-section-0');
  assert.equal(ii.classify(obsOf(sections), opOf('header', 'default')).state, 'APPLY', 'default never reads the variant state');
});

test('classify/resolve: an active title variant whose part does not exist is still FAIL_CLOSED (never created)', () => {
  const sections = [sectionOf({}, true, false)];
  assert.equal(ii.classify(obsOf(sections), opOf('header', 'title')).evidence.reason, 'insert-image-header-footer-part-missing');
  const r = ii.resolve(obsOf(sections), opOf('header', 'title').selector);
  assert.equal(r.ok, false);
  assert.equal(r.ambiguity, 'INSERT_IMAGE_HEADER_FOOTER_PART_MISSING');
  const inactive = ii.resolve(obsOf([sectionOf({ 'header.title': hfPart([para('T')]) }, false, false)]), opOf('header', 'title').selector);
  assert.equal(inactive.ambiguity, 'INSERT_IMAGE_TITLE_VARIANT_INACTIVE');
});

test('lower: title/even set meta.bridgeSectionVariantRead and the body re-checks the flag; default does not', () => {
  const sections = [sectionOf({ 'header.title': hfPart([para('T')]), 'header.default': hfPart([para('D')]) }, true, false)];
  const t = opOf('header', 'title');
  const r = ii.resolve(obsOf(sections), t.selector);
  assert.equal(r.ok, true);
  const low = ii.lower(t, r.target);
  assert.equal(low.meta.bridgeSectionVariantRead, true);
  assert.equal(low.meta.variant, 'title');
  assert.match(low.body, /__eoB21Variant\(m\.sectionIndex\)/);
  const d = opOf('header', 'default');
  const lowD = ii.lower(d, ii.resolve(obsOf(sections), d.selector).target);
  assert.equal(lowD.meta.bridgeSectionVariantRead, undefined, 'default stays byte-identical in meta (no bridge prelude)');
});
