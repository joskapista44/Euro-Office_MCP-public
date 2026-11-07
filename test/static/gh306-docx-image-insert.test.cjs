'use strict';

/**
 * GH #306 (write side, offline): the `insert_image` family (src/adapters/docx/families/image-insert.cjs).
 * plan/classify/resolve/verify against fabricated `normalizeInventory` observations -- no editor, no
 * VM. The commandBody itself (the actual callCommand string) is covered by the LIVE probe
 * (test/live/docx-gh306-insert-image-probe.live.cjs, an internal branch @749a4f607) and by the dedicated
 * LIVE product harness, not here.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const ii = require('../../src/adapters/docx/families/image-insert.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');

const ANCHOR = 'GH306 anchor';
const SIBLING = 'GH306 sibling';
const OTHER = 'GH306 other';
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const JPEG = 'data:image/jpeg;base64,/9j/4AAQSkZJRg==';
const W = 914400;
const H = 457200;

const imgRow = (paragraph, cx = W, cy = H, drawingType = 'inline') => ({ index: 0, drawingType, wrapType: drawingType === 'inline' ? null : 'square', cx, cy, paragraph });

function obsOf(html, images = []) {
  return { primary: normalizeInventory({ html, document: { images } }) };
}

const op = (overrides = {}) => ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H, ...overrides }, 0).operation;

// ---- plan -----------------------------------------------------------------------------------------

test('plan: accepts { afterParagraph }, source, widthEmu/heightEmu, and an optional position (defaults to \'after\')', () => {
  const p = ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H }, 0);
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operation.preservationMode, 'PRIMITIVE_INTEGRAL');
  assert.equal(p.operation.position, 'after');
  const end = ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: JPEG, widthEmu: W, heightEmu: H, position: 'end' }, 0);
  assert.equal(end.ok, true, end.detail);
  assert.equal(end.operation.position, 'end');
  assert.equal(end.operation.source, JPEG);
});

test('plan: refuses a wrong-shaped selector, a non-PNG/JPEG or non-data-URI source, a non-positive/non-integer EMU, and an unmeasured position', () => {
  assert.equal(ii.plan({ intent: 'insert_image', selector: {}, source: PNG, widthEmu: W, heightEmu: H }, 0).ok, false);
  assert.equal(ii.plan({ intent: 'insert_image', selector: { afterParagraph: '' }, source: PNG, widthEmu: W, heightEmu: H }, 0).ok, false);
  assert.equal(ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR, extra: 1 }, source: PNG, widthEmu: W, heightEmu: H }, 0).ok, false);
  for (const bad of ['https://example.com/x.png', 'data:image/gif;base64,abcd', 'not a uri', '']) {
    assert.equal(ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: bad, widthEmu: W, heightEmu: H }, 0).ok, false, bad);
  }
  for (const bad of [0, -1, 1.5, 'x', undefined]) {
    assert.equal(ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: bad, heightEmu: H }, 0).ok, false, `widthEmu=${bad}`);
    assert.equal(ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: bad }, 0).ok, false, `heightEmu=${bad}`);
  }
  assert.equal(ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H, position: 'middle' }, 0).ok, false, 'middle is not measured');
});

// ---- classify / resolve -----------------------------------------------------------------------------

test('classify (position: end): APPLY when the anchor paragraph carries no matching image', () => {
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`;
  assert.equal(ii.classify(obsOf(html).primary, op({ position: 'end' })).state, 'APPLY');
  assert.equal(ii.classify(obsOf(html, [imgRow(ANCHOR, 1, 1)]).primary, op({ position: 'end' })).state, 'APPLY', 'wrong size is still APPLY, not a guessed match');
});

test('classify (position: end): SATISFIED when the anchor paragraph already carries an inline image of the requested size (±1 EMU)', () => {
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`;
  assert.equal(ii.classify(obsOf(html, [imgRow(ANCHOR)]).primary, op({ position: 'end' })).state, 'SATISFIED');
  assert.equal(ii.classify(obsOf(html, [imgRow(ANCHOR, W - 1, H + 1)]).primary, op({ position: 'end' })).state, 'SATISFIED', '±1 EMU drift tolerated');
  assert.equal(ii.classify(obsOf(html, [imgRow(ANCHOR, W, H, 'anchor')]).primary, op({ position: 'end' })).state, 'APPLY', 'a floating image at the same paragraph/size is not a match');
});

test('classify (position: after): APPLY when the next block is not an empty paragraph with a matching image', () => {
  assert.equal(ii.classify(obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p>`).primary, op()).state, 'APPLY', 'anchor is last, nothing follows');
  assert.equal(ii.classify(obsOf(`<p>${ANCHOR}</p><p>${SIBLING}</p>`).primary, op()).state, 'APPLY', 'next block is not empty');
  assert.equal(ii.classify(obsOf(`<p>${ANCHOR}</p><p></p>`, [imgRow('', 1, 1)]).primary, op()).state, 'APPLY', 'empty next block but wrong size');
});

test('classify (position: after): SATISFIED when the block right after the anchor is an empty paragraph carrying a matching inline image', () => {
  const c = ii.classify(obsOf(`<p>${ANCHOR}</p><p></p>`, [imgRow('')]).primary, op());
  assert.equal(c.state, 'SATISFIED');
});

test('classify: FAIL_CLOSED when the anchor is missing or ambiguous, never guessed', () => {
  assert.equal(ii.classify(obsOf(`<p>${OTHER}</p>`).primary, op()).state, 'FAIL_CLOSED');
  assert.equal(ii.classify(obsOf(`<p>${ANCHOR}</p><p>${ANCHOR}</p>`).primary, op()).state, 'FAIL_CLOSED');
});

test('resolve / preDispatchCheck: round-trip through the real planOperations pipeline', () => {
  const planned = planOperations({ steps: [{ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H }] });
  assert.equal(planned.ok, true, planned.detail);
  const o = planned.operations[0];
  const r = ii.resolve(obsOf(`<p>${ANCHOR}</p>`).primary, o.selector);
  assert.equal(r.ok, true);
  assert.equal(r.target.anchorText, ANCHOR);
  assert.equal(ii.preDispatchCheck(o, r.target).ok, true);
  assert.equal(ii.preDispatchCheck(o, { kind: 'insert_image_anchor', anchorText: 'wrong' }).ok, false);
});

// ---- sizeMatches ------------------------------------------------------------------------------------

test('sizeMatches: ±1 EMU is tolerated on both dimensions, ±2 is not, and an UNOBSERVABLE field never matches', () => {
  const o = op();
  assert.equal(ii.sizeMatches({ widthEmu: W, heightEmu: H }, o), true);
  assert.equal(ii.sizeMatches({ widthEmu: W - 1, heightEmu: H + 1 }, o), true);
  assert.equal(ii.sizeMatches({ widthEmu: W - 2, heightEmu: H }, o), false);
  assert.equal(ii.sizeMatches({ widthEmu: 'UNOBSERVABLE:x', heightEmu: H }, o), false);
});

// ---- verify ---------------------------------------------------------------------------------------

test('verifyPrimitive (end): ok only when the after-state carries a matching image at the anchor paragraph', () => {
  const html = `<p>${ANCHOR}</p>`;
  const before = obsOf(html);
  const good = ii.verifyPrimitive(op({ position: 'end' }), before, obsOf(html, [imgRow(ANCHOR)]));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  const bad = ii.verifyPrimitive(op({ position: 'end' }), before, obsOf(html));
  assert.equal(bad.ok, false);
});

test('verifyPrimitive (after): ok only when the block right after the anchor is the empty paragraph carrying the new image', () => {
  const before = obsOf(`<p>${ANCHOR}</p>`);
  const good = ii.verifyPrimitive(op(), before, obsOf(`<p>${ANCHOR}</p><p></p>`, [imgRow('')]));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  const bad = ii.verifyPrimitive(op(), before, obsOf(`<p>${ANCHOR}</p>`));
  assert.equal(bad.ok, false);
});

test('verifyPreservation (end): OK when every prior image survives (by key) and the block sequence is unchanged', () => {
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`;
  const before = obsOf(html, [imgRow(OTHER, 100, 100)]);
  const good = ii.verifyPreservation(op({ position: 'end' }), before, obsOf(html, [imgRow(OTHER, 100, 100), imgRow(ANCHOR)]));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

test('verifyPreservation (end): FAILS if an existing image is lost, if the count does not grow by exactly one, or if a body block changes', () => {
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p>`;
  const before = obsOf(html, [imgRow(OTHER, 100, 100)]);
  const lost = ii.verifyPreservation(op({ position: 'end' }), before, obsOf(html, [imgRow(ANCHOR)]));
  assert.equal(lost.ok, false, 'OTHER\'s image vanished');
  const noNew = ii.verifyPreservation(op({ position: 'end' }), before, obsOf(html, [imgRow(OTHER, 100, 100)]));
  assert.equal(noNew.ok, false, 'count did not grow');
  const changed = ii.verifyPreservation(op({ position: 'end' }), before, obsOf(`<p>${OTHER} CHANGED</p><p>${ANCHOR}</p>`, [imgRow(OTHER, 100, 100), imgRow(ANCHOR)]));
  assert.equal(changed.ok, false, 'an unrelated block text changed');
});

test('verifyPreservation (after): OK when exactly the one new empty paragraph appears right after the anchor and everything else (prefix + rest) is unchanged', () => {
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const good = ii.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`, [imgRow('')]));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

test('verifyPreservation (after): FAILS if the new block is missing, non-empty, or if anything outside it shifted', () => {
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const missing = ii.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`, [imgRow('')]));
  assert.equal(missing.ok, false, 'no new block at all');
  const notEmpty = ii.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>stray text</p><p>${SIBLING}</p>`, [imgRow('stray text')]));
  assert.equal(notEmpty.ok, false, 'the new block carries text, not just the image');
  const shifted = ii.verifyPreservation(op(), before, obsOf(`<p>${OTHER} CHANGED</p><p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`, [imgRow('')]));
  assert.equal(shifted.ok, false, 'a block before the anchor changed');
});

test('wholeTaskChecks: one named check per position, true only on a matching final image', () => {
  const end = ii.wholeTaskChecks(op({ position: 'end' }), obsOf(`<p>${ANCHOR}</p>`, [imgRow(ANCHOR)]));
  assert.equal(end.length, 1);
  assert.equal(end[0].ok, true);
  const endFail = ii.wholeTaskChecks(op({ position: 'end' }), obsOf(`<p>${ANCHOR}</p>`));
  assert.equal(endFail[0].ok, false);
  const after = ii.wholeTaskChecks(op(), obsOf(`<p>${ANCHOR}</p><p></p>`, [imgRow('')]));
  assert.equal(after[0].ok, true);
});

test('producedText: an image is not a text unit a later step can ref', () => {
  assert.equal(ii.producedText({}), null);
});
