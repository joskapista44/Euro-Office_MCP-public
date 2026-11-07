'use strict';

/**
 * GH #317 follow-up: the insert_image header/footer RETRY IDENTITY must name the
 * target paragraph by its INDEX inside the part, never by its text. The read_images `paragraph`
 * text comes from the paragraph's own GetText(), which DROPS a real tab (the GH #317 real-tab
 * LIVE probe), while read_headers_footers keeps it (GH #317, run-text join) -- so a text match
 * either collides ('\tA' vs 'A' both normalize to 'A') or never matches ('A\tB' vs 'AB').
 * Offline: fabricated normalizeInventory observations, no editor.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const ii = require('../../src/adapters/docx/families/image-insert.cjs');

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const W = 914400;
const H = 457200;
const HDR = 'header:0:default';

const para = text => ({ text, format: {}, runs: [{ text, bold: false, hasDrawing: false }], inline: [] });
const hfPart = paras => ({ paragraphs: paras.length, simple: true, paras, tables: [] });
const obsOf = (paras, images) => normalizeInventory({ html: '<p>Body</p>', document: { images, sections: [{ header: { default: hfPart(paras), even: null, title: null }, footer: { default: null, even: null, title: null } }] } });
// `paragraph` is what GetText() returns (tab dropped); `paragraphIndex` is the part-walk index.
const row = (paragraph, paragraphIndex) => ({ index: 0, part: HDR, drawingType: 'inline', wrapType: null, cx: W, cy: H, paragraph, paragraphIndex });
const op = selector => ii.plan({ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header', ...selector }, source: PNG, widthEmu: W, heightEmu: H }, 0).operation;

test('inventory: paragraphIndex passes through (integer), anything else reads as null or its own UNOBSERVABLE reason', () => {
  const p = normalizeInventory({ html: '<p>x</p>', document: { images: [row('A', 1), row('A', null), row('A', 'UNOBSERVABLE:part-docpr-id-set-mismatch'), { ...row('A', 0), paragraphIndex: 1.5 }] } });
  assert.deepEqual(p.document.images.map(r => r.paragraphIndex), [1, null, 'UNOBSERVABLE:part-docpr-id-set-mismatch', null]);
});

test('retry: an image in paragraph 1 (\'A\') does NOT satisfy a retry aimed at paragraph 0 (\'\\tA\') -- same tab-free text, different paragraph', () => {
  const c = ii.classify(obsOf([para('\tA'), para('A')], [row('A', 1)]), op({ paragraph: { index: 0, of: 2 } }));
  assert.equal(c.state, 'APPLY');
});

test('retry: an image already in a mid-tab paragraph (\'A\\tB\', GetText \'AB\') SATISFIES the retry -- never a second image', () => {
  const c = ii.classify(obsOf([para('A\tB')], [row('AB', 0)]), op({}));
  assert.equal(c.state, 'SATISFIED');
});

test('retry: the matching index satisfies; an unreadable index never does', () => {
  assert.equal(ii.classify(obsOf([para('\tA'), para('A')], [row('A', 0)]), op({ paragraph: { index: 0, of: 2 } })).state, 'SATISFIED');
  assert.equal(ii.classify(obsOf([para('A')], [row('A', null)]), op({})).state, 'APPLY');
});
