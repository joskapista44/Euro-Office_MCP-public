'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');

/**
 * Per-element guards in the rich read. The REAL propertyQuery() string run in a vm against a fake Api in which ONE call throws
 * (a bookmark range over a hyperlink, a hyperlink getter, a child): before the guards the outer catch returned { error } and EVERY paragraph's rich entry and the whole bookmark list were lost
 * (g266v1 ROUND4B: hyperlinks and bookmarks null); now only the throwing element names its own call ('UNOBSERVABLE:<call>: <message>') and the rest stays readable.
 * RED/GREEN: the same cases are asserted against the pre-fix driver source in the RED test below (the old catch-all returns { error } for the whole read).
 */

const run = (doc, src = propertyQuery()) => {
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${src} })()`, sandbox)));
};

const link = (over = {}) => ({ GetClassType: () => 'hyperlink', GetLinkedText: () => 'https://example.invalid/a', GetScreenTipText: () => 'tip', GetDisplayedText: () => 'shown',
  GetRange: () => ({ GetTextPr: () => ({ GetStyle: () => ({ GetName: () => 'Hyperlink' }) }) }), ...over });
const para = (text, children) => ({ GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, GetElementsCount: () => children.length, GetElement: j => { const c = children[j]; if (c instanceof Error) throw c; return c; } });
const baseDoc = (paragraphs, bookmarks) => {
  const names = Object.keys(bookmarks);
  return {
    GetElementsCount: () => paragraphs.length, GetElement: i => paragraphs[i],
    Search: () => [],                       // the per-paragraph lookup finds nothing: the entry is { text, props: null, ...runInfo }
    GetAllBookmarksNames: () => names, GetBookmarkRange: n => { const b = bookmarks[n]; if (b.rangeThrows) throw new Error('range over hyperlink'); return { GetText: () => { if (b.textThrows) throw new Error('text over hyperlink'); return b.text; } }; },
    GetFootnotesFirstParagraphs: () => [], GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null,
  };
};
const entry = (r, text) => r.paragraphs.find(p => p.text === text);

test('vm: ONE bookmark whose range / text throws names its own call in its own text; the other bookmarks stay readable and the list is NOT absent', () => {
  const r = run(baseDoc([], { ok1: { text: 'fine' }, badRange: { rangeThrows: true }, badText: { textThrows: true }, ok2: { text: 'also fine' } }));
  assert.equal(r.error, undefined);
  assert.deepEqual(r.document.bookmarks, [
    { name: 'ok1', text: 'fine' },
    { name: 'badRange', text: 'UNOBSERVABLE:GetBookmarkRange: range over hyperlink' },
    { name: 'badText', text: 'UNOBSERVABLE:GetText: text over hyperlink' },
    { name: 'ok2', text: 'also fine' },
  ]);
});

test('vm: a hyperlink whose getter throws names ITS call in ITS field; the paragraph, its other fields and every OTHER paragraph are still read', () => {
  const doc = baseDoc([
    para('P1 plain link', [link()]),
    para('P2 bad link', [link({ GetLinkedText: () => { throw new Error('no linked text'); } })]),
    para('P3 plain link', [link({ GetDisplayedText: () => 'third' })]),
  ], {});
  const r = run(doc);
  assert.equal(r.error, undefined);
  assert.deepEqual(entry(r, 'P1 plain link').hyperlinks, [{ url: 'https://example.invalid/a', tip: 'tip', text: 'shown', style: 'Hyperlink' }]);
  assert.deepEqual(entry(r, 'P2 bad link').hyperlinks, [{ url: 'UNOBSERVABLE:GetLinkedText: no linked text', tip: 'tip', text: 'shown', style: 'Hyperlink' }]);
  assert.equal(entry(r, 'P3 plain link').hyperlinks[0].text, 'third', 'a later paragraph is read');
});

test('vm: a child that cannot even be fetched names itself for THAT paragraph (hyperlinks is the named string, not null, not []); the other paragraphs are untouched', () => {
  const doc = baseDoc([para('P1 ok', [link()]), para('P2 broken child', [new Error('child fetch')]), para('P3 ok', [link()])], {});
  const r = run(doc);
  assert.equal(r.error, undefined);
  assert.equal(entry(r, 'P2 broken child').hyperlinks, 'UNOBSERVABLE:child 0: child fetch');
  assert.equal(entry(r, 'P1 ok').hyperlinks.length, 1); assert.equal(entry(r, 'P3 ok').hyperlinks.length, 1);
});

test('inventory: a NAMED hyperlinks string stays named (not null); an absent field stays null; an array stays an array -- three states, never conflated', () => {
  const mk = hyperlinks => normalizeInventory({ html: '<p>x</p>', paragraphs: [{ text: 'x', props: null, runs: [], nonRunChildren: 0, runsCoverText: true, hyperlinks, noteRefs: { footnote: 0, endnote: 0 } }], document: { bookmarks: [{ name: 'b', text: 'UNOBSERVABLE:GetText: boom' }] } });
  const node = inv => inv.blocks.find(b => b.text === 'x') || inv.blocks[0];
  assert.equal(node(mk('UNOBSERVABLE:child 0: boom')).hyperlinks, 'UNOBSERVABLE:child 0: boom');
  assert.deepEqual(node(mk([])).hyperlinks, []);
  assert.equal(node(mk(undefined)).hyperlinks, null);
  assert.equal(mk([]).document.bookmarks[0].text, 'UNOBSERVABLE:GetText: boom');
});
