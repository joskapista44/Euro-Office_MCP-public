'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * The per-paragraph doc.Search(text, { MatchCase: true }) lookup is its own failure domain. gbolv2 LIVE (DS 9.3.5): Search throws
 * "this.Content[nPos].ConvertParaContentPosToRangePos is not a function" for a paragraph whose hyperlink holds a bookmark (S3 / S5), and
 * that one throw used to take the WHOLE query down (paragraphs and document null). Now ONLY that paragraph names the throw, the query
 * reports which paragraph in `searchFailures`, and every other paragraph and the bookmark list stay readable.
 * The REAL propertyQuery() string runs in a vm against a scripted Api (the gh266 / per-element-guard pattern).
 */

const run = (doc, src = propertyQuery()) => {
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${src} })()`, sandbox)));
};

const para = text => {
  const self = { GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, GetElementsCount: () => 0, GetElement: () => null,
    GetTextPr: () => ({ GetBold: () => true }), GetStyle: () => null, GetParaPr: () => null };
  return self;
};
const docOf = (paragraphs, searchThrowsFor) => ({
  GetElementsCount: () => paragraphs.length, GetElement: i => paragraphs[i],
  Search: text => {
    if (searchThrowsFor.includes(text)) throw new TypeError('this.Content[nPos].ConvertParaContentPosToRangePos is not a function');
    const p = paragraphs.find(x => x.GetText().replace(/[\r\n]+$/, '') === text);
    return p ? [{ GetText: () => text, GetTextPr: () => p.GetTextPr() }] : [];
  },
  GetAllBookmarksNames: () => ['BLB3'], GetBookmarkRange: () => ({ GetText: () => 'zlinkd' }),
  GetFootnotesFirstParagraphs: () => [], GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null,
});
const entry = (r, text) => r.paragraphs.find(p => p.text === text);

test('vm: ONE paragraph whose Search throws names the throw on ITS entry; the other paragraphs keep their props and the bookmark list stays', () => {
  const r = run(docOf([para('P1 ok'), para('P2 bad'), para('P3 ok')], ['P2 bad']));
  assert.equal(r.error, undefined, 'the outer catch is not reached');
  assert.equal(entry(r, 'P1 ok').props.bold, true);
  assert.equal(entry(r, 'P3 ok').props.bold, true);
  assert.equal(entry(r, 'P2 bad').props, null);
  assert.equal(entry(r, 'P2 bad').searchUnobservable, 'UNOBSERVABLE:search-throws: this.Content[nPos].ConvertParaContentPosToRangePos is not a function');
  assert.equal(entry(r, 'P1 ok').searchUnobservable, undefined, 'a healthy paragraph carries no marker');
  assert.deepEqual(r.document.bookmarks, [{ name: 'BLB3', text: 'zlinkd' }]);
  assert.deepEqual(r.searchFailures, [{ index: 1, error: 'this.Content[nPos].ConvertParaContentPosToRangePos is not a function' }]);
});

test('vm: NEG -- no paragraph throws -> no searchFailures key content and no marker anywhere', () => {
  const r = run(docOf([para('P1 ok'), para('P2 ok')], []));
  assert.deepEqual(r.searchFailures || [], []);
  assert.equal(r.paragraphs.some(p => p.searchUnobservable !== undefined), false);
});

test('vm: RED against the pre-fix source -- the old per-paragraph Search is unguarded, so one throw returns { error } and EVERY paragraph is lost', () => {
  const fixed = propertyQuery();
  const old = fixed.replace(/let hits;[\s\S]*?\n    \}\n(?=\s*const get = )/, 'const propsVia = null;\n    const hits = (doc.Search(text, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || \'\') === text);\n    if (hits.length !== 1) { out.push({ text, props: null, ...runInfo }); continue; }\n    const p = hits[0].GetTextPr();\n    if (!p) { out.push({ text, props: null, ...runInfo }); continue; }\n');
  assert.notEqual(old, fixed, 'the pre-fix reconstruction differs from the fixed source (else this RED proves nothing)');
  const r = run(docOf([para('P1 ok'), para('P2 bad'), para('P3 ok')], ['P2 bad']), old);
  assert.match(String(r.error), /ConvertParaContentPosToRangePos/);
  assert.equal(r.paragraphs, undefined);
});

test('richOutcome: searchFailures -> a NAMED richUnobservable with the paragraph indexes; the payload keeps paragraphs + document (ok:true, HTML reads unaffected)', () => {
  const { richOutcome } = require('../../src/drivers/eurooffice/docx-driver.cjs');
  const res = richOutcome({ ok: true, result: { paragraphs: [{ text: 'a' }], document: {}, searchFailures: [{ index: 3, error: 'boom' }, { index: 7, error: 'later' }] } });
  assert.equal(res.richUnobservable, 'UNOBSERVABLE:search-throws: paragraph 3, 7: boom');
  assert.equal(res.paragraphs.length, 1);
  assert.deepEqual(res.document, {});
  assert.equal(richOutcome({ ok: true, result: { paragraphs: [], document: {}, searchFailures: [] } }).richUnobservable, undefined, 'NEG: no failure -> absent');
});

// ---- the paragraph-range fallback (docx-bookmark-over-link S3 / S5: Search throws, the paragraph's own range still answers) ---------------------

const rangePara = (text, range) => ({ ...para(text), GetRange: range });

test('vm: a paragraph whose Search throws but whose own range answers keeps its props, marked propsVia paragraph-range, and is NOT a search failure', () => {
  const bad = rangePara('S3 bad', () => ({ GetTextPr: () => ({ GetBold: () => true, GetItalic: () => false }) }));
  const r = run(docOf([para('P1 ok'), bad, para('P3 ok')], ['S3 bad']));
  assert.equal(r.error, undefined);
  const e = entry(r, 'S3 bad');
  assert.equal(e.props.bold, true);
  assert.equal(e.props.italic, false);
  assert.equal(e.propsVia, 'paragraph-range');
  assert.equal(e.searchUnobservable, undefined);
  assert.deepEqual(r.searchFailures || [], []);
  assert.equal(entry(r, 'P1 ok').propsVia, undefined, 'a Search-read paragraph carries no marker');
});

test('vm: NEG -- the range read throws or answers no text properties -> the paragraph names the Search throw exactly as before', () => {
  for (const range of [() => { throw new Error('range boom'); }, () => ({ GetTextPr: () => null }), () => null, () => ({})]) {
    const r = run(docOf([rangePara('S5 bad', range)], ['S5 bad']));
    const e = entry(r, 'S5 bad');
    assert.equal(e.props, null);
    assert.equal(e.searchUnobservable, 'UNOBSERVABLE:search-throws: this.Content[nPos].ConvertParaContentPosToRangePos is not a function');
    assert.equal(e.propsVia, undefined);
    assert.deepEqual(r.searchFailures, [{ index: 0, error: 'this.Content[nPos].ConvertParaContentPosToRangePos is not a function' }]);
  }
});

test('vm: RED against the pre-fallback source -- the Search throw leaves the paragraph props null although its range answers', () => {
  const fixed = propertyQuery();
  const old = fixed.replace(/      try \{ const pr0 = [^\n]*\n/, '      p = null;\n');
  assert.notEqual(old, fixed, 'the pre-fallback reconstruction differs from the fixed source (else this RED proves nothing)');
  const bad = rangePara('S3 bad', () => ({ GetTextPr: () => ({ GetBold: () => true }) }));
  const r = run(docOf([bad], ['S3 bad']), old);
  assert.equal(entry(r, 'S3 bad').props, null);
  assert.equal(entry(r, 'S3 bad').propsVia, undefined);
});
