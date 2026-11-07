'use strict';

/**
 * GH #312 / #311 (read side): a paragraph whose text also occurs inside ANOTHER paragraph (several doc.Search hits),
 * or whose text holds a tab (doc.Search never matches a tab -- 0 hits), read back with props null: the property
 * query required exactly one hit. Neither is an ambiguity for the paragraph being read: its properties now come
 * from its own TEXT range -- exactly the text, without the paragraph mark (measured LIVE 2026-10-27,
 * docs/evidence/gh311-312-paragraph-lookup-probe-20261027.json): the smallest GetRange(0, length + d) whose text is the
 * paragraph's text (a hyperlink, and a reopened two-run paragraph, make GetRange(0, length) short -- measured LIVE
 * 2026-10-30, docs/evidence/link-paragraph-props-read-probe-20261030.json). The REAL PROPERTY_QUERY string runs
 * via vm against a fake Api whose Search behaves like the editor's: a hit in every paragraph that contains the
 * needle, none when the needle holds a tab.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');

function textPr(bold) {
  return { GetBold: () => bold, GetItalic: () => false, GetUnderline: () => false, GetStrikeout: () => false,
    GetFontFamily: () => 'Arial', GetFontSize: () => 22, GetLanguage: () => 'hu-HU' };
}

/**
 * `calls` records every GetRange argument list, so the test can see WHICH range the query read. A paragraph's `shift`
 * (default 0) is how many extra range positions its text spans (a hyperlink takes two, measured): GetRange(0, length +
 * shift) is its exact text; a shorter range misses the last characters (read with `prefixBold`), a longer one adds the mark.
 */
function makeFakeApi(texts, calls = []) {
  const elements = texts.map(({ text, bold, shift = 0, prefixBold = bold, elements = 0 }) => {
    const pr = textPr(bold);
    // the paragraph MARK is never bold here, so a read that includes it would see a mixed (absent) bold
    const markMixed = { ...pr, GetBold: () => undefined };
    const end = text.length + shift;
    const rangeOf = args => {
      if (args.length === 0 || args[1] > end) return { GetText: () => `${text}\r\n`, GetTextPr: () => markMixed };
      if (args[1] < end) return { GetText: () => text.slice(0, Math.max(0, args[1] - shift)), GetTextPr: () => textPr(prefixBold) };
      return { GetText: () => text, GetTextPr: () => pr };
    };
    return { GetClassType: () => 'paragraph', GetText: () => text, GetElementsCount: () => elements, GetElement: () => null, GetNumbering: () => null,
      GetParaPr: () => null, GetStyle: () => null, Paragraph: null, __pr: pr,
      GetRange: (...args) => { calls.push({ text, args }); return rangeOf(args); } };
  });
  const doc = {
    GetElementsCount: () => elements.length,
    GetElement: i => elements[i],
    Search: needle => (needle.includes('\t') ? [] : elements.filter(e => e.GetText().includes(needle)).map(e => ({ GetText: () => needle, GetTextPr: () => e.__pr }))),
    GetAllBookmarksNames: () => [],
    GetBookmarkRange: () => null,
    GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [],
    GetCurrentFootEndnote: () => null,
    GetAllTables: () => [],
    GetSections: () => [],
    GetDefaultTextPr: () => null,
  };
  return { GetDocument: () => doc };
}

function run(api, query = propertyQuery({})) {
  const sandbox = { Api: api };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${query} })()`, sandbox)));
}
const para = (r, text) => r.paragraphs.find(p => p.text === text);
const TAB = 'Left\tright';
const TEXTS = [{ text: 'Alpha', bold: true }, { text: 'Alpha beta', bold: false }, { text: 'Unique line', bold: false }, { text: TAB, bold: true }];

test('GH #312: the paragraph whose text is a substring of another reads its own props from its own text range', () => {
  const calls = [];
  const r = run(makeFakeApi(TEXTS, calls));
  assert.equal(r.error, undefined, `PROPERTY_QUERY threw: ${r.error}`);
  const a = para(r, 'Alpha');
  assert.ok(a.props, 'props read');
  assert.equal(a.format.bold, true, 'its OWN bold, from the text range (not the mark, not the other paragraph)');
  assert.equal(a.propsVia, 'paragraph-text-range');
  assert.ok(calls.some(c => c.text === 'Alpha' && c.args[0] === 0 && c.args[1] === 'Alpha'.length), 'GetRange(0, length)');
  const u = para(r, 'Unique line');
  assert.equal(u.format.bold, false);
  assert.equal(u.propsVia, undefined, 'a unique Search hit still reads through Search');
});

test('GH #311: a paragraph with a tab (doc.Search finds nothing) reads its props from its own text range', () => {
  const t = para(run(makeFakeApi(TEXTS)), TAB);
  assert.ok(t.props, 'props read');
  assert.equal(t.format.bold, true);
  assert.equal(t.propsVia, 'paragraph-text-range');
});

test('RED CONTROL: the old "exactly one hit or null" rule, on the same query text, leaves both paragraphs unreadable', () => {
  const cur = propertyQuery({});
  assert.ok(cur.includes('// GH #312 / #311:'), 'the fix is present in the query');
  const old = cur.replace(/if \(hits\.length !== 1\) \{\n\s*try \{\n[\s\S]*?\} else p = hits\[0\]\.GetTextPr\(\);/,
    'if (hits.length !== 1) { out.push({ text, props: null, ...runInfo }); continue; }\n      p = hits[0].GetTextPr();');
  assert.notEqual(old, cur, 'the control must actually rewrite the query');
  const r = run(makeFakeApi(TEXTS), old);
  assert.equal(para(r, 'Alpha').props, null);
  assert.equal(para(r, TAB).props, null);
  assert.ok(para(r, 'Unique line').props, 'the unique paragraph is readable either way');
});

test('RED CONTROL 2: reading the WHOLE paragraph range (GetRange() with the mark) would lose the bold', () => {
  const cur = propertyQuery({});
  const whole = cur.replace('const c = el.GetRange(0, text.length + d); if (c && String(c.GetText(textOpts) || \'\') === text) pr1 = c;', 'pr1 = el.GetRange();');
  assert.notEqual(whole, cur, 'the control must actually rewrite the query');
  assert.equal(para(run(makeFakeApi(TEXTS), whole), 'Alpha').format.bold, undefined);
});

const MIXED = [
  { text: 'Pre link ZQ', bold: undefined, prefixBold: true, shift: 3, elements: 4 }, // a link: + 2, and one more (measured)
  { text: 'Pre link ZQ and more', bold: false },
  { text: 'Two run ZQ', bold: undefined, prefixBold: true, shift: 1, elements: 2 },
  { text: 'Two run ZQ and more', bold: false },
];

test('a paragraph whose text spans more range positions (a link, a reopened two-run paragraph) reads its WHOLE text, not a prefix', () => {
  const calls = [];
  const r = run(makeFakeApi(MIXED, calls));
  assert.equal(r.error, undefined, `PROPERTY_QUERY threw: ${r.error}`);
  for (const [t, shift] of [['Pre link ZQ', 3], ['Two run ZQ', 1]]) {
    const p = para(r, t);
    assert.ok(p.props, `${t}: props read`);
    assert.equal(p.format.bold, undefined, `${t}: mixed, not the bold of a prefix`);
    assert.equal(p.propsVia, 'paragraph-text-range');
    assert.deepEqual(calls.filter(c => c.text === t).map(c => c.args[1]).slice(-1), [t.length + shift], `${t}: the smallest exact end`);
  }
});

test('NEG: no exact text range within the element bound -> props unread, never a prefix', () => {
  const r = run(makeFakeApi([{ text: 'Far ZQ', bold: undefined, prefixBold: true, shift: 9, elements: 1 }, { text: 'Far ZQ more', bold: false }]));
  assert.equal(para(r, 'Far ZQ').props, null);
});

test('RED CONTROL 3: the previous fixed GetRange(0, length) reads the prefix and reports a mixed paragraph as bold', () => {
  const cur = propertyQuery({});
  const NEW = /if \(hits\.length !== 1\) \{\n\s*try \{\n[\s\S]*?p = pr1 && typeof pr1\.GetTextPr === 'function' \? pr1\.GetTextPr\(\) : null;\n\s*\} catch \(_\) \{ p = null; \}/;
  const prev = cur.replace(NEW, "if (hits.length !== 1) {\n        try { const pr1 = typeof el.GetRange === 'function' ? el.GetRange(0, text.length) : null; p = pr1 && typeof pr1.GetTextPr === 'function' ? pr1.GetTextPr() : null; } catch (_) { p = null; }");
  assert.notEqual(prev, cur, 'the control must actually rewrite the query');
  const r = run(makeFakeApi(MIXED), prev);
  assert.equal(para(r, 'Pre link ZQ').format.bold, true, 'the misread the fix removes');
  assert.equal(para(r, 'Two run ZQ').format.bold, true);
});
