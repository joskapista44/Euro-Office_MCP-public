'use strict';

/**
 * GH #311 / #312 (write side): the DOCX write families locate their target with the shared body lookup
 * (src/adapters/docx/families/paragraph-lookup.cjs) instead of doc.Search. The REAL format_text command body runs via vm
 * against a fake document whose doc.Search behaves like the editor's (measured LIVE,
 * docs/evidence/gh311-312-paragraph-lookup-probe-20261027.json): a hit in every paragraph containing the needle, none
 * when the needle holds a tab. Each case records WHICH characters got the bold.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const formatting = require('../../src/adapters/docx/families/formatting.cjs');

function editor(texts) {
  const bolded = [];
  const range = (pi, start, end) => ({
    GetText: () => texts[pi].slice(start, end),
    GetParagraph: () => paragraphs[pi],
    SetBold: v => { bolded.push({ paragraph: pi, start, end, v }); },
  });
  const paragraphs = texts.map((t, pi) => ({
    GetClassType: () => 'paragraph',
    GetText: () => t,
    GetRange: (start = 0, end = t.length) => {
      if (start < 0 || end > t.length || end < start) throw new Error('range');
      return range(pi, start, end);
    },
    GetParaPr: () => null,
  }));
  const doc = {
    GetElementsCount: () => paragraphs.length,
    GetElement: i => paragraphs[i],
    Search: needle => (needle.includes('\t') ? [] : texts.flatMap((t, pi) => {
      const out = [];
      for (let at = t.indexOf(needle); at !== -1; at = t.indexOf(needle, at + needle.length)) out.push(range(pi, at, at + needle.length));
      return out;
    })),
    GetStyle: () => null,
  };
  return { Api: { GetDocument: () => doc, HexColor: x => x }, bolded };
}

function boldOn(texts, target, body = formatting.commandBody(target, { bold: true })) {
  const e = editor(texts);
  const sandbox = { Api: e.Api };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${body} })()`, sandbox);
  return e.bolded;
}

test('GH #312: the paragraph whose text is contained in another paragraph is the one written, the other untouched', () => {
  assert.deepEqual(boldOn(['Alpha', 'Alpha beta'], 'Alpha'), [{ paragraph: 0, start: 0, end: 5, v: true }]);
});

test('GH #311: a paragraph holding a tab is found and written over its whole text', () => {
  const t = 'Left\tright';
  assert.deepEqual(boldOn(['Intro', t], t), [{ paragraph: 1, start: 0, end: t.length, v: true }]);
});

test('a unit that is part of ONE paragraph (the case doc.Search covered) is still written at its own offset', () => {
  assert.deepEqual(boldOn(['Before GH228Alpha after', 'Other'], 'GH228Alpha'), [{ paragraph: 0, start: 7, end: 17, v: true }]);
});

test('NEG: a unit found in two paragraphs, or twice in one, is not written (no guess)', () => {
  assert.deepEqual(boldOn(['x Beta y', 'z Beta w'], 'Beta'), []);
  assert.deepEqual(boldOn(['Beta and Beta'], 'Beta'), []);
  assert.deepEqual(boldOn(['Same', 'Same'], 'Same'), [], 'two paragraphs with the same full text');
});

test('RED CONTROL: the old doc.Search lookup (exactly one hit) writes neither the contained nor the tab paragraph', () => {
  const t = 'Left\tright';
  const old = target => `
try {
  const targetText = ${JSON.stringify(target)};
  const doc = Api.GetDocument();
  const ranges = doc.Search(targetText, { MatchCase: true }) || [];
  if (ranges.length !== 1) return;
  ranges[0].SetBold(true);
} catch (_) {}`;
  assert.deepEqual(boldOn(['Alpha', 'Alpha beta'], 'Alpha', old('Alpha')), []);
  assert.deepEqual(boldOn(['Intro', t], t, old(t)), []);
  assert.equal(boldOn(['Unique'], 'Unique', old('Unique')).length, 1, 'the control can say yes: a unique text is written');
});
