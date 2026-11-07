'use strict';

/**
 * add_link regression after the GH #311 / #312 body-paragraph lookup: the REAL add_link command body runs via vm
 * against a fake editor that behaves like the measured one (test/live/docx-add-link-lookup-regression-probe.live.cjs,
 * docs/evidence/add-link-lookup-regression-probe-20261029.json): once a hyperlink is added, the hyperlink takes one
 * paragraph position (paragraph.GetRange(0, length) no longer covers the whole text) and paragraph.Search no longer
 * finds the linked text; the range object the link was added on still is the target and takes SetTextPr. The link must
 * leave the target with its own text properties (the "Hyperlink" character style sets underline and colour).
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const rich = require('../../src/adapters/docx/families/rich.cjs');

const T = 'W010 HYPERLINK TARGET';
const PARENT = 'ZQWlnk link target Zlnk word';
const WORD = 'Zlnk';
const PLAIN = { underline: false, color: 'auto' };
const LINKED = { underline: true, color: 'theme11' };

function editor(texts) {
  const state = texts.map(() => ({ linked: false, props: { ...PLAIN } }));
  const textAt = (pi, start, end) => (state[pi].linked ? `\u0000${texts[pi]}` : texts[pi]).slice(start, end).replace('\u0000', '');
  const range = (pi, text) => ({
    GetText: () => text(),
    GetTextPr: () => ({ ...state[pi].props }),
    SetTextPr: p => { state[pi].props = { ...p }; },
    AddHyperlink: () => { state[pi].linked = true; state[pi].props = { ...LINKED }; return true; },
  });
  const paragraphs = texts.map((t, pi) => ({
    GetClassType: () => 'paragraph',
    GetText: () => t,
    GetRange: (start = 0, end = t.length + (state[pi].linked ? 1 : 0)) => range(pi, () => textAt(pi, start, end)),
    Search: w => (state[pi].linked || !t.includes(w) ? [] : [range(pi, () => w)]),
  }));
  const doc = {
    GetCurrentFootEndnote: () => null,
    GetElementsCount: () => paragraphs.length,
    GetElement: i => paragraphs[i],
  };
  return { Api: { GetDocument: () => doc }, state };
}

const linkBody = meta => rich.commandBody({ action: 'link', url: 'https://example.com', tooltip: '', ...meta });
// RED control: the 1ea78ca7 shape -- re-look the unit up after the link, restore only on a single hit
const relookup = body => body.replace('hits[0].SetTextPr(before);', 'const again = exact(); if (again.length !== 1) return; again[0].SetTextPr(before);');

function run(texts, body) {
  const e = editor(texts);
  const sandbox = { Api: e.Api };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${body} })()`, sandbox);
  return e.state;
}

test('whole unit: add_link restores the target text properties after the link (the hyperlink takes one paragraph position)', () => {
  const [s] = run([T], linkBody({ targetText: T }));
  assert.equal(s.linked, true, 'the link is written');
  assert.deepEqual(s.props, PLAIN, 'the seven compared properties are restored, not left on the Hyperlink style');
});

test('matchText (GH #266): the linked word gets its own text properties back (paragraph.Search no longer finds it)', () => {
  const [, s] = run(['Intro', PARENT], linkBody({ targetText: PARENT, matchText: WORD }));
  assert.equal(s.linked, true);
  assert.deepEqual(s.props, PLAIN);
});

test('RED CONTROL: the re-lookup after the link leaves the Hyperlink style on both paths', () => {
  const whole = relookup(linkBody({ targetText: T }));
  const ranged = relookup(linkBody({ targetText: PARENT, matchText: WORD }));
  assert.match(whole, /const again = exact\(\);/, 'the control body carries the re-lookup');
  assert.deepEqual(run([T], whole)[0].props, LINKED);
  assert.deepEqual(run(['Intro', PARENT], ranged)[1].props, LINKED);
});
