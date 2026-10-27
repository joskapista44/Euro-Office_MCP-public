'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { commandBody } = require('../../src/adapters/docx/families/text.cjs');

/**
 * GH #228 (zero LIVE-lelet, #ref): the rewrite_paragraph command body inserted the replacement BEFORE the
 * intact target and then re-found the old text with `afterInsert.indexOf(targetText)`. When the NEW text CONTAINS
 * the old one ('Total: 5' -> 'Total: 5 EUR'), the paragraph then holds TWO occurrences, the second-occurrence guard
 * fired and the body rolled back silently -- the write never happened (DISPATCHED_OUTCOME_UNKNOWN, VERIFY_CANNOT_PROVE_TERMINAL_STATE).
 * The old text stands at a KNOWN offset after the insert (targetOffset + replacement.length): it is checked THERE, with an exact range,
 * not searched for. This is the mandatory vm-test (runs the ACTUAL generated string) against a fake one-paragraph editor
 * that reproduces the measured semantics: Search -> non-overlapping exact matches, AddText(...,'before') shifts the (stale) range,
 * GetRange(start,end) gives a fresh range, Delete removes the characters.
 */

function fakeEditor(initial) {
  const state = { text: initial, writes: 0 };
  const paragraph = {
    GetClassType: () => 'paragraph',
    GetText: () => state.text,
    GetRange(start, end) {
      if (start < 0 || end > state.text.length || end < start) throw new Error('range');
      return mkRange(start, end);
    },
  };
  function mkRange(start, end) {
    return {
      GetText: () => state.text.slice(start, end),
      GetParagraph: () => paragraph,
      GetTextPr: () => ({ bold: true }),
      SetTextPr: pr => pr,
      AddText(text, where) {
        if (where !== 'before') return false;
        state.text = state.text.slice(0, start) + text + state.text.slice(start);
        state.writes += 1;
        return true;
      },
      Delete() {
        state.text = state.text.slice(0, start) + state.text.slice(end);
        state.writes += 1;
        return true;
      },
    };
  }
  const doc = {
    Search(needle) {
      const out = [];
      for (let i = state.text.indexOf(needle); i !== -1; i = state.text.indexOf(needle, i + needle.length)) out.push(mkRange(i, i + needle.length));
      return out;
    },
  };
  return { state, doc };
}

function run(initial, oldText, newText) {
  const { state, doc } = fakeEditor(initial);
  const sandbox = { Api: { GetDocument: () => doc } };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${commandBody(oldText, newText)} })()`, sandbox);
  return state;
}

test('vm: a plain replacement (the new text does not contain the old one) still rewrites the paragraph', () => {
  assert.equal(run('Before GH228Alpha after', 'GH228Alpha', 'GH228Beta').text, 'Before GH228Beta after');
});

test('vm RED/GREEN: the new text is a SUPERSET of the old one -> the rewrite happens (was: silent rollback, text unchanged)', () => {
  assert.equal(run('GH228Alpha', 'GH228Alpha', 'GH228AlphaRewritten').text, 'GH228AlphaRewritten');
  assert.equal(run('Osszesen: 5', 'Osszesen: 5', 'Osszesen: 5 EUR').text, 'Osszesen: 5 EUR');
  assert.equal(run('x GH228Alpha y', 'GH228Alpha', 'Pre GH228Alpha Post').text, 'x Pre GH228Alpha Post y');
});

test('vm: the REVERSE case (the new text is a PART of the old one) rewrites too', () => {
  assert.equal(run('GH228Alpha', 'GH228Alpha', 'Alpha').text, 'Alpha');
  assert.equal(run('Osszesen: 5 EUR', 'Osszesen: 5 EUR', 'Osszesen: 5').text, 'Osszesen: 5');
});

test('vm: the replacement overlaps the old text at a DIFFERENT offset (a second occurrence of the old text inside inserted+old) still rewrites', () => {
  assert.equal(run('xaby', 'ab', 'aab').text, 'xaaby');
});

test('vm: an empty replacement deletes the target', () => {
  assert.equal(run('Keep GH228Alpha gone', 'GH228Alpha', '').text, 'Keep  gone');
});

test('vm: fail-closed is unchanged -- a target occurring TWICE in the document is refused before any write (0 writes)', () => {
  const s = run('Alpha and Alpha', 'Alpha', 'Beta');
  assert.equal(s.text, 'Alpha and Alpha');
  assert.equal(s.writes, 0);
});

test('vm: a target that is absent writes nothing', () => {
  const s = run('Something else', 'Alpha', 'Beta');
  assert.equal(s.text, 'Something else');
  assert.equal(s.writes, 0);
});
