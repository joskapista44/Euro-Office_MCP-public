'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

/**
 * GH #359 (2/2): rewrite_paragraph across a plain-run boundary. Measured LIVE (internal tracker,
 * tracker comment #ref, MULTIRUN trap -- bold 'RW8Trap_MULTIRUN_A' + plain 'B'): a run boundary
 * the unit spans claims one GetRange position that GetText does not count -- two runs needed the
 * range end widened by +1 to recover the full text (N runs -> +(N-1), measured for N=2, extended
 * here to N=3). text.cjs's commandBody and paragraph-lookup.cjs's rangeAt both widen their
 * exact-match probe window using the paragraph's own run count (plainRunPadBudget) to find it.
 *
 * This fixture models GetRange's exact position-space quirk (one dead position per element
 * boundary, the same mechanism linkedSpanUnsafe already prices for hyperlinks at 2 positions) and
 * runs the REAL commandBody() string against it with `new Function`, the pattern
 * docx-b18w-bridge-write-cutover.test.cjs uses for the Bridge word commands.
 */
const text = require('../../src/adapters/docx/families/text.cjs');

/**
 * A paragraph built from elements (default kind 'run'); GetRange's position space costs one dead
 * index per element boundary (measured LIVE for run-run boundaries), which GetText does not count.
 * A non-'run' element (e.g. a field/bookmark anchor) also claims that boundary position but is
 * NOT counted by plainRunPadBudget's run-only tally -- the fixture for the negative control below.
 */
function makeParagraph(elements) {
  let els = elements.map(e => ({ text: e.text, kind: e.kind || 'run' }));
  const fullText = () => els.map(e => e.text).join('');
  const rangePosToTextPos = rangePos => {
    let rs = 0; let ts = 0;
    for (let i = 0; i < els.length; i += 1) {
      const len = els[i].text.length;
      const re = rs + len;
      if (rangePos <= re) return ts + Math.min(rangePos - rs, len);
      rs = re + 1; ts += len;
    }
    return ts;
  };
  const rangeText = (start, end) => {
    let rs = 0; let out = '';
    for (let i = 0; i < els.length; i += 1) {
      const len = els[i].text.length;
      const re = rs + len;
      const ovStart = Math.max(start, rs); const ovEnd = Math.min(end, re);
      if (ovStart < ovEnd) out += els[i].text.slice(ovStart - rs, ovEnd - rs);
      rs = re + 1;
    }
    return out;
  };
  const insertAt = (textPos, newText) => {
    if (!newText) return;
    let acc = 0;
    for (let i = 0; i < els.length; i += 1) {
      const len = els[i].text.length;
      if (textPos <= acc) { els.splice(i, 0, { text: newText, kind: 'run' }); return; }
      if (textPos < acc + len) {
        const left = els[i].text.slice(0, textPos - acc);
        const right = els[i].text.slice(textPos - acc);
        const repl = [left && { text: left, kind: els[i].kind }, { text: newText, kind: 'run' }, right && { text: right, kind: els[i].kind }].filter(Boolean);
        els.splice(i, 1, ...repl);
        return;
      }
      acc += len;
    }
    els.push({ text: newText, kind: 'run' });
  };
  const deleteAt = (textPos, length) => {
    if (length <= 0) return;
    let acc = 0; const next = [];
    for (let i = 0; i < els.length; i += 1) {
      const len = els[i].text.length;
      const rStart = acc; const rEnd = acc + len;
      if (textPos < rEnd && textPos + length > rStart) {
        const cutStart = Math.max(textPos, rStart) - rStart;
        const cutEnd = Math.min(textPos + length, rEnd) - rStart;
        const keep = els[i].text.slice(0, cutStart) + els[i].text.slice(cutEnd);
        if (keep) next.push({ text: keep, kind: els[i].kind });
      } else next.push(els[i]);
      acc += len;
    }
    els = next.filter(e => e.text !== '');
  };
  const range = (start, end) => ({
    GetText: () => rangeText(start, end),
    AddText: t => { insertAt(rangePosToTextPos(start), t); return true; },
    Delete: () => { deleteAt(rangePosToTextPos(start), rangeText(start, end).length); return true; },
    GetTextPr: () => undefined,
    SetTextPr: () => ({}),
  });
  return {
    GetClassType: () => 'paragraph',
    GetText: () => `${fullText()}\r\n`,
    GetRange: (start, end) => range(start, end),
    GetElementsCount: () => els.length,
    GetElement: i => ({ GetClassType: () => els[i].kind, GetText: () => els[i].text }),
    _texts: () => els.map(e => e.text),
  };
}

function editorFor(elements) {
  const paragraph = makeParagraph(elements);
  const doc = { GetElementsCount: () => 1, GetElement: () => paragraph };
  const Api = { GetDocument: () => doc };
  return { Api, paragraph };
}

const run = (Api, body) => new Function('Api', body)(Api);

test('GH #359 (2/2): a two-run boundary (bold "A" + plain "B", the measured MULTIRUN shape) rewrites cleanly -- APPLIED, old runs replaced', () => {
  const { Api, paragraph } = editorFor([{ text: 'RW8Trap_MULTIRUN_A' }, { text: 'B' }]);
  const r = run(Api, text.commandBody('RW8Trap_MULTIRUN_AB', 'REWRITTEN'));
  assert.equal(r, 'ok');
  assert.deepEqual(paragraph._texts(), ['REWRITTEN']);
});

test('GH #359 (2/2): three runs (the +(N-1) hypothesis at N=3) also rewrite cleanly', () => {
  const { Api, paragraph } = editorFor([{ text: 'RW8_' }, { text: 'Multi' }, { text: 'Run3' }]);
  const r = run(Api, text.commandBody('RW8_MultiRun3', 'MERGED'));
  assert.equal(r, 'ok');
  assert.deepEqual(paragraph._texts(), ['MERGED']);
});

test('GH #359 (2/2) NEG. KONTROLL: an element kind plainRunPadBudget does not count (a field/bookmark anchor) still claims a boundary position -- stays fail-closed at range-text-mismatch, nothing written', () => {
  const { Api, paragraph } = editorFor([{ text: 'RW8Trap_MULTIRUN_A' }, { text: '', kind: 'fieldCode' }, { text: 'B' }]);
  const r = run(Api, text.commandBody('RW8Trap_MULTIRUN_AB', 'REWRITTEN'));
  assert.equal(r, 'rewrite-live-range-text-mismatch');
  assert.deepEqual(paragraph._texts(), ['RW8Trap_MULTIRUN_A', '', 'B']);
});

test('GH #359 (2/2) regression: a single-run paragraph (no boundary at all) is unaffected by the widened window', () => {
  const { Api, paragraph } = editorFor([{ text: 'RW8Trap_PLAIN' }]);
  const r = run(Api, text.commandBody('RW8Trap_PLAIN', 'SOLO'));
  assert.equal(r, 'ok');
  assert.deepEqual(paragraph._texts(), ['SOLO']);
});
