'use strict';

/**
 * A unit after a hyperlink in the same paragraph: the shared body lookup (families/paragraph-lookup.cjs) used the
 * unit's GetText index as its range position, but a hyperlink takes two range positions, one on each side of its
 * text (measured LIVE, docs/evidence/paragraph-lookup-hyperlink-shift-probe-20261030.json: + 2 after one link, + 4
 * after two, + 1 inside a link). The REAL format_text and replace_text command bodies run via vm against a fake
 * paragraph that keeps that position model; each case records which characters were written. A run+link unit (text
 * around a link) is written by format_text / add_footnote, which opt in (whole links only); the rewrite refuses it.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const formatting = require('../../src/adapters/docx/families/formatting.cjs');
const text = require('../../src/adapters/docx/families/text.cjs');
const notes = require('../../src/adapters/docx/families/notes.cjs');

/** `segs`: [{ kind: 'run'|'hyperlink', text }]. Range positions count a hyperlink's two edges; text indexes do not. */
function editor(segs) {
  const bolded = [];
  const plain = () => segs.map(s => s.text).join('');
  // range position -> text index of each character position (an edge has no character)
  const charsOf = (a, b) => {
    const out = []; let pos = 0; let idx = 0;
    for (const s of segs) {
      if (s.kind === 'hyperlink') pos += 1;
      for (let i = 0; i < s.text.length; i += 1) { if (pos >= a && pos < b) out.push(idx); pos += 1; idx += 1; }
      if (s.kind === 'hyperlink') pos += 1;
    }
    if (a < 0 || b > pos || b < a) throw new Error('range');
    return out;
  };
  // insert / delete at a text index, inside the segment that holds it
  const splice = (at, del, ins) => {
    let idx = 0;
    for (const s of segs) {
      if (at <= idx + s.text.length) { const o = at - idx; s.text = s.text.slice(0, o) + ins + s.text.slice(o + del); return; }
      idx += s.text.length;
    }
  };
  const range = (a, b) => {
    const idx = charsOf(a, b);
    return {
      GetText: () => idx.map(i => plain()[i]).join(''),
      SetBold: v => { bolded.push({ text: idx.map(i => plain()[i]).join(''), v }); },
      GetTextPr: () => ({}),
      SetTextPr: () => true,
      AddText: (s, where) => { if (where !== 'before' || !idx.length) return false; splice(idx[0], 0, s); return true; },
      Delete: () => { if (!idx.length) return false; splice(idx[0], idx.length, ''); return true; },
    };
  };
  const element = s => (s.kind === 'run'
    ? { GetClassType: () => 'run', GetText: () => s.text }
    : { GetClassType: () => 'hyperlink', GetDisplayedText: () => s.text });
  const paragraph = {
    GetClassType: () => 'paragraph',
    GetText: () => plain(),
    GetElementsCount: () => segs.length,
    GetElement: j => element(segs[j]),
    GetRange: (a, b) => range(a, b),
    GetParaPr: () => null,
  };
  const doc = { GetElementsCount: () => 1, GetElement: () => paragraph, GetStyle: () => null, Search: () => [] };
  return { Api: { GetDocument: () => doc, HexColor: x => x }, bolded, plain };
}

function runBody(e, body) {
  const sandbox = { Api: e.Api };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${body} })()`, sandbox);
}
const R = t => ({ kind: 'run', text: t });
const L = t => ({ kind: 'hyperlink', text: t });

test('format_text: a unit after one link, after two links, glued to a link and inside a link is the one written', () => {
  const cases = [
    [[R('Pre '), L('link'), R(' unit after')], 'unit after'],
    [[L('A'), R(' and '), L('B'), R(' unit two')], 'unit two'],
    [[L('link'), R('glued unit')], 'glued unit'],
    [[R('Pre '), L('inner link text')], 'link text'],
    [[R('unit before '), L('link'), R(' tail')], 'unit before'],
  ];
  for (const [segs, unit] of cases) {
    const e = editor(segs);
    runBody(e, formatting.commandBody(unit, { bold: true }));
    assert.deepEqual(e.bolded, [{ text: unit, v: true }], unit);
  }
});

test('format_text: a paragraph that IS one link (the whole-paragraph unit a product intent passes) is written', () => {
  const e = editor([L('only link paragraph')]);
  runBody(e, formatting.commandBody('only link paragraph', { bold: true }));
  assert.deepEqual(e.bolded, [{ text: 'only link paragraph', v: true }]);
});

test('rewrite_paragraph: a paragraph holding a hyperlink is not rewritten -- refused before any insert (a rewrite would delete the link)', () => {
  for (const [segs, unit] of [[[L('only link paragraph')], 'only link paragraph'], [[R('Pre '), L('link'), R(' old unit end')], 'old unit']]) {
    const before = segs.map(s => s.text);
    const e = editor(segs);
    runBody(e, text.commandBody(unit, 'new text'));
    assert.deepEqual(segs.map(s => s.text), before, unit);
  }
  const plain = [R('Before old unit after')];
  runBody(editor(plain), text.commandBody('old unit', 'new text'));
  assert.deepEqual(plain.map(s => s.text), ['Before new text after'], 'a paragraph without a link is still rewritten');
});

test('format_text: a paragraph with text around a link (the whole-paragraph unit a product intent passes) is written whole', () => {
  for (const segs of [[R('Pre '), L('link'), R(' post')], [R('A '), L('one'), R(' and '), L('two'), R(' end')], [L('lead link'), R(' then text')]]) {
    const unit = segs.map(x => x.text).join('');
    const e = editor(segs);
    runBody(e, formatting.commandBody(unit, { bold: true }));
    assert.deepEqual(e.bolded, [{ text: unit, v: true }], unit);
  }
});

test('add_footnote: the note goes at the end of a run+link unit', () => {
  const e = editor([R('Pre '), L('link'), R(' post')]);
  const calls = [];
  const doc = e.Api.GetDocument();
  doc.GetCurrentFootEndnote = () => null;
  doc.MoveCursorToStart = () => calls.push('home');
  doc.AddFootnote = () => { calls.push('footnote'); return { GetElement: () => ({ AddText: t => calls.push(`text:${t}`) }) }; };
  const orig = doc.GetElement();
  const getRange = orig.GetRange;
  orig.GetRange = (a, b) => { const r = getRange(a, b); r.Select = () => calls.push(`select:${r.GetText()}`); return r; };
  runBody(e, notes.commandBody({ kind: 'footnote', targetText: 'Pre link post', text: 'N1' }));
  assert.deepEqual(calls, ['select:Pre link post', 'footnote', 'text:N1', 'home']);
});

test('rewrite_paragraph: a target holding a hyperlink is FAIL_CLOSED at classification (no dispatch, not UNKNOWN)', () => {
  const block = (t, hyperlinks) => ({ kind: 'paragraph', text: t, hyperlinks });
  const op = { id: 's1', intent: 'rewrite_paragraph', selector: { text: 'Pre link post' }, text: 'new text' };
  const linked = text.classify({ blocks: [block('Pre link post', [{ text: 'link', url: 'https://x' }])] }, op);
  assert.equal(linked.state, 'FAIL_CLOSED');
  assert.equal(linked.evidence.reason, 'rewrite-target-holds-hyperlink');
  assert.equal(text.classify({ blocks: [block('Pre link post', [])] }, op).state, 'APPLY', 'no link: applied as before');
  assert.equal(text.classify({ blocks: [block('Pre link post', null)] }, op).state, 'APPLY', 'an unread link list is not a refusal here (the command guard stays)');
  assert.equal(text.classify({ blocks: [block('new text', [{ text: 'new text' }])] }, { ...op, selector: { text: 'new text' } }).state, 'SATISFIED', 'already terminal stays terminal');
});

test('NEG: a unit that covers a link only partly is not written (not measured for a write); the rewrite never crosses', () => {
  for (const unit of ['ink rest', 'Pre li', 'in']) {
    const e = editor([R('Pre '), L('link'), R(' rest of it')]);
    runBody(e, formatting.commandBody(unit, { bold: true }));
    assert.deepEqual(e.bolded, unit === 'in' ? [{ text: 'in', v: true }] : [], unit);
  }
  const t = editor([R('Pre '), L('link'), R(' rest of it')]);
  runBody(t, text.commandBody('link rest', 'X'));
  assert.equal(t.plain(), 'Pre link rest of it');
});

test('a paragraph without a hyperlink keeps the plain index', () => {
  const e = editor([R('Before GH228Alpha after')]);
  runBody(e, formatting.commandBody('GH228Alpha', { bold: true }));
  assert.deepEqual(e.bolded, [{ text: 'GH228Alpha', v: true }]);
});

test('RED CONTROL: the previous lookup (plain text index as the range position) writes no unit after a link', () => {
  const { EXACT_PARAGRAPH_LOOKUP } = require('../../src/adapters/docx/families/paragraph-lookup.cjs');
  // the previous unitRangeOf line, byte for byte (0daecbcd): the plain text index is the range position
  const NEW_LINE = 'return rangeAt(u.paragraph, u.start, t, crossLinks);';
  const OLD_LINE = "try { const r = u.paragraph.GetRange(u.start, u.start + t.length); if (r && String(r.GetText(lookupOpts) || '') === t) return r; } catch (_) {}\n    return null;";
  assert.equal(EXACT_PARAGRAPH_LOOKUP.split(NEW_LINE).length, 2, 'the current unitRangeOf line is found once');
  const previous = body => body.replace(EXACT_PARAGRAPH_LOOKUP, EXACT_PARAGRAPH_LOOKUP.replace(NEW_LINE, OLD_LINE));
  const body = formatting.commandBody('unit after', { bold: true });
  assert.ok(body.includes(EXACT_PARAGRAPH_LOOKUP), 'the command body inlines the shared lookup');
  const fresh = editor([R('Pre '), L('link'), R(' unit after')]);
  runBody(fresh, body);
  assert.deepEqual(fresh.bolded, [{ text: 'unit after', v: true }], 'the current lookup writes it');
  const old = editor([R('Pre '), L('link'), R(' unit after')]);
  runBody(old, previous(body));
  assert.deepEqual(old.bolded, []);
  const plainPara = editor([R('Before GH228Alpha after')]);
  runBody(plainPara, previous(formatting.commandBody('GH228Alpha', { bold: true })));
  assert.deepEqual(plainPara.bolded, [{ text: 'GH228Alpha', v: true }], 'the control lookup itself works without a link');
});

test('RED CONTROL: without the crossLinks opt-in (the previous lookup) a run+link paragraph is not written', () => {
  const { EXACT_PARAGRAPH_LOOKUP } = require('../../src/adapters/docx/families/paragraph-lookup.cjs');
  const OPT_IN = 'const range = unitRangeOf(targetText, true);';
  const body = formatting.commandBody('Pre link post', { bold: true });
  assert.equal(body.split(OPT_IN).length, 2, 'format_text opts in once');
  assert.ok(body.includes(EXACT_PARAGRAPH_LOOKUP));
  const off = editor([R('Pre '), L('link'), R(' post')]);
  runBody(off, body.replace(OPT_IN, 'const range = unitRangeOf(targetText);'));
  assert.deepEqual(off.bolded, []);
  const on = editor([R('Pre '), L('link'), R(' post')]);
  runBody(on, body);
  assert.deepEqual(on.bolded, [{ text: 'Pre link post', v: true }]);
});
