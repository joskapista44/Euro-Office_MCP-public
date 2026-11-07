'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { commandBody, rangeCommandBody, compare } = require('../../src/adapters/docx/families/formatting.cjs');

/**
 * W-10-9: the SAME `doc.GetStyle(<unknown name>)` Normal-fallback W10-7 fixed in
 * style-definition.cjs's `commandBody` also broke `formatting.cjs`'s two command
 * generators (`commandBody`, the whole-paragraph write; `rangeCommandBody`, the W1.1 substring
 * write) -- `requestedStyle = doc.GetStyle(wanted.style); if (!requestedStyle) return;` was
 * ALWAYS truthy for ANY name, misspelled or not, so a nonexistent style name would have silently
 * applied the Normal fallback via `paragraph.SetStyle(requestedStyle)`. This is the MANDATORY
 * vm-test (the "W04-5 lesson", #ref) running the ACTUAL generated strings via `vm`, against a
 * fake `doc` whose `GetStyle` reproduces the measured fallback.
 */

function makeFakeDoc({ existingNames = [] } = {}) {
  const styles = {};
  const mk = n => ({ GetName: () => n, GetTextPr: () => ({}), GetParaPr: () => ({}) });
  for (const n of existingNames) styles[n] = mk(n);
  const normalObj = mk('Normal');
  const target = { setStyleCalls: [], SetBold() {}, SetItalic() {}, SetUnderline() {}, SetStrikeout() {}, SetFontFamily() {}, SetFontSize() {}, SetColor() {}, SetHighlight() {}, SetVertAlign() {}, GetTextPr: () => ({ SetLanguage() {} }) };
  // GH #314: each SetXBorder call recorded verbatim (style, sz-eighths, space, r, g, b) -- a
  // single shared array so a test can assert on whichever side(s) it wrote, in call order.
  const borderCalls = [];
  const paraPr = {
    SetJc() {}, SetSpacingBefore() {}, SetSpacingAfter() {}, SetSpacingLine() {}, SetIndLeft() {}, SetIndRight() {}, SetIndFirstLine() {}, SetShd() {},
    SetTopBorder: (...a) => borderCalls.push(['Top', ...a]),
    SetBottomBorder: (...a) => borderCalls.push(['Bottom', ...a]),
    SetLeftBorder: (...a) => borderCalls.push(['Left', ...a]),
    SetRightBorder: (...a) => borderCalls.push(['Right', ...a]),
  };
  const paragraph = {
    GetParaPr: () => paraPr,
    borderCalls,
    SetStyle: st => { paragraph.setStyleCalls.push(st); },
    setStyleCalls: [],
    Search: (needle) => (String(needle) === 'MATCH' ? [target] : []),
    // GH #311/#312: the body finds its paragraph by iterating the document (paragraph-lookup.cjs) and writes through
    // GetRange(start, end); `text` is the paragraph's own text (a range test sets the parent sentence)
    text: 'TARGET',
    GetClassType: () => 'paragraph',
    GetText: () => paragraph.text,
    GetRange: (start, end) => ({ ...target, GetText: () => paragraph.text.slice(start, end), GetParagraph: () => paragraph }),
  };
  target.GetParagraph = () => paragraph;
  const range = { ...target, GetText: () => 'TARGET', GetParagraph: () => paragraph };
  const parentRange = { GetText: () => 'PARENT TARGET sentence', GetParagraph: () => paragraph };
  const doc = {
    // W-10-9: any name NOT already in `styles` resolves to the Normal fallback, exactly as
    // a maintainer's W10-7 LIVE run4 measured for style-definition.cjs's own GetStyle call.
    GetStyle: n => styles[n] || normalObj,
    GetElementsCount: () => 1,
    GetElement: () => paragraph,
    Search: (needle) => {
      if (String(needle) === 'TARGET') return [range];
      if (String(needle) === 'PARENT TARGET sentence') return [parentRange];
      return [];
    },
  };
  return { doc, paragraph };
}

function runBody(body, doc) {
  const sandbox = { Api: { GetDocument: () => doc, HexColor: () => ({}) } };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${body} })()`, sandbox);
}

test('vm: commandBody applies an EXISTING style -- paragraph.SetStyle is called with the genuinely matching style object', () => {
  const { doc, paragraph } = makeFakeDoc({ existingNames: ['Heading1'] });
  const body = commandBody('TARGET', { style: 'Heading1', bold: true });
  runBody(body, doc);
  assert.equal(paragraph.setStyleCalls.length, 1);
  assert.equal(paragraph.setStyleCalls[0].GetName(), 'Heading1');
});

test('vm: commandBody on a NONEXISTENT style name never calls SetStyle -- the Normal fallback must not be mistaken for a match', () => {
  const { doc, paragraph } = makeFakeDoc({ existingNames: ['Heading1'] });
  const body = commandBody('TARGET', { style: 'DoesNotExist', bold: true });
  runBody(body, doc);
  assert.equal(paragraph.setStyleCalls.length, 0, 'a refused style request must never write, not even the Normal fallback');
});

test('RED/GREEN: reverting the GetName()===name guard reproduces the exact bug -- a nonexistent style name silently applies the Normal fallback', () => {
  const { doc, paragraph } = makeFakeDoc({ existingNames: ['Heading1'] });
  // a hand-reduced copy of the PRE-FIX commandBody style branch (bare truthiness on GetStyle)
  const preFixBody = `
try {
  const wanted = ${JSON.stringify({ style: 'DoesNotExist' })};
  const doc = Api.GetDocument();
  const ranges = doc.Search('TARGET') || [];
  if (ranges.length !== 1) return;
  const range = ranges[0];
  const paragraph = range.GetParagraph(0);
  if (!paragraph) return;
  let requestedStyle = null;
  if (Object.prototype.hasOwnProperty.call(wanted, 'style')) { requestedStyle = doc.GetStyle(wanted.style); if (!requestedStyle) return; }
  if (requestedStyle) paragraph.SetStyle(requestedStyle);
} catch (_) {}
`;
  runBody(preFixBody, doc);
  assert.equal(paragraph.setStyleCalls.length, 1, 'reproducing the measured bug: the pre-fix branch silently applies the Normal fallback');
  assert.equal(paragraph.setStyleCalls[0].GetName(), 'Normal');
});

test('vm: rangeCommandBody (the W1.1 substring write) applies an EXISTING style the same way', () => {
  const { doc, paragraph } = makeFakeDoc({ existingNames: ['Heading1'] });
  paragraph.Search = () => [{ GetText: () => 'TARGET', SetBold() {}, SetItalic() {}, SetUnderline() {}, SetStrikeout() {}, SetFontFamily() {}, SetFontSize() {}, SetColor() {}, SetHighlight() {}, SetVertAlign() {}, GetTextPr: () => ({ SetLanguage() {} }) }];
  paragraph.text = 'PARENT TARGET sentence';
  doc.Search = (needle) => (String(needle) === 'PARENT TARGET sentence' ? [{ GetText: () => 'PARENT TARGET sentence', GetParagraph: () => paragraph }] : []);
  const body = rangeCommandBody('PARENT TARGET sentence', 'TARGET', { style: 'Heading1' });
  runBody(body, doc);
  assert.equal(paragraph.setStyleCalls.length, 1);
  assert.equal(paragraph.setStyleCalls[0].GetName(), 'Heading1');
});

test('vm: rangeCommandBody on a NONEXISTENT style name never calls SetStyle', () => {
  const { doc, paragraph } = makeFakeDoc({ existingNames: ['Heading1'] });
  paragraph.Search = () => [{ GetText: () => 'TARGET', SetBold() {}, SetItalic() {}, SetUnderline() {}, SetStrikeout() {}, SetFontFamily() {}, SetFontSize() {}, SetColor() {}, SetHighlight() {}, SetVertAlign() {}, GetTextPr: () => ({ SetLanguage() {} }) }];
  paragraph.text = 'PARENT TARGET sentence';
  doc.Search = (needle) => (String(needle) === 'PARENT TARGET sentence' ? [{ GetText: () => 'PARENT TARGET sentence', GetParagraph: () => paragraph }] : []);
  const body = rangeCommandBody('PARENT TARGET sentence', 'TARGET', { style: 'DoesNotExist' });
  runBody(body, doc);
  assert.equal(paragraph.setStyleCalls.length, 0);
});

/**
 * GH #314 (LIVE-measured, internal tracker/#ref): the DS stores a paragraph border's `space`
 * rounded down to whole millimeters then back to the nearest point, silently -- 12 is the
 * measured write-side INPUT that round-trips to exactly 11 points; sending 11 directly would land
 * on 9 (see formatting.cjs's own BORDER_SPACE_WRITE_INPUT comment). These vm tests run the ACTUAL
 * generated script (not a string match) against a spy SetXBorder, so a translation that computed
 * the wrong number, or forgot to apply it at all, fails here, not just at a later LIVE run.
 */
// color is the POST-normalizeBorderSide shape (no leading '#', uppercase) -- the SAME contract
// commandBody actually receives in production, never the raw user-facing '#FF0000' input.
const border = (extra = {}) => ({ style: 'single', size: 1, color: 'FF0000', ...extra });

test('vm: commandBody sends the TRANSLATED input (12), never the requested target (11), as SetTopBorder\'s 3rd argument', () => {
  const { doc, paragraph } = makeFakeDoc();
  const body = commandBody('TARGET', { borderTop: border({ space: 11 }) });
  runBody(body, doc);
  assert.deepEqual(paragraph.borderCalls, [['Top', 'single', 8, 12, 255, 0, 0]]);
});

test('vm: commandBody -- every one of the 12 reachable targets translates to its own measured input, on all four sides independently', () => {
  const { BORDER_SPACE_WRITE_INPUT } = require('../../src/adapters/docx/families/formatting.cjs');
  const sides = [['borderTop', 'Top'], ['borderBottom', 'Bottom'], ['borderLeft', 'Left'], ['borderRight', 'Right']];
  for (const [key, label] of sides) {
    for (const [target, input] of Object.entries(BORDER_SPACE_WRITE_INPUT)) {
      const { doc, paragraph } = makeFakeDoc();
      const body = commandBody('TARGET', { [key]: border({ space: Number(target) }) });
      runBody(body, doc);
      assert.deepEqual(paragraph.borderCalls, [[label, 'single', 8, input, 255, 0, 0]], `${key} target ${target}`);
    }
  }
});

test('vm: commandBody -- space OMITTED still sends 0 (the prior, unchanged behaviour) -- GH #314 only ever ADDS a 3rd-argument path, never changes the default', () => {
  const { doc, paragraph } = makeFakeDoc();
  const body = commandBody('TARGET', { borderTop: border() });
  runBody(body, doc);
  assert.deepEqual(paragraph.borderCalls, [['Top', 'single', 8, 0, 255, 0, 0]]);
});

test('vm: rangeCommandBody (the W1.1 substring write) translates space the SAME way as commandBody', () => {
  const { doc, paragraph } = makeFakeDoc({ existingNames: ['Heading1'] });
  paragraph.Search = () => [{ GetText: () => 'TARGET', SetBold() {}, SetItalic() {}, SetUnderline() {}, SetStrikeout() {}, SetFontFamily() {}, SetFontSize() {}, SetColor() {}, SetHighlight() {}, SetVertAlign() {}, GetTextPr: () => ({ SetLanguage() {} }) }];
  paragraph.text = 'PARENT TARGET sentence';
  doc.Search = (needle) => (String(needle) === 'PARENT TARGET sentence' ? [{ GetText: () => 'PARENT TARGET sentence', GetParagraph: () => paragraph }] : []);
  const body = rangeCommandBody('PARENT TARGET sentence', 'TARGET', { borderBottom: border({ space: 28 }) });
  runBody(body, doc);
  assert.deepEqual(paragraph.borderCalls, [['Bottom', 'single', 8, 29, 255, 0, 0]]);
});

/**
 * GH #314: compare()/borderSideEqual()'s own null-vs-0 guard. docx-driver.cjs's borderSideFormat
 * reports an unreadable/absent `space` as `null` (the SAME convention `size`/`color` already use
 * for an unobservable field) -- Number(null) is 0, so a naive `Number(actual.space) ===
 * wanted.space` would FALSE-MATCH a requested space of 0 against an UNKNOWN value. This is a
 * regression test for that exact fix, not a hypothetical: it failed before the typeof guard was
 * added (verified by hand, reverting the guard reproduces this exact false pass).
 */
test('compare: a requested space of 0 does NOT match an actual.space of null (unreadable) -- the null-vs-0 coercion trap', () => {
  const wanted = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: 0 } };
  const unknownSpace = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: null } };
  const m = compare(unknownSpace, wanted);
  assert.equal(m.ok, false, JSON.stringify(m));
  assert.equal(m.missing.some(x => x.property === 'borderTop'), true);
});

test('compare: a requested space DOES match when actual.space is the SAME real number', () => {
  const wanted = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: 11 } };
  const matching = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: 11 } };
  const foreign = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: 9 } };
  assert.equal(compare(matching, wanted).ok, true);
  assert.equal(compare(foreign, wanted).ok, false, 'a DIFFERENT real number must still mismatch');
});

/**
 * GH #314 R1 review (LIVE-measured, a real write of space 11 verified L1 FAIL in-session while
 * the saved XML already read "11" correctly): ParaPr.ToJSON() echoes the RAW INPUT sent to
 * SetTopBorder/etc (12, BORDER_SPACE_WRITE_INPUT[11]) until the NEXT save+reopen, never the
 * rounded target value in-session. verifyPrimitive runs in-session (right after the write, before
 * ANY reopen) -- so compare() must accept BOTH shapes a correctly-landed write can be caught in.
 */
test('compare: a requested space of 11 ALSO matches 12 (BORDER_SPACE_WRITE_INPUT[11], the IN-SESSION pre-roundtrip value) -- the R1 review fix', () => {
  const { BORDER_SPACE_WRITE_INPUT } = require('../../src/adapters/docx/families/formatting.cjs');
  const wanted = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: 11 } };
  const inSession = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: BORDER_SPACE_WRITE_INPUT[11] } };
  assert.equal(BORDER_SPACE_WRITE_INPUT[11], 12, 'sanity: the measured input for target 11');
  assert.equal(compare(inSession, wanted).ok, true, 'the pre-reopen, in-session value must ALSO verify');
});

test('compare: a requested space of 11 does NOT match an UNRELATED foreign value (neither 11 nor its own pre-roundtrip input 12)', () => {
  const wanted = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: 11 } };
  for (const foreignSpace of [9, 14, 10, 13, 0]) {
    const actual = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: foreignSpace } };
    assert.equal(compare(actual, wanted).ok, false, `foreign space ${foreignSpace} must not match target 11`);
  }
});

test('compare: every one of the 12 targets ALSO matches its own BORDER_SPACE_WRITE_INPUT value (not just 11)', () => {
  const { BORDER_SPACE_WRITE_INPUT } = require('../../src/adapters/docx/families/formatting.cjs');
  for (const [target, input] of Object.entries(BORDER_SPACE_WRITE_INPUT)) {
    const wanted = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: Number(target) } };
    const postReopen = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: Number(target) } };
    const inSession = { borderTop: { style: 'single', size: 1, color: 'FF0000', space: input } };
    assert.equal(compare(postReopen, wanted).ok, true, `target ${target} (post-reopen)`);
    assert.equal(compare(inSession, wanted).ok, true, `target ${target} (in-session, input ${input})`);
  }
});

test('compare: a border side request WITHOUT space is unaffected by whatever actual.space reads (no regression on the pre-GH #314 fields)', () => {
  const wanted = { borderTop: { style: 'single', size: 1, color: 'FF0000' } };
  for (const space of [null, 0, 11, 'UNOBSERVABLE:no-such-getter']) {
    const actual = { borderTop: { style: 'single', size: 1, color: 'FF0000', space } };
    assert.equal(compare(actual, wanted).ok, true, `space=${JSON.stringify(space)}`);
  }
});
