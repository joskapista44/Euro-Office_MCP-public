'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { commandBody, rangeCommandBody } = require('../../src/adapters/docx/families/formatting.cjs');

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
  const paragraph = {
    GetParaPr: () => ({ SetJc() {}, SetSpacingBefore() {}, SetSpacingAfter() {}, SetSpacingLine() {}, SetIndLeft() {}, SetIndRight() {}, SetIndFirstLine() {}, SetShd() {} }),
    SetStyle: st => { paragraph.setStyleCalls.push(st); },
    setStyleCalls: [],
    Search: (needle) => (String(needle) === 'MATCH' ? [target] : []),
  };
  target.GetParagraph = () => paragraph;
  const range = { ...target, GetText: () => 'TARGET', GetParagraph: () => paragraph };
  const parentRange = { GetText: () => 'PARENT TARGET sentence', GetParagraph: () => paragraph };
  const doc = {
    // W-10-9: any name NOT already in `styles` resolves to the Normal fallback, exactly as
    // a maintainer's W10-7 LIVE run4 measured for style-definition.cjs's own GetStyle call.
    GetStyle: n => styles[n] || normalObj,
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
  doc.Search = (needle) => (String(needle) === 'PARENT TARGET sentence' ? [{ GetText: () => 'PARENT TARGET sentence', GetParagraph: () => paragraph }] : []);
  const body = rangeCommandBody('PARENT TARGET sentence', 'TARGET', { style: 'Heading1' });
  runBody(body, doc);
  assert.equal(paragraph.setStyleCalls.length, 1);
  assert.equal(paragraph.setStyleCalls[0].GetName(), 'Heading1');
});

test('vm: rangeCommandBody on a NONEXISTENT style name never calls SetStyle', () => {
  const { doc, paragraph } = makeFakeDoc({ existingNames: ['Heading1'] });
  paragraph.Search = () => [{ GetText: () => 'TARGET', SetBold() {}, SetItalic() {}, SetUnderline() {}, SetStrikeout() {}, SetFontFamily() {}, SetFontSize() {}, SetColor() {}, SetHighlight() {}, SetVertAlign() {}, GetTextPr: () => ({ SetLanguage() {} }) }];
  doc.Search = (needle) => (String(needle) === 'PARENT TARGET sentence' ? [{ GetText: () => 'PARENT TARGET sentence', GetParagraph: () => paragraph }] : []);
  const body = rangeCommandBody('PARENT TARGET sentence', 'TARGET', { style: 'DoesNotExist' });
  runBody(body, doc);
  assert.equal(paragraph.setStyleCalls.length, 0);
});
