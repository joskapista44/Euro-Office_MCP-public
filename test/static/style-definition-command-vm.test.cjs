'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { commandBody } = require('../../src/adapters/docx/families/style-definition.cjs');

/**
 * W-10-7 (LIVE run4): the SAME `doc.GetStyle(<unknown name>)` Normal-fallback that
 * broke the targeted read (docx-driver-style-read-vm.test.cjs) also broke `commandBody`'s
 * `define`/`modify` branches, independently -- `let style = doc.GetStyle(m.name); if (style)
 * return;` was ALWAYS true for a brand-new name (the fallback is truthy), so `define` NEVER
 * reached `doc.CreateStyle`. a maintainer's own diagnostic: a NEG_CONTROL GetStyle on a never-created
 * name came back `{found:true, identity:{name:'Normal'}}`. This is the MANDATORY vm-test (the
 * "W04-5 lesson", #ref) running the ACTUAL `commandBody` string (imported, not retyped) via
 * `vm`, against a fake `doc` whose `GetStyle` reproduces the measured fallback.
 */

function makeFakeDoc({ existingNames = [] } = {}) {
  const styles = {};
  for (const n of existingNames) {
    styles[n] = { GetName: () => n, GetTextPr: () => ({ SetBold() {}, SetFontSize() {} }), GetParaPr: () => ({ SetJc() {}, SetIndLeft() {}, SetSpacingAfter() {} }) };
  }
  const normalObj = { GetName: () => 'Normal', GetTextPr: () => ({ SetBold() {}, SetFontSize() {} }), GetParaPr: () => ({ SetJc() {}, SetIndLeft() {}, SetSpacingAfter() {} }) };
  const createdNames = [];
  const doc = {
    // W-10-7: any name NOT already in `styles` resolves to the Normal fallback, exactly as LIVE
    // run4 measured -- never null.
    GetStyle: n => styles[n] || normalObj,
    CreateStyle: (n, type) => { createdNames.push(n); const st = { GetName: () => n, GetTextPr: () => ({ SetBold() {}, SetFontSize() {} }), GetParaPr: () => ({ SetJc() {}, SetIndLeft() {}, SetSpacingAfter() {} }) }; styles[n] = st; return st; },
  };
  return { doc, createdNames };
}

function runCommand(body, doc) {
  const sandbox = { Api: { GetDocument: () => doc, RGB: (r, g, b) => ({ r, g, b }) } };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${body} })()`, sandbox);
}

test('vm: define on a brand-new name reaches CreateStyle -- the Normal fallback must not be mistaken for an existing style', () => {
  const { doc, createdNames } = makeFakeDoc({ existingNames: [] });
  const body = commandBody({ family: 'styleDefinition', action: 'styleDefinition', mode: 'define', name: 'W10NewStyle', fields: { text: { bold: true } } });
  runCommand(body, doc);
  assert.deepEqual(createdNames, ['W10NewStyle'], 'CreateStyle must be called exactly once, for the requested name');
});

test('RED/GREEN: reverting the GetName()===name guard reproduces the exact LIVE bug -- define never creates a style, because the Normal fallback reads as "already exists"', () => {
  const { doc, createdNames } = makeFakeDoc({ existingNames: [] });
  // a hand-reduced copy of the PRE-FIX commandBody define branch (bare truthiness on GetStyle)
  const preFixBody = `
try {
  const m = ${JSON.stringify({ name: 'W10NewStyle', mode: 'define', fields: {} })};
  const doc = Api.GetDocument();
  let style = doc.GetStyle(m.name);
  if (m.mode === 'define') {
    if (style) return;
    style = doc.CreateStyle(m.name, 'paragraph');
    if (!style) return;
  }
} catch (_) {}
`;
  runCommand(preFixBody, doc);
  assert.deepEqual(createdNames, [], 'reproducing the measured bug: the pre-fix branch never calls CreateStyle, because the Normal fallback is truthy');
});

test('vm: define on a name that genuinely already exists is still a defensive no-op (never overwrites)', () => {
  const { doc, createdNames } = makeFakeDoc({ existingNames: ['W10Existing'] });
  const body = commandBody({ family: 'styleDefinition', action: 'styleDefinition', mode: 'define', name: 'W10Existing', fields: { text: { bold: true } } });
  runCommand(body, doc);
  assert.deepEqual(createdNames, [], 'CreateStyle must never be called for a name that already exists');
});

test('vm: modify on a name that does not exist is a defensive no-op -- the Normal fallback must not be read as the target style', () => {
  const { doc } = makeFakeDoc({ existingNames: [] });
  let boldCalled = false;
  const normalObj = doc.GetStyle('AnyUnknownName');
  normalObj.GetTextPr = () => ({ SetBold: () => { boldCalled = true; } });
  const body = commandBody({ family: 'styleDefinition', action: 'styleDefinition', mode: 'modify', name: 'W10DoesNotExist', fields: { text: { bold: true } } });
  runCommand(body, doc);
  assert.equal(boldCalled, false, 'modify on an unknown name must never silently mutate the Normal fallback style');
});

test('vm: modify on a name that genuinely exists reaches its own SetBold, never the Normal fallback', () => {
  const { doc } = makeFakeDoc({ existingNames: ['W10Existing'] });
  let targetBold = null;
  let normalBold = null;
  doc.GetStyle('W10Existing').GetTextPr = () => ({ SetBold: v => { targetBold = v; } });
  const normalObj = doc.GetStyle('W10SomeUnrelatedName');
  normalObj.GetTextPr = () => ({ SetBold: v => { normalBold = v; } });
  const body = commandBody({ family: 'styleDefinition', action: 'styleDefinition', mode: 'modify', name: 'W10Existing', fields: { text: { bold: true } } });
  runCommand(body, doc);
  assert.equal(targetBold, true, 'the genuinely existing style must receive the write');
  assert.equal(normalBold, null, 'the Normal fallback object must never receive a write meant for another name');
});
