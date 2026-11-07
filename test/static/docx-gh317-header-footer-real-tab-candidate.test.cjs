'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const hf = require('../../src/adapters/docx/families/header-footer.cjs');

/**
 * GH #317 "real tab character" CANDIDATE -- offline coverage for header-footer.cjs's own logic
 * (the VM transport test for the docx-driver.cjs read-side fix lives in its own file,
 * docx-gh317-header-footer-tab-text-read-vm.test.cjs). Per a maintainer's probe/decision (internal tracker
 * comment #ref/#ref):
 *   1. plan(): a literal tab is now ALLOWED in `text`/`lines`; `\r`/`\n` remain forbidden.
 *   2. the loss guard (replacementLoss/replacementLossLines): a `tab` inline marker is not a loss.
 *   3. the terminal-state comparisons (classify) and the paragraph-mode structural guard
 *      (paragraphGuard) use a tab-preserving normalizer (hfNorm) instead of inventory.cjs's own
 *      norm(), which would silently strip every tab before comparing.
 */

// ---- fixtures -------------------------------------------------------------------------------

const DEFAULT_RUN_FORMAT = Object.freeze({ bold: false, italic: false, underline: false, strike: false, fontFamily: 'Calibri', fontSize: 11 });
const DEFAULTS = Object.freeze({ fontFamily: 'Calibri', fontSize: 11 });

function textRun(text) { return { text, ...DEFAULT_RUN_FORMAT }; }

function para(text, { inline = [], runs, format = { alignment: 'left' } } = {}) {
  return { text, format, runs: runs || [textRun(text)], inline };
}

function part(paras, overrides = {}) {
  return { paragraphs: paras.length, paras, simple: true, tables: [], withTables: false, ...overrides };
}

/** One section, header.default = the given part (or null); every other slot null; flags both false (known). */
function onePartSection(headerDefault) {
  return { header: { default: headerDefault, even: null, title: null }, footer: { default: null, even: null, title: null }, titlePage: false, evenAndOdd: false };
}

function primaryOf(sectionsArr) {
  return { document: { sections: sectionsArr, defaultText: DEFAULTS } };
}

function planOp(step) {
  const r = hf.plan(step, 0);
  assert.ok(r.ok, `plan() should accept this step: ${r.detail}`);
  return r.operation;
}

// ---- 1. plan(): tab allowed, CR/LF still forbidden -----------------------------------------

test('plan: a literal tab in text is ALLOWED', () => {
  const r = hf.plan({ intent: 'edit_header', section: { index: 0, of: 1 }, text: '\tSor1' }, 0);
  assert.ok(r.ok, r.detail);
  assert.equal(r.operation.text, '\tSor1');
});

test('plan: a literal tab in each line of lines is ALLOWED', () => {
  const r = hf.plan({ intent: 'edit_header', section: { index: 0, of: 1 }, lines: ['\tSor1', '\tSor2'] }, 0);
  assert.ok(r.ok, r.detail);
  assert.deepEqual(r.operation.lines, ['\tSor1', '\tSor2']);
});

test('plan NEG: a carriage return in text is still FORBIDDEN', () => {
  const r = hf.plan({ intent: 'edit_header', section: { index: 0, of: 1 }, text: 'Sor1\r' }, 0);
  assert.equal(r.ok, false);
  assert.match(r.detail, /line break/);
});

test('plan NEG: a line feed inside one of the lines is still FORBIDDEN', () => {
  const r = hf.plan({ intent: 'edit_header', section: { index: 0, of: 1 }, lines: ['Sor1', 'Sor2\n'] }, 0);
  assert.equal(r.ok, false);
  assert.match(r.detail, /line break/);
});

test('plan NEG: a text that is only a tab is still rejected as empty (nothing to align after it, unchanged from before this card)', () => {
  const r = hf.plan({ intent: 'edit_header', section: { index: 0, of: 1 }, text: '\t' }, 0);
  assert.equal(r.ok, false);
  assert.match(r.detail, /non-empty/);
});

// ---- 2. the loss guard: a `tab` inline marker is not a loss ---------------------------------

test('replacementLoss: a part whose only inline marker is "tab" is NOT a loss', () => {
  const onePara = part([para('\tSor1', { inline: ['tab'] })]).paras[0];
  const loss = hf.replacementLoss(part([onePara]), DEFAULTS, undefined);
  assert.equal(loss, null);
});

test('replacementLoss: a part with "tab" AND a real loss (image) still refuses, naming only the real loss', () => {
  const onePara = para('\tSor1', { inline: ['tab', 'image'] });
  const loss = hf.replacementLoss(part([onePara]), DEFAULTS, undefined);
  assert.ok(loss, 'expected a refusal');
  assert.match(loss, /inline:image/);
  assert.doesNotMatch(loss, /tab/);
});

test('replacementLoss: unchanged behaviour for a part with NO inline content at all', () => {
  const onePara = para('Sor1', { inline: [] });
  const loss = hf.replacementLoss(part([onePara]), DEFAULTS, undefined);
  assert.equal(loss, null);
});

test('replacementLossLines: a "tab" inline marker on any of the existing paragraphs is NOT a loss', () => {
  const paras = [para('\tSor1', { inline: ['tab'] }), para('\tSor2', { inline: ['tab'] })];
  const loss = hf.replacementLossLines(part(paras), DEFAULTS, undefined);
  assert.equal(loss, null);
});

test('replacementLossLines: a real loss (pgNum) on one paragraph still refuses, "tab" on the other does not add to it', () => {
  const paras = [para('\tSor1', { inline: ['tab'] }), para('Sor2', { inline: ['pgNum'] })];
  const loss = hf.replacementLossLines(part(paras), DEFAULTS, undefined);
  assert.ok(loss);
  assert.match(loss, /inline:pgNum/);
  assert.doesNotMatch(loss, /tab/);
});

// ---- 3. classify(): the terminal-state comparison tolerates a real tab ---------------------

test('classify: a tab-bearing header already holding the exact requested text (with the tab) is SATISFIED', () => {
  const op = planOp({ intent: 'edit_header', section: { index: 0, of: 1 }, text: '\tSor1' });
  const existing = part([para('\tSor1', { inline: ['tab'] })]);
  const primary = primaryOf([onePartSection(existing)]);
  const c = hf.classify(primary, op);
  assert.equal(c.state, 'SATISFIED', JSON.stringify(c));
});

test('classify: the SAME tab-bearing header with a DIFFERENT text is APPLY (not refused), because the loss guard now ignores the "tab" marker', () => {
  const op = planOp({ intent: 'edit_header', section: { index: 0, of: 1 }, text: '\tSorX' });
  const existing = part([para('\tSor1', { inline: ['tab'] })]);
  const primary = primaryOf([onePartSection(existing)]);
  const c = hf.classify(primary, op);
  assert.equal(c.state, 'APPLY', JSON.stringify(c));
});

test('classify "lines": a tab-bearing two-line header already holding the exact lines (both with their own tab) is SATISFIED', () => {
  const op = planOp({ intent: 'edit_header', section: { index: 0, of: 1 }, lines: ['\tSor1', '\tSor2'] });
  const existing = part([para('\tSor1', { inline: ['tab'] }), para('\tSor2', { inline: ['tab'] })]);
  const primary = primaryOf([onePartSection(existing)]);
  const c = hf.classify(primary, op);
  assert.equal(c.state, 'SATISFIED', JSON.stringify(c));
});

// ---- 3b. paragraphGuard (through classify, paragraph mode): the structural check on an
// already-tabbed single-run paragraph must not refuse it ------------------------------------

test('classify, paragraph mode: editing an already-tabbed single-run paragraph to a NEW tabbed text is APPLY, not a structural refusal', () => {
  const op = planOp({ intent: 'edit_header', section: { index: 0, of: 1 }, paragraph: { index: 0, of: 1 }, text: '\tSorX' });
  // the existing part's one paragraph: ONE run, text === the paragraph's own text (both carry the
  // tab -- the SAME shape docx-driver.cjs's run-join fix produces), as the structural check expects
  const existing = part([para('\tSor1', { inline: ['tab'] })]);
  const primary = primaryOf([onePartSection(existing)]);
  const c = hf.classify(primary, op);
  assert.equal(c.state, 'APPLY', JSON.stringify(c));
});

test('classify, paragraph mode NEG: a paragraph whose one run text does not actually match its own reported text is still refused (the structural guard still catches a real mismatch)', () => {
  const op = planOp({ intent: 'edit_header', section: { index: 0, of: 1 }, paragraph: { index: 0, of: 1 }, text: '\tSorX' });
  const mismatched = part([para('\tSor1', { inline: ['tab'], runs: [textRun('somethingElse')] })]);
  const primary = primaryOf([onePartSection(mismatched)]);
  const c = hf.classify(primary, op);
  assert.equal(c.state, 'FAIL_CLOSED', JSON.stringify(c));
  assert.match(c.evidence.reason, /text-not-in-one-run/);
});

// ---- 4. commandBody's own embedded script: the paragraph-mode old-text re-check tolerates a
// real tab too (the generated runtime JS, verified directly rather than through a vm run) --------

test('commandBody: the embedded paragraph-mode old-text check is tab-preserving, not inventory.cjs-style norm()', () => {
  const body = hf.commandBody({ family: 'headerFooter', action: 'headerFooter', part: 'header', variant: 'default', sectionIndex: 0, sectionCount: 1, create: false, activate: false, text: 'x', inheritedBy: [], mode: 'paragraph', paragraphIndex: 0, paragraphCount: 1, oldText: '\tSor1' });
  const m = body.match(/const norm = t => (.*?);\n/);
  assert.ok(m, 'the embedded local norm() definition must exist in commandBody paragraph mode');
  // m[1] is the arrow function's BODY expression (already references `t`) -- not a callable on its own.
  const norm = new Function('t', `return ${m[1]};`);
  assert.equal(norm('\tSor1'), '\tSor1', 'a tab must survive the on-page old-text re-check');
  assert.equal(norm('  Sor1  '), 'Sor1', 'plain whitespace still collapses/trims exactly like before');
});
