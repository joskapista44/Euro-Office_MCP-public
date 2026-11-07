'use strict';

/**
 * GH #308 (write side): the REAL `insert_table` commandBody runs via vm against a fake document
 * modeled on the LIVE probe's own measured shape (test/live/docx-gh308-create-table-probe.live.cjs,
 * evidence agents/an internal branch/gh308-probe-live-20261028/probe1.json, LIVE PASS zero tracker comment
 * #ref on an internal branch):
 *   - doc.GetElement(i) returns a NEW wrapper object per call (identity never matches) --
 *     the fake below does this on purpose, the same way gh311-312's own fake editor() does for
 *     ranges.
 *   - Api.CreateTable(cols, rows) returns null on the 0/negative boundary.
 *   - doc.AddElement(index, table) is the only mutating call the commandBody makes; this fake
 *     records WHERE it was asked to insert and WHAT, as the real editor's own behaviour already
 *     measured LIVE (anchorIndex 7, table immediately between the anchor and its next sibling).
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');

const ti = require('../../src/adapters/docx/families/table-insert.cjs');

function editor(texts) {
  const added = [];
  let createCalls = [];
  // a minimal paragraph/table-cell content stub: GetElement(0) returns the SAME paragraph stub
  // every time it is asked for THIS cell (so a run added to it is observable), but a FRESH
  // wrapper per doc.GetElement(i) call (the measured identity-never-matches shape).
  const para = { runs: [], AddElement: run => para.runs.push(run) };
  const makeParagraph = (pi, t) => ({
    GetClassType: () => 'paragraph',
    GetText: () => t,
  });
  const table = { GetRow: () => ({ GetCell: () => ({ GetContent: () => ({ GetElement: () => para }) }) }) };
  const doc = {
    GetElementsCount: () => texts.length,
    GetElement: i => makeParagraph(i, texts[i]), // a NEW object every call -- identity never matches
    AddElement: (index, el) => { added.push({ index, el }); return true; },
  };
  const Api = {
    GetDocument: () => doc,
    CreateTable: (cols, rows) => { createCalls.push({ cols, rows }); return (cols >= 1 && rows >= 1) ? table : null; },
    CreateRun: () => { const r = { texts: [] }; r.AddText = t => r.texts.push(t); return r; },
  };
  return { Api, added, createCalls, table, para };
}

function run(texts, meta) {
  const e = editor(texts);
  const sandbox = { Api: e.Api };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${ti.commandBody(meta)} })()`, sandbox);
  return e;
}

const ANCHOR_B = 'ZQW gh308 anchor B';

test('GH #308 LIVE PASS reproduction: the anchor is found by TEXT (never identity), and AddElement is called at anchorIndex + 1', () => {
  // the probe's own BEFORE body: four W010 fixture paragraphs, then start/anchorA/siblingA/anchorB/siblingB/end
  const texts = ['W010 BOOKMARK TARGET', 'W010 HYPERLINK TARGET', 'W010 FOOTNOTE TARGET', 'W010 ENDNOTE TARGET', 'ZQW gh308 probe start', 'ZQW gh308 anchor A', 'ZQW gh308 sibling A', ANCHOR_B, 'ZQW gh308 sibling B', 'ZQW gh308 probe end'];
  const e = run(texts, { anchorText: ANCHOR_B, rows: 2, cols: 2, cells: [['TB00', 'TB01'], ['TB10', 'TB11']] });
  assert.equal(e.createCalls.length, 1);
  assert.deepEqual(e.createCalls[0], { cols: 2, rows: 2 }); // Api.CreateTable(cols, rows) -- cols first, measured
  assert.equal(e.added.length, 1);
  assert.equal(e.added[0].index, 8); // anchorIndex 7 (measured) + 1
  assert.equal(e.added[0].el, e.table);
});

test('cells, when given, are written through CreateRun + AddText on the content\'s first element', () => {
  const e = run(['x', ANCHOR_B, 'y'], { anchorText: ANCHOR_B, rows: 1, cols: 1, cells: [['hello']] });
  assert.deepEqual(e.para.runs.map(r => r.texts), [['hello']]);
});

test('no cells given: Api.CreateTable is still called, but no run is written (every cell stays the engine\'s own default empty paragraph)', () => {
  const e = run(['x', ANCHOR_B, 'y'], { anchorText: ANCHOR_B, rows: 1, cols: 1, cells: null });
  assert.equal(e.para.runs.length, 0);
  assert.equal(e.createCalls.length, 1);
});

test('NEG: anchor text absent -> no CreateTable call, no AddElement (refused, not guessed)', () => {
  const e = run(['x', 'y'], { anchorText: 'nowhere', rows: 1, cols: 1, cells: null });
  assert.equal(e.createCalls.length, 0);
  assert.equal(e.added.length, 0);
});

test('NEG: anchor text ambiguous at dispatch time -> refused rather than picking one', () => {
  const e = run([ANCHOR_B, ANCHOR_B], { anchorText: ANCHOR_B, rows: 1, cols: 1, cells: null });
  assert.equal(e.createCalls.length, 0);
  assert.equal(e.added.length, 0);
});

test('NEG: Api.CreateTable returns null on the measured 0/negative boundary -> doc.AddElement is never called', () => {
  const e = run(['x', ANCHOR_B, 'y'], { anchorText: ANCHOR_B, rows: 0, cols: 2, cells: null });
  assert.equal(e.createCalls.length, 1);
  assert.equal(e.added.length, 0);
});

test('NEG: doc.AddElement not a function (an older build without the candidate) -> the whole body returns before any lookup', () => {
  const texts = [ANCHOR_B, 'y'];
  const doc = { GetElementsCount: () => texts.length, GetElement: i => ({ GetClassType: () => 'paragraph', GetText: () => texts[i] }) };
  const sandbox = { Api: { GetDocument: () => doc, CreateTable: () => { throw new Error('must not be called'); } } };
  vm.createContext(sandbox);
  assert.doesNotThrow(() => vm.runInContext(`(function(){ ${ti.commandBody({ anchorText: ANCHOR_B, rows: 1, cols: 1, cells: null })} })()`, sandbox));
});

test('RED CONTROL: an identity-walk-only lookup (doc.GetElement(i) === anchor) never finds the anchor on this measured wrapper shape', () => {
  const texts = ['x', ANCHOR_B, 'y'];
  const e = editor(texts);
  const identityOnly = `
try {
  const doc = Api.GetDocument();
  const anchor = doc.GetElement(1);
  let anchorIndex = -1;
  for (let i = 0; i < doc.GetElementsCount(); i += 1) { if (doc.GetElement(i) === anchor) { anchorIndex = i; break; } }
  if (anchorIndex < 0) return;
  doc.AddElement(anchorIndex + 1, 'should-not-happen');
} catch (_) {}
`;
  const sandbox = { Api: e.Api };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${identityOnly} })()`, sandbox);
  assert.equal(e.added.length, 0, 'the identity walk never matches a new-per-call wrapper -- the measured GH #308 finding');
});
