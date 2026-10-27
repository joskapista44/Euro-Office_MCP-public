'use strict';

const test = require('../helpers/trim-skip.cjs').testFor('set_numbering');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #293 L7 (owner GO B, 2026-10-22): set_numbering scope 'list' -- the EXPLICIT whole-list change -- and the affected set taken from the
 * editor's own numbering collection (word /5 Bridge, docx.numbering.listMembers.read/v1) for BOTH scopes and for set_list_type.
 * Measured (docs/evidence/gh293-l7-internals-probe-run2-released-table-cell-comember-20261022.json): the inventory does not see a list
 * item in a table cell, and the released set_numbering changed such a co-member silently. Offline: the REAL product handler, adapter,
 * executor and lifecycle on test/fakes/fake-docx-driver.cjs (its word /5 members / detach model; `hiddenMembers` = items the inventory
 * misses). The LIVE acceptance is test/live/gh293l7-docx-list-scope-product.live.cjs.
 */

const V3 = '9.3.5|euro-office-bridge-word-v1/9.3.5.0/3';
const V4 = '9.3.5|euro-office-bridge-word-v1/9.3.5.0/4';
const V5 = '9.3.5|euro-office-bridge-word-v1/9.3.5.0/5';
const FILE = '1303627';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const item = (text, numId, level = 0) => ({ text, props: PLAIN, list: { type: 'numbered', level, numId } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const sn = (text, extra = {}) => ({ intent: 'set_numbering', selector: { text }, code: 'a.', ...extra });

function product(paras, doc = {}, script = {}) {
  const store = createStore(paras, { doc });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, runtime: V5, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const numberingOf = async text => { const r = await read.handler({ fileId: FILE, operation: 'read_numbering' }); const it = r.result.items.find(x => x.text === text); return it ? it.numbering : null; };
  const run = steps => tool.handler({ fileId: FILE, task: { steps } });
  return { store, run, numberingOf, saves: () => store.saves, drivers };
}
const op = r => (r.operations || [])[0] || {};
const reasonOf = r => (op(r).failure || {}).reason || r.reason || null;
const durableNum = (store, text) => (store.durable.find(p => p.text === text) || {}).list || null;

// Word's shape: list A first (no override), list B after it with a level-0 startOverride, both on ONE definition
const WORD = () => [item('L7 A1', 'n1'), item('L7 A2', 'n1'), item('L7 B1', 'n2'), item('L7 B2', 'n2'), { text: 'L7 end', props: PLAIN }];
const WORD_DOC = (extra = {}) => ({ abstractOf: { n1: 'abs-s', n2: 'abs-s' }, startOverrides: { n2: [0] }, ...extra });

test('SCHEMA: scope is item | list (default item); any other value is refused pre-dispatch', () => {
  for (const scope of ['item', 'list']) assert.equal(D.validateDocxWriteRequest({ fileId: FILE, task: { steps: [sn('x', { scope })] } }).ok, true, scope);
  const v = D.validateDocxWriteRequest({ fileId: FILE, task: { steps: [sn('x', { scope: 'all' })] } });
  assert.equal(v.ok, false); assert.equal(v.dispatched, false);
});

test('RUNTIME: off word /5 the affected set is not provable -> set_numbering (both scopes) and set_list_type fail closed BEFORE any write', async () => {
  for (const runtime of [V3, V4]) {
    for (const step of [sn('L7 solo'), sn('L7 solo', { scope: 'list' }), { intent: 'set_list_type', selector: { text: 'L7 solo' }, type: 'bullet' }]) {
      const p = product([item('L7 solo', 'n1'), { text: 'L7 end', props: PLAIN }], {}, { runtime });
      const r = await p.run([step]);
      assert.equal(r.ok, false, `${runtime} ${step.intent}`);
      assert.match(reasonOf(r), /members-unprovable:PATCH_VERSION_MISMATCH/);
      assert.equal(p.saves(), 0);
      assert.equal(durableNum(p.store, 'L7 solo').type, 'numbered');
    }
  }
});

test('THE RELEASED HOLE, CLOSED: a one-item list whose numbering also sits on a table-cell item the inventory does not see -> the default scope refuses, the hidden item named, 0 writes', async () => {
  const p = product([item('L7 X1', 'n1'), { text: 'L7 end', props: PLAIN }], { hiddenMembers: [{ text: 'L7 X2 in a cell', numId: 'n1', level: 0, pos: 500 }] });
  const r = await p.run([sn('L7 X1')]);
  assert.equal(r.ok, false);
  assert.equal(reasonOf(r), 'set-numbering-shared-definition:n1:level0:1');
  assert.deepEqual(op(r).conflictingItems, ['L7 X2 in a cell']);
  assert.equal(p.saves(), 0);
  const lt = await product([item('L7 X1', 'n1')], { hiddenMembers: [{ text: 'L7 X2 in a cell', numId: 'n1', level: 0, pos: 500 }] }).run([{ intent: 'set_list_type', selector: { text: 'L7 X1' }, type: 'bullet' }]);
  assert.match(reasonOf(lt), /list-shared-definition:n1:level0:1/);
});

test('scope list, a list alone on its definition: every member at the level changes (the table-cell one included), each NAMED; another level and another list unchanged; the retry is SATISFIED with 0 writes', async () => {
  const paras = [item('L7 solo', 'n1'), item('L7 L1', 'n2'), item('L7 L2', 'n2'), item('L7 L2.1', 'n2', 1), { text: 'L7 end', props: PLAIN }];
  const p = product(paras, { hiddenMembers: [{ text: 'L7 L3 in a cell', numId: 'n2', level: 0, pos: 500 }] });
  const r = await p.run([sn('L7 L1', { scope: 'list' })]);
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  assert.equal(op(r).status, 'APPLIED');
  assert.deepEqual(op(r).affectedItems, ['L7 L1', 'L7 L2', 'L7 L3 in a cell']);
  assert.equal((await p.numberingOf('L7 L2')).format.numFmt, 'lowerLetter');
  assert.notEqual(((await p.numberingOf('L7 L2.1')).format || {}).numFmt, 'lowerLetter');
  assert.notEqual(((await p.numberingOf('L7 solo')).format || {}).numFmt, 'lowerLetter');
  const saves = p.saves();
  const again = await p.run([sn('L7 L1', { scope: 'list' })]);
  assert.equal(again.ok, true); assert.equal(op(again).status, 'SATISFIED'); assert.equal(p.saves(), saves, 'the retry writes nothing');
  assert.deepEqual(op(again).affectedItems, ['L7 L1', 'L7 L2', 'L7 L3 in a cell']);
});

test('the DEFAULT scope is not widened: the same multi-item list without scope list is still refused, the co-members named', async () => {
  const p = product([item('L7 L1', 'n2'), item('L7 L2', 'n2')]);
  const r = await p.run([sn('L7 L1')]);
  assert.equal(r.ok, false); assert.match(reasonOf(r), /set-numbering-shared-definition:n2:level0:1/); assert.deepEqual(op(r).conflictingItems, ['L7 L2']); assert.equal(p.saves(), 0);
});

test('scope list on a SHARED definition (Word): list A is detached first, then written; list B keeps its numbering and format; retry SATISFIED', async () => {
  const p = product(WORD(), WORD_DOC());
  const r = await p.run([sn('L7 A1', { scope: 'list' })]);
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 800));
  assert.equal(op(r).status, 'APPLIED');
  assert.deepEqual(op(r).affectedItems, ['L7 A1', 'L7 A2']);
  const a1 = durableNum(p.store, 'L7 A1'); const a2 = durableNum(p.store, 'L7 A2');
  assert.notEqual(a1.numId, 'n1'); assert.equal(a1.numId, a2.numId);
  assert.equal(durableNum(p.store, 'L7 B1').numId, 'n2');
  assert.equal(p.store.durableDoc.abstractOf.n2, 'abs-s');
  assert.equal((await p.numberingOf('L7 A2')).format.numFmt, 'lowerLetter');
  assert.notEqual(((await p.numberingOf('L7 B1')).format || {}).numFmt, 'lowerLetter');
  const saves = p.saves();
  const again = await p.run([sn('L7 A1', { scope: 'list' })]);
  assert.equal(op(again).status, 'SATISFIED'); assert.equal(p.saves(), saves);
});

test('scope list on the LATER list B (it carries its own startOverride) is isolated too -- detached and written, list A untouched', async () => {
  const p = product(WORD(), WORD_DOC());
  const r = await p.run([sn('L7 B1', { scope: 'list' })]);
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  assert.deepEqual(op(r).affectedItems, ['L7 B1', 'L7 B2']);
  assert.equal(durableNum(p.store, 'L7 A1').numId, 'n1');
  assert.notEqual(durableNum(p.store, 'L7 B1').numId, 'n2');
});

test('not isolatable -> FAIL_CLOSED before any write, the other list named: a count that continues across the lists, a list inside another, a wrapped list, an unordered read', async () => {
  const cases = [
    // B has no startOverride: it continues A's count -- detaching EITHER changes labels
    ['L7 A1', WORD(), WORD_DOC({ startOverrides: {} }), 'a-following-list-continues-it'],
    ['L7 B1', WORD(), WORD_DOC({ startOverrides: {} }), 'continues-a-preceding-list'],
    // an item of A (a table cell) after B: B sits inside A
    ['L7 A1', WORD(), WORD_DOC({ hiddenMembers: [{ text: 'L7 A3 in a cell', numId: 'n1', level: 0, pos: 500 }] }), 'another-list-inside-it'],
    // B has items before AND after A
    ['L7 A1', [item('L7 B0', 'n2'), ...WORD()], WORD_DOC({ startOverrides: { n1: [0], n2: [0] } }), 'wrapped-by-another-list'],
  ];
  for (const [text, paras, doc, why] of cases) {
    const p = product(paras, doc);
    const r = await p.run([sn(text, { scope: 'list' })]);
    assert.equal(r.ok, false, `${text} ${why}`);
    assert.equal(reasonOf(r), `set-numbering-list-not-isolatable:${why}`, JSON.stringify(op(r)));
    assert.ok((op(r).conflictingItems || []).length > 0, 'the other list named');
    assert.equal(p.saves(), 0);
  }
  const unordered = product(WORD(), WORD_DOC(), { membersOrderUnreadable: true });
  assert.match(reasonOf(await unordered.run([sn('L7 A1', { scope: 'list' })])), /set-numbering-list-order-unreadable/);
});

test('scope list keeps the other refusals: an unreadable member set, a deeper level that quotes the target level, an unreadable deeper level text', async () => {
  assert.match(reasonOf(await product(WORD(), WORD_DOC(), { membersUnreadable: true }).run([sn('L7 A1', { scope: 'list' })])), /set-numbering-list-members-unprovable:BRIDGE:paragraphs-unreadable/);
  const quoting = [item('L7 Q1', 'n3'), item('L7 Q1.1', 'n3', 1)];
  const q = await product(quoting, { numberingFormats: { n3: [{ numFmt: 'decimal', lvlText: '%1.' }, { numFmt: 'decimal', lvlText: '%1.%2.' }] } }).run([sn('L7 Q1', { scope: 'list' })]);
  assert.match(reasonOf(q), /set-numbering-level-referenced-by-other-levels:n3:level0:1/); assert.deepEqual(op(q).conflictingItems, ['L7 Q1.1']);
  const u = await product(quoting, { numberingFormats: { n3: [{ numFmt: 'decimal', lvlText: '%1.' }, { numFmt: 'decimal', lvlText: '%2.' }] } }, { membersSigUnreadableLevel: 1 }).run([sn('L7 Q1', { scope: 'list' })]);
  assert.match(reasonOf(u), /level-referenced-by-other-levels/);
  const ok = await product(quoting, { numberingFormats: { n3: [{ numFmt: 'decimal', lvlText: '%1.' }, { numFmt: 'decimal', lvlText: '%2.' }] } }).run([sn('L7 Q1', { scope: 'list' })]);
  assert.equal(op(ok).status, 'APPLIED', 'a deeper level that does not quote it is no refusal (negative control)');
});

test('SABOTAGE: a write that also changes ANOTHER numbering is never APPLIED (L3); a detach that is not observed is never APPLIED', async () => {
  // the fault rewrites level 0 of the FIRST other numbering in numberingFormats (n1 here)
  const p = product([item('L7 solo', 'n1'), item('L7 L1', 'n2'), item('L7 L2', 'n2')], { numberingFormats: { n1: [{ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left' }], n2: [{ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left' }] } }, { fault: 'numbering-touch-other-numid' });
  const r = await p.run([sn('L7 L1', { scope: 'list' })]);
  assert.equal(r.ok, false); assert.notEqual(op(r).status, 'APPLIED'); assert.equal(op(r).l3, 'FAIL');
  const d = product(WORD(), WORD_DOC(), { fault: 'detach-not-observed' });
  const rd = await d.run([sn('L7 A1', { scope: 'list' })]);
  assert.equal(rd.ok, false); assert.notEqual(op(rd).status, 'APPLIED');
  assert.equal(durableNum(d.store, 'L7 A1').numId, 'n1');
});
