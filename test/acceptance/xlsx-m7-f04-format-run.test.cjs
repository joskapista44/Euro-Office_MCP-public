'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const formatRun = require('../../src/adapters/xlsx/families/format-run.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M7 backlog X-F04 (format_run): per-run (character-range) font writes inside one cell -- all 5
 * MEASURED axes (bold/italic/name/size/color) on a run, the neighbour run and the control
 * untouched (owner-approved design doc, docs/XF04-RUN-FONT-DESIGN-20261007.md). */

const RED = '#FF0000';
const book = () => [{ name: 'S', cells: { A1: { v: 'abcdef' }, A2: { v: 'control' }, A3: { f: '=1+1', v: '2' } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const runA = (font = { bold: true, color: RED }) => ({ start: 1, length: 3, font });
const runB = (font = { name: 'Calibri', size: 11 }) => ({ start: 4, length: 3, font });
const task = (runs, cell = 'A1') => ({ id: 'm7-f04', steps: [{ intent: 'format_run', sheet: 'S', cell, runs }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const cellOf = store => store.durable.find(s => s.name === 'S').cells.A1;
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('two runs, one write: APPLIED, both runs hold, retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, task([runA(), runB()]));
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.appliedOperations[0].evidence]));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  const c = cellOf(store);
  assert.deepEqual(c.runs[0], { start: 1, length: 3, font: { bold: true, italic: null, name: 'Calibri', size: 11, color: 16711680 } });
  assert.deepEqual(c.runs[1], { start: 4, length: 3, font: { bold: null, italic: null, name: 'Calibri', size: 11, color: 0 } });
  const retry = await run(store, task([runA(), runB()]));
  assert.equal(retry.m.commands, 0);
  assert.equal(retry.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('L1 FAIL: the engine applies the WRONG axis (format-run-wrong-axis)', async () => {
  const { result } = await run(createStore(book()), task([runA()]), { fault: 'format-run-wrong-axis' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('L3 FAIL: an untouched character changes (format-run-other-char-changed) or the text grows (format-run-text-changed)', async () => {
  const other = await run(createStore(book()), task([runA()]), { fault: 'format-run-other-char-changed', faultCharIndex: 5 });
  assert.equal(other.result.ok, false);
  assert.equal(other.result.execution.primitiveVerification[0].verdict, 'PASS', 'the requested run itself is right');
  assert.equal(other.result.execution.preservationVerification[0].verdict, 'FAIL', 'an untouched character changed: L3');
  const grown = await run(createStore(book()), task([runA()]), { fault: 'format-run-text-changed' });
  assert.equal(grown.result.ok, false);
  assert.equal(grown.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('L3 FAIL: a side-effect on another cell', async () => {
  const { result } = await run(createStore(book()), task([runA()]), { fault: 'side-effect' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
});

for (const [name, cell, why] of [
  ['a formula cell', 'A3', 'xlsx-format-run-on-formula-unmeasured'],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const { result, driver } = await run(createStore(book()), task([runA()], cell));
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('refused BEFORE dispatch: a run past the end of the text (run-out-of-bounds)', async () => {
  const { result, driver } = await run(createStore(book()), task([{ start: 1, length: 50, font: { bold: true } }]));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-format-run-out-of-bounds');
});

test('refused BEFORE dispatch: an empty cell has no run to address', async () => {
  const b = book();
  b[0].cells.A4 = { v: '' };
  const { result, driver } = await run(createStore(b), task([runA()], 'A4'));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-format-run-on-empty-cell');
});

test('refused BEFORE dispatch: text over the 256-char ceiling (xlsx-format-run-text-too-long)', async () => {
  const b = book();
  b[0].cells.A5 = { v: 'x'.repeat(300) };
  const { result, driver } = await run(createStore(b), task([runA()], 'A5'));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-format-run-text-too-long');
});

test('planner: shape validation only (no cell data needed)', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse(task([]), /runs must be a non-empty array/);
  refuse({ id: 'x', steps: [{ intent: 'format_run', sheet: 'S', cell: 'A1:A2', runs: [runA()] }] }, /single A1 cell/);
  refuse(task([{ start: 0, length: 1, font: { bold: true } }]), /start must be an integer/);
  refuse(task([{ start: 1, length: 0, font: { bold: true } }]), /length must be an integer/);
  refuse(task([{ start: 1, length: 3, font: { bold: true } }, { start: 2, length: 2, font: { italic: true } }]), /must not overlap/);
  refuse(task([{ start: 1, length: 1, font: {} }]), /must be a non-empty object/);
  refuse(task([{ start: 1, length: 1, font: { underline: 'single' } }]), /key\(s\) not supported/);
  refuse(task([{ start: 1, length: 1, font: { color: 'red' } }]), /font.color must be/);
  refuse(task([{ start: 1, length: 1, font: { size: 500 } }]), /font.size must be/);
});

//: the FONT object's MEASURED setter names are SetBold/SetItalic/SetName/SetSize/
// SetColor (NOT the RANGE-level SetFontName/SetFontSize/SetFontColor `format.cjs` uses on a
// DIFFERENT object) -- the stand-in below uses the MEASURED names, not the names a buggy
// implementation might expect, so this test actually catches a wrong-name regression.
function measuredFontApi(log) {
  const font = () => ({ SetBold: v => log.push(['bold', v]), SetItalic: v => log.push(['italic', v]), SetName: v => log.push(['name', v]), SetSize: v => log.push(['size', v]), SetColor: v => log.push(['color', v]) });
  const range = { GetCharacters: (start, len) => { log.push(['chars', start, len]); return { GetFont: () => font() }; } };
  const ws = { GetName: () => 'S', GetRange: () => range };
  return { GetSheet: () => ws, CreateColorFromRGB: (...x) => x.join(',') };
}

test('the command body: get-ONE-handle, write-IMMEDIATELY per run -- a run\'s own GetCharacters sits directly before its OWN setters, no other run\'s GetCharacters wedged between', () => {
  const log = [];
  const Api = measuredFontApi(log);
  const op = planOperations(task([runA({ bold: true, color: RED }), runB({ name: 'Calibri' })])).operations[0];
  assert.deepEqual(new Function('Api', formatRun.lower(op).body)(Api), { ok: true });
  //: run 0 reuses the SAME handle the pre-check already resolved (no second
  // GetCharacters(1,3) call) and writes to it immediately; run 1 gets its OWN handle right
  // before ITS OWN setters -- never a handle fetched ahead of its own writes.
  assert.deepEqual(log, [['chars', 1, 3], ['bold', true], ['color', '255,0,0'], ['chars', 4, 3], ['name', 'Calibri']]);
});

test(' point 1 NEGATIVE: a font object with the OLD, WRONG names (SetFontName/SetFontSize/SetFontColor) fails closed, setter-missing, before any write', () => {
  const log = [];
  const font = () => ({ SetBold: v => log.push(['bold', v]), SetItalic: v => log.push(['italic', v]), SetFontName: v => log.push(['name', v]), SetFontSize: v => log.push(['size', v]), SetFontColor: v => log.push(['color', v]) });
  const range = { GetCharacters: (start, len) => { log.push(['chars', start, len]); return { GetFont: () => font() }; } };
  const ws = { GetName: () => 'S', GetRange: () => range };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (...x) => x.join(',') };
  const op = planOperations(task([runA({ bold: true, color: RED })])).operations[0];
  const result = new Function('Api', formatRun.lower(op).body)(Api);
  assert.equal(result.ok, false);
  assert.equal(result.stage, 'setter-missing');
  assert.deepEqual(log, [['chars', 1, 3]], 'NO setter was called -- the guard fired before any write (the exact shape of the run1 bug)');
});

test(' NEGATIVE: the pre-check uses ONLY run 0\'s handle (the UNION of every requested setter name) -- a setter run 1 alone needs, but run 0\'s handle lacks, refuses with ONE GetCharacters call, no second handle ever fetched', () => {
  const log = [];
  const font = () => ({ SetBold: v => log.push(['bold', v]) }); // run 0's own handle is MISSING SetName -- a key only run 1 (runB) requests
  const range = { GetCharacters: (start, len) => { log.push(['chars', start, len]); return { GetFont: () => font() }; } };
  const ws = { GetName: () => 'S', GetRange: () => range };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (...x) => x.join(',') };
  const op = planOperations(task([runA({ bold: true }), runB({ name: 'Calibri' })])).operations[0];
  const result = new Function('Api', formatRun.lower(op).body)(Api);
  assert.equal(result.ok, false);
  assert.equal(result.stage, 'setter-missing');
  assert.deepEqual(result.which, ['bold', 'name'], 'the UNION across both runs, checked on run 0\'s single handle');
  assert.deepEqual(log, [['chars', 1, 3]], 'exactly ONE GetCharacters call -- run 1\'s handle is never even fetched once the union check fails on run 0\'s');
});
