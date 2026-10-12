'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const format = require('../../src/adapters/xlsx/families/format.cjs');
const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { readQuery } = drv;
const { installBridge, withPatch } = require('../fakes/ds-bridge-runtime.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 4: format_range through the target path on the fake transport. The fake reads in the W4a / W4a2 / W4c
 * shapes: per cell fill / wrap / number format (public), the font through the characters (X-F02,
 * m7-s31b-probe1/2 MEASURED: an EMPTY cell's font is observable the SAME way, at (1, 1) -- only a reader that
 * throws for that cell stays unobservable), alignment and borders through the build-gated internal reader, a
 * border side showing the neighbour's shared edge.
 */

const book = () => [
  { name: 'Data', cells: {
    A1: { v: 'Name' }, B1: { v: 'Qty' }, A2: { v: 'x' }, B2: { v: 5 }, C2: { v: 0, f: '=B2*0' }, D5: { v: 'end' },
    E1: { v: 'm' },
  }, merges: ['E1:F1'] },
  { name: 'Other', cells: { A1: { v: 'keep me' } } },
];

const task = (range, spec, sheet = 'Data') => ({ id: 'm6-w4-format', steps: [{ intent: 'format_range', sheet, range, format: spec }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const cell = (store, a) => store.durable.find(s => s.name === 'Data').cells[a];
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

for (const [name, range, spec, check] of [
  ['fill', 'A1:B2', { fill: '#00C800' }, s => assert.equal(cell(s, 'B2').fill, 51200)],
  ['wrap', 'A1:B1', { wrap: true }, s => assert.equal(cell(s, 'A1').wrap, true)],
  ['number format', 'B2', { numberFormat: '0.00' }, s => assert.equal(cell(s, 'B2').nf, '0.00')],
  ['font on text and number cells', 'A2:B2', { font: { bold: true, italic: true, name: 'Courier New', size: 14, color: '#FF0000' } }, s => assert.deepEqual(cell(s, 'B2').font, { bold: true, italic: true, name: 'Courier New', size: 14, color: 16711680 })],
  // X-F02 (m7-s31b-probe1/2 MEASURED): an EMPTY cell's font is observable through the
  // SAME public reader an occupied cell's is -- no longer a separate, refused case.
  ['font on an EMPTY cell too', 'C3', { font: { bold: true, italic: true, name: 'Courier New', size: 14, color: '#FF0000' } }, s => assert.deepEqual(cell(s, 'C3').font, { bold: true, italic: true, name: 'Courier New', size: 14, color: 16711680 })],
  ['alignment on an EMPTY cell too', 'C3:D4', { align: { horizontal: 'center', vertical: 'top' } }, s => assert.deepEqual(cell(s, 'C3').align, { h: 2, v: 4 })],
  ['borders (the shared edge inside the range is allowed)', 'B3:B4', { border: { sides: ['top', 'bottom'], style: 'medium', color: '#0000FF' } }, s => assert.deepEqual(cell(s, 'B4').border.t, { s: 12, c: 255 })],
  ['a black border', 'C4', { border: { sides: ['left'], style: 'thick', color: '#000000' } }, s => assert.deepEqual(cell(s, 'C4').border.l, { s: 13, c: 0 })],
  ['everything at once', 'A1:A2', { fill: '#112233', wrap: false, numberFormat: '@', font: { bold: true }, align: { horizontal: 'right' }, border: { sides: ['left'], style: 'thin', color: '#00FF00' } }, s => assert.equal(cell(s, 'A2').nf, '@')],
]) {
  test(`${name}: FIRST APPLIED (1 command, L1 + L3, barrier), reopened retry SATISFIED 0 write / 0 command`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task(range, spec));
    assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.preservationVerification, result.execution.appliedOperations[0].evidence]));
    assert.equal(m.writes, 1);
    assert.equal(m.commands, 1);
    assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
    assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
    assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
    check(store);
    const again = await run(store, task(range, spec));
    assert.equal(again.result.ok, true);
    assert.equal(again.m.writes, 0);
    assert.equal(again.m.commands, 0);
    assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  });
}

test('a black border reading null after a reopen (the measured shape) is still SATISFIED', async () => {
  const store = createStore(book());
  const spec = { border: { sides: ['left'], style: 'thick', color: '#000000' } };
  await run(store, task('C4', spec));
  const again = await run(store, task('C4', spec), { blackReadsNull: true });
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('a neighbour\'s facing side may change with a border (the shared edge) -- nothing else of it', async () => {
  const { result } = await run(createStore(book()), task('B3', { border: { sides: ['top'], style: 'thin', color: '#0000FF' } }));
  assert.equal(result.ok, true, 'B2 reads the new edge as its bottom; allowed');
});

for (const [name, t, script, why] of [
  // X-F02: an EMPTY cell's font is no longer refused by default -- only when the
  // engine's reader itself throws for that cell (emptyFontUnobservable, below).
  ['the font of an EMPTY cell whose reader throws', task('C3', { font: { bold: true } }), { emptyFontUnobservable: ['C3'] }, 'xlsx-format-font-unobservable'],
  ['the font of a FORMULA cell (not measured)', task('C2', { font: { italic: true } }), {}, 'xlsx-format-font-on-formula-unmeasured'],
  ['the font of a mixed-run cell', task('A1', { font: { bold: true } }), { mixedRuns: ['A1'] }, 'xlsx-format-font-unobservable'],
  ['alignment on an unverified build', task('A1', { align: { horizontal: 'left' } }), { appearanceBuild: false }, 'xlsx-format-align-unobservable'],
  ['borders on an unverified build', task('A1', { border: { sides: ['top'], style: 'thin', color: '#000000' } }), { appearanceBuild: false }, 'xlsx-format-border-unobservable'],
  ['a merged area', task('E1:F1', { fill: '#FF0000' }), {}, 'xlsx-format-in-merged-area-refused'],
  ['a missing sheet', task('A1', { wrap: true }, 'Nope'), {}, 'xlsx-sheet-not-found'],
  ['a truncated read', task('A1', { wrap: true }), { readTruncated: true }, 'xlsx-observation-truncated'],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const store = createStore(book());
    const { result, driver } = await run(store, t, script);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
    assert.deepEqual(store.durable, book());
  });
}

test('fill / wrap / number format still work on an unverified build (public readers only)', async () => {
  const { result } = await run(createStore(book()), task('A1', { fill: '#010203', wrap: true }), { appearanceBuild: false });
  assert.equal(result.ok, true);
});

test('RED: a font.bold that turns an already-bold cell\'s cell-level bold FALSE -> L1 FAIL (the allowance is only toward true)', async () => {
  const b = book();
  b[0].cells.A1 = { ...b[0].cells.A1, bold: true };
  const { result, driver } = await run(createStore(b), task('A1', { font: { bold: true } }), { fault: 'bold-cleared' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
});

// X-F01/S07 (owner GO owner channel, 2026-10-04 --): font.bold:false
// targets the EXISTING cell-level internal reader ('bold') as a new consumer, not the public character
// reader (font.bold) -- which per S07 point 3 never reports a hard false, only true | null.

test('GREEN: font.bold false turns an already-bold cell\'s internal bold off -- the public font.bold reader drops to null, nothing else moves', async () => {
  const b = book();
  b[0].cells.A1 = { ...b[0].cells.A1, bold: true, font: { bold: true, name: 'Calibri', size: 11, color: 0 } };
  const store = createStore(b);
  const { result, m } = await run(store, task('A1', { font: { bold: false } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.writes, 1);
  assert.equal(m.commands, 1);
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(cell(store, 'A1').bold, false);
  assert.equal(cell(store, 'A1').font.bold, null);
});

test('GREEN: font.bold false reaches a FORMULA cell -- unlike font.bold:true, the internal reader is not gated by the font-on-formula refusal', async () => {
  const b = book();
  b[0].cells.C2 = { ...b[0].cells.C2, bold: true, font: { bold: true, name: 'Calibri', size: 11, color: 0 } };
  const store = createStore(b);
  const { result } = await run(store, task('C2', { font: { bold: false } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(cell(store, 'C2').bold, false);
});

test('refused BEFORE the write: font.bold false where the internal reader is not verified (another build) -- fail-closed, no dispatch', async () => {
  const b = book();
  b[0].cells.A1 = { ...b[0].cells.A1, bold: true, font: { bold: true, name: 'Calibri', size: 11, color: 0 } };
  const { result, driver } = await run(createStore(b), task('A1', { font: { bold: false } }), { boldUnobservable: 'always' });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.match(JSON.stringify(result), /xlsx-format-bold-unobservable/);
});

test('RED: a font.bold:false write that silently no-ops (the internal reader stays true) -> L1 FAIL, nothing persisted', async () => {
  const b = book();
  b[0].cells.A1 = { ...b[0].cells.A1, bold: true, font: { bold: true, name: 'Calibri', size: 11, color: 0 } };
  const { result, driver } = await run(createStore(b), task('A1', { font: { bold: false } }), { fault: 'bold-false-not-applied' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
});

// X-F02 (audit S08, font on EMPTY cells): the 'bold' key is a cell-level internal read, independent of
// `cell.font`. Widened by (m7-s31b-probe1/2 MEASURED): font.name on the same empty cell now
// reaches it too, through the SAME public reader as bold's -- the independence still holds the other
// way: if the engine's reader throws for THIS cell, font.name is refused but bold (a separate internal
// reader) is not.

test('GREEN: font.bold false reaches an EMPTY cell (the internal reader is independent of the font axis); font.name on the SAME cell now reaches it too (X-F02), but stays refused if the reader throws for it', async () => {
  const b = book();
  b[0].cells.A1 = { v: '', bold: true, font: { bold: true, name: 'Calibri', size: 11, color: 0 } };
  const store = createStore(b);
  const { result } = await run(store, task('A1', { font: { bold: false } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(cell(store, 'A1').bold, false);

  const store2 = createStore(b);
  const nameOk = await run(store2, task('A1', { font: { name: 'Arial' } }));
  assert.equal(nameOk.result.ok, true, JSON.stringify(nameOk.result.execution.primitiveVerification));
  assert.equal(cell(store2, 'A1').font.name, 'Arial');

  const store3 = createStore(b);
  const refused = await run(store3, task('A1', { font: { name: 'Arial' } }), { emptyFontUnobservable: ['A1'] });
  assert.equal(refused.result.ok, false);
  assert.equal(refused.result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-font-unobservable');
  const boldStill = await run(createStore(b), task('A1', { font: { bold: false } }), { emptyFontUnobservable: ['A1'] });
  assert.equal(boldStill.result.ok, true, JSON.stringify(boldStill.result.execution.primitiveVerification));
});

// X-F03 (S09 LIVE m7-s09-probe1; spec): bold/name/size/color are MEASURED to work on a
// formula cell (byte-stable through a value-length change and a reopen, same as a non-formula
// control, D5) -- the font-on-formula refusal LIFTS for these four axes. font.italic:true stays
// refused (not part of the four measured axes; see the dedicated italic test below).

test('GREEN: font.bold:true now reaches a FORMULA cell (X-F03 lifted the refusal for this axis)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task('C2', { font: { bold: true } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(cell(store, 'C2').font.bold, true);
});

test('GREEN: font.name/size/color now reach a FORMULA cell (X-F03 lifted the refusal for these axes)', async () => {
  for (const spec of [{ name: 'Arial' }, { size: 14 }, { color: '#FF0000' }]) {
    const { result } = await run(createStore(book()), task('C2', { font: spec }));
    assert.equal(result.ok, true, JSON.stringify([spec, result.execution.primitiveVerification]));
  }
});

// X-F01/S07 continued (owner GO owner channel, 2026-10-04 --): font.italic:false is
// the SAME shape as bold:false, but on a BRAND NEW internal read (getFont(false).getItalic(), build-gated
// to EO_INTERNAL_ITALIC_BUILDS) rather than an existing one. The same six axes as bold, mirrored.

test('RED: a font.italic that turns an already-italic cell\'s cell-level italic FALSE -> L1 FAIL', async () => {
  const b = book();
  b[0].cells.A1 = { ...b[0].cells.A1, italic: true, font: { italic: true, name: 'Calibri', size: 11, color: 0 } };
  const { result, driver } = await run(createStore(b), task('A1', { font: { italic: false } }), { fault: 'italic-false-not-applied' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
});

test('GREEN: font.italic false turns an already-italic cell\'s internal italic off -- the public font.italic reader drops to null, nothing else moves', async () => {
  const b = book();
  b[0].cells.A1 = { ...b[0].cells.A1, italic: true, font: { italic: true, name: 'Calibri', size: 11, color: 0 } };
  const store = createStore(b);
  const { result, m } = await run(store, task('A1', { font: { italic: false } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.writes, 1);
  assert.equal(m.commands, 1);
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(cell(store, 'A1').italic, false);
  assert.equal(cell(store, 'A1').font.italic, null);
});

test('GREEN: font.italic false reaches a FORMULA cell -- unlike font.italic:true, the internal reader is not gated by the font-on-formula refusal', async () => {
  const b = book();
  b[0].cells.C2 = { ...b[0].cells.C2, italic: true, font: { italic: true, name: 'Calibri', size: 11, color: 0 } };
  const store = createStore(b);
  const { result } = await run(store, task('C2', { font: { italic: false } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(cell(store, 'C2').italic, false);
});

test('refused BEFORE the write: font.italic false where the internal reader is not verified (another build) -- fail-closed, no dispatch', async () => {
  const b = book();
  b[0].cells.A1 = { ...b[0].cells.A1, italic: true, font: { italic: true, name: 'Calibri', size: 11, color: 0 } };
  const { result, driver } = await run(createStore(b), task('A1', { font: { italic: false } }), { italicUnobservable: 'always' });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.match(JSON.stringify(result), /xlsx-format-italic-unobservable/);
});

test('GREEN: font.italic false reaches an EMPTY cell (the internal reader is independent of the font axis); font.name on the SAME cell now reaches it too (X-F02), but stays refused if the reader throws for it', async () => {
  const b = book();
  b[0].cells.A1 = { v: '', italic: true, font: { italic: true, name: 'Calibri', size: 11, color: 0 } };
  const store = createStore(b);
  const { result } = await run(store, task('A1', { font: { italic: false } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(cell(store, 'A1').italic, false);

  const store2 = createStore(b);
  const nameOk = await run(store2, task('A1', { font: { name: 'Arial' } }));
  assert.equal(nameOk.result.ok, true, JSON.stringify(nameOk.result.execution.primitiveVerification));
  assert.equal(cell(store2, 'A1').font.name, 'Arial');

  const store3 = createStore(b);
  const refused = await run(store3, task('A1', { font: { name: 'Arial' } }), { emptyFontUnobservable: ['A1'] });
  assert.equal(refused.result.ok, false);
  assert.equal(refused.result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-font-unobservable');
  const italicStill = await run(createStore(b), task('A1', { font: { italic: false } }), { emptyFontUnobservable: ['A1'] });
  assert.equal(italicStill.result.ok, true, JSON.stringify(italicStill.result.execution.primitiveVerification));
});

test('font.italic:true (axis font, unchanged) stays refused on a FORMULA cell -- the asymmetry with italic:false (axis italic) is intentional, not a regression', async () => {
  const { result } = await run(createStore(book()), task('C2', { font: { italic: true } }));
  assert.equal(result.ok, false);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-font-on-formula-unmeasured');
});

// X-F05 (LIVE m7-s21-probe1/2): underline/strikeout/subscript/superscript, the SAME
// public GetCharacters(1, len).GetFont() path bold/italic/name/size/color already use.

for (const type of ['single', 'singleAccounting', 'double', 'doubleAccounting']) {
  test(`GREEN: font.underline '${type}' reads back as its enum form, on both a TEXT and a NUMBER cell`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task('A2:B2', { font: { underline: type } }));
    assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
    assert.deepEqual([m.writes, m.commands], [1, 1]);
    const enumWord = { single: 'xlUnderlineStyleSingle', singleAccounting: 'xlUnderlineStyleSingleAccounting', double: 'xlUnderlineStyleDouble', doubleAccounting: 'xlUnderlineStyleDoubleAccounting' }[type];
    assert.equal(cell(store, 'A2').font.underline, enumWord, 'TEXT cell');
    assert.equal(cell(store, 'B2').font.underline, enumWord, 'NUMBER cell');
  });
}

test('GREEN: font.underline "none" resets a cell that HAD a type set -- an actual reset, not just an accepted word', async () => {
  const store = createStore(book());
  await run(store, task('A2', { font: { underline: 'double' } }));
  assert.equal(cell(store, 'A2').font.underline, 'xlUnderlineStyleDouble');
  const { result } = await run(store, task('A2', { font: { underline: 'none' } }));
  assert.equal(result.ok, true);
  assert.equal(cell(store, 'A2').font.underline, 'xlUnderlineStyleNone');
});

test('GREEN: font.strikeout true -> true; false -> null (the same null/false ambiguity bold/italic already have)', async () => {
  const store = createStore(book());
  const on = await run(store, task('A2', { font: { strikeout: true } }));
  assert.equal(on.result.ok, true);
  assert.equal(cell(store, 'A2').font.strike, true);
  const off = await run(store, task('A2', { font: { strikeout: false } }));
  assert.equal(off.result.ok, true);
  assert.equal(cell(store, 'A2').font.strike, null);
});

test('GREEN: font.strikeout false on a NEVER-touched cell is SATISFIED, 0 writes (false and never-touched are the same null reading, LIVE m7-s21-probe1)', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task('A2', { font: { strikeout: false } }));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.deepEqual([m.writes, m.commands], [0, 0]);
});

test('GREEN: font.subscript true -> its own field true AND drops a pre-existing superscript to null (they exclude each other, LIVE-measured, not flagged as drift)', async () => {
  const store = createStore(book());
  await run(store, task('A2', { font: { superscript: true } }));
  assert.equal(cell(store, 'A2').font.sup, true);
  const { result } = await run(store, task('A2', { font: { subscript: true } }));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(cell(store, 'A2').font.sub, true);
  assert.equal(cell(store, 'A2').font.sup, null, 'the pre-existing superscript dropped, tolerated as a shadow effect, not drift');
});

test('GREEN: font.superscript true -> its own field true AND drops a pre-existing subscript to null (the symmetric case)', async () => {
  const store = createStore(book());
  await run(store, task('A2', { font: { subscript: true } }));
  const { result } = await run(store, task('A2', { font: { superscript: true } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.equal(cell(store, 'A2').font.sup, true);
  assert.equal(cell(store, 'A2').font.sub, null);
});

test('GREEN: font.subscript false -> null, on a cell that was never touched (SATISFIED, 0 writes)', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task('A2', { font: { subscript: false } }));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.deepEqual([m.writes, m.commands], [0, 0]);
});

test('planning (a maintainer review): font.underline accepts only its OWN five words -- an inherited name such as toString is refused', () => {
  for (const u of ['toString', 'constructor', '__proto__', 'hasOwnProperty']) {
    const p = planOperations(task('A2', { font: { underline: u } }));
    assert.equal(p.ok, false, u);
    assert.match(p.detail, /font\.underline must be one of/, u);
  }
  assert.equal(planOperations(task('A2', { font: { underline: 'doubleAccounting' } })).ok, true);
});

test('refused BEFORE the write, 0/0: font.subscript:true AND font.superscript:true requested together (plan-time -- they cannot both hold)', () => {
  const p = planOperations(task('A2', { font: { subscript: true, superscript: true } }));
  assert.equal(p.ok, false);
  assert.match(p.detail, /cannot both be true/);
});

for (const [label, range] of [['a NUMBER cell', 'B2'], ['a FORMULA cell', 'C2']]) {
  test(`refused BEFORE the write, 0/0: font.subscript:true on ${label} -- a silent no-op nothing else would catch`, async () => {
    const { result, m } = await run(createStore(book()), task(range, { font: { subscript: true } }));
    assert.equal(result.ok, false);
    assert.deepEqual([m.writes, m.commands], [0, 0]);
    assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-font-effects-not-text-cell');
  });
  test(`refused BEFORE the write, 0/0: font.superscript:true on ${label}`, async () => {
    const { result, m } = await run(createStore(book()), task(range, { font: { superscript: true } }));
    assert.equal(result.ok, false);
    assert.deepEqual([m.writes, m.commands], [0, 0]);
    assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-font-effects-not-text-cell');
  });
}

// a maintainer review: clearing one script while the OTHER is set (and not requested) is NOT measured -- whether it clears the other
// too would be hidden by the shadow tolerance; refused before the write. Requesting both (false + true) is fine.
for (const [set, clear] of [['superscript', 'subscript'], ['subscript', 'superscript']]) {
  test(`refused BEFORE the write, 0/0: font.${clear}:false while ${set} is set on the cell and not requested (cross-script, not measured)`, async () => {
    const store = createStore(book());
    await run(store, task('A2', { font: { [set]: true } }));
    const { result, m } = await run(store, task('A2', { font: { [clear]: false } }));
    assert.equal(result.ok, false);
    assert.deepEqual([m.writes, m.commands], [0, 0]);
    assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-font-effects-cross-script-unmeasured');
  });
}
test('GREEN: { subscript: true, superscript: false } on a superscript cell -- both requested, no cross-script refusal', async () => {
  const store = createStore(book());
  await run(store, task('A2', { font: { superscript: true } }));
  const { result } = await run(store, task('A2', { font: { subscript: true, superscript: false } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.deepEqual([cell(store, 'A2').font.sub, cell(store, 'A2').font.sup], [true, null]);
});
test('X-F05 (a maintainer review): the script setters run a false BEFORE a true, whatever the spec key order', () => {
  for (const font of [{ subscript: true, superscript: false }, { superscript: false, subscript: true }]) {
    const calls = [];
    const rec = a => new Proxy({}, { get: (_t, m) => {
      if (m === 'GetValue') return () => 'x';
      if (m === 'GetCharacters') return () => ({ GetFont: () => rec(`${a}.font`) });
      return (...args) => calls.push([m, ...args]);
    } });
    const ws = { GetName: () => 'Data', GetRange: a => rec(a) };
    new Function('Api', format.lower(planOperations(task('A1', { font })).operations[0]).body)({ GetSheet: () => ws, CreateColorFromRGB: () => null });
    assert.deepEqual(calls, [['SetSuperscript', false], ['SetSubscript', true]], JSON.stringify(font));
  }
});

test('GREEN: font.subscript:false on a NUMBER cell does NOT refuse (only :true needs a text cell)', async () => {
  const { result, m } = await run(createStore(book()), task('B2', { font: { subscript: false } }));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.deepEqual([m.writes, m.commands], [0, 0]);
});

test('GREEN (X-F02): font.strikeout on an EMPTY cell now reaches it too, through the same reader every other font key does', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task('C3', { font: { strikeout: true } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.commands, 1);
  assert.equal(cell(store, 'C3').font.strike, true);
});

test('refused BEFORE the write, 0/0: font.strikeout on an EMPTY cell whose reader throws', async () => {
  const { result, m } = await run(createStore(book()), task('C3', { font: { strikeout: true } }), { emptyFontUnobservable: ['C3'] });
  assert.equal(result.ok, false);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-font-unobservable');
});

// X-F07a (LIVE m7-s24-probe1/2): align.rotation, the SAME public
// SetOrientation/GetOrientation path the S24 probes measured. Unlike subscript/superscript,
// rotation is NOT restricted to text cells (the probe measured it on every cell type).

for (const [label, range, rotation] of [['a TEXT cell', 'A2', 45], ['a NUMBER cell', 'B2', -90], ['a FORMULA cell', 'C2', 30], ['vertical text', 'A2', 255]]) {
  test(`GREEN: align.rotation ${rotation} on ${label} is byte-exact, retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task(range, { align: { rotation } }));
    assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
    assert.deepEqual([m.writes, m.commands], [1, 1]);
    assert.equal(cell(store, range).rot, rotation);
    const again = await run(store, task(range, { align: { rotation } }));
    assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
    assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
  });
}

test('GREEN: align.rotation 0 resets a cell that HAD a different rotation -- an actual reset, not just an accepted value', async () => {
  const store = createStore(book());
  await run(store, task('A2', { align: { rotation: 45 } }));
  assert.equal(cell(store, 'A2').rot, 45);
  const { result } = await run(store, task('A2', { align: { rotation: 0 } }));
  assert.equal(result.ok, true);
  assert.equal(cell(store, 'A2').rot, 0);
});

for (const [label, rotation] of [['a fractional value', 10.5], ['91 (just above the measured domain)', 91], ['-91', -91], ['180', 180], ['254', 254], ['256', 256], ['a string', '45'], ['a named value', 'xlDownward']]) {
  test(`refused BEFORE the write, 0/0: align.rotation ${label} -- measured to silently reset to 0, not a no-op or a clamp`, () => {
    const p = planOperations(task('A2', { align: { rotation } }));
    assert.equal(p.ok, false);
    assert.match(p.detail, /align\.rotation must be an integer in \[-90, 90\] or exactly 255/);
  });
}

test('GREEN: align.rotation on an ordinary EMPTY cell is NOT refused -- rotation reads through GetOrientation() directly, not through the font/GetCharacters path font-effects depend on', async () => {
  const { result, m } = await run(createStore(book()), task('C3', { align: { rotation: 45 } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.deepEqual([m.writes, m.commands], [1, 1]);
});

test('refused BEFORE the write, 0/0: align.rotation on the NARROWER, measured gap -- a never-formatted cell whose GetOrientation() itself throws', async () => {
  const { result, m } = await run(createStore(book()), task('C3', { align: { rotation: 45 } }), { rotUnobservable: ['C3'] });
  assert.equal(result.ok, false);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-rot-unobservable');
});

/**
 *'s own regression concern: a NEW STYLE_KEYS entry enters every format step's L1/L3
 * comparison. A never-formatted, truly blank cell's rotation reader THROWS (LIVE-measured) --
 * this must read as UNOBSERVABLE, like every other gap, and must NOT turn an UNRELATED,
 * already-supported write (fill, here) into a post-write FAIL just because its ring happens to
 * contain such a cell. `rotUnobservable` simulates the measured throw on C3 (already in the
 * window as a genuinely empty cell); the fill step on A1:B2 does not touch C3 at all.
 */
test('REGRESSION GUARD: an unrelated fill step is unaffected by a ring neighbour whose rotation reader throws (the measured empty-cell gap)', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task('A1:B2', { fill: '#00C800' }), { rotUnobservable: ['C3'] });
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.preservationVerification, result.execution.appliedOperations[0].evidence]));
  assert.deepEqual([m.writes, m.commands], [1, 1]);
  assert.equal(cell(store, 'B2').fill, 51200);
});

// X-F10 (owner GO owner channel, m7-s15-probe2 -- MEASURED/plannable): numberFormat
// was already an accepted spec value for ANY non-empty string, 'General' included; what was
// unproven was the RESET EFFECT -- does writing 'General' back onto a cell that has a DIFFERENT
// format actually clear it (not just pass validation vacuously)? LIVE confirmed: yes, and it
// survives a reopen. These pin the SAME behaviour offline, through the fake driver's own
// `c.nf = spec.numberFormat` write (no product-side special-casing needed for this value).

test('GREEN: numberFormat General resets a cell that HAD a different format -- not just an accepted string, an actual reset', async () => {
  const b = book();
  b[0].cells.B2 = { ...b[0].cells.B2, nf: '0.00' };
  const store = createStore(b);
  const { result, m } = await run(store, task('B2', { numberFormat: 'General' }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.writes, 1);
  assert.equal(m.commands, 1);
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(cell(store, 'B2').nf, 'General');
});

test('GREEN: numberFormat General on an ALREADY-General cell is SATISFIED, 0 writes -- not a special case, the same shape as any already-true request', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task('B2', { numberFormat: 'General' }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.writes, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('RED: ignored / partial / an extra key touched -> L1 FAIL, nothing persisted', async () => {
  for (const [t, fault] of [[task('A1:B1', { fill: '#FF0000' }), 'ignore'], [task('A1:B2', { wrap: true }), 'partial'], [task('A1', { fill: '#FF0000' }), 'format-extra-key']]) {
    const store = createStore(book());
    const { result, driver } = await run(store, t, { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
    assert.deepEqual(store.durable, book(), fault);
  }
  const { result } = await run(createStore(book()), task('A1', { fill: '#FF0000' }), { fault: 'format-extra-key' });
  assert.deepEqual(result.execution.primitiveVerification[0].evidence.evidence.otherKeysChanged, [{ cell: 'A1', keys: ['nf'] }]);
});

test('RED: the format leaks onto a neighbour, or a cell elsewhere changes -> L3 FAIL', async () => {
  const leak = await run(createStore(book()), task('A1', { fill: '#FF0000' }), { fault: 'format-leaks', leakCell: 'A2' });
  assert.equal(leak.result.ok, false);
  const pv = leak.result.execution.preservationVerification[0];
  assert.equal(pv.verdict, 'FAIL');
  assert.deepEqual(pv.evidence.evidence.styleChanged, [{ cell: 'A2', keys: ['fill'] }]);
  const side = await run(createStore(book()), task('A1', { fill: '#FF0000' }), { fault: 'side-effect' });
  assert.equal(side.result.ok, false);
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('RED: a border leaking to a neighbour\'s NON-facing side is an L3 FAIL (the shared-edge exception is exact)', async () => {
  const store = createStore(book());
  const leak = await run(store, task('B3', { border: { sides: ['top'], style: 'thin', color: '#0000FF' } }), { fault: 'format-leaks', leakCell: 'C3' });
  assert.equal(leak.result.ok, false);
  assert.deepEqual(leak.result.execution.preservationVerification[0].evidence.evidence.styleChanged.map(x => x.cell).sort(), ['C2', 'C3'], 'the leaked C3 top is also C2\'s bottom (the shared edge)');
});

test('planner: the measured-only values, removal, false, shapes', () => {
  const refuse = (spec, re) => { const p = planOperations(task('A1', spec)); assert.equal(p.ok, false, JSON.stringify(spec)); assert.match(p.detail, re); };
  refuse({}, /format is empty/);
  refuse(null, /format must be an object/);
  refuse({ colour: '#000000' }, /not supported: colour/);
  refuse({ fill: 'red' }, /fill must be '#RRGGBB'/);
  refuse({ fill: null }, /fill must be '#RRGGBB'/);
  refuse({ wrap: 'yes' }, /wrap must be true or false/);
  refuse({ numberFormat: '' }, /numberFormat must be a non-empty string/);
  refuse({ font: { bold: 'x' } }, /font.bold must be true or false/);
  const boldFalse = format.expectationsOf({ font: { bold: false } });
  assert.deepEqual(boldFalse.want, [{ key: 'bold', expect: false }], 'targets the existing internal reader, not font.bold');
  refuse({ font: { italic: 'x' } }, /font.italic must be true or false/);
  const italicFalse = format.expectationsOf({ font: { italic: false } });
  assert.deepEqual(italicFalse.want, [{ key: 'italic', expect: false }], 'targets the new internal reader, not font.italic');
  refuse({ font: { strikethrough: true } }, /font key\(s\) not supported: strikethrough/); // the real key is 'strikeout' (X-F05)
  refuse({ font: { underline: 'triple' } }, /font\.underline must be one of single, singleAccounting, double, doubleAccounting, none/);
  refuse({ font: { strikeout: 'x' } }, /font\.strikeout must be true or false/);
  refuse({ font: { subscript: 'x' } }, /font\.subscript must be true or false/);
  refuse({ font: { superscript: 'x' } }, /font\.superscript must be true or false/);
  refuse({ font: { subscript: true, superscript: true } }, /cannot both be true/);
  refuse({ font: { size: 0 } }, /font.size must be/);
  refuse({ align: { horizontal: 'fill' } }, /align.horizontal must be one of/);
  refuse({ align: { vertical: 'middle' } }, /align.vertical must be one of/);
  refuse({ align: { bogus: 1 } }, /align key\(s\) not supported/);
  // X-F07b/c (owner GO 2026-10-05): indent/readingOrder/shrink are now supported align keys.
  refuse({ align: { indent: 2 } }, /align\.indent requires align\.horizontal 'left' or 'right'/);
  refuse({ align: { indent: -1, horizontal: 'left' } }, /align\.indent must be an integer in \[0, 15\]/);
  refuse({ align: { readingOrder: 'up' } }, /align\.readingOrder must be one of/);
  refuse({ align: { shrink: 'yes' } }, /align\.shrink must be true or false/);
  refuse({ border: { sides: ['top'], style: 'none', color: '#000000' } }, /border\.style 'none' is not a style/);
  refuse({ border: { sides: ['top'], style: 'triple', color: '#000000' } }, /the measured styles/);
  refuse({ border: { sides: ['insideHorizontal'], style: 'thin', color: '#000000' } }, /border.sides/);
  refuse({ border: { sides: ['top', 'top'], style: 'thin', color: '#000000' } }, /border.sides/);
  const big = planOperations(task('A1:Z100', { wrap: true })); // 26*100=2600, over the X-B11 2000 ceiling
  assert.match(big.detail, /over the 2000 ceiling/);
});

test('X-B11: the cell-count ceiling sits at exactly 2000 -- 2000 plans, 2001 refuses', () => {
  const atLimit = planOperations(task('A1:A2000', { wrap: true })); // 1 col * 2000 rows = 2000 cells
  assert.equal(atLimit.ok, true, JSON.stringify(atLimit));
  assert.equal(atLimit.operations[0].cells.length, 2000);
  const overLimit = planOperations(task('A1:A2001', { wrap: true })); // 2001 cells
  assert.equal(overLimit.ok, false);
  assert.match(overLimit.detail, /range has 2001 cells, over the 2000 ceiling/);
});

test('the command body compiles: range-level public setters, per-cell alignment / borders, exact-name sheet', () => {
  const op = planOperations(task('A1:B1', { fill: '#00C800', wrap: true, numberFormat: '0.00', font: { bold: true, size: 14, color: '#FF0000' }, align: { horizontal: 'center' }, border: { sides: ['top'], style: 'thin', color: '#0000FF' } })).operations[0];
  const cmd = format.lower(op);
  const calls = [];
  const rec = a => new Proxy({}, { get: (_t, m) => (...args) => calls.push([a, m, ...args]) });
  const ws = { GetName: () => 'Data', GetRange: a => rec(a) };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (r, g, b) => `rgb(${r},${g},${b})` };
  assert.deepEqual(new Function('Api', cmd.body)(Api), { ok: true });
  assert.deepEqual(calls, [
    ['A1:B1', 'SetFillColor', 'rgb(0,200,0)'], ['A1:B1', 'SetWrap', true], ['A1:B1', 'SetNumberFormat', '0.00'],
    ['A1:B1', 'SetBold', true], ['A1:B1', 'SetFontSize', 14], ['A1:B1', 'SetFontColor', 'rgb(255,0,0)'],
    ['A1', 'SetAlignHorizontal', 'center'], ['A1', 'SetBorders', 'Top', 'Thin', 'rgb(0,0,255)'],
    ['B1', 'SetAlignHorizontal', 'center'], ['B1', 'SetBorders', 'Top', 'Thin', 'rgb(0,0,255)'],
  ]);
  assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ({ GetName: () => 'data' }) }), { ok: false, stage: 'sheet-identity' });
});

test('X-F01: the command body calls SetBold(false), not SetBold(true) or a skipped call -- the fake driver reads the spec directly and cannot catch a body-generation regression here', () => {
  const op = planOperations(task('A1', { font: { bold: false } })).operations[0];
  const cmd = format.lower(op);
  const calls = [];
  const rec = a => new Proxy({}, { get: (_t, m) => (...args) => calls.push([a, m, ...args]) });
  const ws = { GetName: () => 'Data', GetRange: a => rec(a) };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (r, g, b) => `rgb(${r},${g},${b})` };
  assert.deepEqual(new Function('Api', cmd.body)(Api), { ok: true });
  assert.deepEqual(calls, [['A1', 'SetBold', false]]);
});

test('X-F01: the command body calls SetItalic(false), not SetItalic(true) or a skipped call', () => {
  const op = planOperations(task('A1', { font: { italic: false } })).operations[0];
  const cmd = format.lower(op);
  const calls = [];
  const rec = a => new Proxy({}, { get: (_t, m) => (...args) => calls.push([a, m, ...args]) });
  const ws = { GetName: () => 'Data', GetRange: a => rec(a) };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (r, g, b) => `rgb(${r},${g},${b})` };
  assert.deepEqual(new Function('Api', cmd.body)(Api), { ok: true });
  assert.deepEqual(calls, [['A1', 'SetItalic', false]]);
});

test('X-F05: the command body calls SetUnderline/SetStrikeout at the RANGE level, and SetSubscript/SetSuperscript through GetCharacters(1, len).GetFont() per cell', () => {
  const op = planOperations(task('A1:B1', { font: { underline: 'double', strikeout: true, subscript: true } })).operations[0];
  const cmd = format.lower(op);
  const calls = [];
  const rec = a => new Proxy({}, { get: (_t, m) => {
    if (m === 'GetValue') return () => 'x';
    if (m === 'GetCharacters') return (...args) => { calls.push([a, 'GetCharacters', ...args]); return { GetFont: () => rec(`${a}.font`) }; };
    return (...args) => calls.push([a, m, ...args]);
  } });
  const ws = { GetName: () => 'Data', GetRange: a => rec(a) };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (r, g, b) => `rgb(${r},${g},${b})` };
  assert.deepEqual(new Function('Api', cmd.body)(Api), { ok: true });
  assert.deepEqual(calls, [
    ['A1:B1', 'SetUnderline', 'double'], ['A1:B1', 'SetStrikeout', true],
    ['A1', 'GetCharacters', 1, 1], ['A1.font', 'SetSubscript', true],
    ['B1', 'GetCharacters', 1, 1], ['B1.font', 'SetSubscript', true],
  ]);
});

test('X-F05: superscript:false compiles to SetSuperscript(false), not a skipped call -- the fake driver reads the spec directly and cannot catch a body-generation regression here', () => {
  const op = planOperations(task('A1', { font: { superscript: false } })).operations[0];
  const cmd = format.lower(op);
  const calls = [];
  const rec = a => new Proxy({}, { get: (_t, m) => {
    if (m === 'GetValue') return () => 'x';
    if (m === 'GetCharacters') return (...args) => { calls.push([a, 'GetCharacters', ...args]); return { GetFont: () => rec(`${a}.font`) }; };
    return (...args) => calls.push([a, m, ...args]);
  } });
  const ws = { GetName: () => 'Data', GetRange: a => rec(a) };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (r, g, b) => `rgb(${r},${g},${b})` };
  assert.deepEqual(new Function('Api', cmd.body)(Api), { ok: true });
  assert.deepEqual(calls, [['A1', 'GetCharacters', 1, 1], ['A1.font', 'SetSuperscript', false]]);
});

test('X-F07a: the command body calls SetOrientation PER CELL, through the same c-level alignment loop as SetAlignHorizontal/Vertical, not the range-level r', () => {
  const op = planOperations(task('A1:B1', { align: { rotation: 45 } })).operations[0];
  const cmd = format.lower(op);
  const calls = [];
  const rec = a => new Proxy({}, { get: (_t, m) => (...args) => calls.push([a, m, ...args]) });
  const ws = { GetName: () => 'Data', GetRange: a => rec(a) };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (r, g, b) => `rgb(${r},${g},${b})` };
  assert.deepEqual(new Function('Api', cmd.body)(Api), { ok: true });
  assert.deepEqual(calls, [['A1', 'SetOrientation', 45], ['B1', 'SetOrientation', 45]]);
});

test('the driver style read compiles on a stand-in: empty-cell font observable (X-F02), mixed-run fonts unobservable, align/border and align-ext through the DS Bridge', () => {
  const font = (b = null) => ({ GetBold: () => b, GetItalic: () => null, GetName: () => 'Calibri', GetSize: () => 11, GetColor: () => ({ GetRGB: () => 0 }), GetUnderline: () => 'xlUnderlineStyleNone', GetStrikethrough: () => null, GetSubscript: () => null, GetSuperscript: () => null });
  const mk = (a, v, chars) => ({
    GetAddress: () => a, GetValue: () => v, GetFormula: () => v, MergeArea: { GetAddress: () => a }, GetNumberFormat: () => 'General', GetFillColor: () => null, GetWrapText: () => false,
    GetCharacters: (s, n) => ({ GetFont: () => chars(s, n) }),
    // X-F07b/c (owner GO 2026-10-05): getReadingOrder/getIndent/getShrinkToFit on the SAME
    // getAlign() object align.h/v already reads -- readingOrder OMITTED here (undefined), the
    // measured untouched-cell shape, to confirm it maps to 0 ('context'), not UNOBSERVABLE.
    range: { getFont: () => ({ getBold: () => false }), getAlign: () => ({ getAlignHorizontal: () => 2, getAlignVertical: () => 4, getIndent: () => 3, getShrinkToFit: () => true }), getBorderFull: () => ({ t: { s: 7, c: { getRgb: () => 255 } }, b: { s: 0, c: null }, l: { s: 0, c: null }, r: { s: 0, c: null }, d: { s: 0, c: null }, dd: false, du: false }) },
  });
  // B1: empty, GetCharacters(1, 1) succeeds (X-F02, m7-s31b-probe1/2 MEASURED) -- the
  // same `chars` reader as A1's, since it ignores its own (start, len) args here.
  const cellsBy = { A1: mk('A1', 'ab', () => font()), B1: mk('B1', '', () => font()), C1: mk('C1', 'xy', (s, n) => font(s === 1 && n === 1 ? true : null)) };
  // B04 cutover (internal tracker point 8): align/border come through range.appearance.read/v1 of the REAL served bridge block,
  // which resolves its own cell with Api.GetSheet(sheet).GetRange(a1).range -- so a single-cell GetRange is the cell itself.
  const rng = a => (a === 'A1:C1' ? { GetAddress: () => a, ForEach: fn => ['A1', 'B1', 'C1'].forEach(x => fn(cellsBy[x])) } : { ...cellsBy[a], ForEach: fn => fn(cellsBy[a]) });
  const ws = { GetName: () => 'Data', GetIndex: () => 0, GetUsedRange: () => rng('A1:C1'), GetRange: rng };
  const Api = { GetSheets: () => [ws], GetSheet: () => ws, WorksheetFunction: { COUNT: () => 0 } };
  installBridge(Api);
  const read = (api = Api) => new Function('Api', readQuery([{ sheet: 'Data', range: 'A1:C1', style: true }]))(api).cells;
  const [a1, b1, c1] = read();
  const expectedFont = { bold: null, italic: null, name: 'Calibri', size: 11, color: 0, underline: 'xlUnderlineStyleNone', strike: null, sub: null, sup: null };
  assert.deepEqual(a1.font, expectedFont);
  assert.deepEqual(b1.font, expectedFont, 'X-F02: an empty cell\'s font is observable, same shape as an occupied cell\'s');
  assert.equal(c1.font, 'UNOBSERVABLE:mixed-runs');
  assert.deepEqual(a1.align, { h: 2, v: 4 });
  assert.deepEqual(a1.border.t, { s: 7, c: 255 });
  assert.equal(a1.wrap, false);
  assert.equal(a1.nf, 'General', 'style implies the format read');
  // X-F07b/c: readingOrder omitted by the stand-in (undefined) -> maps to the semantic default 0
  // ('context'); indent/shrink pass through their own getters.
  assert.equal(a1.readingOrder, 0, "an omitted getReadingOrder() maps to 'context' (0), never UNOBSERVABLE");
  assert.equal(a1.indent, 3);
  assert.equal(a1.shrink, true);
  const otherBuild = read(withPatch(Api))[0];
  assert.equal(otherBuild.align, 'UNOBSERVABLE:PATCH_VERSION_MISMATCH');
  assert.equal(otherBuild.border, 'UNOBSERVABLE:PATCH_VERSION_MISMATCH');
  for (const k of ['readingOrder', 'indent', 'shrink']) assert.equal(otherBuild[k], 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', `B05 READ: a /3 bridge refuses ${k}, no fallback`);
  // B05 READ cutover: ONE capability -- a gap in it makes all three align-ext axes unobservable, align.h/v untouched
  const noShrink = { ...Api, GetSheet: () => ({ ...ws, GetRange: a => ({ ...rng(a), range: { ...cellsBy.A1.range, getAlign: () => ({ getAlignHorizontal: () => 2, getAlignVertical: () => 4, getIndent: () => 3 }) } }) }) };
  installBridge(noShrink);
  const extGap = read(noShrink)[0];
  assert.deepEqual(extGap.align, { h: 2, v: 4 }, 'align.h/v unaffected by the align-ext capability');
  for (const k of ['readingOrder', 'indent', 'shrink']) assert.equal(extGap[k], 'UNOBSERVABLE:BRIDGE:shrink-missing', k);
  assert.equal('EO_INTERNAL_APPEARANCE_BUILDS' in drv, false, 'B04: MIGRATED to the DS Bridge');
  assert.equal('EO_INTERNAL_ALIGN_EXT_RUNTIMES' in drv, false, 'B05 READ: MIGRATED to the DS Bridge');

  // D1: empty, GetCharacters(1, 1) THROWS -- stays UNOBSERVABLE, never a guessed default.
  const throwCell = mk('D1', '', () => { throw new Error('boom'); });
  const wsD = { GetName: () => 'Data', GetIndex: () => 0, GetUsedRange: () => ({ GetAddress: () => 'D1', ForEach: fn => fn(throwCell) }), GetRange: () => ({ GetAddress: () => 'D1', ForEach: fn => fn(throwCell) }) };
  const ApiD = { GetSheets: () => [wsD], GetSheet: () => wsD, WorksheetFunction: { COUNT: () => 0 } };
  const d1 = new Function('Api', readQuery([{ sheet: 'Data', range: 'D1', style: true }], {}))(ApiD).cells[0];
  assert.equal(d1.font, 'UNOBSERVABLE:threw');
});

test('a border of the right style but ANOTHER colour is not SATISFIED: the colour is part of the proof', async () => {
  const b = book();
  b[0].cells.C4 = { v: null, border: { l: { s: 13, c: 16711680 } } };
  const store = createStore(b);
  const { result, m } = await run(store, task('C4', { border: { sides: ['left'], style: 'thick', color: '#000000' } }));
  assert.equal(result.ok, true);
  assert.equal(m.commands, 1, 'dispatched, not taken as already satisfied');
  assert.deepEqual(cell(store, 'C4').border.l, { s: 13, c: 0 });
});
