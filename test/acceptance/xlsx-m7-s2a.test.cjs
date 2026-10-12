'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { layoutDiffs } = require('../../src/adapters/xlsx/sheet-layout.cjs');
const { colName } = require('../../src/adapters/xlsx/a1.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { installBridge, withPatch, BRIDGE_RUNTIME_ID } = require('../fakes/ds-bridge-runtime.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 S2a (owner GO 2026-09-30, owner channel): a structural step is refused before the write when an AutoFilter, a
 * page break or a frozen pane of the shifted sheet is not BEFORE the cut on the step's axis, or when any sheet's
 * layout axis is unreadable (a split / unknown pane, the column breaks unread); the L3 keeps every sheet's layout
 * objects exactly. The READs are bound (owner decision, owner channel 2026-10-01, msg 30096): the D2 column-break READ
 * (since the B12 cutover through the DS Bridge sheet.columnBreaks.read/v1, asked by `colBreaks: true`), the strict row-break reader and the tri-state pane reader.
 * An automatic break (man false) is not measured: present on any sheet, the step refuses.
 */

const book = extra => {
  const cells = {};
  for (let i = 1; i <= 20; i += 1) cells[`A${i}`] = { v: `r${i}` };
  for (let c = 2; c <= 12; c += 1) cells[`${colName(c)}1`] = { v: `c${colName(c)}` };
  return [{ name: 'S', cells, ...extra }, { name: 'O', cells: { A1: { v: 'o' } } }];
};
const rowsStep = (intent, at, count = 1) => ({ intent, sheet: 'S', at, count, anchor: { column: 'A', value: `r${intent.startsWith('insert') ? at : at + count}` } });
const colsStep = (intent, at, count = 1) => ({ intent, sheet: 'S', at, count, anchor: { row: 1, value: `c${colName((at.charCodeAt(0) - 64) + (intent.startsWith('insert') ? 0 : count))}` } });
const task = step => ({ id: 'm7s2a', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}
const op0 = r => r.execution.appliedOperations[0];
const reason = r => { const o = op0(r); return o && o.evidence && o.evidence.evidence ? o.evidence.evidence.reason : null; };

async function applied(extra, step, script = {}) {
  const store = createStore(book(extra));
  const first = await run(store, task(step), script);
  assert.equal(first.result.ok, true, JSON.stringify([op0(first.result), first.result.execution.preservationVerification]));
  assert.equal(first.m.commands, 1);
  const again = await run(store, task(step), script);
  assert.equal(again.m.commands, 0);
  assert.equal(op0(again.result).status, OPERATION_STATUS.SATISFIED);
  return store;
}
async function refused(extra, step, want, script = {}) {
  const store = createStore(book(extra));
  const { result, m } = await run(store, task(step), script);
  assert.equal(result.ok, false);
  assert.equal(reason(result), want, JSON.stringify(op0(result)));
  assert.equal(m.commands, 0);
}

test('liveScope: a structural step reads every sheet\'s AutoFilter, print READ and pane', () => {
  const s = liveScope(task(rowsStep('insert_rows', 3)));
  assert.deepEqual(s.filterSheets, ['*']);
  assert.deepEqual(s.printSheets, ['*']);
  assert.equal(s.colBreaks, true);
  assert.equal(s.panes, true);
});

// ---- the index cases (plan section 6) ----
const NOT_BEFORE = 'xlsx-structure-layout-object-not-before-the-cut';
const F15 = { autoFilter: { range: 'A1:B5', filters: [] } };
const F1014 = { autoFilter: { range: 'A10:B14', filters: [] } };
const RB10 = { print: { rowBreaks: [{ id: 10, man: true }], colBreaks: [] } };
const CB2 = { print: { rowBreaks: [], colBreaks: [{ id: 2, man: true }] } };
const FR2 = { pane: { row0: 2, col0: 0 } };
const FC1 = { pane: { row0: 0, col0: 1 } };

const CASES = [
  ['filter A1:B5, insert_rows at 7 (BEFORE)', F15, rowsStep('insert_rows', 7), null],
  ['filter A1:B5, insert_rows at 6 (BOUNDARY)', F15, rowsStep('insert_rows', 6), NOT_BEFORE],
  ['filter A1:B5, insert_rows at 3 (INSIDE)', F15, rowsStep('insert_rows', 3), NOT_BEFORE],
  ['filter A10:B14, insert_rows at 5 (AFTER)', F1014, rowsStep('insert_rows', 5), NOT_BEFORE],
  ['filter A1:B5, delete_rows 7 (BEFORE)', F15, rowsStep('delete_rows', 7), null],
  ['filter A1:B5, delete_rows 6 (BOUNDARY)', F15, rowsStep('delete_rows', 6), NOT_BEFORE],
  ['filter A1:B5, delete_columns D (BEFORE on columns)', F15, colsStep('delete_columns', 'D'), null],
  ['filter A1:B5, delete_columns C (BOUNDARY on columns)', F15, colsStep('delete_columns', 'C'), NOT_BEFORE],
  ['filter A1:B5, insert_columns B (INSIDE on columns)', F15, colsStep('insert_columns', 'B'), NOT_BEFORE],
  ['row break id 10, insert_rows at 12 (BEFORE)', RB10, rowsStep('insert_rows', 12), null],
  ['row break id 10, insert_rows at 11 (BOUNDARY)', RB10, rowsStep('insert_rows', 11), NOT_BEFORE],
  ['row break id 10, delete_rows 11..12 (BOUNDARY)', RB10, rowsStep('delete_rows', 11, 2), NOT_BEFORE],
  ['row break id 10, insert_rows at 5 (AFTER)', RB10, rowsStep('insert_rows', 5), NOT_BEFORE],
  ['row break id 10, insert_columns C (the other axis: not classified)', RB10, colsStep('insert_columns', 'C'), null],
  ['column break id 2, delete_columns E (BEFORE)', CB2, colsStep('delete_columns', 'E'), null],
  ['column break id 2, delete_columns D (BEFORE)', CB2, colsStep('delete_columns', 'D'), null],
  ['column break id 2, delete_columns C (BOUNDARY)', CB2, colsStep('delete_columns', 'C'), NOT_BEFORE],
  ['column break id 2, insert_rows at 3 (the other axis: not classified)', CB2, rowsStep('insert_rows', 3), null],
  ['FreezeRows(2), insert_rows at 4 (BEFORE)', FR2, rowsStep('insert_rows', 4), null],
  ['FreezeRows(2), insert_rows at 3 (BOUNDARY)', FR2, rowsStep('insert_rows', 3), NOT_BEFORE],
  ['FreezeRows(2), insert_rows at 2 (INSIDE)', FR2, rowsStep('insert_rows', 2), NOT_BEFORE],
  ['FreezeRows(2), insert_columns B (col0 0, at column 2: BEFORE)', FR2, colsStep('insert_columns', 'B'), null],
  ['FreezeColumns(1), insert_rows at 2 (row0 0: BEFORE)', FC1, rowsStep('insert_rows', 2), null],
  ['FreezeColumns(1), insert_rows at 1 (row0 0, the frozen cell at row index 0)', FC1, rowsStep('insert_rows', 1), NOT_BEFORE],
];

for (const [name, extra, step, want] of CASES) {
  test(`${want ? 'refused 0/0' : 'APPLIED + reopened 0/0'}: ${name}`, async () => {
    if (want) await refused(extra, step, want);
    else {
      const store = await applied(extra, step);
      const S = store.durable.find(x => x.name === 'S');
      if (extra.autoFilter) assert.deepEqual(S.autoFilter, extra.autoFilter, 'the AutoFilter kept');
      if (extra.print) assert.deepEqual(S.print, extra.print, 'the breaks kept');
      if (extra.pane) assert.deepEqual(S.pane, extra.pane, 'the pane kept');
    }
  });
}

// ---- unreadable axes: refused (never an empty set), on the shifted sheet AND on another sheet ----
const UNOBS = 'xlsx-structure-layout-unobservable';
for (const [name, extra, script] of [
  ['the column-break READ build not proven', {}, { colBreakBuild: false }],
  ['the print READ build not proven', {}, { printBuild: false }],
  ['the pane READ build not proven', {}, { paneBuild: false }],
  ['the AutoFilter field READ not proven', {}, { filterBuild: false }],
  ['a split pane on the shifted sheet (the proposed tri-state reader)', { pane: { row0: 0, col0: 0, split: true } }, {}],
]) {
  test(`unobservable, refused 0/0: ${name}`, async () => {
    await refused(extra, rowsStep('insert_rows', 15), UNOBS, script);
  });
}

test('unobservable on ANOTHER sheet refuses too (the L3 claims every sheet)', async () => {
  const b = book({});
  b[1].pane = { row0: 0, col0: 0, split: true };
  const { result, m } = await run(createStore(b), task(rowsStep('insert_rows', 15)));
  assert.equal(reason(result), UNOBS);
  assert.equal(m.commands, 0);
});

// ---- an AUTOMATIC break is not measured: present on any sheet -> refused 0/0 (Talk 30096) ----
const AUTO = 'xlsx-structure-layout-automatic-break-not-measured';
test('an automatic row break on the shifted sheet, even BEFORE the cut, refuses 0/0', async () => {
  await refused({ print: { rowBreaks: [{ id: 3, man: false }], colBreaks: [] } }, rowsStep('insert_rows', 15), AUTO);
});
test('an automatic column break on ANOTHER sheet refuses 0/0', async () => {
  const b = book({});
  b[1].print = { rowBreaks: [], colBreaks: [{ id: 4, man: false }] };
  const { result, m } = await run(createStore(b), task(rowsStep('insert_rows', 15)));
  assert.equal(reason(result), AUTO);
  assert.equal(m.commands, 0);
});

// ---- L3 negative controls: a layout object moved as a side effect -> not APPLIED ----
for (const [fault, extra, axis, other] of [
  ['layout-moves-filter', F15, 'autofilter', false],
  ['layout-moves-row-break', RB10, 'row-breaks', false],
  ['layout-moves-col-break', CB2, 'column-breaks', false],
  ['layout-moves-pane', FR2, 'pane', false],
  ['layout-moves-row-break', null, 'row-breaks', true],
]) {
  test(`L3 catches ${fault}${other ? ' on ANOTHER sheet' : ''}`, async () => {
    const b = book(other ? {} : extra);
    if (other) b[1].print = RB10.print;
    const store = createStore(b);
    const { result } = await run(store, task(rowsStep('insert_rows', 15)), { fault, ...(other ? { layoutSheet: 'O' } : {}) });
    assert.equal(result.ok, false);
    assert.notEqual(op0(result).status, OPERATION_STATUS.APPLIED);
    assert.match(JSON.stringify(result.execution.preservationVerification), new RegExp(`"axis":"${axis}"`));
  });
}

// ---- the L3 comparison itself: identity + position, never a count ----
test('layoutDiffs: same count, another break position -> a difference', () => {
  const p = (breaks, pane = { state: 'none' }) => ({ sheets: [{ name: 'S' }], filterReads: { S: { range: null, filters: [] } }, printReads: { S: { rowBreaks: breaks, colBreaks: [] } }, panes: { S: pane } });
  assert.deepEqual(layoutDiffs(p([{ id: 3, man: true }]), p([{ id: 3, man: true }])), []);
  assert.deepEqual(layoutDiffs(p([{ id: 3, man: true }]), p([{ id: 4, man: true }])).map(d => d.axis), ['row-breaks']);
  assert.deepEqual(layoutDiffs(p([{ id: 3, man: true }]), p([{ id: 3, man: false }])).map(d => d.axis), ['row-breaks']);
  assert.deepEqual(layoutDiffs(p([]), p([], { state: 'frozen', row0: 1, col0: 0 })).map(d => d.axis), ['pane']);
  assert.deepEqual(layoutDiffs(p([]), p([], 'UNOBSERVABLE:x')).map(d => d.axis), ['unobservable']);
});

test('sheetLayout: only none / frozen with integer row0 / col0 are readable pane shapes', () => {
  const { sheetLayout } = require('../../src/adapters/xlsx/sheet-layout.cjs');
  const at = pane => sheetLayout({ sheets: [{ name: 'S' }], filterReads: { S: { range: null, filters: [] } }, printReads: { S: { rowBreaks: [], colBreaks: [] } }, panes: { S: pane } }, 'S').gaps.map(g => g.axis);
  assert.deepEqual(at({ state: 'none' }), []);
  assert.deepEqual(at({ state: 'frozen', row0: 2, col0: 0 }), []);
  for (const bad of [{ state: 'split' }, { state: 'frozen', row0: '2', col0: 0 }, { state: 'frozen', row0: -1, col0: 0 }, {}, null, 'UNOBSERVABLE:x']) assert.deepEqual(at(bad), ['pane'], JSON.stringify(bad));
});

test('sheetLayout: column breaks not asked / unproven / malformed are gaps, never an empty set', () => {
  const { sheetLayout } = require('../../src/adapters/xlsx/sheet-layout.cjs');
  const at = cb => sheetLayout({ sheets: [{ name: 'S' }], filterReads: { S: { range: null, filters: [] } }, printReads: { S: { rowBreaks: [], ...(cb === undefined ? {} : { colBreaks: cb }) } }, panes: { S: { state: 'none' } } }, 'S').gaps;
  assert.deepEqual(at([]), []);
  assert.deepEqual(at(undefined), [{ axis: 'column-breaks', why: 'not-read' }]);
  assert.deepEqual(at('UNOBSERVABLE:runtime-not-verified'), [{ axis: 'column-breaks', why: 'UNOBSERVABLE:runtime-not-verified' }]);
  assert.deepEqual(at([{ id: 2 }]).map(g => g.why), ['unexpected-shape'], 'man missing');
});

// ---- the D2 column-break READ in the frame (the M7 S2a READ probe's measured shapes) ----
test('the driver column-break READ (B12 cutover): only through the DS Bridge, asked only by colBreaks, the measured contract', () => {
  const ps = { asc_getFitToWidth: () => 0, asc_getFitToHeight: () => 0, asc_getScale: () => 100, asc_getWidth: () => 210, asc_getHeight: () => 297, useFirstPageNumber: false, firstPageNumber: null };
  const brk = list => ({ getBreaks: () => list });
  // the full model surface sheet.printSettings.read/v1 (B09, Wave B READ /4) needs; a gap there refuses the whole print read
  const hf = Object.fromEntries(['OddHeader', 'OddFooter', 'FirstHeader', 'FirstFooter', 'EvenHeader', 'EvenFooter', 'DifferentFirst', 'DifferentOddEven'].map(k => [`get${k}`, () => null]));
  const model = over => ({ PagePrintOptions: { asc_getPageSetup: () => ps, asc_getHorizontalCentered: () => null, asc_getVerticalCentered: () => null }, headerFooter: hf, rowBreaks: null, colBreaks: null, sheetPr: null, ...over });
  const mk = (name, m) => ({ GetName: () => name, GetIndex: () => 0, GetUsedRange: () => null, worksheet: m });
  const sheets = [mk('E', model({})),
    mk('M', model({ colBreaks: brk([{ id: 2, man: true, min: null, max: 1048575, pt: null }, { id: 7, man: true, min: null, max: 1048575, pt: null }]) })),
    mk('A', (() => { const m = model({}); delete m.colBreaks; return m; })()), mk('X', model({ colBreaks: {} })), mk('T', model({ colBreaks: { getBreaks: () => { throw new Error('boom'); } } })),
    mk('B', model({ colBreaks: brk([{ id: 16384, man: true }]) })), mk('W', model({ colBreaks: brk([{ id: 2, man: true, max: 1048576 }]) }))];
  const Api = { GetSheets: () => sheets, GetSheet: n => sheets.find(s => s.GetName() === n), WorksheetFunction: { COUNT: () => 0 } };
  const noBridge = new Function('Api', readQuery([], { printSheets: ['*'], colBreaks: true }))(Api).printReads;
  assert.equal(noBridge.M, 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', 'no bridge -> the whole print read (B09) is unobservable, never a direct model read');
  installBridge(Api);
  const r = new Function('Api', readQuery([], { printSheets: ['*'], colBreaks: true }))(Api).printReads;
  assert.deepEqual(r.E.colBreaks, [], 'a present null is a proven empty list');
  assert.deepEqual(r.M.colBreaks, [{ id: 2, man: true, min: null, max: 1048575 }, { id: 7, man: true, min: null, max: 1048575 }]);
  assert.equal(r.A.colBreaks, 'UNOBSERVABLE:BRIDGE:colBreaks-absent');
  assert.equal(r.X.colBreaks, 'UNOBSERVABLE:BRIDGE:colBreaks-unexpected-shape');
  assert.equal(r.T.colBreaks, 'UNOBSERVABLE:BRIDGE:colBreaks-unexpected-shape', 'a throwing getBreaks');
  assert.equal(r.B.colBreaks, 'UNOBSERVABLE:BRIDGE:colBreaks-element-invalid', 'a column id past XFD');
  assert.equal(r.W.colBreaks, 'UNOBSERVABLE:BRIDGE:colBreaks-element-invalid', 'a row span past the last row');
  const otherBuild = new Function('Api', readQuery([], { printSheets: ['*'], colBreaks: true }))(withPatch(Api)).printReads;
  assert.equal(otherBuild.M, 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', 'another patch: the print read itself refuses (B09)');
  assert.equal(new Function('Api', readQuery([], { printSheets: ['*'] }))(Api).printReads.M.colBreaks, undefined, 'not asked -> not read');
  assert.doesNotMatch(readQuery([], { printSheets: ['*'] }), /colBreaks'/, 'the column-break getter is not in the query unless asked');
});
