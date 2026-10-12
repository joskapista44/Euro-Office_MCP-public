'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const ms = require('../../src/adapters/docx/families/table-merge-split.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M5 gates 7 and 8 (owner decision 7A + 8A after 28880), offline on the fake transport whose model
 * follows the 2026-09-25 LIVE probes: a numbering's list type is PER LEVEL (set_list_type re-templates
 * one level); a column split cuts the merged cell at HALF of its calculated width. The internal
 * per-level format and calculated grid reads are the EO_INTERNAL_DEPENDENCY READ inputs of the
 * pre-dispatch guards; every refusal is 0 dispatch / 0 write for the refused step.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const P = 'Pee item.';
const Q = 'Queue item.';
const Z = 'Closing paragraph.';
const doc = () => [{ text: 'Head', level: 1, props: { ...PLAIN, bold: true } }, { text: P, props: { ...PLAIN } }, { text: Q, props: { ...PLAIN } }, { text: Z, props: { ...PLAIN } }];

async function run(store, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'm5-g78', steps });
const reason = (x, i) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const stateOf = (store, text) => { const p = store.durable.find(x => x.text === text); return p && p.list ? { type: p.list.type, level: p.list.level, numId: p.list.numId } : { type: 'normal' }; };

async function firstAndSecond(store, t, script = {}) {
  const first = await run(store, t, script);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => e.status)));
  assert.equal(first.m.writes, t.steps.length);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.primitiveVerification));
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', JSON.stringify(first.x.wholeTaskVerification));
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  const second = await run(store, t, script);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.stopDetail));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0], 'new-session retry: zero write');
  return { first, second };
}

// the exact LIVE reproduction of gate 7 (list_convert_then_indent) with a companion item in the same list
const CONVERT = [
  { id: 'lsP', intent: 'apply_list', selector: { text: P }, type: 'bullet', level: 0 },
  { id: 'lsQ', intent: 'apply_list', selector: { text: Q }, type: 'bullet', level: 0 },
  { id: 'lvP', intent: 'set_list_level', selector: { text: P }, level: 1, dependsOn: ['lsP'] },
  { id: 'tyP', intent: 'set_list_type', selector: { text: P }, type: 'numbered', dependsOn: ['lvP'] },
  { id: 'inP', intent: 'indent_list_item', selector: { text: P }, dependsOn: ['tyP'] },
];

// ---- gate 7: list type per level -----------------------------------------------------------------

test('gate 7: the task projection follows the per-level types of a numbering the task created (indent lands on a bullet level)', () => {
  const p = planOperations(task(...CONVERT));
  assert.equal(p.ok, true, p.detail);
  assert.deepEqual(p.operations.find(o => o.id === 'inP').terminal, { type: 'bullet', level: 2 });
  assert.deepEqual(p.operations.find(o => o.id === 'lsQ').terminal, { type: 'bullet', level: 0 });
  // a pre-existing numbering (no in-task apply_list): a level move projects NO type
  const q = planOperations(task({ intent: 'set_list_level', selector: { text: P }, level: 1 }));
  assert.deepEqual(q.operations[0].terminal, { type: null, level: 1 });
});

test('gate 7 (REPRODUCER, now PASS): convert -> numbered -> indent expects the DESTINATION level type (bullet@2); the companion item and the other levels are kept; retry 0 write', async () => {
  const store = createStore(doc());
  const { first } = await firstAndSecond(store, task(...CONVERT));
  assert.deepEqual(stateOf(store, P), { type: 'bullet', level: 2, numId: stateOf(store, Q).numId });
  assert.deepEqual(stateOf(store, Q).type, 'bullet');
  assert.equal(stateOf(store, Q).level, 0);
  const ty = first.x.preservationVerification[3].evidence.evidence;
  assert.equal(ty.otherLevels.applies, true);
  assert.deepEqual(ty.otherLevels.changed, [], 'set_list_type re-templated its own level only');
});

test('gate 7: a set_list_type that re-templates ANOTHER level of the numbering fails L3 (collateral level change)', async () => {
  const store = createStore(doc());
  const { x } = await run(store, task(...CONVERT.slice(0, 4)), { fault: 'list-retemplate-other-level', faultOn: [4] });
  const l3 = x.preservationVerification[3];
  assert.equal(l3.verdict === 'PASS', false, JSON.stringify(l3));
  assert.equal(l3.evidence.evidence.otherLevels.ok, false);
  assert.deepEqual(l3.evidence.evidence.otherListItems.ok, true, 'the collateral change sits on an item-less level: only the LEVEL axis can catch it');
});

test('gate 7 guard: reader absent / disagreeing / unmeasured destination format -> refused BEFORE dispatch, 0 write for the step', async () => {
  const cases = [
    [{ levelFormatsUnreadable: true }, 'eo-internal-list-level-format-unavailable', 2],
    [{ levelFormatsDisagree: true }, 'eo-internal-list-level-format-disagrees', 2],
    [{ levelFormatUnknownAt: 1 }, 'list-destination-level-type-unknown', 2],
  ];
  for (const [script, want, at] of cases) {
    const store = createStore(doc());
    const { x, driver } = await run(store, task(...CONVERT.slice(0, 3)), script);
    assert.equal(x.appliedOperations[at].status, OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x, at), want);
    assert.equal(driver.calls.dispatch, at, `${want}: only the ${at} steps before it dispatched`);
    assert.deepEqual(stateOf(store, P).level, 0, `${want}: the level move was not written`);
  }
  // CONTROL: the same reader faults do not touch a step that needs no level read (apply_list)
  const store = createStore(doc());
  const { x } = await run(store, task(CONVERT[0]), { levelFormatsUnreadable: true });
  assert.equal(x.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
});

// ---- gate 8: split on the existing grid ----------------------------------------------------------

const cell = text => ({ span: 1, paras: [{ text, props: { ...PLAIN } }] });
const grid = (tag, widths) => ({ rows: [0, 1, 2].map(r => [0, 1, 2].map(c => cell(`${tag} r${r}c${c}`))), ...(widths ? { gridCalc: widths } : {}) });
const tstore = table => createStore([{ text: 'Merge heading', level: 1, props: { ...PLAIN, bold: true } }], { doc: { tables: [table] } });
const MG = { intent: 'merge_table_cells', selector: { tableCell: { text: 'G r1c0', tableAnchor: 'G r2c2' } }, endSelector: { tableCell: { text: 'G r1c1' } }, terminalMergedText: 'G r1c0 G r1c1', terminalMergedSpan: { rows: 1, cols: 2 } };
const SP = (rows = 1, cols = 2) => ({ intent: 'split_table_cell', selector: { tableCell: { text: 'G r1c0 G r1c1', tableAnchor: 'G r2c2' } }, splitRows: rows, splitCols: cols });

test('gate 8 (POSITIVE): merge -> split on EQUAL spanned widths keeps the grid; retry 0 write', async () => {
  const s = tstore(grid('G', [14.586, 14.586, 18.9]));
  const { first } = await firstAndSecond(s, task(MG, SP()));
  assert.equal(first.x.appliedOperations[1].evidence.evidence.shape, true, 'L1: the grid shape is unchanged');
});

test('gate 8 (NEGATIVE REGRESSION): merge -> split on UNEQUAL spanned widths is refused BEFORE the split write (adapter boundary)', async () => {
  const s = tstore(grid('G', [40.27, 14.59, 18.9]));
  const { x, driver } = await run(s, task(MG, SP()));
  assert.equal(x.appliedOperations[0].status, OPERATION_STATUS.APPLIED, 'the merge itself is applied');
  assert.equal(x.appliedOperations[1].status, OPERATION_STATUS.FAILED);
  assert.equal(reason(x, 1), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-would-cut-a-new-grid-line');
  assert.equal(driver.calls.dispatch, 1, '0 dispatch / 0 write for the split');
  assert.deepEqual(x.appliedOperations[1].evidence.evidence.cuts.map(v => Math.round(v * 100) / 100), [27.43]);
});

test('gate 8 guard: grid unreadable / grid disagrees / an unmeasured column count -> refused before dispatch; a row-only split needs no column read', async () => {
  const cases = [
    [grid('G'), { gridUnreadable: true }, 'eo-internal-table-grid-unavailable'],
    [grid('G', [20, 20]), {}, 'eo-internal-table-grid-disagrees-with-logical-grid'],
  ];
  for (const [table, script, want] of cases) {
    const s = tstore(table);
    const { x, driver } = await run(s, task(MG, SP()), script);
    assert.equal(reason(x, 1), want);
    assert.equal(driver.calls.dispatch, 1, want);
  }
  // the model is measured for 1 and 2 columns only -- never generalized (owner decision 8A)
  const o = { r: 0, c: 0, rs: 1, cs: 3 };
  assert.equal(ms.splitGridGuard({ width: 3, gridCalc: [10, 10, 10] }, o, 3).reason, 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-boundary-model-not-measured-for-3-columns');
  assert.deepEqual(ms.splitGridGuard({ width: 3, gridCalc: null }, { r: 0, c: 0, rs: 2, cs: 1 }, 1), { ok: true, model: 'no-column-cut' });
});
