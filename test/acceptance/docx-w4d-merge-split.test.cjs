'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const ms = require('../../src/adapters/docx/families/table-merge-split.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 Wave 4d offline acceptance: table-merge-split (merge_table_cells, split_table_cell) through the
 * TARGET path on the fake transport whose model follows merge/split probe 1 (2026-09-24): a merged
 * owner holds the region's paragraphs in row-major order with their own props, the covered rows
 * keep one empty continuation cell; a split of the owner's span leaves all the content top-left.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (text, props = PLAIN, extra = {}) => ({ span: 1, paras: [{ text, props: { ...props } }], ...extra });
const grid = (tag, rows, cols = rows) => ({ rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (__, c) => cell(`${tag} r${r}c${c}`))) });
const cont = span => ({ span, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] });
const owner = (texts, span, rowspan) => ({ span, ...(rowspan > 1 ? { rowspan } : {}), paras: texts.map(t => ({ text: t, props: { ...PLAIN } })) });
const paragraphs = () => [{ text: 'Merge heading', level: 1, props: { ...PLAIN, bold: true } }, { text: 'Between.', props: { ...PLAIN } }];
const store = (tables = [grid('A', 3)], paras = paragraphs()) => createStore(paras, { doc: { tables } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w4d', steps });
const mg = (start, end, merged, rows, cols, tableAnchor) => ({ intent: 'merge_table_cells', selector: { tableCell: { text: start, tableAnchor } }, endSelector: { tableCell: { text: end } }, terminalMergedText: merged, terminalMergedSpan: { rows, cols } });
const sp = (target, rows, cols, tableAnchor) => ({ intent: 'split_table_cell', selector: { tableCell: { text: target, tableAnchor } }, splitRows: rows, splitCols: cols });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const texts = (s, t = 0) => s.durableDoc.tables[t].rows.map(r => r.map(c => (c.vContinue ? '^' : c.paras.map(p => p.text).join('/'))));
const spans = (s, t = 0) => s.durableDoc.tables[t].rows.map(r => r.map(c => `${c.rowspan || 1}x${c.span || 1}`));

async function firstAndSecond(s, t) {
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(first.m.writes, t.steps.length);
  assert.equal(first.m.commands, t.steps.length, 'FIRST_COMMANDS: one command per operation');
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.primitiveVerification));
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', JSON.stringify(first.x.wholeTaskVerification));
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  const second = await run(s, t);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.stopDetail || second.x.outcome || second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

test('merge shapes (probe M1-M7 + a region holding a WHOLE merged cell): owner placed with the region span, paragraphs row-major; SECOND 0/0', async () => {
  const withColspan = () => { const t = grid('C', 3); t.rows[0] = [owner(['C r0c0', 'C r0c1'], 2), cell('C r0c2')]; return t; };
  const withRowspan = () => { const t = grid('W', 3); t.rows[0][2] = owner(['W r0c2', 'W r1c2'], 1, 2); t.rows[1][2] = cont(1); return t; };
  const cases = [
    ['1x2', grid('A', 3), mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0/A r1c1', 'A r1c2'], ['A r2c0', 'A r2c1', 'A r2c2']]],
    ['2x1', grid('A', 3), mg('A r0c1', 'A r1c1', 'A r0c1 A r1c1', 2, 1, 'A r2c2'), [['A r0c0', 'A r0c1/A r1c1', 'A r0c2'], ['A r1c0', '^', 'A r1c2'], ['A r2c0', 'A r2c1', 'A r2c2']]],
    ['2x2', grid('A', 3), mg('A r0c0', 'A r1c1', 'A r0c0 A r0c1 A r1c0 A r1c1', 2, 2, 'A r2c2'), [['A r0c0/A r0c1/A r1c0/A r1c1', 'A r0c2'], ['^', 'A r1c2'], ['A r2c0', 'A r2c1', 'A r2c2']]],
    ['edge 2x2 (corners given bottom-right first)', grid('A', 3), mg('A r2c2', 'A r1c1', 'A r1c1 A r1c2 A r2c1 A r2c2', 2, 2, 'A r0c0'), [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1/A r1c2/A r2c1/A r2c2'], ['A r2c0', '^']]],
    ['middle of 4x4', grid('A', 4), mg('A r1c1', 'A r1c2', 'A r1c1 A r1c2', 1, 2, 'A r0c0'), null],
    ['colspan elsewhere', withColspan(), mg('C r2c1', 'C r2c2', 'C r2c1 C r2c2', 1, 2, 'C r0c2'), null],
    ['rowspan elsewhere', withRowspan(), mg('W r2c0', 'W r2c1', 'W r2c0 W r2c1', 1, 2, 'W r0c0'), null],
    ['whole merged cell inside', withColspan(), mg('C r0c0 C r0c1', 'C r1c1', 'C r0c0 C r0c1 C r1c0 C r1c1', 2, 2, 'C r2c2'), [['C r0c0/C r0c1/C r1c0/C r1c1', 'C r0c2'], ['^', 'C r1c2'], ['C r2c0', 'C r2c1', 'C r2c2']]],
  ];
  for (const [label, t, st, want] of cases) {
    const s = store([t]);
    await firstAndSecond(s, task(st));
    if (want) assert.deepEqual(texts(s), want, label);
  }
});

test('MERGE_FORMAT_PRESERVATION: each merged paragraph keeps its own runs and alignment (probe 1)', async () => {
  const t = grid('F', 3);
  t.rows[1][0] = { span: 1, paras: [{ text: 'F A', props: { ...PLAIN, bold: true, fontFamily: 'Courier New', fontSize: 28 }, alignment: 'center' }] };
  t.rows[1][1] = { span: 1, paras: [{ text: 'F B', props: { ...PLAIN, italic: true, fontFamily: 'Times New Roman', fontSize: 20 }, alignment: 'right' }] };
  const s = store([t]);
  await firstAndSecond(s, task(mg('F A', 'F B', 'F A F B', 1, 2, 'F r0c0')));
  const o = s.durableDoc.tables[0].rows[1][0];
  assert.deepEqual(o.paras.map(p => [p.text, p.props.fontFamily, p.props.bold, p.props.italic, p.alignment]), [['F A', 'Courier New', true, false, 'center'], ['F B', 'Times New Roman', false, true, 'right']]);
});

test('split: horizontal / vertical / 2x2 -> 1x1 cells, ALL content top-left, the rest empty; formatted content kept; SECOND 0/0', async () => {
  const h = () => { const t = grid('S', 3); t.rows[1] = [owner(['S r1c0', 'S r1c1'], 2), cell('S r1c2')]; return t; };
  const v = () => { const t = grid('S', 3); t.rows[0][1] = owner(['S r0c1', 'S r1c1'], 1, 2); t.rows[1][1] = cont(1); return t; };
  const q = () => { const t = grid('S', 3); t.rows[0] = [owner(['S r0c0', 'S r0c1', 'S r1c0', 'S r1c1'], 2, 2), cell('S r0c2')]; t.rows[1] = [cont(2), cell('S r1c2')]; return t; };
  const cases = [
    [h(), sp('S r1c0 S r1c1', 1, 2, 'S r0c0'), [['S r0c0', 'S r0c1', 'S r0c2'], ['S r1c0/S r1c1', '', 'S r1c2'], ['S r2c0', 'S r2c1', 'S r2c2']]],
    [v(), sp('S r0c1 S r1c1', 2, 1, 'S r2c2'), [['S r0c0', 'S r0c1/S r1c1', 'S r0c2'], ['S r1c0', '', 'S r1c2'], ['S r2c0', 'S r2c1', 'S r2c2']]],
    [q(), sp('S r0c0 S r0c1 S r1c0 S r1c1', 2, 2, 'S r2c2'), [['S r0c0/S r0c1/S r1c0/S r1c1', '', 'S r0c2'], ['', '', 'S r1c2'], ['S r2c0', 'S r2c1', 'S r2c2']]],
  ];
  for (const [t, st, want] of cases) {
    const s = store([t]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), want, JSON.stringify(st));
    assert.ok(spans(s).flat().every(x => x === '1x1'), 'every cell 1x1 after the split');
  }
  const f = h();
  f.rows[1][0].paras[0].props = { ...PLAIN, bold: true, fontFamily: 'Courier New' };
  const s = store([f]);
  await firstAndSecond(s, task(sp('S r1c0 S r1c1', 1, 2, 'S r0c0')));
  assert.equal(s.durableDoc.tables[0].rows[1][0].paras[0].props.fontFamily, 'Courier New');
});

test('MERGE -> SPLIT round trip in one session: the split resolves the merged cell from the fresh state; the merge is SATISFIED by its successor on retry', async () => {
  const s = store([grid('R', 3)]);
  const { second } = await firstAndSecond(s, task(mg('R r1c0', 'R r1c1', 'R r1c0 R r1c1', 1, 2, 'R r2c2'), sp('R r1c0 R r1c1', 1, 2, 'R r2c2')));
  assert.deepEqual(texts(s)[1], ['R r1c0/R r1c1', '', 'R r1c2'], 'the grid is back; the content stays top-left (measured, not redistributed)');
  assert.match(second.x.appliedOperations[0].evidence.proof, /successor-split/);
});

test('merge -> an independent next operation (row insert, the W4b family) on the fresh grid; after a VERTICAL merge the row family decides on the reconciled grid', async () => {
  const s = store([grid('N', 3)]);
  await firstAndSecond(s, task(mg('N r0c0', 'N r0c1', 'N r0c0 N r0c1', 1, 2, 'N r2c2'), { intent: 'insert_table_row_after', selector: { tableCell: { text: 'N r1c2', tableAnchor: 'N r2c2' } }, terminalRowCount: 4 }));
  assert.deepEqual(texts(s)[2], ['', '', ''], 'the row landed after row 1, not at the merged cell');
  const v = store([grid('N', 3)]);
  // M5: after a VERTICAL merge the row family decides on the reconciled grid -- a boundary
  // inside the merge is refused, a row outside it is inserted
  const { x } = await run(v, task(mg('N r0c0', 'N r1c0', 'N r0c0 N r1c0', 2, 1, 'N r2c2'), { intent: 'insert_table_row_after', selector: { tableCell: { text: 'N r0c2', tableAnchor: 'N r2c2' } }, terminalRowCount: 4 }));
  assert.equal(status(x, 0), OPERATION_STATUS.APPLIED);
  assert.equal(reason(x, 1), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-cuts-a-vertical-merge');
  const w = store([grid('N', 3)]);
  await firstAndSecond(w, task(mg('N r0c0', 'N r1c0', 'N r0c0 N r1c0', 2, 1, 'N r2c2'), { intent: 'insert_table_row_after', selector: { tableCell: { text: 'N r2c1', tableAnchor: 'N r0c2' } }, terminalRowCount: 4 }));
  assert.deepEqual(texts(w)[3], ['', '', ''], 'outside the merge: the new last row');
});

test('split -> an independent next operation (replace_table_cell) on another cell', async () => {
  const t = grid('P', 3); t.rows[1] = [owner(['P r1c0', 'P r1c1'], 2), cell('P r1c2')];
  const s = store([t]);
  await firstAndSecond(s, task(sp('P r1c0 P r1c1', 1, 2, 'P r0c0'), { intent: 'replace_table_cell', selector: { text: 'P r2c2' }, text: 'P NEW' }));
  assert.deepEqual(texts(s)[2], ['P r2c0', 'P r2c1', 'P NEW']);
});

test('MULTI_TABLE_TARGETING: the same texts in two tables -- the anchor picks one, the other is unchanged', async () => {
  const two = () => { const a = grid('X', 3); const b = grid('X', 3); b.rows[0][0] = cell('B anchor'); a.rows[2][2] = cell('A anchor'); return [a, b]; };
  const s = store(two());
  const other = texts(s, 1);
  await firstAndSecond(s, task(mg('X r1c0', 'X r1c1', 'X r1c0 X r1c1', 1, 2, 'A anchor'), sp('X r1c0 X r1c1', 1, 2, 'A anchor')));
  assert.deepEqual(texts(s, 1), other);
  const { x, driver } = await run(store(two()), task(mg('X r1c0', 'X r1c1', 'X r1c0 X r1c1', 1, 2, 'X r0c1')));
  assert.equal(reason(x), 'table-anchor-ambiguous');
  assert.equal(driver.calls.dispatch, 0);
});

test('negatives before dispatch -> FAIL_CLOSED or plan-failed, 0 dispatch, nothing written', async () => {
  const withColspan = () => { const t = grid('C', 3); t.rows[0] = [owner(['C r0c0', 'C r0c1'], 2), cell('C r0c2')]; return t; };
  const vInside = () => { const t = grid('V', 3); t.rows[0][0] = owner(['V r0c0', 'V r1c0'], 1, 2); t.rows[1][0] = cont(1); return t; };
  const dup = () => { const t = grid('A', 3); t.rows[2][1] = cell('A r1c0'); return t; };
  const failed = [
    // corners r0c2 .. r1c1: the box (rows 0-1, cols 1-2) cuts through the colspan cell (cols 0-1) of row 0
    [task(mg('C r0c2', 'C r1c1', 'C r0c2 C r1c1', 2, 2, 'C r2c2')), [withColspan()], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:merge-region-partially-covers-a-merged-cell'],
    [task(mg('V r0c0 V r1c0', 'V r1c1', 'V r0c0 V r1c0 V r0c1 V r1c1', 2, 2, 'V r2c2')), [vInside()], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:merge-region-contains-a-vertical-merge'],
    [task(mg('A r0c0', 'A r1c1', 'A r0c0 A r0c1 A r1c0 A r1c1', 2, 2, 'A r0c1')), [grid('A', 3)], 'table-anchor-in-merge-region'],
    [task(mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1 WRONG', 1, 2, 'A r2c2')), [grid('A', 3)], 'merge-terminal-contract-mismatch'],
    [task(mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 2, 2, 'A r2c2')), [grid('A', 3)], 'merge-terminal-contract-mismatch'],
    [task(mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2')), [dup()], 'merge-cell-ambiguous'],
    [task(mg('A r1c0', 'nowhere', 'A r1c0 nowhere', 1, 2, 'A r2c2')), [grid('A', 3)], 'merge-terminal-unprovable'],
    [task(sp('C r0c0 C r0c1', 1, 3, 'C r2c2')), [withColspan()], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-shape-is-not-the-cell-span'],
    [task(sp('A r1c1', 1, 2, 'A r0c0')), [grid('A', 3)], 'split-target-not-merged'],
    [task(sp('nowhere', 1, 2, 'A r0c0')), [grid('A', 3)], 'split-target-not-found'],
    [task(mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2')), [{ rows: [...grid('A', 3).rows.slice(0, 2), [cell('A r2c0'), cell('A r2c2')]] }], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:uneven-logical-width'],
  ];
  for (const [t, tables, want] of failed) {
    const s = store(tables);
    const before = JSON.stringify(texts(s));
    const { x, driver } = await run(s, t);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
    assert.equal(JSON.stringify(texts(s)), before, `${want}: nothing written`);
  }
  for (const st of [
    { ...mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), endSelector: undefined },
    { ...mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), endSelector: { tableCell: { text: 'A r1c1', tableAnchor: 'B other' } } },
    { ...mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), terminalMergedSpan: { rows: 1, cols: 1 } },
    { ...mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), terminalMergedText: undefined },
    { intent: 'merge_table_cells', selector: { text: 'A r1c0' }, endSelector: { text: 'A r1c1' }, terminalMergedText: 'A r1c0 A r1c1', terminalMergedSpan: { rows: 1, cols: 2 } },
    sp('A r1c1', 0, 2, 'A r0c0'), sp('A r1c1', -1, 2, 'A r0c0'), sp('A r1c1', 1.5, 2, 'A r0c0'), sp('A r1c1', 1, 21, 'A r0c0'), sp('A r1c1', 1, 1, 'A r0c0'),
    { intent: 'split_table_cell', selector: { text: 'A r1c1' }, splitRows: 1, splitCols: 2 },
  ]) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('faults at L1: merge ignored / content lost / paragraphs joined (same text) / wrong span; split ignored / paragraphs joined / content in the wrong cell / wrong grid -> never APPLIED', async () => {
  const h = () => { const t = grid('S', 3); t.rows[1] = [owner(['S r1c0', 'S r1c1'], 2), cell('S r1c2')]; return t; };
  for (const [fault, st, tables] of [
    ['ignore', mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), [grid('A', 3)]],
    ['merge-lose-content', mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), [grid('A', 3)]],
    ['merge-wrong-span', mg('A r0c0', 'A r0c2', 'A r0c0 A r0c1 A r0c2', 1, 3, 'A r2c2'), [grid('A', 3)]],
    ['merge-join-paragraphs', mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), [grid('A', 3)]],
    ['ignore', sp('S r1c0 S r1c1', 1, 2, 'S r0c0'), [h()]],
    ['split-join-paragraphs', sp('S r1c0 S r1c1', 1, 2, 'S r0c0'), [h()]],
    ['split-content-wrong-cell', sp('S r1c0 S r1c1', 1, 2, 'S r0c0'), [h()]],
    ['split-wrong-grid', sp('S r1c0 S r1c1', 1, 2, 'S r0c0'), [h()]],
  ]) {
    const { x, driver } = await run(store(tables), task(st), { fault });
    const tag = `${fault} ${st.intent}`;
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, tag);
    assert.equal(x.primitiveVerification[0].verdict, 'FAIL', tag);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE, tag);
    assert.equal(driver.calls.forceSave, 0, tag);
  }
});

test('faults at L3: a merged paragraph reformatted / an unrelated cell / another table; split reformat -> NOT APPLIED (P1)', async () => {
  const h = () => { const t = grid('S', 3); t.rows[1] = [owner(['S r1c0', 'S r1c1'], 2), cell('S r1c2')]; return t; };
  for (const [fault, st, tables] of [
    ['merge-reformat', mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), [grid('A', 3)]],
    ['merge-touch-cell', mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), [grid('A', 3)]],
    ['merge-touch-other-table', mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2'), [grid('A', 3), grid('B', 2)]],
    ['split-reformat', sp('S r1c0 S r1c1', 1, 2, 'S r0c0'), [h()]],
  ]) {
    const { x } = await run(store(tables), task(st), { fault });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', `${fault}: the structure itself is right`);
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN, fault);
  }
});

test('lossless lowering (MERGE_ARGS_PRESERVED / SPLIT_ARGS_PRESERVED); the command re-checks every coordinate by text; never Search / frame.evaluate', async () => {
  const s = store([grid('A', 3)]);
  const driver = createFakeDocxDriver(s);
  await driver.open({ fileId: 'd' });
  const primary = normalizeInventory((await driver.readLive()).payload);
  const op = planOperations(task(mg('A r0c0', 'A r1c1', 'A r0c0 A r0c1 A r1c0 A r1c1', 2, 2, 'A r2c2'))).operations[0];
  const target = ms.resolve(primary, op.selector, { operation: op }).target;
  const low = ms.lower(op, target);
  assert.deepEqual([low.meta.startText, low.meta.endText, low.meta.start, low.meta.end, low.meta.cells], ['A r0c0', 'A r1c1', [0, 0], [1, 1], [[0, 0], [0, 1], [1, 0], [1, 1]]]);
  assert.match(low.body, /t\.MergeCells\(cells\)/);
  assert.match(low.body, /textOf\(s\) !== m\.startText \|\| textOf\(e\) !== m\.endText/);
  const sop = planOperations(task(sp('A r0c0 A r0c1', 1, 2, 'A r2c2'))).operations[0];
  const slow = ms.lower(sop, { kind: 'table_split', tableAnchor: 'A r2c2', targetText: sop.targetText, cell: [0, 0], rows: 1, cols: 2 });
  assert.deepEqual([slow.meta.cell, slow.meta.rows, slow.meta.cols], [[0, 0], 1, 2]);
  assert.match(slow.body, /t\.Split\(cell, m\.rows, m\.cols\)/);
  for (const b of [low.body, slow.body]) assert.doesNotMatch(b, /\.Search\(|evaluate|asc_|SelectRange|private_/);
  // STALE ENDPOINT LOCATOR: the coordinates resolved before a row was inserted above -> the text re-check refuses, nothing changes
  await driver.dispatch(null, { kind: 'callCommand', meta: { family: 'table', action: 'row', rowAction: 'add', before: true, targetText: 'A r0c0', tableAnchor: 'A r2c2' } });
  await driver.dispatch(null, low);
  const now = normalizeInventory((await driver.readLive()).payload);
  assert.ok(now.document.tables[0].rows[1].map(c => c.text).includes('A r0c0'), 'the merge did not happen on the stale coordinates');
});

test('a merge whose L2 names a successor: annotateTask pairs the merge with the later split of its result only (same anchor, same text)', () => {
  const planned = planOperations(task(mg('R r1c0', 'R r1c1', 'R r1c0 R r1c1', 1, 2, 'R r2c2'), sp('R r1c0 R r1c1', 1, 2, 'R r2c2'), sp('other', 1, 2, 'R r2c2')));
  assert.equal(planned.ok, true);
  assert.deepEqual(planned.operations[0].successorSplit, { id: 's1', rows: 1, cols: 2 });
  const lone = planOperations(task(mg('R r1c0', 'R r1c1', 'R r1c0 R r1c1', 1, 2, 'R r2c2'), sp('R r1c0 R r1c1', 1, 2, 'Z anchor')));
  assert.equal(lone.operations[0].successorSplit, undefined, 'another anchor is not the successor');
});
