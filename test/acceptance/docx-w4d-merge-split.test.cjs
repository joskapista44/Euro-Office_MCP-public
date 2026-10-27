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

test('GH #291 T-07: a region that wholly contains a vertical merge (rows 0-1 x cols 0-1, and rows 0-2 x col 0) -> APPLIED, one owner of the region span, the paragraphs in row-major owner order; retry SATISFIED 0/0', async () => {
  const vInside = () => { const t = grid('V', 3); t.rows[0][0] = owner(['V r0c0', 'V r1c0'], 1, 2); t.rows[1][0] = cont(1); return t; };
  for (const [end, merged, rows, cols] of [['V r1c1', 'V r0c0 V r1c0 V r0c1 V r1c1', 2, 2], ['V r2c0', 'V r0c0 V r1c0 V r2c0', 3, 1]]) {
    const s = store([vInside()]);
    await firstAndSecond(s, task(mg('V r0c0 V r1c0', end, merged, rows, cols, 'V r2c2')));
    const owner = s.durableDoc.tables[0].rows[0][0];
    assert.deepEqual([owner.span || 1, owner.rowspan || 1, owner.paras.map(p => p.text).join(' ')], [cols, rows, merged], end);
  }
});

test('negatives before dispatch -> FAIL_CLOSED or plan-failed, 0 dispatch, nothing written', async () => {
  const withColspan = () => { const t = grid('C', 3); t.rows[0] = [owner(['C r0c0', 'C r0c1'], 2), cell('C r0c2')]; return t; };
  const dup = () => { const t = grid('A', 3); t.rows[2][1] = cell('A r1c0'); return t; };
  const failed = [
    // corners r0c2 .. r1c1: the box (rows 0-1, cols 1-2) cuts through the colspan cell (cols 0-1) of row 0
    [task(mg('C r0c2', 'C r1c1', 'C r0c2 C r1c1', 2, 2, 'C r2c2')), [withColspan()], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:merge-region-partially-covers-a-merged-cell'],
    // GH #291 T-07: a region WHOLLY containing a vertical merge is no longer a failure (positive test below)
    [task(mg('A r0c0', 'A r1c1', 'A r0c0 A r0c1 A r1c0 A r1c1', 2, 2, 'A r0c1')), [grid('A', 3)], 'table-anchor-in-merge-region'],
    [task(mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1 WRONG', 1, 2, 'A r2c2')), [grid('A', 3)], 'merge-terminal-contract-mismatch'],
    [task(mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 2, 2, 'A r2c2')), [grid('A', 3)], 'merge-terminal-contract-mismatch'],
    [task(mg('A r1c0', 'A r1c1', 'A r1c0 A r1c1', 1, 2, 'A r2c2')), [dup()], 'merge-cell-ambiguous'],
    [task(mg('A r1c0', 'nowhere', 'A r1c0 nowhere', 1, 2, 'A r2c2')), [grid('A', 3)], 'merge-terminal-unprovable'],
    [task(sp('C r0c0 C r0c1', 1, 3, 'C r2c2')), [withColspan()], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-shape-is-not-the-cell-span'],
    // W05-3: 'split-target-not-merged' is GONE -- a never-merged (1x1) cell now has its
    // own APPLY path (docx-w4d-merge-split.test.cjs's own new positive tests, below). What is
    // STILL refused for a never-merged target is an UNMEASURED geometry (same measured sets as
    // the restore case): 1x7 columns here (1-6 measured, GH #291 S2x).
    [task(sp('A r1c1', 1, 7, 'A r0c0')), [grid('A', 3)], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-boundary-model-not-measured-for-7-columns'],
    [task(sp('nowhere', 1, 2, 'A r0c0')), [grid('A', 3)], 'split-target-not-found'],
    // GH #291 H5m: a row short WITHOUT a marker is an implicit trailing gap now -- a region reaching into the short row's gap is not one rectangle of cells
    [task(mg('A r0c1', 'A r1c2', 'A r0c1 A r1c1 A r1c2', 2, 2, 'A r2c2')), [{ rows: [[cell('A r0c0'), cell('A r0c1')], ...grid('A', 3).rows.slice(1)] }], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:merge-region-covers-a-grid-gap'],
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

test('W05-3: split on a NEVER-MERGED cell grows the grid -- 1x2, 1x3, 2x1, 2x2 (measured on the fake transport, matching W05-1a/1b); SECOND 0/0', async () => {
  // MEASURED shapes (live-reproduced via a throwaway diagnostic against this same fake driver,
  // not guessed): a column split shifts same-row cells to the right and widens the untouched
  // rows' aligned cell by (cols-1); a row split inserts new row(s) with vContinue placeholders
  // for same-row neighbours (rowspan grown to cover them) and a fresh, independent blank cell at
  // the target's own column in each inserted row.
  const cases = [
    ['1x2', 1, 2,
      [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1', '', 'A r1c2'], ['A r2c0', 'A r2c1', 'A r2c2']],
      [['1x1', '1x2', '1x1'], ['1x1', '1x1', '1x1', '1x1'], ['1x1', '1x2', '1x1']]],
    ['1x3', 1, 3,
      [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1', '', '', 'A r1c2'], ['A r2c0', 'A r2c1', 'A r2c2']],
      [['1x1', '1x3', '1x1'], ['1x1', '1x1', '1x1', '1x1', '1x1'], ['1x1', '1x3', '1x1']]],
    ['2x1', 2, 1,
      [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1', 'A r1c2'], ['^', '', '^'], ['A r2c0', 'A r2c1', 'A r2c2']],
      [['1x1', '1x1', '1x1'], ['2x1', '1x1', '2x1'], ['1x1', '1x1', '1x1'], ['1x1', '1x1', '1x1']]],
    ['2x2', 2, 2,
      [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1', '', 'A r1c2'], ['^', '', '', '^'], ['A r2c0', 'A r2c1', 'A r2c2']],
      [['1x1', '1x2', '1x1'], ['2x1', '1x1', '1x1', '2x1'], ['1x1', '1x1', '1x1', '1x1'], ['1x1', '1x2', '1x1']]],
  ];
  for (const [label, rows, cols, wantTexts, wantSpans] of cases) {
    const s = store([grid('A', 3)]);
    const { first } = await firstAndSecond(s, task(sp('A r1c1', rows, cols, 'A r0c0')));
    assert.deepEqual(texts(s), wantTexts, label);
    assert.deepEqual(spans(s), wantSpans, label);
    assert.equal(first.x.primitiveVerification[0].evidence.evidence.wasMerged, false, label);
    assert.equal(first.x.preservationVerification[0].evidence.evidence.wasMerged, false, label);
  }
});

test('GH #291 S2 (LIVE, gh291-s-split-probe-20261022.json): 4 columns and 3 rows -- a restore of a 1x4 / 3x1 merged cell and a split of a plain cell APPLY; SECOND 0/0', async () => {
  const m14 = () => { const t = grid('M', 2, 4); t.rows[0] = [owner(['M r0c0', 'M r0c1', 'M r0c2', 'M r0c3'], 4)]; return t; };
  const m31 = () => { const t = grid('N', 4, 2); t.rows[0][0] = owner(['N r0c0', 'N r1c0', 'N r2c0'], 1, 3); t.rows[1][0] = cont(1); t.rows[2][0] = cont(1); return t; };
  const r14 = store([m14()]);
  await firstAndSecond(r14, task(sp('M r0c0 M r0c1 M r0c2 M r0c3', 1, 4, 'M r1c3')));
  assert.deepEqual(texts(r14)[0], ['M r0c0/M r0c1/M r0c2/M r0c3', '', '', ''], '1x4 restore: the content top-left, the grid unchanged');
  const r31 = store([m31()]);
  await firstAndSecond(r31, task(sp('N r0c0 N r1c0 N r2c0', 3, 1, 'N r3c1')));
  assert.deepEqual(texts(r31).map(r => r[0]), ['N r0c0/N r1c0/N r2c0', '', '', 'N r3c0'], '3x1 restore');
  assert.deepEqual(spans(r31).map(r => r[0]), ['1x1', '1x1', '1x1', '1x1']);
  const n14 = store([grid('A', 3)]);
  await firstAndSecond(n14, task(sp('A r1c1', 1, 4, 'A r0c0')));
  assert.deepEqual(texts(n14)[1], ['A r1c0', 'A r1c1', '', '', '', 'A r1c2'], '1x4 split of a plain cell: the grid grows');
  assert.deepEqual(spans(n14)[0], ['1x1', '1x4', '1x1'], 'the other rows\' cell of that column spans the new columns');
  const n31 = store([grid('A', 3)]);
  await firstAndSecond(n31, task(sp('A r1c1', 3, 1, 'A r0c0')));
  assert.deepEqual(texts(n31).map(r => r[1]), ['A r0c1', 'A r1c1', '', '', 'A r2c1'], '3x1 split of a plain cell: two new rows');
  assert.equal(spans(n31)[1][0], '3x1', 'the same-row neighbours span the new rows');
});

test('GH #291 S3 (LIVE, gh291-s-split-probe-20261022.json): a 2x2 merged cell split 1x2 / 2x1 gives two equal pieces (2x1 / 1x2), the content top-left, the grid unchanged; SECOND 0/0; a 1x3 cell -> 1x2 (a new grid line) stays refused', async () => {
  const m22 = tag => { const t = grid(tag, 3, 3); t.rows[0] = [owner([`${tag} r0c0`, `${tag} r0c1`, `${tag} r1c0`, `${tag} r1c1`], 2, 2), cell(`${tag} r0c2`)]; t.rows[1] = [cont(2), cell(`${tag} r1c2`)]; return t; };
  const target = tag => `${tag} r0c0 ${tag} r0c1 ${tag} r1c0 ${tag} r1c1`;
  const a = store([m22('P')]);
  const { first, second } = await firstAndSecond(a, task(sp(target('P'), 1, 2, 'P r2c2')));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED);
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED, 'the retry finds the measured piece block');
  assert.deepEqual(texts(a).slice(0, 2), [['P r0c0/P r0c1/P r1c0/P r1c1', '', 'P r0c2'], ['^', '^', 'P r1c2']]);
  assert.deepEqual(spans(a).slice(0, 2), [['2x1', '2x1', '1x1'], ['1x1', '1x1', '1x1']]);
  const b = store([m22('Q')]);
  await firstAndSecond(b, task(sp(target('Q'), 2, 1, 'Q r2c2')));
  assert.deepEqual(texts(b).slice(0, 2), [['Q r0c0/Q r0c1/Q r1c0/Q r1c1', 'Q r0c2'], ['', 'Q r1c2']]);
  assert.deepEqual(spans(b).slice(0, 2), [['1x2', '1x1'], ['1x2', '1x1']]);
  // not measured / measured unsafe: a 1x3 cell -> 1x2 (a half cut off the grid lines), a 2x2 cell -> 1x4
  for (const [tables, st] of [
    [[(() => { const t = grid('R', 2, 3); t.rows[0] = [owner(['R r0c0', 'R r0c1', 'R r0c2'], 3)]; return t; })()], sp('R r0c0 R r0c1 R r0c2', 1, 2, 'R r1c2')],
    [[m22('S')], sp(target('S'), 1, 4, 'S r2c2')],
  ]) {
    const s = store(tables);
    const before = JSON.stringify(texts(s));
    const { x, driver } = await run(s, task(st));
    assert.equal(reason(x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-shape-is-not-the-cell-span');
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(JSON.stringify(texts(s)), before);
  }
});

test('GH #291 S2x (LIVE, gh291-split-remaining-probe-20261023.json): 5 / 6 columns and 4 rows -- 1x5 / 1x6 / 2x5 / 4x1 restores and plain-cell splits APPLY; SECOND 0/0', async () => {
  const wide = (tag, rs, cs) => { const t = grid(tag, rs + 1, 6); const texts = []; for (let r = 0; r < rs; r += 1) for (let c = 0; c < cs; c += 1) texts.push(`${tag} r${r}c${c}`);
    t.rows[0] = [owner(texts, cs, rs), ...t.rows[0].slice(cs)]; for (let r = 1; r < rs; r += 1) t.rows[r] = [cont(cs), ...t.rows[r].slice(cs)]; return t; };
  const joined = (tag, rs, cs) => { const x = []; for (let r = 0; r < rs; r += 1) for (let c = 0; c < cs; c += 1) x.push(`${tag} r${r}c${c}`); return x; };
  for (const [tag, rs, cs] of [['X5', 1, 5], ['X6', 1, 6], ['XF', 2, 5], ['XB', 4, 1]]) {
    const s = store([wide(tag, rs, cs)]);
    await firstAndSecond(s, task(sp(joined(tag, rs, cs).join(' '), rs, cs, `${tag} r${rs}c5`)));
    assert.equal(texts(s)[0][0], joined(tag, rs, cs).join('/'), `${tag}: the content top-left`);
    assert.ok(spans(s).slice(0, rs).every(r => r.slice(0, cs).every(x => x === '1x1')), `${tag}: every piece 1x1 on the existing grid`);
  }
  const n15 = store([grid('A', 3)]);
  await firstAndSecond(n15, task(sp('A r1c1', 1, 5, 'A r0c0')));
  assert.deepEqual(spans(n15)[0], ['1x1', '1x5', '1x1'], '1x5 split of a plain cell: the grid grows the W05 way');
  const n41 = store([grid('A', 3)]);
  await firstAndSecond(n41, task(sp('A r1c1', 4, 1, 'A r0c0')));
  assert.equal(spans(n41)[1][0], '4x1', '4x1 split of a plain cell: the same-row neighbours span the new rows');
});

test('GH #291 S3x (LIVE, gh291-split-remaining-probe-20261023.json): 2x3 -> 1x3 / 2x1, 3x2 -> 3x1 / 1x2 give equal pieces on the existing grid; SECOND 0/0; a column cut off the grid lines and a 1x4 -> 1x2 (each piece the request\'s own shape: the retry is ambiguous) are refused', async () => {
  const block = (tag, rs, cs) => { const t = grid(tag, rs + 1, 4); const texts = []; for (let r = 0; r < rs; r += 1) for (let c = 0; c < cs; c += 1) texts.push(`${tag} r${r}c${c}`);
    t.rows[0] = [owner(texts, cs, rs), ...t.rows[0].slice(cs)]; for (let r = 1; r < rs; r += 1) t.rows[r] = [cont(cs), ...t.rows[r].slice(cs)]; return { t, text: texts.join(' ') }; };
  for (const [tag, rs, cs, r, c, piece] of [['PA', 2, 3, 1, 3, '2x1'], ['PB', 2, 3, 2, 1, '1x3'], ['PC', 3, 2, 3, 1, '1x2'], ['PD', 3, 2, 1, 2, '3x1']]) {
    const { t, text } = block(tag, rs, cs);
    const s = store([t]);
    const { second } = await firstAndSecond(s, task(sp(text, r, c, `${tag} r${rs}c3`)));
    assert.equal(status(second.x), OPERATION_STATUS.SATISFIED, `${tag}: the retry finds the measured piece block`);
    assert.equal(spans(s)[0][0], piece, `${tag}: the top-left piece`);
    assert.equal(texts(s)[0][0].split('/').length, rs * cs, `${tag}: all content in the top-left piece`);
  }
  // a 2x2 cell -> 1x2 (the released S3a) on UNEQUAL columns: the half cut is not the existing line -> refused, 0 dispatch (release 27)
  const u = block('PU', 2, 2);
  u.t.gridCalc = [30, 10, 10, 10];
  for (const [tables, st, want] of [
    [[u.t], sp(u.text, 1, 2, 'PU r2c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-would-cut-a-new-grid-line'],
    [[block('PE', 1, 4).t], sp(block('PE', 1, 4).text, 1, 2, 'PE r1c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-shape-is-not-the-cell-span'],
  ]) {
    const s = store(tables);
    const { x, driver } = await run(s, task(st));
    assert.equal(reason(x), want, want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
});

test('release 27 L1: a partial split whose pieces come out the wrong size (1x1) is not verified (the piece check)', async () => {
  const t = grid('PW', 3, 4); t.rows[0] = [owner(['PW r0c0', 'PW r0c1', 'PW r1c0', 'PW r1c1'], 2, 2), ...t.rows[0].slice(2)]; t.rows[1] = [cont(2), ...t.rows[1].slice(2)];
  const { x } = await run(store([t]), task(sp('PW r0c0 PW r0c1 PW r1c0 PW r1c1', 1, 2, 'PW r2c3')), { fault: 'split-partial-1x1-pieces' });
  assert.notEqual(status(x), OPERATION_STATUS.APPLIED);
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL');
});

test('GH #291 H5m (LIVE, gh291-gap-remaining-probe-20261023.json): merge on a table with a short row (implicit or gridAfter gap) APPLIES -- inside the short row, in the full row over the missing column, vertically at the short row\'s last real column; SECOND 0/0; a region into the gap and the split of a merged cell on a gap table are refused', async () => {
  const gapCell = () => ({ gap: true, span: 1, paras: [] });
  const u1 = tag => ({ rows: [[cell(`${tag} r0c0`), cell(`${tag} r0c1`), cell(`${tag} r0c2`)], [cell(`${tag} r1c0`), cell(`${tag} r1c1`), cell(`${tag} r1c2`), cell(`${tag} r1c3`)], [cell(`${tag} r2c0`), cell(`${tag} r2c1`), cell(`${tag} r2c2`), cell(`${tag} r2c3`)]] });
  const g1 = tag => { const t = u1(tag); t.rows[0].push(gapCell()); return t; };
  for (const [tag, mk, st, row] of [
    ['UM', u1, (t) => mg(`${t} r0c0`, `${t} r0c1`, `${t} r0c0 ${t} r0c1`, 1, 2, `${t} r2c3`), 0],
    ['UN', u1, (t) => mg(`${t} r1c2`, `${t} r1c3`, `${t} r1c2 ${t} r1c3`, 1, 2, `${t} r2c3`), 1],
    ['UO', u1, (t) => mg(`${t} r0c2`, `${t} r1c2`, `${t} r0c2 ${t} r1c2`, 2, 1, `${t} r2c3`), 0],
    ['GM', g1, (t) => mg(`${t} r0c0`, `${t} r0c1`, `${t} r0c0 ${t} r0c1`, 1, 2, `${t} r2c3`), 0],
    ['GN', g1, (t) => mg(`${t} r1c2`, `${t} r1c3`, `${t} r1c2 ${t} r1c3`, 1, 2, `${t} r2c3`), 1],
    ['GO', g1, (t) => mg(`${t} r0c2`, `${t} r1c2`, `${t} r0c2 ${t} r1c2`, 2, 1, `${t} r2c3`), 0],
  ]) {
    const s = store([mk(tag)]);
    await firstAndSecond(s, task(st(tag)));
    assert.ok(texts(s)[row].some(x => x.includes('/')), `${tag}: one merged owner holds both paragraphs`);
  }
  const lead = tag => ({ rows: [[gapCell(), cell(`${tag} r0c1`), cell(`${tag} r0c2`), cell(`${tag} r0c3`)], [cell(`${tag} r1c0`), cell(`${tag} r1c1`), cell(`${tag} r1c2`), cell(`${tag} r1c3`)]] });
  // run r2 (gh291-gap-remaining-probe-r2-20261023.json): a leading gap merges the same way; a NEW split (a 1x1 cell) on any short-row
  // table grows the grid the W05 way, aligned in every row
  for (const [tag, mk, st] of [
    ['LM', lead, t => mg(`${t} r1c1`, `${t} r1c2`, `${t} r1c1 ${t} r1c2`, 1, 2, `${t} r1c3`)],
    ['LN', lead, t => mg(`${t} r0c1`, `${t} r1c1`, `${t} r0c1 ${t} r1c1`, 2, 1, `${t} r1c3`)],
    ['SU', u1, t => sp(`${t} r1c1`, 1, 2, `${t} r2c3`)],
    ['SG', g1, t => sp(`${t} r1c1`, 1, 2, `${t} r2c3`)],
    ['SL', lead, t => sp(`${t} r1c1`, 1, 2, `${t} r1c3`)],
  ]) {
    const s = store([mk(tag)]);
    await firstAndSecond(s, task(st(tag)));
  }
  const mergedOnGap = (() => { const t = g1('SM'); t.rows[1] = [cell('SM r1c0'), owner(['SM r1c1', 'SM r1c2'], 2), cell('SM r1c3')]; return t; })();
  for (const [tables, st, want] of [
    [[g1('GX')], mg('GX r0c2', 'GX r1c3', 'GX r0c2 GX r1c2 GX r1c3', 2, 2, 'GX r2c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:merge-region-covers-a-grid-gap'],
    [[mergedOnGap], sp('SM r1c1 SM r1c2', 1, 2, 'SM r2c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-of-a-merged-cell-on-a-table-with-a-grid-gap'],
  ]) {
    const s = store(tables);
    const before = JSON.stringify(texts(s));
    const { x, driver } = await run(s, task(st));
    assert.equal(reason(x), want, want);
    assert.equal(driver.calls.dispatch, 0, want);
    assert.equal(JSON.stringify(texts(s)), before);
  }
});

test('W05-3: a NEVER-MERGED target rejects unmeasured geometries -- 5 rows and 7 columns are not generalized (GH #291 S2x measured 4 rows / 6 columns)', async () => {
  for (const [rows, cols, want] of [
    [5, 1, 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-boundary-model-not-measured-for-5-rows'],
    [1, 7, 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-boundary-model-not-measured-for-7-columns'],
  ]) {
    const s = store([grid('A', 3)]);
    const before = JSON.stringify(texts(s));
    const { x, driver } = await run(s, task(sp('A r1c1', rows, cols, 'A r0c0')));
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
    assert.equal(JSON.stringify(texts(s)), before, `${want}: nothing written`);
  }
});

test('W05-3 L3 fault: a new-split reformat is caught (the text-only, position-independent "outside" check still catches corruption)', async () => {
  const { x } = await run(store([grid('A', 3)]), task(sp('A r1c1', 1, 2, 'A r0c0')), { fault: 'split-reformat' });
  assert.equal(x.primitiveVerification[0].verdict, 'PASS', 'the structure itself is right');
  assert.equal(x.preservationVerification[0].verdict, 'FAIL');
  assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
});

// W05-5 (#ref, a maintainer's dispatch): FIXED -- the new-split L3 check is now POSITION-RESPECTING
// (verifyPreservation maps each outside owner to its MEASURED expected after-position/span/
// rowspan, point 0/6 of the design doc), replacing W05-3/W05-4's text-only multiset. This test
// WAS the W05-4 canary documenting the weakness (the swap went undetected, L3 PASSed); it now
// asserts the OPPOSITE, permanently, as a regression guard: if this ever flips back to PASS, the
// position-respecting check regressed to a text-only one, which must not happen silently.
test('W05-5: a new-split\'s position-respecting L3 catches two outside cells swapping text', async () => {
  const { x } = await run(store([grid('A', 3)]), task(sp('A r1c1', 1, 2, 'A r0c0')), { fault: 'outside-swap' });
  assert.equal(x.primitiveVerification[0].verdict, 'PASS', 'the structure itself is right');
  assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'the outside text-swap is now caught');
  assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
});

test('a merge whose L2 names a successor: annotateTask pairs the merge with the later split of its result only (same anchor, same text)', () => {
  const planned = planOperations(task(mg('R r1c0', 'R r1c1', 'R r1c0 R r1c1', 1, 2, 'R r2c2'), sp('R r1c0 R r1c1', 1, 2, 'R r2c2'), sp('other', 1, 2, 'R r2c2')));
  assert.equal(planned.ok, true);
  assert.deepEqual(planned.operations[0].successorSplit, { id: 's1', rows: 1, cols: 2 });
  const lone = planOperations(task(mg('R r1c0', 'R r1c1', 'R r1c0 R r1c1', 1, 2, 'R r2c2'), sp('R r1c0 R r1c1', 1, 2, 'Z anchor')));
  assert.equal(lone.operations[0].successorSplit, undefined, 'another anchor is not the successor');
});
