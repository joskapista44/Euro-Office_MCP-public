'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const row = require('../../src/adapters/docx/families/table-row.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 Wave 4b offline acceptance: table-row (insert_table_row_before / _after, delete_table_row)
 * through the TARGET path on the fake transport whose row model follows row probes 3-4: AddRow puts
 * one empty row of the target row's shape directly on the declared side, RemoveRow removes exactly
 * the target row; a ragged row keeps its own cell count; a vertical merge is a read cell state.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = text => ({ span: 1, paras: [{ text, props: { ...PLAIN } }] });
const grid = (tag, n, m = n) => ({ rows: Array.from({ length: n }, (_, r) => Array.from({ length: m }, (__, c) => cell(`${tag} r${r}c${c}`))) });
const paragraphs = () => [{ text: 'Row heading', level: 1, props: { ...PLAIN, bold: true } }, { text: 'Between.', props: { ...PLAIN } }];
const store = (tables = [grid('A', 3)], paras = paragraphs()) => createStore(paras, { doc: { tables } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  assert.equal(driver.calls.open, 1);
  assert.equal(driver.calls.close, 1);
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w4b', steps });
const ins = (side, text, terminalRowCount, tableAnchor) => ({ intent: `insert_table_row_${side}`, selector: tableAnchor ? { tableCell: { text, tableAnchor } } : { text }, terminalRowCount });
const del = (text, terminalRowCount, tableAnchor) => ({ intent: 'delete_table_row', selector: { tableCell: { text, tableAnchor } }, terminalRowCount });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const texts = (s, t = 0) => s.durableDoc.tables[t].rows.map(r => r.map(c => c.paras.map(p => p.text).join(' ')));

async function firstAndSecond(s, t) {
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(first.m.writes, t.steps.length);
  assert.equal(first.m.commands, t.steps.length, 'FIRST_COMMANDS: one command per row operation');
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), 'L1');
  assert.equal(first.x.preservationVerification.length, t.steps.length);
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', JSON.stringify(first.x.wholeTaskVerification));
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  const second = await run(s, t);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.stopDetail || second.x.outcome));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

const E = ['', '', ''];

test('rectangular 3x3: before/after on the FIRST, MIDDLE and LAST row; delete of each -- FIRST applies, SECOND is 0/0', async () => {
  const cases = [
    [ins('before', 'A r0c0', 4), [E, ['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1', 'A r1c2'], ['A r2c0', 'A r2c1', 'A r2c2']]],
    [ins('after', 'A r1c1', 4), [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1', 'A r1c2'], E, ['A r2c0', 'A r2c1', 'A r2c2']]],
    [ins('after', 'A r2c2', 4), [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1', 'A r1c2'], ['A r2c0', 'A r2c1', 'A r2c2'], E]],
    [ins('before', 'A r2c0', 4), [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1', 'A r1c2'], E, ['A r2c0', 'A r2c1', 'A r2c2']]],
    [del('A r0c1', 2, 'A r2c2'), [['A r1c0', 'A r1c1', 'A r1c2'], ['A r2c0', 'A r2c1', 'A r2c2']]],
    [del('A r1c1', 2, 'A r0c0'), [['A r0c0', 'A r0c1', 'A r0c2'], ['A r2c0', 'A r2c1', 'A r2c2']]],
    [del('A r2c1', 2, 'A r0c0'), [['A r0c0', 'A r0c1', 'A r0c2'], ['A r1c0', 'A r1c1', 'A r1c2']]],
  ];
  for (const [st, want] of cases) {
    const s = store();
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), want, JSON.stringify(st));
  }
});

test('ragged (a split row with 4 cells) and a colspan row: the new row takes the target row\'s shape', async () => {
  const t = grid('R', 3);
  t.rows[1] = [cell('R r1c0'), cell('R r1c1'), cell('R r1cX'), cell('R r1c2')];
  const s = store([t]);
  await firstAndSecond(s, task(ins('after', 'R r1c0', 4)));
  assert.equal(s.durableDoc.tables[0].rows[2].length, 4);
  const c = grid('C', 3);
  c.rows[1] = [{ ...cell('C r1c0 C r1c1'), span: 2 }, cell('C r1c2')];
  const s2 = store([c]);
  await firstAndSecond(s2, task(ins('before', 'C r2c0', 4)));
  assert.deepEqual([s2.durableDoc.tables[0].rows[1].length, s2.durableDoc.tables[0].rows[1][0].span, s2.durableDoc.tables[0].rows[2].length], [2, 2, 3]);
});

test('MULTI_TABLE_TARGETING: the same cell text in two tables -- the anchor picks one, the other is unchanged; no anchor -> FAIL_CLOSED', async () => {
  const two = () => { const b = grid('B', 2); b.rows[1][1] = cell('A r1c1'); return [grid('A', 3), b]; };
  const s = store(two());
  const before = texts(s, 1);
  await firstAndSecond(s, task(ins('after', 'A r1c1', 4, 'A r0c0')));
  assert.equal(s.durableDoc.tables[0].rows.length, 4);
  assert.deepEqual(texts(s, 1), before, 'the other table is unchanged');
  const s2 = store(two());
  await firstAndSecond(s2, task(del('A r1c1', 1, 'B r0c0')));
  assert.deepEqual([s2.durableDoc.tables[0].rows.length, s2.durableDoc.tables[1].rows.length], [3, 1]);
  const { x, driver } = await run(store(two()), task(ins('after', 'A r1c1', 4)));
  assert.equal(reason(x), 'table-row-target-ambiguous');
  assert.equal(driver.calls.dispatch, 0);
});

test('mixed one session: insert_before -> insert_after -> delete, each resolved from the FRESH state after the previous mutation; retry 0/0', async () => {
  const s = store([grid('M', 4)]);
  const t = task(ins('before', 'M r1c0', 5, 'M r0c0'), ins('after', 'M r2c0', 6, 'M r0c0'), del('M r3c0', 5, 'M r0c0'));
  const { first } = await firstAndSecond(s, t);
  // FRESH_STRUCTURAL_RETARGETING: every later target row moved because of an earlier insert
  assert.deepEqual(first.x.primitiveVerification.map(v => v.evidence && v.evidence.evidence && v.evidence.evidence.targetRowBefore), [1, 3, 5]);
  assert.deepEqual(texts(s).map(r => r[0]), ['M r0c0', '', 'M r1c0', 'M r2c0', '']);
});

test('delete retry identity: target absent + terminal count in the anchor table -> SATISFIED; recovery -> FAIL_CLOSED', async () => {
  const s = store();
  await run(s, task(del('A r1c1', 2, 'A r0c0')));
  const r = await run(s, task(del('A r1c1', 2, 'A r0c0')));
  assert.equal(status(r.x), OPERATION_STATUS.SATISFIED);
  const rec = await run(s, { ...task(del('A r1c1', 2, 'A r0c0')), recovery: { unresolvedUncertainOperations: ['s0'] } });
  assert.equal(reason(rec.x), 'preservation-baseline-unavailable');
  assert.equal(rec.driver.calls.dispatch, 0);
});

test('negatives before dispatch -> FAIL_CLOSED or plan-failed, 0 dispatch', async () => {
  const dup = grid('A', 3); dup.rows[2][2] = cell('A r1c1');
  // a vertical merge in the LIVE projection's shape (row probe 4): rowspan on the top, the continuation cell omitted from the HTML
  const vm = grid('A', 3); vm.rows[0][2] = { ...vm.rows[0][2], rowspan: 2 }; vm.rows[1][2] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] };
  const twoAnchor = () => { const b = grid('B', 2); b.rows[0][0] = cell('A r0c0'); return [grid('A', 3), b]; };
  const failed = [
    [task(ins('after', 'A r1c1', 4)), store([dup]), {}, 'table-row-target-ambiguous'],
    [task(ins('after', 'A r1c1', 4, 'A r0c0')), store(twoAnchor()), {}, 'table-anchor-ambiguous'],
    [task(ins('after', 'nowhere', 4)), store(), {}, 'table-row-target-not-found'],
    [task(ins('after', 'A r1c1', 9)), store(), {}, 'table-row-terminal-unprovable'],
    [task(del('A r1c1', 1, 'A r0c0')), store(), {}, 'delete-target-present-but-row-count-not-terminal-plus-one'],
    [task(del('A r1c1', 2, 'A r1c0')), store(), {}, 'table-anchor-in-target-row'],
    // W06-3: the per-table blanket narrowed to per-row -- 'A r0c0' is row 0, INSIDE the
    // vm fixture's own merge span [0,1] (col 2), so this STILL refuses, with the new, row-delete-
    // specific reason (distinct from the insert-side reasons, never conflated).
    [task(del('A r0c0', 2, 'A r1c1')), store([vm]), {}, 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-delete-inside-a-vertical-merge'],
    // M5: on a vertically merged table the INSERT is decided by the reconciled grid
    [task(ins('after', 'A r1c1', 4)), store([vm]), {}, 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-target-row-inside-a-vertical-merge'],
    [task(ins('before', 'A r1c1', 4)), store([vm]), {}, 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-cuts-a-vertical-merge'],
    [task(ins('after', 'A r0c0', 4)), store([vm]), {}, 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-cuts-a-vertical-merge'],
    [task(ins('before', 'A r0c0', 4, 'A r2c2'), ins('after', 'A r2c2', 5)), store(), {}, 'row-operations-without-table-anchor-cannot-be-grouped'],
  ];
  for (const [t, s, script, want] of failed) {
    const { x, driver } = await run(s, t, script);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
  for (const st of [
    { intent: 'delete_table_row', selector: { text: 'A r1c1' }, terminalRowCount: 2 },
    { ...ins('after', 'A r1c1', 4), terminalRowCount: undefined },
    { ...ins('after', 'A r1c1', 1) },
    { intent: 'insert_table_row_after', selector: { tableCell: { text: 'A r1c1', tableAnchor: 'A r1c1' } }, terminalRowCount: 4 },
    { ...del('A r1c1', 2, 'A r0c0'), terminalTableAnchor: 'A r2c2' },
    { intent: 'insert_table_row_after', selector: { text: 'A r1c1', index: 1 }, terminalRowCount: 4 },
  ]) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('the source form is accepted: combined selector + terminalRowCount (+ terminalTableAnchor for delete)', async () => {
  await firstAndSecond(store(), task({ intent: 'insert_table_row_after', selector: { tableCell: { text: 'A r1c1' }, text: 'A r1c1' }, terminalRowCount: 4 }));
  await firstAndSecond(store(), task({ intent: 'delete_table_row', selector: { tableCell: { text: 'A r1c1' }, text: 'A r1c1' }, terminalRowCount: 2, terminalTableAnchor: 'A r0c0' }));
});

test('faults at L1: wrong side / twice / wrong shape / not empty / reordered / delete ignored / other row deleted / stale index -> never APPLIED', async () => {
  for (const [fault, st] of [
    ['row-wrong-side', ins('before', 'A r1c1', 4)], ['row-wrong-side', ins('after', 'A r1c1', 4)], ['row-twice', ins('after', 'A r1c1', 4)],
    ['row-new-shape', ins('after', 'A r1c1', 4)], ['row-new-not-empty', ins('after', 'A r1c1', 4)], ['row-reorder', ins('before', 'A r0c0', 4)],
    ['ignore', del('A r1c1', 2, 'A r0c0')], ['row-delete-other', del('A r1c1', 2, 'A r0c0')], ['row-stale-index', ins('after', 'A r1c1', 4)],
    ['row-stale-index', del('A r2c1', 2, 'A r0c0')],
  ]) {
    const { x, driver } = await run(store(), task(st), { fault });
    const tag = `${fault} ${st.intent}`;
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, tag);
    assert.equal(x.primitiveVerification[0].verdict, 'FAIL', tag);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE, tag);
    assert.equal(driver.calls.forceSave, 0, tag);
  }
});

test('faults at L3: an unrelated cell / another table / text outside the tables -> NOT APPLIED (P1)', async () => {
  const two = () => [grid('A', 3), grid('B', 2)];
  for (const [fault, st, tables] of [
    ['row-touch-cell', ins('after', 'A r1c1', 4), undefined],
    ['row-touch-other-table', ins('after', 'A r1c1', 4, 'A r0c0'), two()],
    ['row-change-outside', del('A r1c1', 2, 'A r0c0'), undefined],
  ]) {
    const { x } = await run(store(tables), task(st), { fault });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', `${fault}: the row itself is right`);
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN, fault);
  }
});

test('the command finds the table by its anchor and the cell by its full text, then AddRow / RemoveRow -- never Search, never frame.evaluate', () => {
  const body = row.commandBody({ family: 'table', action: 'row', rowAction: 'add', before: true, targetText: 'x', tableAnchor: 'a' });
  assert.match(body, /\.AddRow\(hits\[0\]\.cell, m\.before === true\)/);
  assert.match(row.commandBody({ family: 'table', action: 'row', rowAction: 'remove', before: null, targetText: 'x', tableAnchor: 'a' }), /\.RemoveRow\(/);
  assert.doesNotMatch(body, /\.Search\(|evaluate|asc_|addRowAbove|remRow/);
});

test('mixed: a partial first run (the 2nd operation did not happen) -> the retry SATISFIES the 1st and APPLIES the rest', async () => {
  const s = store([grid('M', 4)]);
  const t = task(ins('before', 'M r1c0', 5, 'M r0c0'), ins('after', 'M r2c0', 6, 'M r0c0'), del('M r3c0', 5, 'M r0c0'));
  const first = await run(s, t, { fault: 'ignore', faultOn: [2] });
  assert.equal(status(first.x, 0), OPERATION_STATUS.APPLIED);
  assert.notEqual(status(first.x, 1), OPERATION_STATUS.APPLIED);
  // the uncertain prefix persisted (the first insert); the retry resolves every operation from the fresh durable state
  const second = await run(s, t);
  assert.deepEqual(second.x.appliedOperations.map(e => e.status), [OPERATION_STATUS.SATISFIED, OPERATION_STATUS.APPLIED, OPERATION_STATUS.APPLIED]);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  assert.deepEqual(texts(s).map(r => r[0]), ['M r0c0', '', 'M r1c0', 'M r2c0', '']);
});

test('group final state reached but THIS operation\'s claim missing -> FAIL_CLOSED, not a blind re-apply', () => {
  const planned = require('../../src/adapters/docx/adapter.cjs').planOperations(task(ins('after', 'M r0c0', 4, 'M r2c0'), del('M r1c0', 3, 'M r2c0')));
  assert.equal(planned.ok, true);
  const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
  const { htmlOf } = require('../fakes/fake-docx-driver.cjs');
  // 3 rows, M r1c0 gone (the delete's claim holds), and no empty row after M r0c0 (the insert's claim does not)
  const tables = [{ rows: [[cell('M r0c0')], [cell('M r2c0')], [cell('M r3c0')]] }];
  const primary = normalizeInventory({ html: htmlOf([], [], { tables }), paragraphs: [], document: { tables: tables.map(t => ({ rows: t.rows.map(r => r.map(c => ({ text: `${c.paras[0].text}\t`, paragraphs: 1, simple: true, vMerge: 'none', format: {} }))) })) } });
  const c = row.classify(primary, planned.operations[0]);
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'group-final-state-without-this-operation-claim');
});

// ---- M5 (owner decision after Talk 28963, internal tracker): row insert on a table with a vertical merge

/** A 4x4 table whose column 1 is merged over rows 1-2 (the LIVE projection's shape: rowspan on the
 *  owner, the continuation cell omitted from the HTML, kept as an empty cell by the driver read). */
const vm4 = () => { const t = grid('V', 4); t.rows[1][1] = { ...cell('V r1c1'), rowspan: 2 }; t.rows[2][1] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }; return t; };
const mergeKept = (s, ownerRow) => {
  const rows = s.durableDoc.tables[0].rows;
  assert.equal(rows[ownerRow][1].rowspan, 2, 'the merge owner keeps its rowspan');
  assert.equal(rows[ownerRow + 1][1].vContinue, true, 'the continuation cell stays under it');
  assert.equal(rows[ownerRow][1].paras[0].text, 'V r1c1', 'the merged content is kept');
};

test('M5 POSITIVE: a row insert before / after a vertical merge (outside it, both directions) -> APPLIED, content and merge kept; retry SATISFIED 0/0', async () => {
  const cases = [
    [ins('before', 'V r0c0', 5, 'V r3c3'), 0, 1 + 1], // new row 0; the owner moves to row 2
    [ins('after', 'V r0c0', 5, 'V r3c3'), 1, 2],      // new row 1 between row 0 and the merge
    [ins('before', 'V r3c0', 5), 3, 1],               // new row 3 right after the merge's bottom
    [ins('after', 'V r3c0', 5), 4, 1],                // new last row
  ];
  for (const [step, newRow, ownerRow] of cases) {
    const s = store([vm4()]);
    await firstAndSecond(s, task(step));
    const t = texts(s);
    assert.equal(t.length, 5, JSON.stringify(step));
    assert.deepEqual(t[newRow], ['', '', '', ''], `${JSON.stringify(step)}: the new empty row at ${newRow}`);
    mergeKept(s, ownerRow);
  }
});

// M5 (owner decision A, owner channel after 28972): the anchor-in-target-row refusal is for DELETE only.
// Three shapes, the anchor 'A r2c2' in row 2: plain | a colspan in row 1 | a vertical merge (column 1,
// rows 0-1, outside the new row's boundary on both sides).
const anchorShapes = () => {
  const colspan = grid('A', 3); colspan.rows[1] = [{ ...cell('A r1c0 A r1c1'), span: 2 }, cell('A r1c2')];
  const vmerge = grid('A', 3); vmerge.rows[0][1] = { ...vmerge.rows[0][1], rowspan: 2 }; vmerge.rows[1][1] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] };
  return [['plain', grid('A', 3)], ['colspan', colspan], ['vmerge', vmerge]];
};

test('M5 POSITIVE: an insert in the anchor\'s own row (before / after) on plain / colspan / vertical-merge -> APPLIED, every original row (content, spans, merges) kept; a fresh session SATISFIED 0 write / 0 command', async () => {
  for (const [shape, table] of anchorShapes()) {
    for (const [step, newRow] of [[ins('before', 'A r2c0', 4, 'A r2c2'), 2], [ins('after', 'A r2c0', 4, 'A r2c2'), 3]]) {
      const s = store([JSON.parse(JSON.stringify(table))]);
      const before = JSON.stringify(s.durableDoc.tables[0].rows);
      await firstAndSecond(s, task(step)); // the second run is a new session: SATISFIED, [writes, commands, forceSave] = [0, 0, 0]
      const rows = s.durableDoc.tables[0].rows;
      assert.equal(rows.length, 4, `${shape} ${JSON.stringify(step)}`);
      assert.deepEqual(texts(s)[newRow], ['', '', ''], `${shape}: the new empty row at ${newRow}`);
      assert.equal(JSON.stringify(rows.filter((_, i) => i !== newRow)), before, `${shape}: every original row, its content, spans and merges unchanged, in order`);
      assert.equal(texts(s).flat().filter(x => x === 'A r2c2').length, 1, `${shape}: the anchor stays unique`);
    }
  }
});

test('M5 NEGATIVE CONTROL: the DELETE of the anchor\'s own row (same anchor, same shapes) still fails before any write, 0 dispatch, the document unchanged', async () => {
  for (const [shape, table] of anchorShapes()) {
    const s = store([table]);
    const before = JSON.stringify(s.durableDoc);
    const d = await run(s, task(del('A r2c0', 2, 'A r2c2')));
    assert.equal(status(d.x), OPERATION_STATUS.FAILED, shape);
    assert.equal(reason(d.x), 'table-anchor-in-target-row', shape);
    assert.deepEqual([d.driver.calls.dispatch, d.m.writes, d.m.commands], [0, 0, 0], shape);
    assert.equal(JSON.stringify(s.durableDoc), before, `${shape}: nothing written`);
  }
});

test('GH #291 T-06: deleting a CONTINUATION row of a vertical merge -> APPLIED, the owner keeps its content, its rowspan shrinks (the two-row merge ends); retry SATISFIED 0/0; an editor that leaves the rowspan -> never APPLIED', async () => {
  const s = store([vm4()]);
  const first = await run(s, task(del('V r2c0', 3, 'V r3c3')));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED, JSON.stringify(first.x.appliedOperations.map(e => e.evidence)));
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  const rows = s.durableDoc.tables[0].rows;
  assert.equal(rows.length, 3);
  assert.equal(rows[1][1].paras[0].text, 'V r1c1', 'the merged content is kept');
  assert.equal(rows[1][1].rowspan, undefined, 'the two-row merge ended');
  assert.equal(rows.flat().some(c => c.vContinue), false);
  const second = await run(s, task(del('V r2c0', 3, 'V r3c3')));
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([second.m.writes, second.m.commands], [0, 0]);
  const bad = await run(store([vm4()]), task(del('V r2c0', 3, 'V r3c3')), { fault: 'row-delete-merge-not-shortened' });
  assert.notEqual(status(bad.x), OPERATION_STATUS.APPLIED);
});

test('GH #291 T-06: the MIDDLE continuation row of a three-row merge -> APPLIED, the owner rowspan 3 -> 2, the content kept', async () => {
  const t = grid('V', 4); t.rows[0][1] = { ...cell('V r0c1'), rowspan: 3 };
  for (const r of [1, 2]) t.rows[r][1] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] };
  const s = store([t]);
  const r = await run(s, task(del('V r1c0', 3, 'V r3c3')));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => e.evidence)));
  assert.ok(r.x.preservationVerification.every(v => v.verdict === 'PASS'));
  const rows = s.durableDoc.tables[0].rows;
  assert.deepEqual([rows.length, rows[0][1].rowspan, rows[1][1].vContinue, rows[0][1].paras[0].text], [3, 2, true, 'V r0c1']);
});

test('GH #291 T-05: extendMerge -- an insert INSIDE a merge (before / after the boundary that cuts it) -> APPLIED, the merge grows by one row, its content kept; retry SATISFIED 0/0; an editor that leaves the merge -> never APPLIED', async () => {
  for (const [side, text] of [['after', 'V r1c0'], ['before', 'V r2c0']]) {
    const s = store([vm4()]);
    const step = { ...ins(side, text, 5, 'V r3c3'), extendMerge: true };
    const first = await run(s, task(step));
    assert.equal(status(first.x), OPERATION_STATUS.APPLIED, `${side}: ${JSON.stringify(first.x.appliedOperations.map(e => e.evidence))}`);
    assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
    const rows = s.durableDoc.tables[0].rows;
    assert.deepEqual([rows.length, rows[1][1].rowspan, rows[1][1].paras[0].text, rows[2][1].vContinue, rows[3][1].vContinue], [5, 3, 'V r1c1', true, true]);
    const second = await run(s, task(step));
    assert.equal(status(second.x), OPERATION_STATUS.SATISFIED);
    assert.deepEqual([second.m.writes, second.m.commands], [0, 0]);
    const bad = await run(store([vm4()]), task(step), { fault: 'row-insert-merge-not-extended' });
    assert.notEqual(status(bad.x), OPERATION_STATUS.APPLIED);
  }
});

test('GH #291 T-05: extendMerge on an insert that cuts NO merge (the merge edge, or a plain row) -> refused before dispatch, 0/0; extendMerge on a delete or as false -> refused at plan', async () => {
  for (const step of [{ ...ins('before', 'V r1c0', 5, 'V r3c3'), extendMerge: true }, { ...ins('after', 'V r2c0', 5, 'V r3c3'), extendMerge: true }, { ...ins('before', 'V r0c0', 5, 'V r3c3'), extendMerge: true }]) {
    const s = store([vm4()]);
    const before = JSON.stringify(s.durableDoc);
    const r = await run(s, task(step));
    assert.equal(status(r.x), OPERATION_STATUS.FAILED, JSON.stringify(step));
    assert.equal(reason(r.x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:extend-merge-requested-but-the-insert-cuts-no-vertical-merge', JSON.stringify(step));
    assert.deepEqual([r.driver.calls.dispatch, r.m.writes], [0, 0]);
    assert.equal(JSON.stringify(s.durableDoc), before);
  }
  for (const step of [{ ...del('V r2c0', 3, 'V r3c3'), extendMerge: true }, { ...ins('after', 'V r1c0', 5, 'V r3c3'), extendMerge: false }]) {
    const planned = require('../../src/adapters/docx/adapter.cjs').planOperations(task(step));
    assert.equal(planned.ok, false, JSON.stringify(step));
    assert.match(planned.detail, /field-unsupported: step 0 \(delete_table_row\)\.extendMerge|extendMerge is only true, and only on insert_table_row_before \/ insert_table_row_after/);
  }
});

test('M5 NEGATIVE: on the merged table a boundary INSIDE the merge / a target row inside it / a delete -> refused before dispatch, 0/0; an unreconciled grid -> refused', async () => {
  const cases = [
    [ins('after', 'V r1c0', 5, 'V r3c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-cuts-a-vertical-merge'],
    [ins('before', 'V r2c0', 5, 'V r3c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-cuts-a-vertical-merge'],
    [ins('before', 'V r1c0', 5, 'V r3c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-target-row-inside-a-vertical-merge'],
    [ins('after', 'V r2c0', 5, 'V r3c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-target-row-inside-a-vertical-merge'],
    // W06-3: 'V r0c0' (row 0) is OUTSIDE the merge's span [1,2] -- LIVE-measured clean
    // (#ref case A) -- moved to its own POSITIVE test below, no longer a negative case here.
    // the OWNER row of the span [1,2] stays refused (W06-2a, destructive); its continuation row (GH #291 T-06) is a positive test below
    [del('V r1c0', 3, 'V r3c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-delete-inside-a-vertical-merge'],
  ];
  for (const [step, want] of cases) {
    const s = store([vm4()]);
    const before = JSON.stringify(s.durableDoc);
    const { x, driver, m } = await run(s, task(step));
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want, JSON.stringify(step));
    assert.deepEqual([driver.calls.dispatch, m.writes], [0, 0], want);
    assert.equal(JSON.stringify(s.durableDoc), before, `${want}: nothing written`);
  }
  // the continuation cell MISSING from the driver read too: the two reads no longer reconcile
  const bad = vm4(); bad.rows[2] = bad.rows[2].filter(c => !c.vContinue);
  const u = await run(store([bad]), task(ins('after', 'V r3c0', 5)));
  assert.equal(reason(u.x), 'vertical-merge-grid-unreconciled');
  assert.equal(u.driver.calls.dispatch, 0);
});

// W06-3 (#ref, LIVE-measured #ref case A): a row OUTSIDE a vertical merge's span deletes
// cleanly -- the per-table blanket is narrowed to per-row (table-row.cjs:149).
test('M5 POSITIVE (W06-3): a row delete OUTSIDE a vertical merge\'s span -> APPLIED, the merge and its content kept; retry SATISFIED 0/0', async () => {
  const cases = [
    [del('V r0c0', 3, 'V r3c3'), 0],   // row 0, above the merge [1,2] -- owner shifts up to row 0
    [del('V r3c0', 3, 'V r1c0'), 1],   // row 3, below the merge [1,2] -- owner stays at row 1
  ];
  for (const [step, ownerRow] of cases) {
    const s = store([vm4()]);
    await firstAndSecond(s, task(step));
    const t = texts(s);
    assert.equal(t.length, 3, JSON.stringify(step));
    mergeKept(s, ownerRow);
  }
});
