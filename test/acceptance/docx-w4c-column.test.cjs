'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const column = require('../../src/adapters/docx/families/table-column.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 Wave 4c offline acceptance: table-column (insert_table_column_before / _after,
 * delete_table_column) through the TARGET path on the fake transport whose column model follows
 * column probe 1 (2026-09-24): AddColumn puts one empty column at the boundary in every row, a
 * colspan straddling the boundary gets the new cell before it (drift); RemoveColumn removes the
 * column, a colspan covering it is lost (content loss). The adapter refuses both before dispatch.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (text, span = 1) => ({ span, paras: [{ text, props: { ...PLAIN } }] });
const grid = (tag, rows, cols = rows) => ({ rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (__, c) => cell(`${tag} r${r}c${c}`))) });
/** 3x5, row 1 merges columns 1..3 (colspan 3) -- the probe's K table */
const kTable = tag => { const t = grid(tag, 3, 5); t.rows[1] = [cell(`${tag} r1c0`), cell(`${tag} r1c1 ${tag} r1c2 ${tag} r1c3`, 3), cell(`${tag} r1c4`)]; return t; };
/** the public Split's shape (probe G): rows 0 and 2 carry a colspan-2 cell over row 1's two halves */
const gTable = tag => { const t = { rows: [[cell(`${tag} r0c0`), cell(`${tag} r0c1`, 2), cell(`${tag} r0c2`)], [cell(`${tag} r1c0`), cell(`${tag} r1c1`), cell(''), cell(`${tag} r1c2`)], [cell(`${tag} r2c0`), cell(`${tag} r2c1`, 2), cell(`${tag} r2c2`)]] }; return t; };
const paragraphs = () => [{ text: 'Column heading', level: 1, props: { ...PLAIN, bold: true } }, { text: 'Between.', props: { ...PLAIN } }];
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
const task = (...steps) => ({ id: 'w4c', steps });
const ins = (side, text, terminalColumnCount, tableAnchor) => ({ intent: `insert_table_column_${side}`, selector: tableAnchor ? { tableCell: { text, tableAnchor } } : { text }, terminalColumnCount });
const del = (text, terminalColumnCount, tableAnchor) => ({ intent: 'delete_table_column', selector: { tableCell: { text, tableAnchor } }, terminalColumnCount });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const texts = (s, t = 0) => s.durableDoc.tables[t].rows.map(r => r.map(c => c.paras.map(p => p.text).join(' ')));
const spans = (s, t = 0) => s.durableDoc.tables[t].rows.map(r => r.map(c => c.span || 1));

async function firstAndSecond(s, t) {
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(first.m.writes, t.steps.length);
  assert.equal(first.m.commands, t.steps.length, 'FIRST_COMMANDS: one command per column operation');
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

const row = (r, cols) => cols.map(c => (c === null ? '' : `A r${r}c${c}`));
const rows3 = cols => [0, 1, 2].map(r => row(r, cols));

test('rectangular 3x3: before/after on the FIRST, MIDDLE and LAST column; delete of each -- FIRST applies, SECOND is 0/0; order and direction exact', async () => {
  const cases = [
    [ins('before', 'A r0c0', 4), rows3([null, 0, 1, 2])],
    [ins('before', 'A r1c1', 4), rows3([0, null, 1, 2])],
    [ins('before', 'A r2c2', 4), rows3([0, 1, null, 2])],
    [ins('after', 'A r0c0', 4), rows3([0, null, 1, 2])],
    [ins('after', 'A r1c1', 4), rows3([0, 1, null, 2])],
    [ins('after', 'A r2c2', 4), rows3([0, 1, 2, null])],
    [del('A r0c0', 2, 'A r1c1'), rows3([1, 2])],
    [del('A r1c1', 2, 'A r0c0'), rows3([0, 2])],
    [del('A r2c2', 2, 'A r0c0'), rows3([0, 1])],
  ];
  for (const [st, want] of cases) {
    const s = store();
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), want, JSON.stringify(st));
  }
});

test('colspan matrix (probe K): OUTSIDE and span-EDGE inserts, an OUTSIDE delete -> SUPPORTED, the colspan and its content kept', async () => {
  const merged = t => `${t} r1c1 ${t} r1c2 ${t} r1c3`;
  for (const [tag, st, row1] of [
    ['K1', ins('before', 'K1 r0c0', 6), ['', 'K1 r1c0', merged('K1'), 'K1 r1c4']],
    ['K2', ins('after', 'K2 r0c4', 6), ['K2 r1c0', merged('K2'), 'K2 r1c4', '']],
    ['K3', del('K3 r0c4', 4, 'K3 r0c0'), ['K3 r1c0', merged('K3')]],
    ['K4', ins('before', 'K4 r0c1', 6), ['K4 r1c0', '', merged('K4'), 'K4 r1c4']],
    ['K5', ins('after', 'K5 r0c3', 6), ['K5 r1c0', merged('K5'), '', 'K5 r1c4']],
  ]) {
    const s = store([kTable(tag)]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s)[1], row1, tag);
    assert.ok(spans(s)[1].includes(3), `${tag}: the colspan 3 is kept`);
  }
});

test('colspan matrix: an insert INSIDE the span, a delete of a column the span covers, a merged cell as the target -> FAIL_CLOSED_UNSUPPORTED_SHAPE, 0 dispatch', async () => {
  for (const [tag, st, want] of [
    ['K6', del('K6 r0c1', 4, 'K6 r0c0'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-covers-deleted-column'],
    ['K7', ins('after', 'K7 r0c1', 6), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-straddles-insert-boundary'],
    ['K8', ins('before', 'K8 r0c2', 6), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-straddles-insert-boundary'],
    ['K9', del('K9 r0c2', 4, 'K9 r0c0'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-covers-deleted-column'],
    ['KM1', ins('before', 'KM1 r1c1 KM1 r1c2 KM1 r1c3', 6), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:target-spans-columns'],
    ['KM2', ins('after', 'KM2 r1c1 KM2 r1c2 KM2 r1c3', 6), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:target-spans-columns'],
    ['KM3', del('KM3 r1c1 KM3 r1c2 KM3 r1c3', 4, 'KM3 r0c0'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:target-spans-columns'],
  ]) {
    const s = store([kTable(tag)]);
    const before = texts(s);
    const { x, driver } = await run(s, task(st));
    assert.equal(status(x), OPERATION_STATUS.FAILED, tag);
    assert.equal(reason(x), want, tag);
    assert.equal(driver.calls.dispatch, 0, tag);
    assert.deepEqual(texts(s), before, `${tag}: nothing written`);
  }
});

test('the Split shape (probe G) is a colspan grid: span-edge inserts and a span-free delete SUPPORTED; a target on the colspan-2 cell FAIL_CLOSED', async () => {
  const g3 = store([gTable('G')]);
  await firstAndSecond(g3, task(ins('after', 'G r1c2', 5)));
  assert.deepEqual(texts(g3).map(r => r[r.length - 1]), ['', '', '']);
  const g4 = store([gTable('G')]);
  await firstAndSecond(g4, task(del('G r1c0', 3, 'G r0c2')));
  assert.deepEqual(texts(g4), [['G r0c1', 'G r0c2'], ['G r1c1', '', 'G r1c2'], ['G r2c1', 'G r2c2']]);
  const g5 = store([gTable('G')]);
  await firstAndSecond(g5, task(ins('before', 'G r0c2', 5)));
  assert.deepEqual(texts(g5)[1], ['G r1c0', 'G r1c1', '', '', 'G r1c2']);
  for (const st of [ins('after', 'G r0c1', 5), del('G r0c1', 2, 'G r0c0')]) {
    const { x, driver } = await run(store([gTable('G')]), task(st));
    assert.equal(reason(x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:target-spans-columns');
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('LOGICAL vs PHYSICAL column: deleting logical column 3 of the Split shape removes row 1\'s 4th physical cell; a physical-index delete is RED at L1', async () => {
  const s = store([gTable('G')]);
  await firstAndSecond(s, task(del('G r0c2', 3, 'G r0c0')));
  assert.deepEqual(texts(s)[1], ['G r1c0', 'G r1c1', ''], 'the physical index 2 of row 1 (an empty half) survives');
  const { x } = await run(store([gTable('G')]), task(del('G r0c2', 3, 'G r0c0')), { fault: 'col-physical-index' });
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL');
  assert.notEqual(status(x), OPERATION_STATUS.APPLIED);
});

test('MULTI_TABLE_TARGETING: the same cell text in two tables -- the anchor picks one, the other is unchanged; no anchor -> FAIL_CLOSED', async () => {
  const two = () => { const b = grid('B', 3); b.rows[1][1] = cell('ZQW SHARED'); const a = grid('A', 3); a.rows[1][1] = cell('ZQW SHARED'); return [a, b]; };
  const s = store(two());
  const other = texts(s, 1);
  await firstAndSecond(s, task(ins('after', 'ZQW SHARED', 4, 'A r0c0')));
  assert.deepEqual(texts(s, 1), other, 'the other table is unchanged');
  assert.deepEqual(texts(s, 0)[1], ['A r1c0', 'ZQW SHARED', '', 'A r1c2']);
  const s2 = store(two());
  const other2 = texts(s2, 0);
  await firstAndSecond(s2, task(del('ZQW SHARED', 2, 'B r0c0')));
  assert.deepEqual(texts(s2, 0), other2);
  assert.deepEqual(texts(s2, 1), [['B r0c0', 'B r0c2'], ['B r1c0', 'B r1c2'], ['B r2c0', 'B r2c2']]);
  const { x, driver } = await run(store(two()), task(ins('after', 'ZQW SHARED', 4)));
  assert.equal(reason(x), 'table-column-target-ambiguous');
  assert.equal(driver.calls.dispatch, 0);
});

test('mixed one session: insert_before -> insert_after -> delete, each resolved from the FRESH state; retry 0/0', async () => {
  const s = store([grid('M', 3, 4)]);
  const t = task(ins('before', 'M r0c1', 5, 'M r2c0'), ins('after', 'M r0c2', 6, 'M r2c0'), del('M r0c3', 5, 'M r2c0'));
  const { first } = await firstAndSecond(s, t);
  // FRESH_STRUCTURAL_RETARGETING: every later target column moved because of an earlier insert
  assert.deepEqual(first.x.primitiveVerification.map(v => v.evidence && v.evidence.evidence && v.evidence.evidence.targetColumnBefore), [1, 3, 5]);
  assert.deepEqual(texts(s)[0], ['M r0c0', '', 'M r0c1', 'M r0c2', '']);
});

test('mixed: a partial first run (the 2nd operation did not happen) -> the retry SATISFIES the 1st and APPLIES the rest', async () => {
  const s = store([grid('M', 3, 4)]);
  const t = task(ins('before', 'M r0c1', 5, 'M r2c0'), ins('after', 'M r0c2', 6, 'M r2c0'), del('M r0c3', 5, 'M r2c0'));
  const first = await run(s, t, { fault: 'ignore', faultOn: [2] });
  assert.equal(status(first.x, 0), OPERATION_STATUS.APPLIED);
  assert.notEqual(status(first.x, 1), OPERATION_STATUS.APPLIED);
  const second = await run(s, t);
  assert.deepEqual(second.x.appliedOperations.map(e => e.status), [OPERATION_STATUS.SATISFIED, OPERATION_STATUS.APPLIED, OPERATION_STATUS.APPLIED]);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  assert.deepEqual(texts(s)[0], ['M r0c0', '', 'M r0c1', 'M r0c2', '']);
});

test('delete retry identity: target absent + terminal width in the anchor table -> SATISFIED; recovery -> FAIL_CLOSED', async () => {
  const s = store();
  await run(s, task(del('A r1c1', 2, 'A r0c0')));
  const r = await run(s, task(del('A r1c1', 2, 'A r0c0')));
  assert.equal(status(r.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([r.m.writes, r.m.commands], [0, 0]);
  const rec = await run(s, { ...task(del('A r1c1', 2, 'A r0c0')), recovery: { unresolvedUncertainOperations: ['s0'] } });
  assert.equal(reason(rec.x), 'preservation-baseline-unavailable');
  assert.equal(rec.driver.calls.dispatch, 0);
});

test('negatives before dispatch -> FAIL_CLOSED or plan-failed, 0 dispatch', async () => {
  const dup = grid('A', 3); dup.rows[2][2] = cell('A r1c1');
  const vm = grid('A', 3); vm.rows[0][2] = { ...vm.rows[0][2], rowspan: 2 }; vm.rows[1][2] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] };
  const uneven = grid('A', 3); uneven.rows[1].pop();
  const twoAnchor = () => { const b = grid('B', 2); b.rows[0][0] = cell('A r0c0'); return [grid('A', 3), b]; };
  const failed = [
    [task(ins('after', 'A r1c1', 4)), store([dup]), 'table-column-target-ambiguous'],
    [task(ins('after', 'A r1c1', 4, 'A r0c0')), store(twoAnchor()), 'table-anchor-ambiguous'],
    [task(ins('after', 'nowhere', 4)), store(), 'table-column-target-not-found'],
    [task(ins('after', 'A r1c1', 9)), store(), 'table-column-terminal-unprovable'],
    [task(del('A r1c1', 1, 'A r0c0')), store(), 'delete-target-present-but-column-count-not-terminal-plus-one'],
    [task(del('A r1c1', 2, 'A r0c1')), store(), 'table-anchor-in-target-column'],
    [task(ins('after', 'A r1c1', 4)), store([vm]), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:vertical-merge'],
    [task(del('A r0c0', 2, 'A r1c1')), store([vm]), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:vertical-merge'],
    [task(ins('after', 'A r0c0', 4)), store([uneven]), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:uneven-logical-width'],
    [task(ins('before', 'A r0c0', 4, 'A r2c2'), ins('after', 'A r2c2', 5)), store(), 'column-operations-without-table-anchor-cannot-be-grouped'],
  ];
  for (const [t, s, want] of failed) {
    const { x, driver } = await run(s, t);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
  for (const st of [
    { intent: 'delete_table_column', selector: { text: 'A r1c1' }, terminalColumnCount: 2 },
    { ...ins('after', 'A r1c1', 4), terminalColumnCount: undefined },
    { ...ins('after', 'A r1c1', 1) },
    { intent: 'insert_table_column_after', selector: { tableCell: { text: 'A r1c1', tableAnchor: 'A r1c1' } }, terminalColumnCount: 4 },
    { ...del('A r1c1', 2, 'A r0c0'), terminalTableAnchor: 'A r2c2' },
    { intent: 'insert_table_column_after', selector: { text: 'A r1c1', index: 1 }, terminalColumnCount: 4 },
  ]) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('the source form is accepted: combined selector + terminalColumnCount (+ terminalTableAnchor for delete)', async () => {
  await firstAndSecond(store(), task({ intent: 'insert_table_column_after', selector: { tableCell: { text: 'A r1c1' }, text: 'A r1c1' }, terminalColumnCount: 4 }));
  await firstAndSecond(store(), task({ intent: 'delete_table_column', selector: { tableCell: { text: 'A r1c1' }, text: 'A r1c1' }, terminalColumnCount: 2, terminalTableAnchor: 'A r0c0' }));
});

test('faults at L1: wrong side / twice / not empty / a new cell of span 2 / reordered / delete ignored / other column deleted / stale index -> never APPLIED', async () => {
  for (const [fault, st] of [
    ['col-wrong-side', ins('before', 'A r1c1', 4)], ['col-wrong-side', ins('after', 'A r1c1', 4)], ['col-twice', ins('after', 'A r1c1', 4)],
    ['col-new-not-empty', ins('after', 'A r1c1', 4)], ['col-new-span2', ins('after', 'A r1c1', 4)], ['col-reorder', ins('before', 'A r0c0', 4)],
    ['ignore', del('A r1c1', 2, 'A r0c0')], ['col-delete-other', del('A r1c1', 2, 'A r0c0')], ['col-stale-index', ins('after', 'A r1c1', 4)],
    ['col-stale-index', del('A r2c2', 2, 'A r0c0')],
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
  for (const [fault, st, tables] of [
    ['col-touch-cell', ins('after', 'A r1c1', 4), undefined],
    ['col-touch-other-table', ins('after', 'A r1c1', 4, 'A r0c0'), [grid('A', 3), grid('B', 2)]],
    ['col-change-outside', del('A r1c1', 2, 'A r0c0'), undefined],
  ]) {
    const { x } = await run(store(tables), task(st), { fault });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', `${fault}: the column itself is right`);
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN, fault);
  }
});

test('a colspan changed by the write (3 -> 1) -> the row grid turns uneven: RED already at L1, never APPLIED', async () => {
  const { x, driver } = await run(store([kTable('K')]), task(ins('after', 'K r0c4', 6)), { fault: 'col-span-change' });
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(x.primitiveVerification[0].evidence.evidence.reason, 'after-FAIL_CLOSED_UNSUPPORTED_SHAPE:uneven-logical-width');
  assert.equal(driver.calls.forceSave, 0);
});

test('DEFENCE IN DEPTH: with the shape guard bypassed, the measured drift (K7) and content loss (K6) are RED at L1', async () => {
  for (const [tag, st] of [['K7', ins('after', 'K7 r0c1', 6)], ['K6', del('K6 r0c1', 4, 'K6 r0c0')]]) {
    const s = store([kTable(tag)]);
    const driver = createFakeDocxDriver(s);
    await driver.open({ fileId: 'd' });
    const inv = async () => normalizeInventory((await driver.readLive()).payload);
    const op = planOperations(task(st)).operations[0];
    const before = await inv();
    assert.match(column.classify(before, op).evidence.reason, /^FAIL_CLOSED_UNSUPPORTED_SHAPE/, `${tag}: the guard refuses it`);
    const target = column.resolve(before, op.selector, { operation: op }).target;
    await driver.dispatch(null, column.lower(op, target));
    const after = await inv();
    assert.equal(column.verifyPrimitive(op, { primary: before }, { primary: after }).ok, false, `${tag}: L1 would still catch it`);
  }
});

test('the command finds the table by its anchor and the cell by its full text, then AddColumn / RemoveColumn -- never Search, never frame.evaluate', () => {
  const body = column.commandBody({ family: 'table', action: 'column', columnAction: 'add', before: true, targetText: 'x', tableAnchor: 'a' });
  assert.match(body, /\.AddColumn\(hits\[0\]\.cell, m\.before === true\)/);
  assert.match(column.commandBody({ family: 'table', action: 'column', columnAction: 'remove', before: null, targetText: 'x', tableAnchor: 'a' }), /\.RemoveColumn\(/);
  assert.doesNotMatch(body, /\.Search\(|evaluate|asc_|addColumnLeft|remColumn/);
});

test('group final state reached but THIS operation\'s claim missing -> FAIL_CLOSED, not a blind re-apply', async () => {
  const planned = planOperations(task(ins('after', 'M r0c0', 4, 'M r0c2'), del('M r0c1', 3, 'M r0c2')));
  assert.equal(planned.ok, true);
  // 3 columns, M r0c1 gone (the delete's claim holds), and no empty column after M r0c0 (the insert's does not)
  const t = { rows: [[cell('M r0c0'), cell('M r0c2'), cell('M r0c3')], [cell('M r1c0'), cell('M r1c2'), cell('M r1c3')]] };
  const driver = createFakeDocxDriver(store([t]));
  await driver.open({ fileId: 'd' });
  const primary = normalizeInventory((await driver.readLive()).payload);
  const c = column.classify(primary, planned.operations[0]);
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'group-final-state-without-this-operation-claim');
});

/** M5 row 14: the gridAfter probe's S table (row 0 ends in a one-column grid gap) and its R control (the gap next to a colspan) */
const gap = () => ({ gap: true, span: 1, paras: [] });
const sTable = tag => ({ rows: [[cell(`${tag} a0`), cell(`${tag} a1`), cell(`${tag} a2`), gap()], [cell(`${tag} b0`), cell(`${tag} b1`), cell(`${tag} b2`), cell(`${tag} anchor`)]] });
/** a TWO-column gap (unmeasured): row 0 [a0, a1, gap, gap], row 1 four real cells */
const wTable = tag => ({ rows: [[cell(`${tag} a0`), cell(`${tag} a1`), gap(), gap()], [cell(`${tag} b0`), cell(`${tag} b1`), cell(`${tag} b2`), cell(`${tag} b3`)]] });
const rTable = tag => ({ rows: [[cell(`${tag} a0`), cell(`${tag} a1`), cell(`${tag} a2`), gap()], [cell(`${tag} b0`), cell(`${tag} b1`, 2), cell(`${tag} b2`)]] });
const gapFlags = (s, t = 0) => s.durableDoc.tables[t].rows.map(r => r.map(c => c.gap === true));

test('M5 row 14, grid gap (gridAfter): the measured inserts APPLY -- the new column at the boundary in every row, the gap stays trailing; retry 0/0', async () => {
  for (const [tag, st, row0, row1] of [
    ['S1', ins('after', 'S1 a2', 5), ['S1 a0', 'S1 a1', 'S1 a2', '', ''], ['S1 b0', 'S1 b1', 'S1 b2', '', 'S1 anchor']],
    ['S2', ins('before', 'S2 anchor', 5), ['S2 a0', 'S2 a1', 'S2 a2', '', ''], ['S2 b0', 'S2 b1', 'S2 b2', '', 'S2 anchor']],
    ['S3', ins('after', 'S3 a0', 5), ['S3 a0', '', 'S3 a1', 'S3 a2', ''], ['S3 b0', '', 'S3 b1', 'S3 b2', 'S3 anchor']],
    ['S4', ins('after', 'S4 b0', 5), ['S4 a0', '', 'S4 a1', 'S4 a2', ''], ['S4 b0', '', 'S4 b1', 'S4 b2', 'S4 anchor']],
    ['S5', ins('before', 'S5 a0', 5), ['', 'S5 a0', 'S5 a1', 'S5 a2', ''], ['', 'S5 b0', 'S5 b1', 'S5 b2', 'S5 anchor']],
  ]) {
    const s = store([sTable(tag)]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), [row0, row1], tag);
    assert.deepEqual(gapFlags(s), [[false, false, false, false, true], [false, false, false, false, false]], `${tag}: the gap is kept, trailing`);
  }
});

test('M5 row 14, grid gap: a delete of a column that is a real cell in the gap row APPLIES -- the gap kept; retry 0/0 (probe DEL_S_a1, source ...a145c7)', async () => {
  for (const [tag, st, row0, row1] of [
    ['SX1', del('SX1 a1', 3, 'SX1 a0'), ['SX1 a0', 'SX1 a2', ''], ['SX1 b0', 'SX1 b2', 'SX1 anchor']],
    ['SX2', del('SX2 b0', 3, 'SX2 a1'), ['SX2 a1', 'SX2 a2', ''], ['SX2 b1', 'SX2 b2', 'SX2 anchor']],
  ]) {
    const s = store([sTable(tag)]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), [row0, row1], tag);
    assert.deepEqual(gapFlags(s), [[false, false, true], [false, false, false]], `${tag}: the gap is kept, trailing`);
  }
});

test('M5 row 14, grid gap: a delete of a ONE-column gap\'s own column APPLIES -- the gap row loses its gap; retry 0/0 (probe DEL_S_anchor, source ...ff8ea3)', async () => {
  const s = store([sTable('SY')]);
  await firstAndSecond(s, task(del('SY anchor', 3, 'SY a0')));
  assert.deepEqual(texts(s), [['SY a0', 'SY a1', 'SY a2'], ['SY b0', 'SY b1', 'SY b2']]);
  assert.deepEqual(gapFlags(s), [[false, false, false], [false, false, false]], 'the gap is consumed, no cell of row 0 lost');
});

test('M5 row 14, grid gap: a delete inside a wider (unmeasured) gap, a boundary inside the gap, the gap next to a colspan (R control) -> FAIL_CLOSED_UNSUPPORTED_SHAPE, 0 dispatch', async () => {
  for (const [tag, mk, st, want] of [
    ['SD3', wTable, del('SD3 b3', 3, 'SD3 a0'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:grid-gap-column-delete'],
    ['RD1', rTable, del('RD1 a0', 3, 'RD1 a1'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:grid-gap-with-colspan'],
    ['SG1', sTable, ins('after', 'SG1 anchor', 5), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-grid-gap'],
    ['R1', rTable, ins('after', 'R1 a0', 5), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:grid-gap-with-colspan'],
    ['R2', rTable, ins('after', 'R2 a2', 5), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:grid-gap-with-colspan'],
  ]) {
    const s = store([mk(tag)]);
    const before = JSON.stringify(s.durableDoc.tables);
    const { x, driver } = await run(s, task(st));
    assert.equal(status(x), OPERATION_STATUS.FAILED, tag);
    assert.equal(reason(x), want, tag);
    assert.equal(driver.calls.dispatch, 0, tag);
    assert.equal(JSON.stringify(s.durableDoc.tables), before, `${tag}: nothing written`);
  }
});

test('M5 row 14, columnRead: a placeholder that is not trailing, a non-empty or spanning placeholder, a row-count or cell-count mismatch -> grid-gap-unreconciled', () => {
  const td = (text, extra = {}) => ({ text, span: 1, rowspan: 1, ...extra });
  const ph = (extra = {}) => td('', { placeholder: true, ...extra });
  const rd = t => ({ text: t });
  const primary = (hrows, drows) => ({ blocks: [{ kind: 'table', rows: hrows }], document: { tables: [{ rows: drows }] } });
  const ok = column.columnRead(primary([[td('a'), td('b'), ph()], [td('c'), td('d'), td('e')]], [[rd('a'), rd('b')], [rd('c'), rd('d'), rd('e')]]));
  assert.equal(ok.ok, true, 'POSITIVE CONTROL: a trailing empty placeholder reconciles');
  assert.deepEqual(ok.gaps, [[1, 0]]);
  assert.deepEqual(ok.html[0].rows.map(r => r.length), [2, 3]);
  const nbsp = column.columnRead(primary([[td('a'), ph({ text: ' ' })], [td('b'), td('c')]], [[rd('a')], [rd('b'), rd('c')]]));
  assert.equal(nbsp.ok, true, 'a whitespace-only placeholder (the projection\'s &nbsp; paragraph) is still a gap');
  const named = column.columnRead(primary([[td('a'), ph({ text: 'x' })]], [[rd('a')]]));
  assert.equal(named.detail.check, 'gap-shape', 'the refusal names the failed check');
  assert.deepEqual([named.detail.table, named.detail.row, named.detail.driverCells], [0, 0, 1]);
  assert.equal(column.columnRead(primary([[td('a'), td('b'), ph()]], [[rd('a')]])).detail.check, 'cell-count');
  assert.equal(column.columnRead(primary([[td('a'), ph()], [td('b')]], [[rd('a')]])).detail.check, 'row-count');
  for (const [name, p] of [
    ['placeholder in the middle', primary([[td('a'), ph(), td('b')]], [[rd('a'), rd('b')]])],
    ['non-empty placeholder', primary([[td('a'), ph({ text: 'x' })]], [[rd('a')]])],
    ['spanning placeholder', primary([[td('a'), ph({ span: 2 })]], [[rd('a')]])],
    ['row-count mismatch', primary([[td('a'), ph()], [td('b')]], [[rd('a')]])],
    ['cell-count mismatch', primary([[td('a'), td('b'), ph()]], [[rd('a')]])],
  ]) assert.equal(column.columnRead(p).reason, 'FAIL_CLOSED_UNSUPPORTED_SHAPE:grid-gap-unreconciled', name);
});
