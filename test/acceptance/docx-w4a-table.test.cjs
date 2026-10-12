'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const table = require('../../src/adapters/docx/families/table.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 Wave 4a offline acceptance: table-cell (replace_table_cell) through the TARGET path
 * (M1 -> M2 -> M3 -> DOCX adapter) on the fake transport whose table model follows the live table
 * probes 1-2: the driver's public table read (cell text ending in '\t', a merged cell's paragraphs
 * joined by '\r\n', the first paragraph's effective properties) next to the HTML projection
 * (a merged cell is colspan=2, its paragraphs joined without a separator).
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Courier New', fontSize: 28, language: 'hu-HU' };
const H = 'Table heading';
const A = 'Alpha paragraph.';
const Z = 'Closing paragraph.';
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const tables = () => [{
  rows: [
    [cell('r0c0'), cell('r0c1'), { span: 1, paras: [{ text: 'r0c2', props: { ...RICH }, alignment: '2' }] }],
    [cell('r1c0'), cell('r1c1'), cell('r1c2')],
    [{ ...cell('r2c0', 'r2c1'), span: 2 }, cell('r2c2')],
  ],
}];
const paragraphs = () => [{ text: H, level: 1, props: { ...PLAIN, bold: true } }, { text: A, props: { ...PLAIN } }, { text: Z, props: { ...PLAIN } }];
const store = (over = {}) => createStore(over.paragraphs || paragraphs(), { doc: { tables: over.tables || tables() } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  assert.equal(driver.calls.open, 1, 'OPEN_COUNT=1');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT=1');
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w4a', steps });
const step = (old, text, form = 'text') => ({ intent: 'replace_table_cell', selector: form === 'text' ? { text: old } : { tableCell: { text: old } }, text });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => {
  const e = x.appliedOperations[i].evidence;
  return (e && e.evidence && e.evidence.reason) || (e && e.reason);
};
const durableCell = (s, r, c, t = 0) => s.durableDoc.tables[t].rows[r][c];

async function firstAndSecond(s, t) {
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome || first.x.appliedOperations.map(e => e.evidence)));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(first.m.writes, t.steps.length);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), 'L1');
  assert.equal(first.x.preservationVerification.length, t.steps.length, 'integral L3 for every P1 operation');
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), 'L3');
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', 'L2');
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  assert.equal(first.driver.calls.forceSave, 1, 'one barrier, at task end');
  const second = await run(s, t);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.stopDetail || second.x.outcome));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(second.m.writes, 0);
  assert.equal(second.driver.calls.forceSave, 0);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

// ---- inventory -----------------------------------------------------------------------------------

test('inventory: HTML spans and the driver table read (normalized text, merged cell joined by a space)', () => {
  const html = '<body><table><tr><td><p>x</p></td><td><p>y</p></td></tr><tr><td colspan="2"><p>a</p><p>b</p></td></tr></table></body>';
  const document = { tables: [{ rows: [[{ text: 'x\t', paragraphs: 1, simple: true, format: {} }, { text: 'y\r\n', paragraphs: 1, simple: true, format: {} }], [{ text: 'a\r\nb\t', paragraphs: 2, simple: true, format: { bold: true } }]] }] };
  const inv = normalizeInventory({ html, paragraphs: [], document });
  assert.deepEqual(inv.blocks[0].rows.map(r => r.map(c => c.span)), [[1, 1], [2]]);
  assert.deepEqual(inv.document.tables[0].rows.map(r => r.map(c => c.text)), [['x', 'y'], ['a b']]);
  assert.equal(inv.blocks[0].rows[1][0].text, 'ab', 'the HTML joins a merged cell without a separator -- why it is not the identity');
  assert.equal(normalizeInventory({ html, paragraphs: [], document: {} }).document.tables, null, 'unread stays unknown, never []');
});

// ---- first / second --------------------------------------------------------------------------------

test('plain cell: FIRST applies, SECOND is a 0-write no-op; { tableCell: { text } } and the source\'s combined form are the same target', async () => {
  const s = store();
  await firstAndSecond(s, task(step('r1c1', 'NEW r1c1')));
  assert.equal(durableCell(s, 1, 1).paras[0].text, 'NEW r1c1');
  const s2 = store();
  await firstAndSecond(s2, task(step('r0c1', 'NEW r0c1', 'tableCell')));
  assert.equal(durableCell(s2, 0, 1).paras[0].text, 'NEW r0c1');
  // the source's canonical agent selector names the cell twice (its preflight reads tableCell, its classifier text)
  const s3 = store();
  await firstAndSecond(s3, task({ intent: 'replace_table_cell', selector: { tableCell: { text: 'r1c0' }, text: 'r1c0' }, text: 'NEW r1c0' }));
  assert.equal(durableCell(s3, 1, 0).paras[0].text, 'NEW r1c0');
});

test('formatted cell: the first paragraph\'s properties survive (bold, italic, font family, size, alignment)', async () => {
  const s = store();
  const { first } = await firstAndSecond(s, task(step('r0c2', 'NEW formatted')));
  const p = durableCell(s, 0, 2).paras[0];
  assert.deepEqual([p.text, p.props.bold, p.props.italic, p.props.fontFamily, p.props.fontSize, p.alignment], ['NEW formatted', true, true, 'Courier New', 28, '2']);
  assert.deepEqual(first.x.preservationVerification[0].evidence.evidence.targetProperties.expected,
    { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Courier New', fontSize: 14, style: 'Normal', alignment: '2' });
});

test('merged two-paragraph cell: becomes one paragraph, keeps colspan=2 (probe 2 contentMerged)', async () => {
  const s = store();
  await firstAndSecond(s, task(step('r2c0 r2c1', 'NEW merged')));
  const c = durableCell(s, 2, 0);
  assert.deepEqual([c.span, c.paras.length, c.paras[0].text], [2, 1, 'NEW merged']);
});

test('two cells in one task: both APPLIED, L2 holds, retry 0/0', async () => {
  await firstAndSecond(store(), task(step('r0c0', 'A0'), step('r2c2', 'A2')));
});

// ---- fail closed before dispatch ----------------------------------------------------------------

test('negatives -> FAIL_CLOSED, 0 dispatch', async () => {
  const dupCell = tables();
  dupCell[0].rows[1][2] = cell('r1c1');
  const nonPlain = tables();
  nonPlain[0].rows[1][1] = { ...cell('r1c1'), nonPlain: true };
  const cases = [
    [task(step('nowhere', 'x')), store(), {}, 'table-cell-terminal-unprovable'],
    [task(step('r1c1', 'x')), store({ tables: dupCell }), {}, 'table-cell-ambiguous'],
    [task(step(A, 'x')), store(), {}, 'table-cell-terminal-unprovable'],
    [task(step('r1c1', 'r0c0')), store(), {}, 'table-cell-conflicting-old-and-new-state'],
    [task(step('r1c1', 'x')), store({ tables: nonPlain }), {}, 'table-cell-not-plain-paragraphs-of-runs'],
    [task(step('r1c1', 'x')), store(), { tablesUnreadable: true }, 'tables-unreadable'],
    [task(step('r1c1', 'x')), store(), { docUnreadable: true }, 'tables-unreadable'],
    [task(step('r1c1', 'x')), store(), { tableReadMissesOne: true }, 'table-read-shape-mismatch'],
  ];
  for (const [t, s, script, want] of cases) {
    const { x, driver } = await run(s, t, script);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
});

test('retry after unresolved uncertainty: a terminal cell without its baseline -> FAIL_CLOSED, never SATISFIED', async () => {
  const s = store();
  await run(s, task(step('r1c1', 'NEW r1c1')));
  const { x, driver } = await run(s, { ...task(step('r1c1', 'NEW r1c1')), recovery: { unresolvedUncertainOperations: ['s0'] } });
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(reason(x), 'preservation-baseline-unavailable');
  assert.equal(driver.calls.dispatch, 0);
});

test('bad parameters and selectors -> plan fails before any dispatch', async () => {
  for (const s of [
    step('r1c1', ' '),
    step('r1c1', 'a\nb'),
    step('r1c1', 'a\tb'),
    step('r1c1', ' r1c1 '),
    { intent: 'replace_table_cell', selector: { ref: 's0' }, text: 'x' },
    { intent: 'replace_table_cell', selector: { text: 'r1c1', index: 2 }, text: 'x' },
    { intent: 'replace_table_cell', selector: { tableCell: { text: 'r1c1', row: 1 } }, text: 'x' },
    { intent: 'replace_table_cell', selector: { text: '' }, text: 'x' },
    { intent: 'replace_table_cell', selector: { tableCell: { text: 'r1c1' }, text: 'r1c0' }, text: 'x' },
  ]) {
    const { x, driver } = await run(store(), task(s));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(s));
    assert.equal(driver.calls.dispatch, 0);
  }
});

// ---- the write happened but is not the terminal contract ----------------------------------------

test('faults at L1: no write / wrong cell / appended / merged cell not collapsed / an extra empty paragraph -> never APPLIED, no barrier', async () => {
  for (const [fault, st] of [['ignore', step('r1c1', 'N')], ['cell-wrong-cell', step('r1c1', 'N')], ['cell-append', step('r1c1', 'N')], ['cell-keep-rest', step('r2c0 r2c1', 'N')], ['cell-extra-empty-paragraph', step('r1c1', 'N')]]) {
    const { x, result, driver } = await run(store(), task(st), { fault });
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, fault);
    assert.equal(x.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE, fault);
    assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_NOT_PERSISTED, fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

test('faults at L3: lost properties / lost style / another cell touched / unmerged / text outside the tables -> NOT APPLIED (P1)', async () => {
  for (const [fault, st] of [
    ['cell-drop-props', step('r0c2', 'N')], ['cell-drop-style', step('r1c1', 'N')], ['cell-touch-other', step('r1c1', 'N')],
    ['cell-touch-other-format', step('r1c1', 'N')], ['cell-unmerge', step('r2c0 r2c1', 'N')], ['cell-change-outside', step('r1c1', 'N')],
  ]) {
    const { x, driver } = await run(store(), task(st), { fault });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', `${fault}: the cell text itself is right`);
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN, fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

// ---- static shape of the command ------------------------------------------------------------------

test('the command identifies the cell by its full text through the public table API, never by Search', () => {
  const body = table.commandBody({ family: 'table', action: 'cell', targetText: 'r1c1', text: 'N' });
  assert.match(body, /GetAllTables\(\)/);
  assert.match(body, /GetCellsCount\(\)/);
  assert.doesNotMatch(body, /\.Search\(/, 'Search matches substrings (probe 1)');
  assert.doesNotMatch(body, /evaluate|asc_/, 'no frame.evaluate backdoor, no internal asc_ API');
  // NEGATIVE CONTROL: the same patterns see a Search-based body
  assert.match("doc.Search('r1c1')", /\.Search\(/);
});

// ---- W4a closure (owner, owner channel): table identity + cell identity; multi-paragraph controls --------

test('closure 2: the same text OUTSIDE the tables (exact or partial) is not an ambiguity', async () => {
  const s = store({ paragraphs: [...paragraphs(), { text: 'r1c1', props: { ...PLAIN } }, { text: 'r1c1 twin', props: { ...PLAIN } }] });
  await firstAndSecond(s, task(step('r1c1', 'NEW r1c1')));
  assert.equal(durableCell(s, 1, 1).paras[0].text, 'NEW r1c1');
  assert.ok(s.durable.some(p => p.text === 'r1c1'), 'the paragraph with the same text is untouched');
});

const twoTables = () => {
  const a = tables()[0];
  const b = { rows: [[cell('Other head'), cell('r1c1')], [cell('b10'), cell('b11')]] };
  return [a, b];
};

test('closure 2 / W4b 10: the same cell text in TWO tables -- the table anchor discriminates; without it FAIL_CLOSED', async () => {
  const s = store({ tables: twoTables() });
  const anchored = { intent: 'replace_table_cell', selector: { tableCell: { text: 'r1c1', tableAnchor: 'Other head' } }, text: 'NEW in B' };
  await firstAndSecond(s, task(anchored));
  assert.equal(durableCell(s, 0, 1, 1).paras[0].text, 'NEW in B', 'the anchored table changed');
  assert.equal(durableCell(s, 1, 1, 0).paras[0].text, 'r1c1', 'the other table did not');
  // the other table through its own anchor, in the SAME store
  await firstAndSecond(s, task({ intent: 'replace_table_cell', selector: { tableCell: { text: 'r1c1', tableAnchor: 'r0c0' } }, text: 'NEW in A' }));
  assert.equal(durableCell(s, 1, 1, 0).paras[0].text, 'NEW in A');
  // NEGATIVE CONTROLS: the table identity removed / an anchor present in both tables / an absent anchor
  const both = twoTables(); both[1].rows[1][0] = cell('r0c0');
  for (const [t, st, want] of [
    [task(step('r1c1', 'x')), store({ tables: twoTables() }), 'table-cell-ambiguous'],
    [task({ intent: 'replace_table_cell', selector: { tableCell: { text: 'r1c1', tableAnchor: 'r0c0' } }, text: 'x' }), store({ tables: both }), 'table-anchor-ambiguous'],
    [task({ intent: 'replace_table_cell', selector: { tableCell: { text: 'r1c1', tableAnchor: 'nowhere' } }, text: 'x' }), store({ tables: twoTables() }), 'table-anchor-not-found'],
  ]) {
    const { x, driver } = await run(st, t);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
  // a write that lands in the OTHER table (same cell text) is caught by L3
  const { x } = await run(store({ tables: twoTables() }), task(anchored), { fault: 'cell-touch-other', otherTable: true });
  assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'an unrelated table changed');
});

test('closure 2: an anchor equal to the old or the new cell text is refused at planning', async () => {
  for (const tableAnchor of ['r1c1', 'NEW', ' ']) {
    const { x, driver } = await run(store(), task({ intent: 'replace_table_cell', selector: { tableCell: { text: 'r1c1', tableAnchor } }, text: 'NEW' }));
    assert.equal(x.outcome, 'plan-failed', tableAnchor);
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('closure 4 A-E: a multi-paragraph cell -- one terminal paragraph; leftover / other cell / shape / first-paragraph property -> FAIL', async () => {
  const multi = () => { const t = tables(); t[0].rows[1][1] = { span: 1, paras: [{ text: 'm1', props: { ...RICH } }, { text: 'm2', props: { ...PLAIN } }, { text: 'm3', props: { ...PLAIN } }] }; return t; };
  // A: exactly one terminal paragraph with exactly the requested text, the first paragraph's properties kept
  const s = store({ tables: multi() });
  await firstAndSecond(s, task(step('m1 m2 m3', 'ONE')));
  const c = durableCell(s, 1, 1);
  assert.deepEqual([c.paras.length, c.paras[0].text, c.paras[0].props.bold, c.paras[0].props.fontFamily], [1, 'ONE', true, 'Courier New']);
  // B: a leftover old paragraph -> L1 FAIL
  let r = await run(store({ tables: multi() }), task(step('m1 m2 m3', 'ONE')), { fault: 'cell-keep-rest' });
  assert.equal(r.x.primitiveVerification[0].verdict, 'FAIL', 'B');
  // C: another cell changed -> integral L3 FAIL
  r = await run(store({ tables: multi() }), task(step('m1 m2 m3', 'ONE')), { fault: 'cell-touch-other' });
  assert.equal(r.x.preservationVerification[0].verdict, 'FAIL', 'C');
  assert.equal(r.x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN, 'C');
  // D: the table shape changed -> L1 FAIL
  r = await run(store({ tables: multi() }), task(step('m1 m2 m3', 'ONE')), { fault: 'cell-add-row' });
  assert.equal(r.x.primitiveVerification[0].verdict, 'FAIL', 'D');
  // E: a measured first-paragraph property lost -> integral L3 FAIL
  r = await run(store({ tables: multi() }), task(step('m1 m2 m3', 'ONE')), { fault: 'cell-drop-props' });
  assert.equal(r.x.primitiveVerification[0].verdict, 'PASS', 'E: the text is right');
  assert.equal(r.x.preservationVerification[0].verdict, 'FAIL', 'E');
});

test('closure 5: a colspan merge survives (LIVE-measured); a VERTICAL merge (rowspan) is refused before dispatch', async () => {
  // horizontal: the merged r2c0+r2c1 cell keeps colspan=2 and the other cells (see also the merged FIRST/SECOND test)
  const s = store();
  await firstAndSecond(s, task(step('r1c1', 'NEW r1c1')));
  assert.equal(durableCell(s, 2, 0).span, 2);
  // vertical, in the LIVE projection's shape (row probe 4): rowspan on the top, the continuation cell
  // omitted from the HTML but kept (empty) by the driver read -> FAIL_CLOSED_UNSUPPORTED_SHAPE, 0 dispatch
  const vertical = () => { const t = tables(); t[0].rows[0][0] = { ...cell('rs'), rowspan: 2 }; t[0].rows[1][0] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }; return t; };
  for (const target of ['rs', 'r2c2']) {
    const { x, driver } = await run(store({ tables: vertical() }), task(step(target, 'NEW')));
    assert.equal(reason(x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:vertical-merge', target);
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('closure 6: two cells in one session, retry: both SATISFIED, 0 write, 0 command, no barrier', async () => {
  const { second } = await firstAndSecond(store(), task(step('r0c0', 'A0'), { intent: 'replace_table_cell', selector: { tableCell: { text: 'r2c2' } }, text: 'A2' }));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
});
