'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const tf = require('../../src/adapters/docx/families/table-format.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * W04 offline acceptance: format_table through the TARGET path on the fake transport, modeled on
 * test/acceptance/docx-w11-page-setup.test.cjs's own shape. The fake model follows the LIVE
 * probes (W04-1a/1b/1c/2/4, #ref/#ref/#ref/#ref/#ref): six fields, all measured end to
 * end -- borders (6 edges + 'none'), shading, cellMargins (4 edges), indent, width, style (name ->
 * id, resolved against a read-only tableStyles list). See docs/W04-TABLE-FORMAT-DESIGN-
 * 20261007.md (v3).
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = text => ({ span: 1, paras: [{ text, props: { ...PLAIN } }] });
const paragraphs = () => [{ text: 'Body heading', level: 1, props: {} }, { text: 'Body first.', props: {} }];
const oneTable = (anchor = 'Header cell') => [{ rows: [[cell(anchor), cell('r0c1')], [cell('r1c0'), cell('r1c1')]] }];
const twoTables = () => [{ rows: [[cell('T1 anchor'), cell('a1')]] }, { rows: [[cell('T2 anchor'), cell('a2')]] }];
const store = (tables = oneTable(), paras = paragraphs(), tableStyles) => createStore(paras, { doc: { tables, ...(tableStyles ? { tableStyles } : {}) } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w04', steps });
const ft = (fields, table = { text: 'Header cell' }, extra = {}) => ({ intent: 'format_table', table, ...fields, ...extra });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const tableOf = (s, i) => s.durableDoc.tables[i];
const snapshot = s => JSON.stringify([s.durable.map(p => p.text), s.durableDoc.tables]);

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

test('borders: all 6 edges, single write + retry SATISFIED 0/0 -- distinct colors each, no cross-edge bleed', async () => {
  const s = store();
  const borders = { top: { style: 'single', size: 4, color: '#FF0000' }, bottom: { style: 'single', size: 4, color: '#00FF00' }, left: { style: 'single', size: 4, color: '#0000FF' }, right: { style: 'single', size: 4, color: '#FFFF00' }, insideH: { style: 'single', size: 4, color: '#FF00FF' }, insideV: { style: 'single', size: 4, color: '#00FFFF' } };
  await firstAndSecond(s, task(ft({ borders })));
  const got = tableOf(s, 0).format.borders;
  assert.equal(got.top.value, 'single'); assert.deepEqual(got.top.color, { auto: false, r: 255, g: 0, b: 0 });
  assert.deepEqual(got.bottom.color, { auto: false, r: 0, g: 255, b: 0 });
  assert.deepEqual(got.left.color, { auto: false, r: 0, g: 0, b: 255 });
  assert.deepEqual(got.right.color, { auto: false, r: 255, g: 255, b: 0 });
  assert.deepEqual(got.insideH.color, { auto: false, r: 255, g: 0, b: 255 });
  assert.deepEqual(got.insideV.color, { auto: false, r: 0, g: 255, b: 255 });
});

test("border 'none': MEASURED -- forbids size/color at plan time, writes without them", async () => {
  const s = store();
  await firstAndSecond(s, task(ft({ borders: { top: { style: 'none' } } })));
  assert.equal(tableOf(s, 0).format.borders.top.value, 'none');
  const { x, driver } = await run(store(), task(ft({ borders: { top: { style: 'none', size: 4, color: '#FF0000' } } })));
  assert.equal(x.outcome, 'plan-failed', JSON.stringify(x));
  assert.equal(driver.calls.dispatch, 0);
});

test('shading: write + retry SATISFIED 0/0', async () => {
  const s = store();
  await firstAndSecond(s, task(ft({ shading: '#008000' })));
  assert.deepEqual(tableOf(s, 0).format.shading.fill, { r: 0, g: 128, b: 0 });
});

test('cellMargins: all four edges, write + retry SATISFIED 0/0', async () => {
  const s = store();
  await firstAndSecond(s, task(ft({ cellMargins: { top: 100, right: 200, bottom: 300, left: 400 } })));
  const m = tableOf(s, 0).format.cellMargins;
  assert.deepEqual([m.top.w, m.right.w, m.bottom.w, m.left.w], [100, 200, 300, 400]);
});

test('indent + width: write + retry SATISFIED 0/0', async () => {
  const s = store();
  await firstAndSecond(s, task(ft({ indent: 720, width: 3600 })));
  assert.equal(tableOf(s, 0).format.indent, 720);
  assert.deepEqual(tableOf(s, 0).format.width, { type: 'dxa', w: 3600 });
});

test('style: resolved by NAME against the tableStyles list, write + retry SATISFIED 0/0', async () => {
  const s = store();
  await firstAndSecond(s, task(ft({ style: 'Table Grid' })));
  assert.equal(tableOf(s, 0).format.style, '701');
});

test('ALL SIX FIELDS in one step: every axis changes exactly once, no cross-axis bleed', async () => {
  const s = store();
  await firstAndSecond(s, task(ft({
    borders: { top: { style: 'single', size: 4, color: '#FF0000' } },
    shading: '#008000', cellMargins: { top: 100, right: 200, bottom: 300, left: 400 },
    indent: 720, width: 3600, style: 'Table Grid',
  })));
  const f = tableOf(s, 0).format;
  assert.equal(f.borders.top.value, 'single');
  assert.deepEqual(f.shading.fill, { r: 0, g: 128, b: 0 });
  assert.equal(f.cellMargins.left.w, 400);
  assert.equal(f.indent, 720);
  assert.deepEqual(f.width, { type: 'dxa', w: 3600 });
  assert.equal(f.style, '701');
});

test('TABLE_TARGETING: two tables, each anchored by a distinct cell -- writing one leaves the other untouched', async () => {
  const s = store(twoTables());
  await firstAndSecond(s, task(ft({ indent: 720 }, { text: 'T2 anchor' })));
  assert.equal(tableOf(s, 0).format.indent, undefined);
  assert.equal(tableOf(s, 1).format.indent, 720);
});

test('negatives before dispatch -> plan-failed or FAIL_CLOSED, 0 dispatch, nothing written', async () => {
  const planFailed = [
    ft({}), // no field at all
    { intent: 'format_table', indent: 720 }, // no table selector
    ft({ borders: { top: { style: 'single', size: 4 } } }), // partial border edge
    ft({ borders: { top: { style: 'single', size: 4, color: 'red' } } }), // bad color
    ft({ borders: { diagonal: { style: 'single', size: 4, color: '#FF0000' } } }), // unknown edge
    ft({ cellMargins: { top: 100 } }), // partial cellMargins
    ft({ width: 0 }), // width must be positive
    ft({ indent: -1 }), // indent must be non-negative
    ft({ style: '' }), // empty style name
    { intent: 'format_table', table: { text: 'Header cell' }, indent: 720, orientation: 'landscape' }, // unknown field
  ];
  for (const st of planFailed) {
    const s = store();
    const before = snapshot(s);
    const { x, driver } = await run(s, task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(snapshot(s), before, `${JSON.stringify(st)}: nothing written`);
  }
  // table anchor not found -> FAIL_CLOSED, never a guess
  const nf = await run(store(), task(ft({ indent: 720 }, { text: 'No such cell' })));
  assert.equal(status(nf.x), OPERATION_STATUS.FAILED);
  assert.equal(reason(nf.x), 'table-anchor-not-found');
  assert.equal(nf.driver.calls.dispatch, 0);
});

test('FAIL_CLOSED: the whole table-format read unobservable -> inventory.cjs normalizes it to a PER-FIELD unobservable string, caught by the SAME path as a single unobservable field', async () => {
  const { x, driver } = await run(store(), task(ft({ indent: 720 })), { tableFormatUnreadable: true });
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(reason(x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:table-format-field-unobservable:indent');
  assert.equal(driver.calls.dispatch, 0);
});

test('FAIL_CLOSED: one requested border edge specifically unobservable -> FAIL_CLOSED before dispatch (page-setup\'s unobservableFields discipline, by analogy)', async () => {
  const { x, driver } = await run(store(), task(ft({ borders: { top: { style: 'single', size: 4, color: '#FF0000' } } })), { tableBorderEdgeUnreadable: 'top' });
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(reason(x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:table-format-field-unobservable:borders');
  assert.equal(driver.calls.dispatch, 0);
});

test('style resolution: not-found and ambiguous -> FAIL_CLOSED, never a guessed id', async () => {
  const nf = await run(store(), task(ft({ style: 'No Such Style' })));
  assert.equal(reason(nf.x), 'table-style-not-found');
  assert.equal(nf.driver.calls.dispatch, 0);
  const amb = await run(store(oneTable(), paragraphs(), [{ name: 'Table Grid', styleId: 701 }, { name: 'Table Grid', styleId: 702 }]), task(ft({ style: 'Table Grid' })));
  assert.equal(reason(amb.x), 'table-style-ambiguous');
  assert.equal(amb.driver.calls.dispatch, 0);
  const unr = await run(store(), task(ft({ style: 'Table Grid' })), { tableStylesUnreadable: true });
  assert.equal(reason(unr.x), 'table-style-list-unobservable');
  assert.equal(unr.driver.calls.dispatch, 0);
});

test("faults at L3 (P1): a sibling field on the SAME table changing, another table changing, or the body changing -> NOT APPLIED", async () => {
  {
    const { x } = await run(store(), task(ft({ shading: '#008000' })), { fault: 'tf-other-field-changed' });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', 'tf-other-field-changed: the requested field itself is right');
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'tf-other-field-changed');
  }
  {
    const { x } = await run(store(twoTables()), task(ft({ indent: 720 }, { text: 'T1 anchor' })), { fault: 'tf-other-table' });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', 'tf-other-table: the requested field itself is right');
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'tf-other-table');
  }
  {
    const { x } = await run(store(), task(ft({ indent: 720 })), { fault: 'tf-body-change' });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS');
    assert.equal(x.preservationVerification[0].verdict, 'FAIL');
  }
});

test("REGRESSION PROOF: a border edge landing on the WRONG edge (the exact start/end key-mapping class, #ref) must FAIL L1, not read back as correct", async () => {
  const { x } = await run(store(), task(ft({ borders: { left: { style: 'single', size: 4, color: '#00FF00' } } })), { fault: 'tf-border-edge-swapped' });
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL', 'a border written to the wrong edge must not read back as correct on the requested edge');
});

test('retry: a correct terminal table is SATISFIED without a write; an unresolved earlier run -> FAIL_CLOSED', async () => {
  const s = store([{ rows: oneTable()[0].rows, format: { indent: 720 } }]);
  const { x, m } = await run(s, task(ft({ indent: 720 })));
  assert.equal(status(x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  const r = await run(store([{ rows: oneTable()[0].rows, format: { indent: 720 } }]), { ...task(ft({ indent: 720 })), recovery: { unresolvedUncertainOperations: ['s0'] } });
  assert.equal(reason(r.x), 'preservation-baseline-unavailable');
  assert.equal(r.driver.calls.dispatch, 0);
});

test('lowering: public table API only, colors pre-split into raw r/g/b, style resolved by exact name+type, no internal/cursor call', () => {
  const op = planOperations(task(ft({ borders: { top: { style: 'single', size: 4, color: '#FF0000' } }, shading: '#008000', style: 'Table Grid' }))).operations[0];
  const target = { kind: 'table_format', tableIndex: 0, anchorText: 'Header cell', tableAnchor: null, fields: op.fields, expectedStyleId: 701 };
  const low = tf.lower(op, target);
  assert.match(low.body, /table\[setter\]\('single', req\.size, 0, req\.r, req\.g, req\.b\)/);
  assert.match(low.body, /table\.SetShd\('clear', m\.fields\.shading\.r, m\.fields\.shading\.g, m\.fields\.shading\.b\)/);
  assert.match(low.body, /GetType\(\) === 'table'/);
  assert.match(low.body, /GetName\(\)\) === m\.fields\.style\.name/);
  assert.doesNotMatch(low.body, /evaluate|asc_|private_|\.Document\b|MoveCursor|SetDocPosType|CreateRGBColor/);
  assert.equal(op.preservationMode, 'PRIMITIVE_INTEGRAL');
});

test("W04-6: commandBody is ALL-OR-NOTHING -- if ONE requested setter is missing, NO setter is called at all (never a partial write)", () => {
  const vm = require('node:vm');
  const op = planOperations(task(ft({ borders: { top: { style: 'single', size: 4, color: '#FF0000' } }, indent: 720 }))).operations[0];
  const target = { kind: 'table_format', tableIndex: 0, anchorText: 'Header cell', tableAnchor: null, fields: op.fields, expectedStyleId: null };
  const low = tf.lower(op, target);
  const calls = [];
  const cell = { GetContent: () => ({ GetText: () => 'Header cell' }) };
  const table = {
    GetRowsCount: () => 1, GetRow: () => ({ GetCellsCount: () => 1, GetCell: () => cell }),
    SetTableBorderTop: (...a) => calls.push(['SetTableBorderTop', ...a]),
    // SetTableInd deliberately ABSENT -- indent was requested too, and this must block the write.
  };
  const doc = { GetAllTables: () => [table] };
  const Api = { GetDocument: () => doc };
  new vm.Script(`(function(Api){${low.body}})`).runInThisContext()(Api);
  assert.deepEqual(calls, [], 'a missing setter for ONE requested field must prevent EVERY write, not just its own');
});
