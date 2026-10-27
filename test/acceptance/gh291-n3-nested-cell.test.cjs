'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * GH #291 N3 offline acceptance: replace_table_cell on a cell INSIDE a nested table (one level down, bound to its holding cell),
 * on the fake transport whose nested model follows the LIVE probe (docs/evidence/gh291-n3-nested-cell-probe-20261022.json): the
 * write's command reaches a nested cell through GetAllTables(); an anchor of the HOLDING table does not; the holding cell's
 * read text runs on into the nested text.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = text => ({ span: 1, paras: [{ text, props: { ...PLAIN } }] });
const grid = (tag, rows, cols = rows) => ({ rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (__, c) => cell(`${tag} r${r}c${c}`))) });
/** OUT 2x2 whose cell (0,0) holds IN 2x2, and a plain T3 after it */
const tables = (holderText = 'OUT r0c0') => { const out = grid('OUT', 2); out.rows[0][0] = { ...cell(holderText), nested: [grid('IN', 2)] }; return [out, grid('T3', 2)]; };
const store = t => createStore([{ text: 'Nested heading', level: 1, props: { ...PLAIN, bold: true } }, { text: 'Between.', props: { ...PLAIN } }], { doc: { tables: t } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'n3', steps });
const rep = (text, newText, tableAnchor) => ({ intent: 'replace_table_cell', selector: { tableCell: tableAnchor ? { text, tableAnchor } : { text } }, text: newText });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const nestedTexts = s => s.durableDoc.tables[0].rows[0][0].nested[0].rows.map(r => r.map(c => c.paras.map(p => p.text).join(' ')));
const topTexts = s => s.durableDoc.tables.map(t => t.rows.map(r => r.map(c => c.paras.map(p => p.text).join(' '))));

test('GH #291 N3: a nested cell named with an anchor IN the nested table, or without an anchor -> APPLIED (L1/L3/whole PASS), only that cell changes; retry SATISFIED 0/0', async () => {
  for (const [label, target, anchor, want] of [
    ['nested anchor', 'IN r0c0', 'IN r1c1', [['IN new', 'IN r0c1'], ['IN r1c0', 'IN r1c1']]],
    ['no anchor', 'IN r0c1', null, [['IN r0c0', 'IN new'], ['IN r1c0', 'IN r1c1']]],
  ]) {
    const s = store(tables());
    const top = topTexts(s);
    const first = await run(s, task(rep(target, 'IN new', anchor)));
    assert.equal(status(first.x), OPERATION_STATUS.APPLIED, `${label}: ${JSON.stringify(first.x.appliedOperations.map(e => e.evidence))}`);
    assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), `${label}: L1 ${JSON.stringify(first.x.primitiveVerification)}`);
    assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), `${label}: L3 ${JSON.stringify(first.x.preservationVerification)}`);
    assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', label);
    assert.deepEqual(nestedTexts(s), want, label);
    assert.deepEqual(topTexts(s), top, `${label}: every top-level cell (the holder's own text too) unchanged`);
    const second = await run(s, task(rep(target, 'IN new', anchor)));
    assert.equal(status(second.x), OPERATION_STATUS.SATISFIED, label);
    assert.equal(second.m.writes, 0, label);
  }
});

test('GH #291 N3: a nested cell named with an anchor of the HOLDING table -> FAILED, 0 dispatch (the holding table\'s cells do not include it -- measured, the command finds none)', async () => {
  const s = store(tables());
  const before = JSON.stringify(s.durableDoc.tables);
  const { x, driver } = await run(s, task(rep('IN r0c0', 'IN new', 'OUT r1c1')));
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(reason(x), 'table-cell-terminal-unprovable');
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(JSON.stringify(s.durableDoc.tables), before);
});

test('GH #291 N3: the holding cell\'s own text repeating the nested target\'s text -> FAIL_CLOSED nested-table-holder-text-unprovable, 0 dispatch', async () => {
  const s = store(tables('IN r0c0'));
  const { x, driver } = await run(s, task(rep('IN r0c0', 'IN new', 'IN r1c1')));
  assert.equal(reason(x), 'nested-table-holder-text-unprovable');
  assert.equal(driver.calls.dispatch, 0);
});

test('GH #291 N3 L3: a write that also touches another cell of the nested table is caught (never a silent APPLIED)', async () => {
  const s = store(tables());
  const { x } = await run(s, task(rep('IN r0c0', 'IN new', 'IN r1c1')), { fault: 'cell-touch-other' });
  assert.equal(x.preservationVerification[0].verdict, 'FAIL', JSON.stringify(x.preservationVerification));
  assert.notEqual(status(x), OPERATION_STATUS.APPLIED);
});

/** GH #291 N3x: OUT 2x2 whose cell (0,0) holds MID 2x2, whose cell (0,0) holds IN 2x2 (two levels down); a third level DEEP under IN */
const twoLevel = (deep = false) => {
  const inner = grid('IN', 2);
  if (deep) inner.rows[0][0] = { ...cell('IN r0c0'), nested: [grid('DEEP', 2)] };
  const mid = grid('MID', 2); mid.rows[0][0] = { ...cell('MID r0c0'), nested: [inner] };
  const out = grid('OUT', 2); out.rows[0][0] = { ...cell('OUT r0c0'), nested: [mid] };
  return [out, grid('T3', 2)];
};
const inTexts = s => s.durableDoc.tables[0].rows[0][0].nested[0].rows[0][0].nested[0].rows.map(r => r.map(c => c.paras.map(p => p.text).join(' ')));

test('GH #291 N3x (LIVE, gh291-nested-remaining-probe-20261023.json): a cell TWO levels down, with an anchor in its own table or without one -> APPLIED, only that cell changes; retry SATISFIED 0/0', async () => {
  for (const [label, target, anchor, want] of [
    ['nested anchor', 'IN r1c0', 'IN r1c1', [['IN r0c0', 'IN r0c1'], ['IN new', 'IN r1c1']]],
    ['no anchor', 'IN r0c1', null, [['IN r0c0', 'IN new'], ['IN r1c0', 'IN r1c1']]],
  ]) {
    const s = store(twoLevel());
    const top = topTexts(s);
    const first = await run(s, task(rep(target, 'IN new', anchor)));
    assert.equal(status(first.x), OPERATION_STATUS.APPLIED, `${label}: ${JSON.stringify(first.x.appliedOperations.map(e => e.evidence))}`);
    assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), `${label}: L1 ${JSON.stringify(first.x.primitiveVerification)}`);
    assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), `${label}: L3 ${JSON.stringify(first.x.preservationVerification)}`);
    assert.deepEqual(inTexts(s), want, label);
    assert.deepEqual(topTexts(s), top, `${label}: every top-level cell unchanged`);
    const second = await run(s, task(rep(target, 'IN new', anchor)));
    assert.equal(status(second.x), OPERATION_STATUS.SATISFIED, label);
    assert.equal(second.m.writes, 0, label);
  }
});

test('GH #291 N3x: a cell THREE levels down is not offered (not measured) -> FAILED table-cell-terminal-unprovable, 0 dispatch; a middle holder repeating the text -> nested-table-holder-text-unprovable', async () => {
  for (const [anchor, want] of [['DEEP r1c1', 'table-anchor-not-found'], [null, 'table-cell-terminal-unprovable']]) {
    const s = store(twoLevel(true));
    const before = JSON.stringify(s.durableDoc.tables);
    const { x, driver } = await run(s, task(rep('DEEP r0c1', 'DEEP new', anchor)));
    assert.equal(reason(x), want, want);
    assert.equal(driver.calls.dispatch, 0, want);
    assert.equal(JSON.stringify(s.durableDoc.tables), before, want);
  }
  const t = twoLevel(); t[0].rows[0][0].nested[0].rows[0][0].paras = [{ text: 'IN r0c1', props: { ...PLAIN } }];
  const r = await run(store(t), task(rep('IN r0c1', 'IN new', 'IN r1c1')));
  assert.equal(reason(r.x), 'nested-table-holder-text-unprovable');
  assert.equal(r.driver.calls.dispatch, 0);
});
