'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * GH #290 offline acceptance: format_text center tab, border side removal ({ style: 'none' }), and a pagination flag set
 * to false -- proven by an EXPLICIT direct false, never by an absent flag (a style may set it). LIVE basis:
 * docs/evidence/gh290-paragraph-format-public-api-20261021.json.
 */

const RICH = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const T = 'GH290 target paragraph.';
const O = 'GH290 other paragraph.';
const single = { style: 'single', size: 1, color: '000000' };
const doc = (targetFormat = {}, otherFormat = {}) => [
  { text: 'GH290 heading', level: 1, props: { ...RICH, bold: true } },
  { text: T, props: { ...RICH }, format: targetFormat },
  { text: O, props: { ...RICH }, format: otherFormat },
];

async function run(store, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = format => ({ id: 'gh290', steps: [{ intent: 'format_text', selector: { text: T }, format }] });
const status = x => x.appliedOperations[0].status;
const fmt = (s, text) => { const p = s.durable.find(x => x.text === text); return { ...(p.format || {}) }; };

test('F-01: a center tab -- APPLIED, stored as center; retry 0/0', async () => {
  const s = createStore(doc());
  const r = await run(s, task({ tabs: [{ pos: 144, align: 'center' }] }));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual(fmt(s, T).tabs, [{ pos: 144, align: 'center' }]);
  const again = await run(s, task({ tabs: [{ pos: 144, align: 'center' }] }));
  assert.equal(status(again.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
});

test('F-04: borderTop { style: none } on four single borders -- the top side is an explicit none, the other three unchanged', async () => {
  const s = createStore(doc({ borderTop: { ...single }, borderBottom: { ...single }, borderLeft: { ...single }, borderRight: { ...single } }));
  const r = await run(s, task({ borderTop: { style: 'none' } }));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
  const f = fmt(s, T);
  assert.equal(f.borderTop.style, 'none');
  for (const k of ['borderBottom', 'borderLeft', 'borderRight']) assert.equal(f[k].style, 'single', k);
  const again = await run(s, task({ borderTop: { style: 'none' } }));
  assert.equal(status(again.x), OPERATION_STATUS.SATISFIED);
});

test('F-04: an ABSENT side is not a removed one -- { style: none } on a paragraph with no border still writes the explicit none', async () => {
  const s = createStore(doc());
  const r = await run(s, task({ borderTop: { style: 'none' } }));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED);
  assert.equal(fmt(s, T).borderTop.style, 'none');
});

test('F-06: keepNext false where the flag is ABSENT (a style may set it) -- APPLIED as an explicit false, never SATISFIED from the absence; retry 0/0', async () => {
  const s = createStore(doc({ keepNext: false, paginationDirect: { keepNext: null, keepLines: null, pageBreakBefore: null, widowControl: null } }));
  const r = await run(s, task({ keepNext: false }));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(fmt(s, T).paginationDirect.keepNext, false);
  const again = await run(s, task({ keepNext: false }));
  assert.equal(status(again.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
});

test('F-06: pageBreakBefore true -> false, keepLines true -> false in one format', async () => {
  const s = createStore(doc({ pageBreakBefore: true, keepLines: true, paginationDirect: { keepNext: null, keepLines: true, pageBreakBefore: true, widowControl: null } }));
  const r = await run(s, task({ pageBreakBefore: false, keepLines: false }));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
  const f = fmt(s, T);
  assert.deepEqual([f.pageBreakBefore, f.keepLines, f.paginationDirect.pageBreakBefore, f.paginationDirect.keepLines], [false, false, false, false]);
});

test('F-06: a write that leaves the flag ABSENT instead of an explicit false is never APPLIED (L1)', async () => {
  const s = createStore(doc({ keepNext: true, paginationDirect: { keepNext: true, keepLines: null, pageBreakBefore: null, widowControl: null } }));
  const r = await run(s, task({ keepNext: false }), { fault: 'format-flag-not-explicit' });
  assert.notEqual(status(r.x), OPERATION_STATUS.APPLIED);
});

test('F-07 (LIVE, gh289-290-remaining-probe-20261023.json): widowControl false on a direct true and where ABSENT -- APPLIED as an explicit false, never SATISFIED from the absence; retry 0/0', async () => {
  for (const before of [{ widowControl: true, paginationDirect: { keepNext: null, keepLines: null, pageBreakBefore: null, widowControl: true } }, { widowControl: false, paginationDirect: { keepNext: null, keepLines: null, pageBreakBefore: null, widowControl: null } }]) {
    const s = createStore(doc(before));
    const r = await run(s, task({ widowControl: false }));
    assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
    assert.equal(fmt(s, T).paginationDirect.widowControl, false);
    const again = await run(s, task({ widowControl: false }));
    assert.equal(status(again.x), OPERATION_STATUS.SATISFIED);
    assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
  }
});

test('F-02 (LIVE, gh290-paragraph-format-public-api-20261021.json P3): tabs [] removes the paragraph\'s own tabs -- APPLIED, the other paragraph\'s tabs kept; retry 0/0', async () => {
  const s = createStore(doc({ tabs: [{ pos: 72, align: 'left' }, { pos: 216, align: 'right' }] }, { tabs: [{ pos: 36, align: 'left' }] }));
  const r = await run(s, task({ tabs: [] }));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual(fmt(s, T).tabs, []);
  assert.deepEqual(fmt(s, O).tabs, [{ pos: 36, align: 'left' }]);
  const again = await run(s, task({ tabs: [] }));
  assert.equal(status(again.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
});

test('refused before dispatch: decimal tab, a double border, none with a size, a non-array tabs', async () => {
  for (const format of [{ tabs: [{ pos: 36, align: 'decimal' }] }, { borderTop: { style: 'double', size: 1, color: '#000000' } }, { borderTop: { style: 'none', size: 1 } }, { tabs: 'none' }]) {
    const { x, driver } = await run(createStore(doc()), task(format));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(format));
    assert.equal(driver.calls.dispatch, 0);
  }
});
