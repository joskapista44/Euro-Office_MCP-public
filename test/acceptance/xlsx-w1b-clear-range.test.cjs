'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const clear = require('../../src/adapters/xlsx/families/clear.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 1b: clear_range, CONTENT ONLY, through the target path on the fake transport. The primitive
 * choice (SetValue('') per cell) is the wave-1b LIVE measurement: Clear() and ClearContents() also
 * reset number format and fill, which the owner rule forbids for a content-only clear. The fake's
 * 'clear-resets-format' fault reproduces that measured Clear() behaviour -- the proof must catch it.
 */

const book = () => [
  { name: 'Data', cells: {
    A1: { v: 'Name' }, B1: { v: 'Qty' }, A2: { v: 'Alpha' }, B2: { v: 3 }, C2: { v: null, f: '=B2*2' },
    D1: { v: 'old', nf: '0.00', fill: 16776960, bold: true }, D2: { v: 5, nf: '0.00', fill: 16776960, bold: false },
  }, merges: ['E1:F1'] },
  { name: 'Other', cells: { A1: { v: 'keep me' } } },
];

const task = (range = 'D1:D2', sheet = 'Data') => ({ id: 'm6-w1b', steps: [{ intent: 'clear_range', sheet, range }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

test('FIRST: APPLY -> 1 dispatch -> blank AND number format / fill kept -> L3 -> barrier; DURABLE cleared', async () => {
  const store = createStore(book());
  const { result, driver, m } = await run(store, task());
  assert.equal(result.ok, true);
  assert.equal(m.writes, 1);
  assert.equal(m.commands, 1);
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.APPLIED);
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  const d = store.durable.find(s => s.name === 'Data').cells;
  assert.equal(d.D1.v, null);
  assert.equal(d.D1.nf, '0.00', 'number format kept');
  assert.equal(d.D2.fill, 16776960, 'fill kept');
  assert.equal(d.D1.bold, true, 'cell bold kept');
  assert.equal(result.execution.preservationVerification[0].evidence.evidence.targetBold.changed.length, 0);
  assert.deepEqual(driver.reads[0].windows, [{ sheet: 'Data', range: 'D1:D2', format: true }], 'the window is read WITH format');
});

test('SECOND (reopened retry): SATISFIED, 0 write, 0 command, no barrier', async () => {
  const store = createStore(book());
  await run(store, task());
  const { result, m } = await run(store, task());
  assert.equal(result.ok, true);
  assert.equal(result.execution.noOp, true);
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('RED: a clear that also resets formatting (the measured Clear() behaviour) is NOT APPLIED, nothing persisted', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task(), { fault: 'clear-resets-format' });
  assert.equal(result.ok, false);
  assert.notEqual(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const pv = result.execution.primitiveVerification[0];
  assert.equal(pv.verdict, 'FAIL');
  assert.deepEqual(pv.evidence.evidence.formatChanged, ['D1', 'D2']);
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(store.durable.find(s => s.name === 'Data').cells.D1.v, 'old', 'durable untouched');
});

test('ignored or partial clear -> L1 FAIL; side effect elsewhere -> L3 FAIL', async () => {
  for (const fault of ['ignore', 'partial']) {
    const { result } = await run(createStore(book()), task(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
  }
  const { result } = await run(createStore(book()), task(), { fault: 'side-effect' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('a formula cell is cleared too (content includes the formula)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task('C2'));
  assert.equal(result.ok, true);
  const c2 = store.durable.find(s => s.name === 'Data').cells.C2;
  assert.ok(!c2.f && (c2.v === null || c2.v === undefined));
});

for (const [name, t, script, reason] of [
  ['in a merged area', task('E1'), {}, 'xlsx-clear-in-merged-area-refused'],
  ['on a missing sheet', task('A1', 'Nope'), {}, 'xlsx-sheet-not-found'],
  ['on a case-mismatched sheet name', task('A1', 'data'), {}, 'xlsx-sheet-not-found'],
  ['on a truncated read', task(), { readTruncated: true }, 'xlsx-observation-truncated'],
  ['when the number/text type cannot be read', task(), { countUnreadable: true }, 'xlsx-cell-type-unreadable'],
]) {
  test(`refused BEFORE dispatch: clear ${name}`, async () => {
    const store = createStore(book());
    const { result, driver } = await run(store, t, script);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, reason);
    assert.deepEqual(store.durable, book());
  });
}

test('fail closed when the target format was not read (a window without format)', () => {
  const op = planOperations(task()).operations[0];
  const primary = normalizeInventory({ sheets: [{ name: 'Data', index: 0, used: 'D1:D2' }], cells: [
    { sheet: 'Data', a: 'D1', v: 'old', f: 'old', count: 0, merge: 'D1' }, { sheet: 'Data', a: 'D2', v: '5', f: '5', count: 1, merge: 'D2' },
  ], truncated: false, unsupported: [] });
  assert.deepEqual(clear.classify(primary, op), { state: 'FAIL_CLOSED', evidence: { ok: false, reason: 'xlsx-target-format-unread', cell: 'D1' } });
});

test('planner: fields, range shape, ceiling, overlap with a write', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false); assert.match(p.detail, re); };
  refuse({ id: 'x', steps: [{ intent: 'clear_range', sheet: 'Data', range: 'D1', what: 'all' }] }, /does not support field/);
  refuse(task('Data!D1'), /plain A1 range/);
  refuse(task('A1:AL30'), /over the 1000 ceiling/);
  refuse({ id: 'x', steps: [{ intent: 'write_range', sheet: 'Data', range: 'D1:D2', values: [['a'], ['b']] }, { intent: 'clear_range', sheet: 'Data', range: 'D2' }] }, /overlapping-writes-refused/);
});

test('the clear command body compiles and clears through SetValue(\'\') only -- never Clear/ClearContents', () => {
  const op = planOperations(task('D1:E1')).operations[0];
  const cmd = clear.lower(op);
  const calls = [];
  const ws = { GetName: () => 'Data', GetRange: a => ({ SetValue: v => calls.push([a, v]) }) };
  assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: () => ws }), { ok: true, cleared: 2 });
  assert.deepEqual(calls, [['D1', ''], ['E1', '']]);
  assert.deepEqual([...cmd.body.matchAll(/\.([A-Z]\w*)\s*\(/g)].map(x => x[1]).sort(), ['GetName', 'GetRange', 'GetSheet', 'SetValue']);
  assert.doesNotMatch(cmd.body, /\.Clear\w*\s*\(/);
});

test('BOLD not observable BEFORE the write -> refused, 0 dispatch, 0 write', async () => {
  const store = createStore(book());
  const { result, driver, m } = await run(store, task(), { boldUnobservable: 'always' });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(m.writes, 0);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-target-bold-unobservable');
  assert.deepEqual(store.durable, book());
});

test('BOLD unreadable AFTER the write -> preservation NOT proven: never APPLIED, no barrier, no rewrite', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task(), { boldUnobservable: 'after-dispatch' });
  const e = result.execution.appliedOperations[0];
  assert.equal(result.ok, false);
  assert.notEqual(e.status, OPERATION_STATUS.APPLIED);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  const tb = result.execution.preservationVerification[0].evidence.evidence.targetBold;
  assert.equal(tb.unobservable.length, 2);
  assert.equal(tb.changed.length, 0, 'unreadable is never reported as a change to false');
  assert.equal(driver.calls.dispatch, 1, 'exactly one dispatch -- no automatic rewrite');
  assert.equal(driver.calls.forceSave, 0);
});

test('an OBSERVED bold loss -> preservation FAIL, never APPLIED (the Clear() class the probe caught)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(), { fault: 'clear-drops-bold' });
  const e = result.execution.appliedOperations[0];
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'PASS', 'blank and number format / fill are fine');
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
  assert.deepEqual(result.execution.preservationVerification[0].evidence.evidence.targetBold.changed, [{ cell: 'D1', before: true, after: false }]);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
});

test('a clear that SHRINKS the used range: the cells it no longer covers were blank before -> not a change (W3 LIVE #ref)', async () => {
  // the only content outside A1:D2 is the cell being cleared, so the used range grows to F4 and back
  const b = book();
  b[0].cells.F4 = { v: 'x', nf: 'General', fill: 'S:none', bold: false };
  const store = createStore(b);
  const { result } = await run(store, task('F4'));
  assert.equal(result.ok, true, JSON.stringify(result.execution.preservationVerification));
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(store.durable[0].cells.F4.v, null);
});

test('the used-range allowance is narrow: content gone unread, or a blank unread INSIDE the used range, is still a change', () => {
  const cells = require('../../src/adapters/xlsx/families/cells.cjs');
  const obs = (used, cs) => ({ primary: normalizeInventory({ sheets: [{ name: 'S', index: 0, used }], cells: cs.map(([a, v]) => ({ sheet: 'S', a, v, f: '', count: 0, merge: a })) }) });
  const op = { selector: { sheet: 'S', range: 'A1' }, cells: [{ a: 'A1' }] };
  const before = obs('A1:C3', [['A1', 'x'], ['B2', ''], ['C3', 'keep']]);
  // B2 blank, outside the after used range (A1:A1) -> not a change; C3 had content -> a change
  const r1 = cells.verifyPreservation(op, before, obs('A1', [['A1', '']]));
  assert.deepEqual(r1.evidence.changed.map(c => c.cell), ['S!C3']);
  // B2 blank but still INSIDE the after used range and not read -> a change
  const r2 = cells.verifyPreservation(op, obs('A1:C3', [['A1', 'x'], ['B2', '']]), obs('A1:C3', [['A1', '']]));
  assert.deepEqual(r2.evidence.changed.map(c => [c.cell, c.change]), [['S!B2', 'unreadable-after']]);
});

test('NEGATIVE CONTROL for the used-range allowance: a clear that also ERASES content which then drops out of the used range is caught (L3 FAIL, not persisted)', async () => {
  // before: F4 (the target) and G5 (content) stretch the used range; the clear wrongly erases G5 too,
  // so after it neither is read -- G5 had content, so its absence is a change, not an allowance
  const b = book();
  b[0].cells.F4 = { v: 'x', nf: 'General', fill: 'S:none', bold: false };
  b[0].cells.G5 = { v: 'must stay' };
  const store = createStore(b);
  const { result, driver } = await run(store, task('F4'), { fault: 'side-erase', eraseCells: ['G5'] });
  assert.equal(result.ok, false);
  const pv = result.execution.preservationVerification[0];
  assert.equal(pv.verdict, 'FAIL');
  assert.deepEqual(pv.evidence.evidence.changed.map(c => [c.cell, c.change]), [['Data!G5', 'unreadable-after']]);
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(store.durable[0].cells.G5.v, 'must stay', 'durable untouched');
});
