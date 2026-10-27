'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * GH #83 (an internal branch), the LIVE-confirmed perimeter
 * design (docs/GH83-BORDER-MERGED-DESIGN-20261014.md, a maintainer's rounds 1+2): border.mode 'cells' on a range that
 * CONTAINS a merged area now writes each merge member cell ONLY the requested sides on that merge's OWN
 * perimeter (mergeAwareSides, format.cjs), instead of refusing the whole task. Every other border mode
 * (inside/all/remove) stays refused for any merge membership, unchanged (their shared-edge logic has no defined
 * meaning for a multi-cell block). X-B10 (owner GO 36696) lifted the refusal for 'outline' over a WHOLE merge only:
 * xb10-format-merged-axes.test.cjs.
 */

const book = () => [
  { name: 'Data', cells: {
    A1: { v: 'Title' }, B1: { v: '' }, C1: { v: '' }, D1: { v: '' }, E1: { v: 'next' },
    F1: { v: 'g' }, G1: { v: 'h' },
  }, merges: ['A1:D1'] },
];

const bd = (range, format, sheet = 'Data') => ({ id: 'gh83', steps: [{ intent: 'format_range', sheet, range, format: { border: format } }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const cell = (store, a) => store.durable.find(s => s.name === 'Data').cells[a];
const SIDES = { mode: 'cells', sides: ['top', 'bottom', 'left', 'right'], style: 'thin', color: '#000000' };

test("GH #83: border.mode 'cells' on a range CONTAINING a merge -- APPLIED, each member gets its OWN perimeter sides", async () => {
  const store = createStore(book());
  const { result, m } = await run(store, bd('A1:D1', SIDES));
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(m.writes, 1);
  // leftmost (anchor): left + top + bottom; middle: top + bottom only; rightmost: right + top + bottom
  assert.deepEqual(cell(store, 'A1').border, { l: { s: 7, c: 0 }, t: { s: 7, c: 0 }, b: { s: 7, c: 0 } });
  assert.deepEqual(cell(store, 'B1').border, { t: { s: 7, c: 0 }, b: { s: 7, c: 0 } });
  assert.deepEqual(cell(store, 'C1').border, { t: { s: 7, c: 0 }, b: { s: 7, c: 0 } });
  assert.deepEqual(cell(store, 'D1').border, { r: { s: 7, c: 0 }, t: { s: 7, c: 0 }, b: { s: 7, c: 0 } });
  // the merge survives the write
  assert.ok(store.durable.find(s => s.name === 'Data').merges.includes('A1:D1'));
  // INTERIOR edges (A1 right, B1/C1 left+right, D1 left) are untouched -- no border key at all
  assert.equal(cell(store, 'A1').border.r, undefined);
  assert.equal(cell(store, 'B1').border.l, undefined);
  assert.equal(cell(store, 'D1').border.l, undefined);

  const again = await run(store, bd('A1:D1', SIDES));
  assert.equal(again.result.ok, true);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.m.commands, 0, 'retry writes nothing -- the whole perimeter already holds');
});

test("GH #83: a range mixing a merge with ordinary cells -- the merge gets its perimeter, the plain cell gets every side", async () => {
  const store = createStore(book());
  const { result } = await run(store, bd('A1:E1', SIDES)); // E1 is a plain, non-merged cell
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.deepEqual(cell(store, 'E1').border, { t: { s: 7, c: 0 }, b: { s: 7, c: 0 }, l: { s: 7, c: 0 }, r: { s: 7, c: 0 } });
  assert.deepEqual(cell(store, 'D1').border, { r: { s: 7, c: 0 }, t: { s: 7, c: 0 }, b: { s: 7, c: 0 } });
});

test("GH #83 control: the SAME shape with NO merge at all -- every cell gets every requested side (unchanged behaviour)", async () => {
  const store = createStore([{ name: 'Data', cells: { F1: { v: 1 }, G1: { v: 1 } }, merges: [] }]);
  const { result } = await run(store, bd('F1:G1', SIDES));
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.deepEqual(cell(store, 'F1').border, { t: { s: 7, c: 0 }, b: { s: 7, c: 0 }, l: { s: 7, c: 0 }, r: { s: 7, c: 0 } });
  assert.deepEqual(cell(store, 'G1').border, { t: { s: 7, c: 0 }, b: { s: 7, c: 0 }, l: { s: 7, c: 0 }, r: { s: 7, c: 0 } });
});

for (const mode of ['inside', 'all']) {
  test(`GH #83: border.mode '${mode}' on a range containing the SAME merge is STILL refused -- the exception is 'cells' (and, X-B10, a whole-merge 'outline')`, async () => {
    const store = createStore(book());
    const { result } = await run(store, bd('A1:D1', { mode, style: 'thin', color: '#000000' }));
    assert.equal(result.ok, false);
    assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-in-merged-area-refused');
  });
}

test("GH #83 review: a merge that CROSSES the range boundary is refused before any write (a partial frame would pass L1)", async () => {
  for (const range of ['B1', 'A1:B1', 'C1:E2']) {
    const store = createStore(book());
    const { result } = await run(store, bd(range, SIDES));
    assert.equal(result.ok, false, range);
    assert.match(JSON.stringify(result), /xlsx-format-merge-crosses-range-refused/, range);
    assert.equal(cell(store, 'B1').border, undefined, `${range}: nothing written`);
  }
});

test("GH #83 RED: without the merge-aware exception (an 'inside' border on the merge), the task is refused BEFORE any write", async () => {
  const store = createStore(book());
  const before = JSON.stringify(cell(store, 'A1'));
  const { result } = await run(store, bd('A1:D1', { mode: 'inside', style: 'thin', color: '#000000' }));
  assert.equal(result.ok, false);
  assert.equal(JSON.stringify(cell(store, 'A1')), before, '0 mutation on refusal');
});

test('GH #83: RED mutant -- if a cell write leaked onto an INTERIOR edge, L1 catches it (drift not tolerated)', async () => {
  const store = createStore(book());
  const { result } = await run(store, bd('A1:D1', SIDES), { fault: 'gh83-interior-leak' });
  assert.equal(result.ok, false, 'an interior-edge leak must fail L1/L3, not pass silently');
});
