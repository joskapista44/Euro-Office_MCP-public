'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

/**
 * X-B10 / L-15 (owner GO "A", owner channel; LIVE: test/live/xlsx-xb10-merged-format-axes-probe.live.cjs,
 * 1x3 + 2x2 merge, session + reopen): align (h+v), wrap, numberFormat, fill 'none' (written as 'No Fill') and the border
 * OUTLINE on a range covering a WHOLE merged area are no longer refused -- every member is written and verified, the merge and
 * the anchor value survive. ClearFormats() dissolves the merge LIVE, so the merged fill 'none' never runs the X-F08 composite.
 * Everything else keeps refusing merge membership, named.
 */

const book = () => [
  { name: 'Data', cells: {
    A1: { v: 'Title', fill: 0xff0000 }, B1: { v: '', fill: 0xff0000 }, C1: { v: '', fill: 0xff0000 }, // merge #1 (1x3)
    E1: { v: 'Box', fill: 0xff0000 }, F1: { v: '', fill: 0xff0000 }, E2: { v: '', fill: 0xff0000 }, F2: { v: '', fill: 0xff0000 }, // merge #2 (2x2)
    A4: { v: 'plain', fill: 0xff0000 }, B4: { v: 'plain2', fill: 0xff0000 }, // no-merge control
    H1: { v: 'ring' }, I1: { v: '' }, G1: { v: 'left of the ring merge' }, // merge #3, adjacent to G1
  }, merges: ['A1:C1', 'E1:F2', 'H1:I1'] },
];

const MEMBERS = { 'A1:C1': ['A1', 'B1', 'C1'], 'E1:F2': ['E1', 'F1', 'E2', 'F2'] };
const AXES = {
  align: [{ align: { horizontal: 'center', vertical: 'center' } }, (c, a) => assert.ok(c.align && c.align.h !== undefined && c.align.v !== undefined, `${a} align`)],
  wrap: [{ wrap: true }, (c, a) => assert.equal(c.wrap, true, `${a} wrap`)],
  numberFormat: [{ numberFormat: '0.00%' }, (c, a) => assert.equal(c.nf, '0.00%', `${a} nf`)],
  fillNone: [{ fill: 'none' }, (c, a) => assert.equal(c.fill, 'S:No Fill', `${a} fill`)],
  outline: [{ border: { mode: 'outline', style: 'thin', color: '#000000' } }, (c, a) => assert.ok(c.border && Object.keys(c.border).length > 0, `${a} border`)],
};

const fmt = (range, format, sheet = 'Data') => ({ id: 'xb10', steps: [{ intent: 'format_range', sheet, range, format }] });

async function run(store, t, script = {}) {
  let n = 0;
  const bodies = [];
  const driver = createFakeXlsxDriver(store, script);
  const dispatch = driver.dispatch.bind(driver);
  driver.dispatch = async (h, c) => { bodies.push(c && c.body); return dispatch(h, c); };
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, bodies, m: result.metrics };
}

const sheet = store => store.durable.find(s => s.name === 'Data');
const cell = (store, a) => sheet(store).cells[a];
const evidenceOf = result => result.execution.appliedOperations[0].evidence;

for (const [axis, [format, holdsOn]] of Object.entries(AXES)) {
  for (const [range, members] of Object.entries(MEMBERS)) {
    test(`X-B10: ${axis} on the WHOLE merge ${range} -- APPLIED, EVERY member holds it, the merge survives; retry SATISFIED 0 commands`, async () => {
      const store = createStore(book());
      const { result, m } = await run(store, fmt(range, format));
      assert.equal(result.ok, true, JSON.stringify(evidenceOf(result)));
      assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
      for (const a of members) holdsOn(cell(store, a), a);
      assert.ok(sheet(store).merges.includes(range), 'merge preserved');
      assert.equal(cell(store, members[0]).v, range === 'A1:C1' ? 'Title' : 'Box', 'anchor value survives');
      const again = await run(store, fmt(range, format));
      assert.equal(again.result.ok, true);
      assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
      assert.equal(again.m.commands, 0, 'retry writes nothing');
      assert.equal(m.writes, 1);
    });
  }

  test(`X-B10: ${axis} on a range mixing a whole merge with a plain cell -- merge members AND the plain cell all written`, async () => {
    const mixed = createStore([{ name: 'Data', cells: {
      A1: { v: 'Title', fill: 0xff0000 }, B1: { v: '', fill: 0xff0000 }, C1: { v: '', fill: 0xff0000 }, D1: { v: 'plain', fill: 0xff0000 },
    }, merges: ['A1:C1'] }]); // no other merge near the ring, so the outline ring check is not what is under test
    const r = await run(mixed, fmt('A1:D1', format));
    assert.equal(r.result.ok, true, JSON.stringify(evidenceOf(r.result)));
    for (const a of ['A1', 'B1', 'C1', 'D1']) holdsOn(cell(mixed, a), a);
    assert.ok(sheet(mixed).merges.includes('A1:C1'));
  });

  test(`X-B10: ${axis} on a range CROSSING a merge (a part of it) is refused BEFORE any write, named`, async () => {
    for (const range of ['A1', 'A1:B1', 'B1:C1', 'E1:F1', 'E1']) {
      const store = createStore(book());
      const was = JSON.stringify(sheet(store));
      const { result, m } = await run(store, fmt(range, format));
      assert.equal(result.ok, false, range);
      assert.equal(evidenceOf(result).evidence.reason, 'xlsx-format-merge-crosses-range-refused', range);
      assert.deepEqual([m.writes, m.commands], [0, 0], range);
      assert.equal(JSON.stringify(sheet(store)), was, `${range}: nothing written`);
    }
  });

  test(`X-B10 control: ${axis} with NO merge at all -- unchanged, every cell written and verified`, async () => {
    const store = createStore(book());
    const { result } = await run(store, fmt('A4:B4', format));
    assert.equal(result.ok, true, JSON.stringify(evidenceOf(result)));
    for (const a of ['A4', 'B4']) holdsOn(cell(store, a), a);
  });
}

test("X-B10: fill 'none' over a merge is written as 'No Fill', NEVER through ClearFormats (which dissolves the merge LIVE); a plain range keeps the X-F08 composite", async () => {
  const merged = await run(createStore(book()), fmt('A1:C1', { fill: 'none' }));
  assert.equal(merged.result.ok, true);
  assert.equal(merged.bodies.length, 1);
  assert.match(merged.bodies[0], /SetFillColor\('No Fill'\)/);
  assert.doesNotMatch(merged.bodies[0], /ClearFormats/);
  const plain = await run(createStore(book()), fmt('A4:B4', { fill: 'none' }));
  assert.equal(plain.result.ok, true);
  assert.match(plain.bodies[0], /ClearFormats/);
  assert.doesNotMatch(plain.bodies[0], /No Fill/);
});

test('X-B10 RED mutant: an engine whose fill clear DISSOLVES the merge (the LIVE ClearFormats finding) is caught by L1 -- the task fails, not passes', async () => {
  const store = createStore(book());
  const { result } = await run(store, fmt('A1:C1', { fill: 'none' }), { fault: 'fillclear-merge-destroyed' });
  assert.equal(result.ok, false, 'a vanished merge must fail the task');
  assert.match(JSON.stringify(result), /mergeChanged/);
});

test('X-B10 RED mutant: a write that stops at the anchor (partial) is caught -- every member must hold the axis', async () => {
  const store = createStore(book());
  const { result } = await run(store, fmt('A1:C1', { wrap: true }), { fault: 'partial' });
  assert.equal(result.ok, false);
});

test('X-B10: the axes the probe did NOT measure still refuse merge membership, named', async () => {
  for (const format of [
    { align: { rotation: 45 } },
    { align: { horizontal: 'center', indent: 1 } },
    { border: { mode: 'inside', style: 'thin', color: '#000000' } },
    { border: { mode: 'all', style: 'thin', color: '#000000' } },
    // border remove over a merge was released by L-16/3 (C5: UnMerge -> composite -> Merge) -- see xlsx-l16-3-c5-remove-merge.test.cjs
    { fill: '#3A7BC8', wrap: true },
    { wrap: true, numberFormat: '0.00' , font: { bold: true } },
  ]) {
    const store = createStore(book());
    const { result, m } = await run(store, fmt('A1:C1', format));
    assert.equal(result.ok, false, JSON.stringify(format));
    assert.deepEqual([m.writes, m.commands], [0, 0], JSON.stringify(format));
  }
});

test('X-B10: an outline whose RING touches a merge outside the range stays refused (the shared facing edge is not measured)', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, fmt('G1', { border: { mode: 'outline', style: 'thin', color: '#000000' } }));
  assert.equal(result.ok, false);
  assert.equal(evidenceOf(result).evidence.reason, 'xlsx-format-outline-merged-ring-refused');
  assert.deepEqual([m.writes, m.commands], [0, 0]);
});

test('X-B10 RED (pre-change shape): a mixed fill-and-align format over a merge is refused, so the exception stays per-axis', async () => {
  const store = createStore(book());
  const { result } = await run(store, fmt('A1:C1', { fill: 'none', wrap: true }));
  assert.equal(result.ok, false);
});

test('X-B10 release harness seed: a VALUE goes only into the anchor of each merge (a merge drops the others, the product refuses that: xlsx-merge-data-loss-risk)', { skip: needsPrivate('test/live/release-xb10-merged-axes-product.live.cjs') }, () => {
  const h = require('../live/release-xb10-merged-axes-product.live.cjs');
  const nonAnchor = new Set(h.mergeRanges().flatMap(r => h.cellsOfRange(r).slice(1)));
  assert.deepEqual(h.valueCells().filter(a => nonAnchor.has(a)), [], 'no seeded value on a non-anchor merge member');
  for (const r of h.mergeRanges()) assert.ok(h.valueCells().includes(h.cellsOfRange(r)[0]), `${r}: the anchor carries the value`);
  assert.ok(nonAnchor.size > 0, 'sanity: the merges do have non-anchor members to protect');
});
