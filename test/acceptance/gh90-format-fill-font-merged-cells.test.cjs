'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * GH #90 (design #ref, accepted #ref -- the probe is test/live/xlsx-gh90-merge-fill-font-probe.live.cjs,
 * MEASURED R1=R2): format.fill / format.font on a range containing a merged area are no longer refused
 * outright. fill writes and verifies the WHOLE merged area (a plain range-level write already lands on every
 * member, LIVE-measured -- no per-cell filtering needed, unlike border.mode 'cells'). font writes and verifies
 * the ANCHOR only -- every other member's value is blanked by the merge itself, so its font axis is
 * UNOBSERVABLE there, not a gap. A merge that CROSSES the requested range's boundary stays refused, named
 * (xlsx-format-merge-crosses-range-refused), same reasoning GH #83 already uses for border. Every other
 * format key (numberFormat, align, wrap, border outline/inside/all/remove, fill 'none') keeps refusing any
 * merge membership, unchanged.
 */

const book = () => [
  { name: 'Data', cells: {
    B2: { v: 'Title' }, B3: { v: '' }, C2: { v: '' }, C3: { v: '' }, // merge #1 -- fill target
    G2: { v: 'Title2' }, G3: { v: '' }, H2: { v: '' }, H3: { v: '' }, // merge #2 -- font target
    D2: { v: 'plain' }, D3: { v: 'plain2' }, // plain cells, same rows as merge #1, for the mixed-range case
    K2: { v: 'x' }, L2: { v: '' }, // merge #3 -- fill 'none' control (not merge-aware)
    E2: { v: 'plain' }, F2: { v: 'plain2' }, // no-merge control
  }, merges: ['B2:C3', 'G2:H3', 'K2:L2'] },
];

const fmt = (range, format, sheet = 'Data') => ({ id: 'gh90', steps: [{ intent: 'format_range', sheet, range, format }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const cell = (store, a) => store.durable.find(s => s.name === 'Data').cells[a];
const hex = h => parseInt(h, 16);

test("GH #90: fill '#RRGGBB' on a range covering a WHOLE merge -- APPLIED, EVERY member carries the colour; retry SATISFIED", async () => {
  const store = createStore(book());
  const { result, m } = await run(store, fmt('B2:C3', { fill: '#3A7BC8' }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(m.writes, 1);
  for (const a of ['B2', 'B3', 'C2', 'C3']) assert.equal(cell(store, a).fill, hex('3A7BC8'), a);
  assert.ok(store.durable.find(s => s.name === 'Data').merges.includes('B2:C3'));

  const again = await run(store, fmt('B2:C3', { fill: '#3A7BC8' }));
  assert.equal(again.result.ok, true);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.m.commands, 0, 'retry writes nothing -- every member already holds the colour');
});

test("GH #90: fill on a range mixing a merge with plain cells -- the merge gets the colour on every member, the plain cells get it too", async () => {
  const store = createStore(book());
  const { result } = await run(store, fmt('B2:D3', { fill: '#3A7BC8' })); // D2/D3 are plain, not part of the merge
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  for (const a of ['B2', 'B3', 'C2', 'C3', 'D2', 'D3']) assert.equal(cell(store, a).fill, hex('3A7BC8'), a);
});

test('GH #90 control: fill with NO merge at all -- unchanged, every cell written and verified', async () => {
  const store = createStore(book());
  const { result } = await run(store, fmt('E2:F2', { fill: '#3A7BC8' }));
  assert.equal(result.ok, true);
  assert.equal(cell(store, 'E2').fill, hex('3A7BC8'));
  assert.equal(cell(store, 'F2').fill, hex('3A7BC8'));
});

test("GH #90: font on a range covering a WHOLE merge -- APPLIED, the anchor is written+verified, the other members stay UNOBSERVABLE (not a gap); retry SATISFIED", async () => {
  const store = createStore(book());
  const script = { emptyFontUnobservable: ['G3', 'H2', 'H3'] }; // the real engine's UNOBSERVABLE:... on a merge-blanked member, not the fake's default
  const { result } = await run(store, fmt('G2:H3', { font: { color: '#C00000', bold: true } }), script);
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(cell(store, 'G2').font.color, hex('C00000'));
  assert.equal(cell(store, 'G2').bold, true);
  assert.ok(store.durable.find(s => s.name === 'Data').merges.includes('G2:H3'));

  const again = await run(store, fmt('G2:H3', { font: { color: '#C00000', bold: true } }), script);
  assert.equal(again.result.ok, true, JSON.stringify(again.result.execution.appliedOperations[0].evidence));
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test("GH #90 control: without the emptyFontUnobservable simulation (the fake's default empty-cell font reads observable), font on the SAME merge still APPLIES -- the merge-aware skip exists for the REAL engine's unobservable member, not as a precondition for this test to pass", async () => {
  const store = createStore(book());
  const { result } = await run(store, fmt('G2:H3', { font: { color: '#C00000', bold: true } })); // no script -- members read a normal font tuple here
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
});

for (const [range, format] of [['B2', { fill: '#3A7BC8' }], ['B2:B3', { fill: '#3A7BC8' }], ['G2', { font: { bold: true } }], ['G2:G3', { font: { bold: true } }]]) {
  test(`GH #90: a merge that CROSSES the requested range (${range}) is refused before any write -- fill and font alike`, async () => {
    const store = createStore(book());
    const before = JSON.stringify(cell(store, range.slice(0, 2)));
    const { result } = await run(store, fmt(range, format));
    assert.equal(result.ok, false, range);
    assert.match(JSON.stringify(result), /xlsx-format-merge-crosses-range-refused/, range);
    assert.equal(JSON.stringify(cell(store, range.slice(0, 2))), before, `${range}: nothing written`);
  });
}

test("GH #90 control: fill 'none' on a range CROSSING a merge is STILL refused (X-B10 lifted it for a whole merge only)", async () => {
  const store = createStore(book());
  const { result } = await run(store, fmt('K2', { fill: 'none' }));
  assert.equal(result.ok, false);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-merge-crosses-range-refused');
});

test('GH #90 control: numberFormat / align / wrap on a range CROSSING a merge are STILL refused (X-B10: a whole merge is allowed, a partial one is not)', async () => {
  for (const format of [{ numberFormat: '0.00' }, { align: { horizontal: 'center' } }, { wrap: true }]) {
    const store = createStore(book());
    const { result } = await run(store, fmt('B2', format));
    assert.equal(result.ok, false, JSON.stringify(format));
    assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-format-merge-crosses-range-refused', JSON.stringify(format));
  }
});

test("GH #90 RED mutant: a fill write that leaked onto an UNRELATED cell outside the range is caught (L3, drift not tolerated)", async () => {
  const store = createStore(book());
  const { result } = await run(store, fmt('B2:C3', { fill: '#3A7BC8' }), { fault: 'format-leaks', leakCell: 'E2' });
  assert.equal(result.ok, false, 'a leaked write onto an unrelated cell must fail L3, not pass silently');
});
