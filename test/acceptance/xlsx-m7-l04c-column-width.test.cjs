'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { installBridge, withPatch, BRIDGE_RUNTIME_ID } = require('../fakes/ds-bridge-runtime.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 X-L04c (OD-2a + D-W, owner GO owner channel, 2026-10-05; docs/M7-COLUMN-WIDTH-CONTRACT-20261005.md): a fractional
 * column width is verified against the engine's CANONICAL persisted value of the request (the build-gated pure
 * conversion), exactly -- requested 20.5 persists and reads back 20.43, and the retry is 0 write / 0 command. The fake
 * runs the engine's own chain (maxDigitWidth 7, padding 5), so its canonical values are p54's measured ones.
 */

const book = () => [{ name: 'Data', cells: { A1: { v: 'x' }, B2: { v: 5 } } }, { name: 'Other', cells: { A1: { v: 'keep' } } }];
const task = width => ({ id: 'm7-l04c', steps: [{ intent: 'layout_range', sheet: 'Data', columns: 'B:C', width }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}
const cols = store => store.durable.find(s => s.name === 'Data').layout.cols;
const l1 = result => result.execution.primitiveVerification[0];

for (const [requested, canonical] of [[20.5, 20.43], [0.25, 0.5800000000000001], [10.25, 10.29], [1.5, 1.43], [33.33, 33.29], [20, 20]]) {
  test(`width ${requested}: APPLIED, persisted = canonical ${canonical} (never claims ${requested}), reopened retry 0 / 0`, async () => {
    const store = createStore(book());
    const first = await run(store, task(requested));
    assert.equal(first.result.ok, true, JSON.stringify(l1(first.result)));
    assert.deepEqual([first.m.writes, first.m.commands], [1, 1]);
    assert.deepEqual([cols(store).B.width, cols(store).C.width], [canonical, canonical]);
    assert.deepEqual([l1(first.result).evidence.evidence.requested, l1(first.result).evidence.evidence.canonical], [requested, canonical]);
    const retry = await run(store, task(requested));
    assert.equal(retry.result.ok, true);
    assert.deepEqual([retry.m.writes, retry.m.commands], [0, 0], 'the canonical value already holds: no re-apply for the comparison');
    assert.deepEqual([cols(store).B.width, cols(store).C.width], [canonical, canonical], 'no retry drift');
  });
}

// 0.25 -> 0.5800000000000001: the engine's own float result of its chain (S33 LIVE read the same value); GetColumnWidth
// runs the identical chain, so the comparison is exact -- the p54 report printed it rounded as 0.58.

test('the default width 8.43 already IS its canonical value: SATISFIED at once, 0 / 0 (no write for the comparison)', async () => {
  const { result, m } = await run(createStore(book()), task(8.43));
  assert.equal(result.ok, true);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
});

test('RED: the engine stores the REQUESTED number instead of its canonical value -> L1 FAIL (exact, no tolerance)', async () => {
  const { result } = await run(createStore(book()), task(20.5), { fault: 'width-stores-requested' });
  assert.equal(result.ok, false);
  assert.equal(l1(result).verdict, 'FAIL');
});

test('fail-closed: an unverified build or a throwing conversion refuses the width before the write, 0 / 0', async () => {
  for (const script of [{ colWidthCanonBuild: false }, { colWidthCanonThrows: true }]) {
    const { result, m } = await run(createStore(book()), task(20.5), script);
    assert.equal(result.ok, false, JSON.stringify(script));
    assert.deepEqual([m.writes, m.commands], [0, 0]);
    assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-layout-width-canonical-unobservable');
  }
});

// B17 cutover (internal tracker point 8; LIVE parity #ref): the canonical width comes through the DS Bridge
// sheet.columnWidth.canonicalize/v1 -- the REAL served block runs on the stand-in (test/fakes/ds-bridge-runtime.cjs).
test('the read: the canonical width only through the DS Bridge (pure arithmetic, no write), keyed by the requested value, absent without a width ask', () => {
  const ask = { layout: [{ sheet: 'Data', maxCol: 3, maxRow: 0, canon: [20.5, 15] }] };
  const q = readQuery([], { ...ask });
  assert.doesNotMatch(q, /charCountToModelColWidth|modelColWidthToColWidth|colWidthToCharCount/, 'the direct conversion read is gone');
  assert.match(q, /sheet\.columnWidth\.canonicalize\/v1/, 'positive control: the bridge read is in the query');
  const col = { GetColumnWidth: () => 8.43, GetHidden: () => false };
  const ws = { GetName: () => 'Data', GetIndex: () => 0, GetUsedRange: () => null, GetRange: () => ({ GetEntireColumn: () => col, GetEntireRow: () => ({ GetRowHeight: () => 15, GetHidden: () => false }) }),
    worksheet: { charCountToModelColWidth: x => x * 7, modelColWidthToColWidth: x => x + 1, colWidthToCharCount: x => Math.round(x / 7 * 100) / 100 } };
  const apiOf = bridge => { const a = { GetSheets: () => [ws], GetSheet: n => (n === 'Data' ? ws : null), WorksheetFunction: { COUNT: () => 0 } }; if (bridge) installBridge(a); return a; };
  const run = (a, o = ask) => new Function('Api', readQuery([], { ...o }))(a).layout.Data;
  assert.deepEqual(run(apiOf(true)).canon, { 20.5: 20.64, 15: 15.14 }, 'the chain result per requested width, in the bridge');
  assert.equal(run(apiOf(false)).canon, 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', 'no bridge -> never a direct read');
  assert.equal(run(withPatch(apiOf(true))).canon, 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', 'an unaccepted served runtime');
  assert.equal('canon' in run(apiOf(true), { layout: [{ sheet: 'Data', maxCol: 3, maxRow: 0 }] }), false, 'no width asked: the plain { cols, rows }');
});
