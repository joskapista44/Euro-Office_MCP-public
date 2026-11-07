'use strict';

/**
 * GH #352 (H01): set_hyperlink to an INTERNAL target (`location`, no `address`). The engine shape and every rule here come
 * from the LIVE PUBLIC-FIRST probe (test/live/xlsx-gh352-internal-hyperlink-probe.live.cjs,
 * docs/evidence/gh352-internal-hyperlink-probe-live-20261030.json) and the family header (hyperlinks-write.cjs).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/hyperlinks-write.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const OTHER = { ref: 'D4', address: 'https://older.example/', location: null, tooltip: null, kind: 'external' };
const book = () => [
  { name: 'A', cells: { A1: { v: 'x' }, B2: { v: 'linked value' }, C3: { v: 7 } }, hyperlinks: [OTHER] },
  { name: 'Other Sheet', cells: { A1: { v: 'y' } } },
];
const setStep = (extra = {}) => ({ intent: 'set_hyperlink', sheet: 'A', cell: 'B2', location: 'Other Sheet!A1', ...extra });
const task = step => ({ id: 'gh352', steps: [step] });
const planOne = step => fam.plan(step, 0);

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}
const linkAt = (store, cell) => (store.durable[0].hyperlinks || []).find(l => l.ref === cell) || null;

test('plan: "Ref" / "Sheet!Ref" / "\'Sheet\'!Ref" are canonicalised to the STORED quoted form; the engine gets the plain name', () => {
  const cases = [
    ['D9', "'A'!D9", 'A!D9'],
    ['other sheet!a1', "'other sheet'!A1", 'other sheet!A1'],
    ["'Other Sheet'!A1:B2", "'Other Sheet'!A1:B2", 'Other Sheet!A1:B2'],
    ['H352!D9', "'H352'!D9", 'H352!D9'],
  ];
  for (const [location, stored, engine] of cases) {
    const p = planOne(setStep({ location }));
    assert.equal(p.ok, true, p.detail);
    assert.equal(p.operation.kind, 'internal');
    assert.equal(p.operation.address, null);
    assert.equal(p.operation.location, stored, location);
    assert.equal(p.operation.engineLocation, engine, location);
  }
});

test('plan: refused before dispatch, named -- address AND location, neither, a $ marker, an apostrophe sheet, a non-reference', () => {
  assert.match(planOne(setStep({ address: 'https://example.com/a' })).detail, /not both/);
  assert.match(planOne({ intent: 'set_hyperlink', sheet: 'A', cell: 'B2' }).detail, /an address .* or a location/);
  assert.match(planOne(setStep({ location: '$D$9' })).detail, /\$ \(absolute\)/);
  assert.match(planOne(setStep({ location: "O'Brien!A1" })).detail, /apostrophe/);
  assert.match(planOne(setStep({ location: 'A!ZZZ' })).detail, /not an A1 cell or range/);
  assert.match(planOne(setStep({ location: '!A1' })).detail, /empty sheet/);
  assert.equal(planOne(setStep({ location: '' })).ok, false);
});

test('lower: SetHyperlink(cell, "<Sheet>!<Ref>", "<Sheet>!<Ref>", tip, <own text>) -- the measured round-4 shape', () => {
  const p = planOne(setStep({ tooltip: 'tip' }));
  const cmd = fam.lower(p.operation, { sheet: 'A', cell: 'B2', text: 'linked value' });
  assert.match(cmd.body, /ws\.SetHyperlink\("B2", "Other Sheet!A1", "Other Sheet!A1", "tip", "linked value"\)/);
  assert.equal(cmd.meta.kind, 'internal');
});

test('APPLIED: cross-sheet internal link stored with the quoted location and no address, value and the other link kept; retry SATISFIED 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(setStep({ tooltip: 'go there' })));
  assert.equal(result.ok, true, JSON.stringify(result.execution));
  assert.equal(m.commands, 1);
  assert.deepEqual(linkAt(store, 'B2'), { ref: 'B2', address: null, location: "'Other Sheet'!A1", tooltip: 'go there', kind: 'internal' });
  assert.deepEqual(linkAt(store, 'D4'), OTHER);
  assert.equal(store.durable[0].cells.B2.v, 'linked value');
  const again = await run(store, task(setStep({ location: "'Other Sheet'!A1", tooltip: 'go there' })));
  assert.equal(again.m.commands, 0, 'the quoted input names the same target');
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('APPLIED: a bare reference targets the link\'s own sheet (a number cell keeps its value)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(setStep({ cell: 'C3', location: 'A1' })));
  assert.equal(result.ok, true, JSON.stringify(result.execution));
  assert.deepEqual(linkAt(store, 'C3'), { ref: 'C3', address: null, location: "'A'!A1", tooltip: null, kind: 'internal' });
  assert.equal(store.durable[0].cells.C3.v, 7);
});

test('an UNQUOTED stored location (the engine quotes only when it must -- LIVE "GH352_x!C3") satisfies the same target, 0 commands', async () => {
  const store = createStore([{ name: 'A', cells: { A1: { v: 'x' }, B2: { v: 'linked value' } }, hyperlinks: [{ ref: 'B2', address: null, location: 'A!D9', tooltip: null, kind: 'internal' }] }]);
  const { result, m } = await run(store, task(setStep({ location: 'D9' })));
  assert.equal(result.ok, true, JSON.stringify(result.execution));
  assert.equal(m.commands, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  const other = await run(store, task(setStep({ location: 'D10' })));
  assert.equal(other.m.commands, 0);
  assert.match(JSON.stringify(other.result.execution), /xlsx-hyperlink-cell-occupied/, 'a different internal target is not the same link');
});

test('a target sheet that does not exist is refused before the write, named, 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(setStep({ location: 'Nope!A1' })));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-hyperlink-target-sheet-not-found/);
  assert.equal(linkAt(store, 'B2'), null);
});

test('a cell already holding an EXTERNAL link is occupied for an internal one (no overwrite), 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(setStep({ cell: 'D4', location: 'A1' })));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-hyperlink-cell-occupied/);
});

for (const fault of ['hyperlink-wrong-address', 'hyperlink-touches-other', 'ignore', 'side-effect']) {
  test(`a fake internal write with fault ${fault} fails the verify, not passes it`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task(setStep()), { fault });
    assert.equal(m.commands, 1);
    assert.equal(result.ok, false);
  });
}

test('product surface: a location-only (internal) write passes validateWriteRequest', () => {
  const v = step => W.validateWriteRequest(req([{ id: 's0', ...step }]));
  const a = v(setStep());
  assert.equal(a.ok, true, a.reason || a.detail);
});
