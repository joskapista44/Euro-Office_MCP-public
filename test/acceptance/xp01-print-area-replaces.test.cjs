'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const names = require('../../src/adapters/xlsx/families/names.cjs');
const W = require('../../src/server/write-tools.cjs');
const { req } = require('../fakes/write-wiring.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * X-P01 (owner GO "A", owner channel; LIVE: test/live/xlsx-xp01-print-area-repoint-probe.live.cjs -- REPOINT: public AND
 * model reader MOVED, exactly ONE Print_Area, session and after two reopens, no Delete first): set_print_area may OVERWRITE an existing area,
 * as a compare-and-swap -- the step names the area it replaces (`replaces`), and the write happens only while the current area equals it.
 * Without `replaces` the existing refusal (xlsx-name-exists-with-another-reference) is unchanged.
 */

const seeded = ref => { const b = [{ name: 'P', cells: { A1: { v: 1 } } }]; if (ref) b[0].objects = { charts: [], pivots: [], names: [{ name: 'Print_Area', refersTo: ref }] }; return b; };
const task = step => ({ id: 'xp01', steps: [{ sheet: 'P', intent: 'set_print_area', ...step }] });
const areaNames = store => ((store.durable.find(x => x.name === 'P').objects || {}).names || []).filter(n => n.name === 'Print_Area').map(n => n.refersTo);
const op = result => result.execution.appliedOperations[0];

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

test('X-P01: replaces a single area -- APPLIED with NO Delete, exactly ONE Print_Area afterwards; retry SATISFIED 0 commands', async () => {
  const store = createStore(seeded('=P!$A$1:$B$2'));
  const t = task({ range: 'A4:B5', replaces: 'A1:B2' });
  const a = await run(store, t);
  assert.equal(a.result.ok, true, JSON.stringify(op(a.result).evidence));
  assert.equal(op(a.result).status, OPERATION_STATUS.APPLIED);
  assert.equal(a.m.commands, 1);
  assert.deepEqual(areaNames(store), ['=P!$A$4:$B$5']);
  assert.match(a.bodies[0], /ws\.AddDefName\("Print_Area", "P!\$A\$4:\$B\$5"\)/);
  assert.doesNotMatch(a.bodies[0], /\.Delete\(\)/, 'the measured overwrite needs no Delete first');
  const again = await run(store, t);
  assert.equal(again.result.ok, true);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.m.commands, 0);
});

test('X-P01: replaces a multi-area print area with another (order as stored)', async () => {
  const store = createStore(seeded('=P!$A$1:$B$2,P!$D$1:$E$2'));
  const a = await run(store, task({ areas: ['A10:B11', 'D10:E11'], replaces: ['A1:B2', 'D1:E2'] }));
  assert.equal(a.result.ok, true, JSON.stringify(op(a.result).evidence));
  assert.deepEqual(areaNames(store), ['=P!$A$10:$B$11,P!$D$10:$E$11']);
});

test('X-P01: the current area is NOT the one `replaces` names (the premise is stale) -- refused, named, 0 dispatch, nothing written', async () => {
  const store = createStore(seeded('=P!$G$1:$H$2'));
  const a = await run(store, task({ range: 'A4:B5', replaces: 'A1:B2' }));
  assert.equal(a.driver.calls.dispatch, 0);
  assert.equal(op(a.result).evidence.evidence.reason, 'xlsx-print-area-replaces-mismatch');
  assert.deepEqual(areaNames(store), ['=P!$G$1:$H$2']);
});

test('X-P01: `replaces` names an area that is not there -- refused, named, nothing is created', async () => {
  const store = createStore(seeded(null));
  const a = await run(store, task({ range: 'A4:B5', replaces: 'A1:B2' }));
  assert.equal(a.driver.calls.dispatch, 0);
  assert.equal(op(a.result).evidence.evidence.reason, 'xlsx-print-area-replaces-mismatch');
  assert.deepEqual(areaNames(store), []);
});

test('X-P01: WITHOUT `replaces` an existing area with another reference is still refused (no silent re-point) -- unchanged', async () => {
  const store = createStore(seeded('=P!$A$1:$B$2'));
  const a = await run(store, task({ range: 'A4:B5' }));
  assert.equal(a.driver.calls.dispatch, 0);
  assert.equal(op(a.result).evidence.evidence.reason, 'xlsx-name-exists-with-another-reference');
});

test('X-P01: a set onto a sheet with NO area and no `replaces` is unchanged (plain APPLY)', async () => {
  const store = createStore(seeded(null));
  const a = await run(store, task({ range: 'A4:B5' }));
  assert.equal(a.result.ok, true);
  assert.deepEqual(areaNames(store), ['=P!$A$4:$B$5']);
});

test('X-P01 RED mutant: an engine that KEEPS the old area and adds the new one (a duplicate) fails L1 -- not APPLIED', async () => {
  const store = createStore(seeded('=P!$A$1:$B$2'));
  const a = await run(store, task({ range: 'A4:B5', replaces: 'A1:B2' }), { fault: 'name-repoint-duplicates' });
  assert.equal(a.result.ok, false);
  assert.equal(a.result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(a.result.execution.primitiveVerification[0].evidence.evidence.now.length, 2, 'the live read really showed TWO Print_Area names');
});

test('X-P01: planning refuses a malformed or misplaced `replaces`', () => {
  const bad = [
    task({ range: 'A4:B5', replaces: 'A1:B2:C3' }),
    task({ range: 'A4:B5', replaces: [] }),
    task({ range: 'A4:B5', replaces: ['A1:C3', 'B2:D4'] }), // overlapping
    task({ range: 'A4:B5', replaces: 7 }),
    { id: 'x', steps: [{ sheet: 'P', intent: 'set_print_titles', rows: '1:2', replaces: 'A1:B2' }] }, // fields of a different intent
    { id: 'x', steps: [{ sheet: 'P', intent: 'clear_print_area', replaces: 'A1:B2' }] },
  ];
  for (const t of bad) assert.equal(planOperations(t).ok, false, JSON.stringify(t));
  assert.equal(planOperations(task({ range: 'A4:B5', replaces: 'A1:B2' })).ok, true);
  assert.equal(planOperations(task({ areas: ['A4:B5', 'D4:E5'], replaces: ['A1:B2'] })).ok, true);
});

test('X-P01: the product WRITE tool accepts `replaces` as one range or an array and refuses a malformed one before any dispatch', () => {
  const step = extra => ({ intent: 'set_print_area', id: 's', sheet: 'P', range: 'A4:B5', ...extra });
  const validate = s => W.validateWriteRequest(req([s]));
  assert.equal(validate(step({ replaces: 'A1:B2' })).ok, true);
  assert.equal(validate(step({ replaces: ['A1:B2', 'D1:E2'] })).ok, true);
  for (const replaces of ['a1:b2', 'A1:B2:C3', [], 7, ['A1:B2', 'nope']]) assert.equal(validate(step({ replaces })).ok, false, JSON.stringify(replaces));
  assert.equal(validate({ intent: 'set_print_titles', id: 's', sheet: 'P', rows: '1:2', replaces: 'A1:B2' }).ok, false);
});

test('X-P01: lower() is unchanged by `replaces` -- the same measured AddDefName call', () => {
  const body = t => names.lower(planOperations(t).operations[0]).body;
  assert.equal(body(task({ range: 'A4:B5', replaces: 'A1:B2' })), body(task({ range: 'A4:B5' })));
});
