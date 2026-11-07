'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const names = require('../../src/adapters/xlsx/families/names.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * GH #362 2nd unit: set_defined_name / delete_defined_name must refuse PRE-MUTATION when the identifier collides
 * with an already-live HIDDEN name -- scopedHits/otherScopes (names.cjs) only ever saw the driver's visible-name
 * projection (p.definedNames), so a same-identifier hidden name reached dispatch as a plain create/delete. LIVE
 * measured outcome: DISPATCHED_OUTCOME_UNKNOWN (GH #362 comment 6646047708), never a known APPLY.
 * Base: origin/main 3533b01f.
 */

const book = () => [{ name: 'Data', cells: { A1: { v: 7 } } }, { name: 'Other', cells: { A1: { v: 'keep' } } }];
const set = (name, refersTo, extra = {}) => ({ id: 'gh362-u2', steps: [{ intent: 'set_defined_name', name, refersTo, ...extra }] });
const del = (name, extra = {}) => ({ id: 'gh362-u2', steps: [{ intent: 'delete_defined_name', name, ...extra }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const evidence = result => result.execution.appliedOperations[0].evidence.evidence;

test('set: refused PRE-MUTATION when the identifier collides with a workbook-level hidden name (0 dispatch)', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, set('Secret', '=Data!$A$1'), { hiddenNames: [{ name: 'Secret', scopeSheet: null }] });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.deepEqual(evidence(result), { ok: false, reason: 'xlsx-defined-name-collides-hidden', name: 'Secret', scope: null });
});

test('set: case-insensitive collision (the hidden name is "secret", the step asks "SECRET")', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, set('SECRET', '=Data!$A$1'), { hiddenNames: [{ name: 'secret', scopeSheet: null }] });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(evidence(result).reason, 'xlsx-defined-name-collides-hidden');
});

test('set: refused even when the hidden name lives in ANOTHER scope than the step (same identifier, any scope)', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, set('Secret', '=Data!$A$1', { scope: 'Data' }), { hiddenNames: [{ name: 'Secret', scopeSheet: 'Other' }] });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.deepEqual(evidence(result), { ok: false, reason: 'xlsx-defined-name-collides-hidden', name: 'Secret', scope: 'Other' });
});

test('delete: refused PRE-MUTATION the same way (the hidden name is not among the visible names the delete would act on)', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, del('Secret'), { hiddenNames: [{ name: 'Secret', scopeSheet: null }] });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.deepEqual(evidence(result), { ok: false, reason: 'xlsx-defined-name-collides-hidden', name: 'Secret', scope: null });
});

test('an unobservable hidden-name axis refuses closed (never treated as "no hidden names")', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, set('Total', '=Data!$A$1'), { hiddenNamesFault: 'UNOBSERVABLE:PATCH_VERSION_MISMATCH' });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(evidence(result).reason, 'xlsx-hidden-names-unobservable');
});

test('print-area intents are untouched by the hidden-name check (RESERVED already refuses Print_Area as a plain name)', () => {
  const p = planOperations({ id: 'x', steps: [{ intent: 'set_print_area', sheet: 'Data', range: 'A1:A1' }] });
  assert.equal(p.ok, true);
  const op = p.operations[0];
  assert.equal(op.name, 'Print_Area');
  // the hidden-name gate is bypassed for PRINT intents (names.cjs PRINT map) -- a hidden name named
  // 'Print_Area' cannot occur anyway: _xlnm-prefixed entries never enter p.hiddenDefinedNames (driver filter).
  const r = names.classify({ format: 'xlsx', complete: true, definedNames: [], hiddenDefinedNames: [{ name: 'Print_Area', scopeSheet: null }], dependencies: { complete: true, broken: [] }, sheets: [{ name: 'Data' }] }, op);
  assert.equal(r.state, 'APPLY');
});

test('no hidden-name fixture: visible-name behavior stays byte-identical (SATISFIED / APPLY as before)', async () => {
  const store = createStore(book());
  const a = await run(store, set('Total', '=Data!$A$1'));
  assert.equal(a.result.ok, true);
  const again = await run(store, set('Total', '=Data!$A$1'));
  assert.equal(again.m.commands, 0, 'SATISFIED on retry, no re-dispatch');
});

test('pure: classify() refuses PRE-MUTATION with the named reason, {name, scope} evidence, for workbook and sheet scope', () => {
  const base = { format: 'xlsx', complete: true, definedNames: [], dependencies: { complete: true, broken: [] }, sheets: [{ name: 'S1' }] };
  const wb = names.classify({ ...base, hiddenDefinedNames: [{ name: 'X', scopeSheet: null }] }, { intent: 'set_defined_name', name: 'X', refersTo: '=S1!$A$1', scope: null });
  assert.deepEqual(wb, { state: 'FAIL_CLOSED', evidence: { ok: false, reason: 'xlsx-defined-name-collides-hidden', name: 'X', scope: null } });
  const sh = names.classify({ ...base, hiddenDefinedNames: [{ name: 'X', scopeSheet: 'S1' }] }, { intent: 'delete_defined_name', name: 'X', scope: 'S1' });
  assert.deepEqual(sh, { state: 'FAIL_CLOSED', evidence: { ok: false, reason: 'xlsx-defined-name-collides-hidden', name: 'X', scope: 'S1' } });
});

test('pure: an array (even empty) hidden-name axis measures "no collision", distinct from "unobservable"', () => {
  const base = { format: 'xlsx', complete: true, definedNames: [], dependencies: { complete: true, broken: [] }, sheets: [{ name: 'S1' }] };
  const op = { intent: 'set_defined_name', name: 'Fresh', refersTo: '=S1!$A$1', scope: null };
  assert.equal(names.classify({ ...base, hiddenDefinedNames: [] }, op).state, 'APPLY');
  assert.deepEqual(names.classify({ ...base, hiddenDefinedNames: 'UNOBSERVABLE:bridge-names-threw' }, op).evidence, { ok: false, reason: 'xlsx-hidden-names-unobservable' });
  assert.deepEqual(names.classify({ ...base, hiddenDefinedNames: null }, op).evidence, { ok: false, reason: 'xlsx-hidden-names-unobservable' });
});
