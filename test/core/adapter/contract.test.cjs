'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { validateAdapter, REQUIRED, FORBIDDEN } = require('../../../src/core/adapter/contract.cjs');
const { createFakeAdapter } = require('../../fakes/fake-adapter.cjs');

test('the fake adapter satisfies the contract', () => {
  const v = validateAdapter(createFakeAdapter());
  assert.deepEqual(v.problems, []);
  assert.equal(v.ok, true);
});

test('every required member is actually required (each one measured separately)', () => {
  for (const name of REQUIRED) {
    const a = createFakeAdapter();
    delete a[name];
    const v = validateAdapter(a);
    assert.equal(v.ok, false, `removing ${name} must fail validation`);
  }
});

test('an adapter exposing any lifecycle member is refused', () => {
  for (const name of FORBIDDEN) {
    const a = createFakeAdapter();
    a[name] = () => {};
    const v = validateAdapter(a);
    assert.equal(v.ok, false, `an adapter exposing ${name} would be a second lifecycle`);
    assert.ok(v.problems.some(p => p.includes(name)));
  }
});

test('optional members may be null — XLSX returns null for enrichment', () => {
  const a = createFakeAdapter();
  a.enrichObservation = null;
  a.verifyPreservation = null;
  a.terminalContract = null;
  assert.equal(validateAdapter(a).ok, true);
});

test('a non-adapter fails closed rather than throwing', () => {
  for (const bad of [null, undefined, 42, 'adapter']) {
    assert.equal(validateAdapter(bad).ok, false);
  }
});
