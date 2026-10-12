'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const A = require('../../src/core/authority.cjs');

test('the lattice is totally ordered, weakest to strongest', () => {
  assert.deepEqual(A.ORDER, ['PLAN_ONLY', 'DISPATCH_ONLY', 'LIVE_READ', 'LIVE_VERIFY', 'PERSISTED']);
  for (let i = 1; i < A.ORDER.length; i++) {
    assert.ok(A.rank(A.ORDER[i]) > A.rank(A.ORDER[i - 1]), `${A.ORDER[i]} must outrank ${A.ORDER[i - 1]}`);
  }
});

test('combining evidence takes the WEAKEST link', () => {
  assert.equal(A.weakest('PERSISTED', 'LIVE_VERIFY', 'DISPATCH_ONLY'), 'DISPATCH_ONLY');
  assert.equal(A.weakest('LIVE_VERIFY', 'LIVE_VERIFY'), 'LIVE_VERIFY');
  assert.equal(A.weakest(), 'PLAN_ONLY', 'no evidence is the floor, not the ceiling');

  // The direction that matters: one weak step does not get carried by strong neighbours.
  assert.notEqual(A.weakest('PERSISTED', 'PLAN_ONLY'), 'PERSISTED');
});

test('every layer ceiling is enforced, each one measured separately', () => {
  const cases = [
    ['observation', 'LIVE_READ', 'LIVE_VERIFY'],
    ['enrichment', 'LIVE_READ', 'LIVE_VERIFY'],
    ['primitiveVerifier', 'LIVE_VERIFY', 'PERSISTED'],
    ['wholeTaskVerifier', 'LIVE_READ', 'LIVE_VERIFY'],
    ['preservationVerifier', 'LIVE_READ', 'LIVE_VERIFY'],
    ['dispatch', 'DISPATCH_ONLY', 'LIVE_READ'],
    ['plan', 'PLAN_ONLY', 'DISPATCH_ONLY'],
    ['persistence', 'PERSISTED', null],
  ];

  for (const [layer, allowed, forbidden] of cases) {
    assert.equal(A.assertClaim(layer, allowed), allowed, `${layer} may claim ${allowed}`);
    if (forbidden) {
      assert.throws(
        () => A.assertClaim(layer, forbidden),
        A.AuthorityViolationError,
        `${layer} must not be able to claim ${forbidden}`
      );
    }
  }
});

test('an over-claim THROWS rather than being silently clamped', () => {
  // Clamping would let a layer keep reporting an authority it never had, and the caller would
  // never learn its evidence was weaker than it believed.
  try {
    A.assertClaim('observation', 'PERSISTED');
    assert.fail('expected a throw');
  } catch (err) {
    assert.equal(err.name, 'AuthorityViolationError');
    assert.equal(err.layer, 'observation');
    assert.equal(err.claimed, 'PERSISTED');
    assert.equal(err.ceiling, 'LIVE_READ');
  }
});

test('unknown authorities and layers fail loudly', () => {
  assert.throws(() => A.rank('VERY_SURE'), /unknown authority/);
  assert.throws(() => A.assertClaim('vibes', 'LIVE_READ'), /unknown authority layer/);
  assert.equal(A.withinCeiling('vibes', 'LIVE_READ'), false, 'an unknown layer is never within a ceiling');
});

test('the persistence ceiling is the only one that reaches PERSISTED', () => {
  const reaching = Object.entries(A.CEILING).filter(([, c]) => c === 'PERSISTED').map(([l]) => l);
  assert.deepEqual(reaching, ['persistence'], 'nothing but the barrier may claim durability');
});
