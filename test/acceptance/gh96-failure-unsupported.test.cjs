'use strict';

/**
 * GH #96 (owner request via a PA): set_validation on the reporter's file refused xlsx-observation-unreadable-getters
 * with NO hint which reader failed -- the classifier's evidence carries the observation's `unsupported` list
 * ([{ field, reason }]), but failureOf() only passes scalars, so the product answer dropped it. The list is now carried
 * on failure.unsupported (classification + preflight), bounded: at most 10 entries, each reason cut to 160 characters.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');

const ev = (stage, unsupported) => ({ status: 'FAILED', evidence: { stage, evidence: { reason: 'xlsx-observation-unreadable-getters', unsupported } } });

test('GH #96: failureOf carries the observation unsupported list (classification + preflight)', () => {
  for (const stage of ['classification', 'preflight']) {
    const f = W.failureOf(ev(stage, [{ field: 'Data!C18.GetValidation', reason: 'Error: something' }]), { l1: null, l3: null });
    assert.equal(f.reason, 'xlsx-observation-unreadable-getters', stage);
    assert.deepEqual(f.unsupported, [{ field: 'Data!C18.GetValidation', reason: 'Error: something' }], stage);
  }
});

test('GH #96: the carried list is bounded (10 entries, 160-char reasons) and only scalar field/reason survive', () => {
  const many = Array.from({ length: 25 }, (_, i) => ({ field: `F${i}`, reason: 'x'.repeat(500), extra: { nested: true } }));
  const f = W.failureOf(ev('classification', many), { l1: null, l3: null });
  assert.equal(f.unsupported.length, 10);
  assert.equal(f.unsupported[0].reason.length, 160);
  assert.deepEqual(Object.keys(f.unsupported[0]).sort(), ['field', 'reason']);
  assert.equal(f.unsupportedTotal, 25);
});

test('GH #96: no unsupported key when the list is absent, empty or not an array', () => {
  for (const u of [undefined, [], 'UNOBSERVABLE:x', { field: 'a' }]) {
    const f = W.failureOf(ev('classification', u), { l1: null, l3: null });
    assert.equal('unsupported' in f, false, JSON.stringify(u));
  }
});
