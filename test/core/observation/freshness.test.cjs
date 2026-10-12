'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { FRESHNESS, classify, isFresh, requireFresh, sessionIdentity } = require('../../../src/core/observation/freshness.cjs');
const { bumpGeneration, currentGeneration } = require('../../../src/core/mutation-epoch.cjs');

const handle = id => ({ id });
const obs = (handle, generation) => ({ freshness: { sessionId: handle.id, sessionIdentity: sessionIdentity(handle), generation } });

test('a session starts at generation 0 and only a bump moves it', () => {
  const h = handle('s1');
  assert.equal(currentGeneration(h), 0);
  assert.equal(bumpGeneration(h), 1);
  assert.equal(currentGeneration(h), 1);
});

test('same session, same generation -> FRESH', () => {
  const h = handle('s2');
  assert.equal(classify(obs(h, 0), h).state, FRESHNESS.FRESH);
  assert.equal(isFresh(obs(h, 0), h), true);
});

test('a mutation makes every earlier observation STALE, and says how many', () => {
  const h = handle('s3');
  const before = obs(h, 0);
  bumpGeneration(h);
  bumpGeneration(h);

  const v = classify(before, h);
  assert.equal(v.state, FRESHNESS.STALE);
  assert.equal(v.observedAtGeneration, 0);
  assert.equal(v.currentGeneration, 2);
  assert.equal(v.mutationsSince, 2);
});

test('a different session is CROSS_SESSION even at the same generation', () => {
  const a = handle('sA');
  const b = handle('sB');
  const v = classify(obs(a, 0), b);

  assert.equal(v.state, FRESHNESS.CROSS_SESSION);
  assert.equal(v.observedIn, 'sA');
  assert.equal(v.usedIn, 'sB');
  // The distinction matters: a generation-only check would have called this FRESH.
  assert.equal(currentGeneration(a), currentGeneration(b));
});


test('same caller-visible id on two handles is still CROSS_SESSION', () => {
  const a = handle('duplicate-id');
  const b = handle('duplicate-id');
  const o = obs(a, 0);
  const v = classify(o, b);
  assert.equal(v.state, FRESHNESS.CROSS_SESSION);
  assert.equal(currentGeneration(a), currentGeneration(b));
});

test('missing freshness metadata is MALFORMED, not assumed fresh', () => {
  const h = handle('s4');
  for (const bad of [null, undefined, {}, { freshness: null }, 'observation']) {
    assert.equal(classify(bad, h).state, FRESHNESS.MALFORMED);
  }
});

test('requireFresh returns a refusal rather than throwing — staleness is EXPECTED', () => {
  const h = handle('s5');
  const o = obs(h, 0);
  assert.equal(requireFresh(o, h, 'targeting').ok, true);

  bumpGeneration(h);
  const gate = requireFresh(o, h, 'targeting');
  assert.equal(gate.ok, false);
  assert.equal(gate.outcome, 'observation-stale');
  assert.equal(gate.purpose, 'targeting');
  assert.ok(gate.reason.includes('generation 0 -> 1'));
});

test('the outcome name encodes the reason, so a caller can branch on it', () => {
  const h = handle('s6');
  const other = handle('other');
  assert.equal(requireFresh(obs(other, 0), h).outcome, 'observation-cross-session');
  assert.equal(requireFresh({}, h).outcome, 'observation-malformed');
});

test('freshness is NOT time-based — no clock advances it', () => {
  const h = handle('s7');
  const o = obs(h, 0);
  const start = Date.now();
  while (Date.now() - start < 5) { /* let wall-clock time pass */ }
  assert.equal(isFresh(o, h), true, 'time alone never invalidates an observation');

  bumpGeneration(h);
  assert.equal(isFresh(o, h), false, 'only a mutation does');
});
