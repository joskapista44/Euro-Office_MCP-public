'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { STATES, LifecycleState, IllegalTransitionError } = require('../../../src/core/session/state.cjs');

test('the happy path walks the M0 section 6 diagram', () => {
  const m = new LifecycleState();
  for (const s of ['OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING', 'PERSISTING', 'CLOSING', 'CLOSED', 'FINALIZED']) {
    m.transition(s);
  }
  assert.equal(m.state, STATES.FINALIZED);
  assert.equal(m.trace.length, 10);
});

test('the no-write path skips PERSISTING', () => {
  const m = new LifecycleState();
  for (const s of ['OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING', 'CLOSING', 'CLOSED']) m.transition(s);
  assert.equal(m.trace.includes('PERSISTING'), false);
});

test('every pre-close state can reach CLOSING — F1..F8 all close the editor', () => {
  for (const from of ['CREATED', 'OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING', 'PERSISTING']) {
    const m = new LifecycleState();
    const route = { CREATED: [], OPENING: ['OPENING'], READY: ['OPENING', 'READY'],
      OBSERVING: ['OPENING', 'READY', 'OBSERVING'],
      EXECUTING: ['OPENING', 'READY', 'OBSERVING', 'EXECUTING'],
      VERIFYING: ['OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING'],
      PERSISTING: ['OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING', 'PERSISTING'] }[from];
    for (const s of route) m.transition(s);
    assert.ok(m.canTransition('CLOSING'), `${from} must be able to close`);
  }
});

test('an illegal transition fails closed rather than being coerced', () => {
  const m = new LifecycleState();
  assert.throws(() => m.transition('PERSISTING'), IllegalTransitionError);
  assert.throws(() => m.transition('NOT_A_STATE'), IllegalTransitionError);
  assert.equal(m.state, STATES.CREATED, 'a refused transition must not move the state');
});

test('FINALIZED is terminal, and the trace is a copy', () => {
  const m = new LifecycleState();
  for (const s of ['OPENING', 'READY', 'CLOSING', 'CLOSED', 'FINALIZED']) m.transition(s);
  assert.throws(() => m.transition('CLOSING'), IllegalTransitionError);

  const t = m.trace;
  t.push('TAMPERED');
  assert.equal(m.trace.includes('TAMPERED'), false, 'the trace must not be mutable by a caller');
});
