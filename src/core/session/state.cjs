'use strict';

/**
 * Lifecycle state machine — M0 section 6.
 *
 * The editor-session states come from the M0 diagram verbatim. FINALIZED is the task-level
 * terminal state that follows CLOSED: the M0 pseudo-code in 4.1 returns finalizeTask() after the
 * finally block, so the FinalTaskResult exists strictly after the editor is closed. It is the same
 * lifecycle, named at its last step — not a second one.
 */

const STATES = Object.freeze({
  CREATED: 'CREATED',
  OPENING: 'OPENING',
  READY: 'READY',
  OBSERVING: 'OBSERVING',
  EXECUTING: 'EXECUTING',
  VERIFYING: 'VERIFYING',
  PERSISTING: 'PERSISTING',
  CLOSING: 'CLOSING',
  CLOSED: 'CLOSED',
  FINALIZED: 'FINALIZED',
});

/**
 * Legal transitions. Every state except FINALIZED may reach CLOSING, because a failure at any
 * point still closes the editor in the lifecycle owner's finally block (M0 F1-F8).
 */
const TRANSITIONS = Object.freeze({
  CREATED: ['OPENING', 'CLOSING'],
  OPENING: ['READY', 'CLOSING'],
  READY: ['OBSERVING', 'CLOSING'],
  OBSERVING: ['EXECUTING', 'CLOSING'],
  EXECUTING: ['VERIFYING', 'CLOSING'],
  VERIFYING: ['PERSISTING', 'CLOSING'],
  PERSISTING: ['CLOSING'],
  CLOSING: ['CLOSED'],
  CLOSED: ['FINALIZED'],
  FINALIZED: [],
});

class IllegalTransitionError extends Error {
  constructor(from, to) {
    super(`illegal lifecycle transition ${from} -> ${to}`);
    this.name = 'IllegalTransitionError';
    this.from = from;
    this.to = to;
  }
}

/**
 * A recording state machine. The trace it keeps is what every acceptance case asserts against;
 * M1 requires an explicit expected lifecycle trace per test (owner requirement 20).
 */
class LifecycleState {
  constructor() {
    this._state = STATES.CREATED;
    this._trace = [STATES.CREATED];
  }

  get state() {
    return this._state;
  }

  /** The visited states, in order. Never mutated in place by callers. */
  get trace() {
    return this._trace.slice();
  }

  canTransition(to) {
    return (TRANSITIONS[this._state] || []).includes(to);
  }

  /** Fail-closed: an illegal transition throws rather than being coerced into a legal one. */
  transition(to) {
    if (!Object.prototype.hasOwnProperty.call(TRANSITIONS, to)) {
      throw new IllegalTransitionError(this._state, to);
    }
    if (!this.canTransition(to)) {
      throw new IllegalTransitionError(this._state, to);
    }
    this._state = to;
    this._trace.push(to);
    return this;
  }

  /** True once the editor is closed, whether or not the task succeeded. */
  isClosed() {
    return this._state === STATES.CLOSED || this._state === STATES.FINALIZED;
  }
}

module.exports = { STATES, TRANSITIONS, LifecycleState, IllegalTransitionError };
