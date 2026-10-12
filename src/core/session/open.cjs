'use strict';

const { SessionHandle } = require('./handle.cjs');
const { bindDriver } = require('./internal.cjs');

/**
 * openSession / awaitReady — M0 section 4.
 *
 * AUTHORITY: none. Opening an editor proves nothing about the document, so neither function
 * returns anything a verifier may treat as evidence.
 *
 * The transport is injected as `driver`. M0 specifies the contract, not the wire: the DOCX and
 * XLSX sources reach the editor differently, and M1 has no real editor at all. Injection is what
 * lets the core lifecycle be proven against a fake session with zero format or transport knowledge.
 */

async function openSession({ driver, url, user, pass, fileId, adapterKind, metrics }) {
  if (!driver) throw new Error('openSession requires a session driver');
  metrics.clock.mark('openEntry');

  const opened = await driver.open({ url, user, pass, fileId });
  if (!opened || opened.ok !== true) {
    // F1: the open never produced a handle, so there is nothing to close and no barrier to run.
    const error = new Error(`session open failed: ${(opened && opened.detail) || 'unknown'}`);
    error.openFailure = opened || { ok: false, detail: 'unknown' };
    throw error;
  }

  const handle = new SessionHandle({
    id: opened.id,
    fileId,
    adapterKind,
    frame: opened.frame || null,
    apiWhere: opened.apiWhere || null,
    metrics,
  });
  bindDriver(handle, driver);

  metrics.clock.mark('handleReturned').span('sessionOpenMs', 'openEntry', 'handleReturned');
  handle.openedMs = metrics.clock.spans().sessionOpenMs;
  return handle;
}

/**
 * Readiness is a SEMANTIC probe supplied by the adapter — the only adapter call allowed before
 * READY (M0 section 4). It polls a condition; it never sleeps a fixed interval, because a fixed
 * delay is not a correctness mechanism (PROJECT invariant 6, owner requirement 17).
 *
 * EXCEPTION POLICY. The probe is an ADAPTER boundary, and an adapter may throw rather than report
 * not-ready. A raw exception escaping here would leave the caller with no FinalTaskResult, so the
 * probe call is wrapped at its own boundary and normalized into an ordinary not-ready result that
 * carries the throw as evidence. The wrap is around the probe call ONLY — not around the loop —
 * so a bug in this function still propagates.
 */
async function awaitReady(handle, adapter, { timeoutMs = 15000, pollMs = 100, wait } = {}) {
  const metrics = handle.metrics;
  const waitFor = wait || defaultWait;
  const deadlineAt = attempts => attempts * pollMs >= timeoutMs;

  let attempts = 0;
  let last = null;
  // Condition-based wait: the loop exits on the observed predicate, and the elapsed budget is a
  // timeout bound, not a delay that the happy path pays.
  for (;;) {
    try {
      last = await adapter.readinessProbe(handle);
    } catch (err) {
      // A thrown probe is terminal, not retryable: the adapter contract was broken, and polling a
      // broken boundary would only repeat the same failure until the timeout.
      return {
        ok: false,
        readyMs: null,
        readiness: {
          ok: false,
          threw: true,
          detail: err.message,
          errorName: err.name,
          attempt: attempts + 1,
        },
      };
    }
    if (last && last.ok === true) {
      metrics.clock.mark('ready').span('readyMs', 'handleReturned', 'ready');
      return { ok: true, readyMs: metrics.clock.spans().readyMs, readiness: last };
    }
    attempts += 1;
    if (deadlineAt(attempts)) {
      // F2: READY never reached -> close without any write, no barrier.
      return { ok: false, readyMs: null, readiness: last || { ok: false, detail: 'no probe result' } };
    }
    await waitFor(pollMs);
  }
}

/** Yields between polls. Injectable so tests neither sleep nor depend on wall-clock timing. */
function defaultWait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * closeSession — M0 section 4. A close failure never upgrades a task result, and it is never
 * swallowed: the outcome is returned so finalizeTask can report cleanupOutcome explicitly.
 */
async function closeSession(handle) {
  const { driverOf } = require('./internal.cjs');
  try {
    const closed = await driverOf(handle).close(handle);
    // M7 save contract point 3: the driver's forcesave confirmation RE-EVALUATED at close (null: none to report)
    const confirmation = (closed && closed.confirmation) || null;
    if (closed && closed.ok === true) {
      return { ok: true, outcome: 'closed', detail: closed.detail || null, confirmation };
    }
    return { ok: false, outcome: 'close-failed', detail: (closed && closed.detail) || 'unknown', confirmation };
  } catch (err) {
    return { ok: false, outcome: 'close-threw', detail: err.message };
  }
}

module.exports = { openSession, awaitReady, closeSession, defaultWait };
