'use strict';

const { driverOf } = require('../session/internal.cjs');
const { defaultWait } = require('../session/open.cjs');

/**
 * Persistence barrier — M0 section 8.2.
 *
 *   mechanism   forceSave(), then poll canSave() until false, or timeout
 *   parameters  timeoutMs 15000, pollMs 100   (identical in both sources today — kept)
 *
 * RESULT (M7 save contract, owner GO owner channel): ok means the editor reported nothing left to save -- EDITOR_CLEAN,
 * never PERSISTED. The driver's own forcesave confirmation is carried as `confirmation` (confirmed | unconfirmed, with
 * the reason); a driver that reports none (no save observer, e.g. DOCX today) is `driver-provides-no-save-confirmation`
 * -- nothing is invented for it. A persistence failure NEVER reports success (M0 F7).
 *
 * The poll is a condition-based wait bounded by a timeout, not a fixed delay standing in for
 * correctness (invariant 6, owner requirement 17): the loop exits on the observed predicate.
 *
 * EXCEPTION POLICY. A transport can THROW instead of returning a refusal, and a raw exception
 * escaping here would leave the caller with no FinalTaskResult at all — the lifecycle would close
 * the editor and then rethrow. Each driver call is therefore wrapped AT ITS OWN BOUNDARY, where
 * the contract is known, and normalized into a distinct outcome. This is deliberately not a
 * catch-all around the barrier: a programmer error inside this module still propagates.
 */

const DEFAULT_TIMEOUT_MS = 15000;
const DEFAULT_POLL_MS = 100;

async function persist(handle, { timeoutMs = DEFAULT_TIMEOUT_MS, pollMs = DEFAULT_POLL_MS, wait } = {}) {
  const metrics = handle.metrics;
  const waitFor = wait || defaultWait;
  const driver = driverOf(handle);
  const events = [];
  const startedAt = metrics.clock.now();
  let reads = 0;

  const before = await probeCanSave(driver, handle, metrics);
  reads += 1;
  events.push({ event: 'before', canSave: before.canSave, ok: before.ok, threw: before.threw === true });
  if (before.threw) {
    // The very first probe threw: nothing was saved and nothing can be polled.
    return fail('canSave-threw', {
      before,
      after: null,
      events,
      reads,
      waitMs: elapsed(metrics, startedAt),
      exception: { stage: 'canSave-initial', message: before.detail, name: before.errorName },
    });
  }

  let forced;
  try {
    forced = await driver.forceSave(handle);
  } catch (err) {
    // The transport threw. Distinct from a refusal: the editor was never asked successfully, and
    // the detail is kept as evidence rather than being flattened into a timeout.
    events.push({ event: 'forceSave', ok: false, threw: true, detail: err.message });
    return fail('forceSave-threw', {
      before,
      after: null,
      events,
      reads,
      waitMs: elapsed(metrics, startedAt),
      exception: { stage: 'forceSave', message: err.message, name: err.name },
    });
  }
  events.push({ event: 'forceSave', ok: forced && forced.ok === true, detail: (forced && forced.detail) || null });
  if (!forced || forced.ok !== true) {
    // The editor could not be asked to save at all. Fail closed; F7 downgrades the task result.
    return fail('forceSave-unavailable', { before, after: null, events, reads, waitMs: elapsed(metrics, startedAt), confirmation: confirmationOf(forced) });
  }

  let attempts = 0;
  for (;;) {
    const probe = await probeCanSave(driver, handle, metrics);
    reads += 1;
    events.push({ event: 'poll', attempt: attempts + 1, canSave: probe.canSave, ok: probe.ok, threw: probe.threw === true });

    if (probe.threw) {
      // A throw DURING polling is not a timeout: the barrier's own state is unknown from here,
      // so it is reported as its own outcome with the exception kept as evidence.
      return fail('canSave-threw', {
        before,
        after: probe,
        events,
        reads,
        waitMs: elapsed(metrics, startedAt),
        exception: { stage: 'canSave-poll', attempt: attempts + 1, message: probe.detail, name: probe.errorName },
      });
    }

    if (probe.ok === true && probe.canSave === false) {
      // The editor reports nothing left to save: EDITOR_CLEAN. Whether the server confirmed the save is the driver's
      // own paired confirmation, carried as is -- canSave never stands in for it.
      return {
        ok: true,
        outcome: 'forceSave-can-save-false',
        editorClean: true,
        confirmation: confirmationOf(forced),
        before,
        after: probe,
        events,
        reads,
        waitMs: elapsed(metrics, startedAt),
      };
    }

    attempts += 1;
    if (attempts * pollMs >= timeoutMs) {
      return fail('forceSave-timeout', {
        before,
        after: probe,
        events,
        reads,
        waitMs: elapsed(metrics, startedAt),
      });
    }
    await waitFor(pollMs);
  }
}

async function probeCanSave(driver, handle, metrics) {
  metrics.count('persistencePolls');
  let probe;
  try {
    probe = await driver.canSave(handle);
  } catch (err) {
    // Normalized at the boundary, and marked `threw` so the caller can tell a thrown probe from
    // a probe that merely reported "still saving".
    return { ok: false, canSave: null, threw: true, detail: err.message, errorName: err.name };
  }
  return probe && typeof probe.canSave === 'boolean'
    ? probe
    : { ok: false, canSave: null, detail: (probe && probe.detail) || 'no canSave result' };
}

/** The driver's forcesave confirmation, normalized; a driver without one gets an explicit "none", never a guess. */
function confirmationOf(forced) {
  if (!forced || typeof forced.confirmed !== 'boolean') return { state: 'unconfirmed', reason: 'driver-provides-no-save-confirmation' };
  const fs = forced.forceSave || {};
  return { state: forced.confirmed ? 'confirmed' : 'unconfirmed', reason: fs.reason || (forced.confirmed ? 'confirmed' : 'unconfirmed') };
}

function elapsed(metrics, startedAt) {
  return metrics.clock.now() - startedAt;
}

function fail(outcome, rest) {
  // A failed barrier never reports a save — the authority stays at the execution's level.
  return { ok: false, outcome, authority: 'PRIMITIVE_LIVE_VERIFY_ONLY', writeAllowed: false, ...rest };
}

module.exports = { persist, DEFAULT_TIMEOUT_MS, DEFAULT_POLL_MS };
