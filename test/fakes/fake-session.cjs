'use strict';

/**
 * Deterministic fake session driver — owner requirement 19.
 *
 * It simulates the TRANSPORT only: open, close, forceSave, canSave, readLive, dispatch.
 * It contains no DOCX or XLSX semantics, and it never sleeps: every outcome is scripted, so a
 * test asserts on a decision, not on timing.
 *
 * Every call is counted, which is how OPEN_COUNT=1 / CLOSE_COUNT=1 is proven rather than assumed.
 */

function createFakeSession(script = {}) {
  const calls = { open: 0, close: 0, forceSave: 0, canSave: 0, readLive: 0, dispatch: 0, readPackage: 0 };

  // canSave answers, consumed in order. The barrier polls until it sees false.
  const canSaveSequence = (script.canSaveSequence || [false]).slice();

  const driver = {
    calls,

    async open() {
      calls.open += 1;
      if (script.openFails) return { ok: false, detail: script.openFailDetail || 'fake open refused' };
      return { ok: true, id: 'fake-session-1', frame: 'fake-frame', apiWhere: 'fake-api' };
    },

    async close() {
      calls.close += 1;
      if (script.closeThrows) throw new Error(script.closeThrowDetail || 'fake close threw');
      const confirmation = script.closeConfirmation || (calls.forceSave > 0 && !script.noSaveConfirmation && !script.forceSaveUnconfirmed ? { state: 'confirmed', reason: 'forcesave-confirmed' } : null);
      if (script.closeFails) return { ok: false, detail: script.closeFailDetail || 'fake close refused', confirmation };
      return { ok: true, detail: 'closed', confirmation };
    },

    async forceSave() {
      calls.forceSave += 1;
      // A transport can THROW rather than return a refusal. The core must normalize this.
      if (script.forceSaveThrows) throw new Error(script.forceSaveThrowDetail || 'fake forceSave threw');
      if (script.forceSaveUnavailable) return { ok: false, detail: 'fake forceSave unavailable' };
      // M7 save contract: a driver WITH a save observer reports its own paired confirmation; one without reports none
      if (script.noSaveConfirmation) return { ok: true, detail: 'forceSave dispatched' };
      if (script.forceSaveUnconfirmed) return { ok: true, confirmed: false, detail: 'forceSave dispatched', forceSave: { state: 'unconfirmed', reason: script.forceSaveUnconfirmed } };
      return { ok: true, confirmed: true, detail: 'forceSave dispatched', forceSave: { state: 'confirmed', reason: 'forcesave-confirmed' } };
    },

    async canSave() {
      calls.canSave += 1;
      // Throwing on the FIRST probe and throwing DURING polling are different code paths.
      if (script.canSaveThrowsAt === calls.canSave) {
        throw new Error(`fake canSave threw on probe ${calls.canSave}`);
      }
      if (script.canSaveThrows) throw new Error(script.canSaveThrowDetail || 'fake canSave threw');
      const next = canSaveSequence.length > 1 ? canSaveSequence.shift() : canSaveSequence[0];
      if (script.canSaveBroken) return { ok: false, detail: 'fake canSave unavailable' };
      return { ok: true, canSave: next };
    },

    async readLive(_handle, request) {
      calls.readLive += 1;
      if (script.readLiveThrows) throw new Error(script.readLiveThrowDetail || 'fake readLive threw');
      if (script.readLiveFails) return { ok: false, detail: 'fake readLive refused' };
      return { ok: true, authority: 'LIVE_READ', payload: script.livePayload || { units: [] }, request };
    },

    async readPackage(_handle, request) {
      calls.readPackage += 1;
      if (script.readPackageFails) return { ok: false, detail: 'fake package read refused' };
      return { ok: true, parts: script.packageParts || {}, request };
    },

    async dispatch(_handle, command) {
      calls.dispatch += 1;
      return { ok: script.dispatchFails !== true, dispatched: true, command };
    },
  };

  return driver;
}

module.exports = { createFakeSession };
