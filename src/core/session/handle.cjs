'use strict';

/**
 * SessionHandle — M0 section 4, `openSession` OUTPUT.
 *
 * The handle is what adapters receive. It deliberately carries NO open, close, save or persist
 * method: lifecycle ownership belongs to core/session/lifecycle.cjs alone (owner requirement 16,
 * M0 4.1). What an adapter may do with the handle is read the live projection and dispatch a
 * primitive; both go through the driver, which the handle does not expose.
 */

let seq = 0;

class SessionHandle {
  constructor({ id, fileId, adapterKind, frame = null, apiWhere = null, metrics }) {
    this.id = id || `session-${++seq}`;
    this.fileId = fileId;
    this.adapterKind = adapterKind;
    this.frame = frame;
    this.apiWhere = apiWhere;
    this.openedMs = null;
    this.metrics = metrics;
    Object.seal(this);
  }
}

/**
 * The read/dispatch surface an adapter is allowed to use. It is a plain function pair rather than
 * methods on the handle so that the set of things an adapter can do is enumerable here, in one
 * place, instead of being whatever the handle happens to expose.
 */
function readLive(handle, request) {
  const { driverOf } = require('./internal.cjs');
  handle.metrics.count('freshReads');
  return driverOf(handle).readLive(handle, request);
}

function dispatch(handle, command) {
  const { driverOf } = require('./internal.cjs');
  const { bumpGeneration } = require('../mutation-epoch.cjs');
  handle.metrics.count('commands');
  // Every dispatch invalidates every observation taken before it. Bumping HERE, at the single
  // point through which a mutation must pass, is what makes staleness structural: a caller
  // cannot forget to invalidate, because it is not the caller's job.
  bumpGeneration(handle);
  return driverOf(handle).dispatch(handle, command);
}

/**
 * Common-mediated PACKAGE read (M4, the physical package transport M2 left at 0).
 *
 * The one place a package is read. It is counted HERE, at the physical I/O boundary -- never
 * inferred from an enrichment hook invocation. It is READ-ONLY by construction: it does not bump
 * the mutation generation, and it is handed only to the enrichment hook (core/observation), never
 * to a primitive, so package evidence cannot become a mutation path.
 */
function readPackage(handle, request) {
  const { driverOf } = require('./internal.cjs');
  const driver = driverOf(handle);
  if (typeof driver.readPackage !== 'function') {
    // No transport, no I/O: nothing is counted, and the caller gets an explicit refusal.
    return Promise.resolve({ ok: false, detail: 'session driver has no package transport' });
  }
  handle.metrics.count('packageReads');
  return driver.readPackage(handle, request);
}

module.exports = { SessionHandle, readLive, dispatch, readPackage };
