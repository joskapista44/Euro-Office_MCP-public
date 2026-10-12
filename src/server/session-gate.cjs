'use strict';

/**
 * M11.3b.2 -- the per-process editor-session bound (owner decision, owner channel 2026-10-10: release boundary).
 *
 * MAX_ACTIVE_SESSIONS_PER_PROCESS=1. One gate per product process, shared by every capability tool:
 *   - a request that finds a session active is refused BUSY at once: no queue, no second browser, no shared editor
 *     session between two MCP requests;
 *   - the lease is released when its request finishes (success or failure) -> the next request is allowed;
 *   - shutdown() closes the active session through the driver's OWN close() (the existing lifecycle), bounded by a
 *     timeout -- it never waits forever.
 */

const MAX_ACTIVE_SESSIONS_PER_PROCESS = 1;

function createSessionGate({ max = MAX_ACTIVE_SESSIONS_PER_PROCESS } = {}) {
  const leases = new Set();
  let closing = false;

  /** tryEnter() -> lease | null. A lease: { track(driver), release() }. */
  function tryEnter() {
    if (closing || leases.size >= max) return null;
    const lease = {
      driver: null,
      track(driver) { lease.driver = driver; },
      release() { leases.delete(lease); },
    };
    leases.add(lease);
    return lease;
  }

  /** shutdown({ timeoutMs }) -> { sessions, cleanup: 'NONE'|'CLOSED'|'TIMEOUT', ms }. New requests are refused from now on. */
  async function shutdown({ timeoutMs }) {
    closing = true;
    const started = Date.now();
    const active = [...leases].filter(l => l.driver);
    if (active.length === 0) return { sessions: 0, cleanup: 'NONE', ms: 0 };
    let timer;
    const timeout = new Promise(res => { timer = setTimeout(() => res('TIMEOUT'), timeoutMs); });
    const closed = Promise.allSettled(active.map(l => l.driver.close())).then(() => 'CLOSED');
    const cleanup = await Promise.race([closed, timeout]);
    clearTimeout(timer);
    return { sessions: active.length, cleanup, ms: Date.now() - started };
  }

  return { tryEnter, shutdown, active: () => leases.size, max };
}

module.exports = { createSessionGate, MAX_ACTIVE_SESSIONS_PER_PROCESS };
