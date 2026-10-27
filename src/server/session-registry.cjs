'use strict';

const crypto = require('node:crypto');

/**
 * GH #100 U2 -- the per-process WARM-SESSION registry (docs/GH100-SESSION-REUSE-DESIGN.md sections 3.1, 3.5, 3.6). It HOLDS at most one idle editor
 * session between calls and decides, by rule, whether a later call may take it. It does not open, drive or verify a session: opening stays the driver's
 * and the lifecycle's, the call-end claim is the driver's release() (U1), the per-call freshness check is U3, the lifecycle reuse path is U4. Until U4
 * nothing calls hold()/take(), and with the switch OFF (the default) hold() refuses and take() answers null, so the product behaves exactly as before.
 *
 *   concurrency  unchanged: MAX_ACTIVE_SESSIONS_PER_PROCESS = 1 (session-gate.cjs) -- one call at a time, BUSY otherwise. "Held" is not "active":
 *                a held session is IDLE; take() makes it ACTIVE (the registry no longer owns it, the gate's lease does).
 *   one browser  at most ONE held session, ever. A hold() while another is held closes the older first; a take() for another key closes the held one.
 *   the key      (fileId, principal, config signature): a warm session NEVER serves another file, requester or config.
 *   limits       idle timeout, maximum age, maximum number of calls -- all three REQUIRED when the switch is on (no default is invented: the open cost
 *                split and the lock behaviour of a long-lived page are not measured); over any of them the session is closed, never served.
 *   shutdown     bounded; closes the held session through the driver's own close().
 */

const ENV = Object.freeze({ SWITCH: 'EURO_MCP_SESSION_REUSE', IDLE: 'EURO_MCP_SESSION_IDLE_MS', AGE: 'EURO_MCP_SESSION_MAX_AGE_MS', CALLS: 'EURO_MCP_SESSION_MAX_CALLS' });
// hard ceilings: a typo can never hold a lock for hours; they are safety caps, NOT recommended values
const CEILING = Object.freeze({ idleMs: 600000, maxAgeMs: 3600000, maxCalls: 1000 });

const intIn = (raw, max) => {
  if (typeof raw !== 'string' || !/^[1-9]\d*$/.test(raw)) return null;
  const n = Number(raw);
  return Number.isSafeInteger(n) && n <= max ? n : null;
};

/** sessionReuseConfig(env) -> { enabled:false, reason } | { enabled:true, idleMs, maxAgeMs, maxCalls }. Anything unclear is OFF (fail-closed). */
function sessionReuseConfig(env = process.env) {
  const sw = env[ENV.SWITCH];
  if (sw === undefined || sw === '' || sw === '0') return Object.freeze({ enabled: false, reason: 'switch-off' });
  if (sw !== '1') return Object.freeze({ enabled: false, reason: 'switch-invalid' });
  const idleMs = intIn(env[ENV.IDLE], CEILING.idleMs);
  if (idleMs === null) return Object.freeze({ enabled: false, reason: 'idle-ms-required-or-invalid' });
  const maxAgeMs = intIn(env[ENV.AGE], CEILING.maxAgeMs);
  if (maxAgeMs === null) return Object.freeze({ enabled: false, reason: 'max-age-ms-required-or-invalid' });
  const maxCalls = intIn(env[ENV.CALLS], CEILING.maxCalls);
  if (maxCalls === null) return Object.freeze({ enabled: false, reason: 'max-calls-required-or-invalid' });
  return Object.freeze({ enabled: true, idleMs, maxAgeMs, maxCalls });
}

/** The identity of what a warm session may serve. Strings only; the principal is a uid, never a credential. */
function sessionKey({ fileId, principal, configSignature }) {
  if ([fileId, principal, configSignature].some(x => typeof x !== 'string' || x === '')) return null;
  return JSON.stringify([fileId, principal, configSignature]);
}

const digest = key => crypto.createHash('sha256').update(String(key)).digest('hex').slice(0, 8);

function createSessionRegistry(config = { enabled: false, reason: 'switch-off' }, { now = Date.now, setTimer = setTimeout, clearTimer = clearTimeout, log = null } = {}) {
  const enabled = config.enabled === true;
  let held = null; // { key, driver, openedAt, heldAt, calls, timer }

  const note = (event, data) => { if (log && typeof log.info === 'function') { try { log.info(event, data); } catch (_e) { /* logging never changes the outcome */ } } };

  /** Closes the held session through the driver's OWN close(); never throws, never leaves `held` set. */
  async function discard(reason) {
    const h = held;
    if (!h) return { discarded: false };
    held = null;
    if (h.timer) clearTimer(h.timer);
    let closed;
    try { closed = await h.driver.close(); } catch (err) { closed = { ok: false, detail: String(err && err.message || err).slice(0, 200) }; }
    note('warm session discarded', { reason, keyDigest: digest(h.key), calls: h.calls, ageMs: now() - h.openedAt, closeOk: !!(closed && closed.ok) });
    return { discarded: true, reason, closeOk: !!(closed && closed.ok) };
  }

  /**
   * hold(key, driver, { openedAt, calls, lastRelease, opened }) -> { held: true } | { held: false, reason }. The caller offers a session that ENDED A CALL CLEANLY (that
   * judgement is the lifecycle's, U4); on { held: false } the caller closes it itself -- the registry never closes a session it did not accept.
   */
  async function hold(key, driver, { openedAt = now(), calls = 1, lastRelease = null, opened = null } = {}) {
    if (!enabled) return { held: false, reason: 'reuse-disabled' };
    if (typeof key !== 'string' || !driver || typeof driver.close !== 'function') return { held: false, reason: 'invalid-offer' };
    if (now() - openedAt >= config.maxAgeMs) return { held: false, reason: 'max-age-reached' };
    if (calls >= config.maxCalls) return { held: false, reason: 'max-calls-reached' };
    if (held) await discard('replaced-by-newer');
    const entry = { key, driver, openedAt, heldAt: now(), calls, lastRelease, opened, timer: null };
    entry.timer = setTimer(() => { if (held === entry) discard('idle-timeout'); }, config.idleMs);
    if (entry.timer && typeof entry.timer.unref === 'function') entry.timer.unref();
    held = entry;
    return { held: true };
  }

  /**
   * take(key) -> { driver, openedAt, calls, lastRelease, opened } | null. null = open a fresh session. The held session is removed (the caller owns it again); a key
   * mismatch closes it (one browser per process), and so does an expired one -- an expired session is never served.
   */
  async function take(key) {
    if (!enabled || !held) return null;
    const h = held;
    if (h.key !== key) { await discard('other-key'); return null; }
    if (now() - h.heldAt >= config.idleMs) { await discard('idle-expired'); return null; }
    if (now() - h.openedAt >= config.maxAgeMs) { await discard('max-age-reached'); return null; }
    if (h.calls >= config.maxCalls) { await discard('max-calls-reached'); return null; }
    held = null;
    if (h.timer) clearTimer(h.timer);
    return { driver: h.driver, openedAt: h.openedAt, calls: h.calls, lastRelease: h.lastRelease, opened: h.opened };
  }

  /** shutdown({ timeoutMs }) -> { sessions, cleanup: 'NONE'|'CLOSED'|'TIMEOUT', ms }. Bounded: it never waits forever. */
  async function shutdown({ timeoutMs = 5000 } = {}) {
    if (!held) return { sessions: 0, cleanup: 'NONE', ms: 0 };
    const started = now();
    let timer;
    const timeout = new Promise(res => { timer = setTimeout(() => res('TIMEOUT'), timeoutMs); });
    const cleanup = await Promise.race([discard('shutdown').then(() => 'CLOSED'), timeout]);
    clearTimeout(timer);
    return { sessions: 1, cleanup, ms: now() - started };
  }

  /** Non-secret state (a digest of the key, never the key). */
  function status() {
    if (!enabled) return { enabled: false, reason: config.reason || 'switch-off', held: false };
    return held ? { enabled: true, held: true, keyDigest: digest(held.key), calls: held.calls, ageMs: now() - held.openedAt, idleMs: now() - held.heldAt }
      : { enabled: true, held: false };
  }

  return { enabled, hold, take, discard, shutdown, status };
}

module.exports = { ENV, CEILING, sessionReuseConfig, sessionKey, createSessionRegistry };
