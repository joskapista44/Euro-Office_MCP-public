'use strict';
/**
 * save-ack — pairs the editor's last saveChanges with the Document Server's unSaveLock acknowledgement,
 * read from the editor <-> DS websocket frames (public Playwright websocket events; no editor-internal API).
 *
 * The protocol, as read from the DS cell SDK 9.3.4 (DocsCoApi, sdk-all-min.js):
 *   client auth  ->  DS auth{result:1}                  (the session handshake)
 *   client isSaveLock{syncChangesIndex}  ->  DS saveLock{saveLock:false|true}
 *   client saveChanges{startSaveChanges,endSaveChanges,...}  (chunked by websocketMaxPayloadSize)
 *     a non-final chunk is answered by DS savePartChanges, the final chunk by DS unSaveLock{index,time,syncChangesIndex}
 *   the client keeps ONE save in flight: it sends the next chunk / the next cycle only after that answer.
 *
 * Owner review points (owner channel), each held by an offline control:
 *   1. Only the editor <-> DS connection counts: a socket on the DS host whose path is a /doc/ socket and whose
 *      `shardkey` query value is this session's document key. Every connection has its OWN state; an answer on one
 *      socket never pairs with a send on another. A DS socket counts as observed only after the auth handshake.
 *   2. verified needs the LAST send (across all DS sockets) to be a final chunk with its own paired ack; the
 *      sync index must advance from a cycle start index that was itself observed.
 *   3. "no save traffic" is a success only for a session the driver knows made no mutation; any DS frame this
 *      module cannot read (unparsable text, binary) is never read as "nothing was saved".
 *
 * The forcesave pairing (owner GO, owner channel), on the SAME per-connection state:
 *   client forceSaveStart  ->  DS forceSaveStart{messages:{code,time}}   code 0 = started, 4 = NotModified, else error
 *                          ->  DS forceSave{messages:{type:1,time,success}}            type 1 = the button/API kind
 *   A request is confirmed only when it was sent AFTER the last send was an acked final chunk, its own start answer is
 *   code 0 with `time` equal to that ack's save time, and its own completion carries the same time and success:true --
 *   all on the request's socket, with no newer send since the request. FORCESAVE_CONFIRMED is the DS's statement about
 *   that save; it is not PERSISTED on its own.
 *   The DS answers carry no request id, so OUR request is identified by the operation window, not by a sequence number:
 *   exactly ONE forceSaveStart may leave between the operation's start and its verdict, no earlier request may still be
 *   open, and no unmatched (stray) answer may arrive meanwhile. Anything else is ambiguous, hence never confirmed.
 */

/** The socket.io "message" event of a text frame: { msg } | { skip } (engine.io control) | { unreadable }. */
function readFrame(payload) {
  if (typeof payload !== 'string') return { unreadable: 'binary' };
  if (!/^4[2-6]/.test(payload)) return { skip: true }; // engine.io open/ping/pong and socket.io connect
  let data = null;
  try {
    const arr = JSON.parse(payload.slice(payload.indexOf('[')));
    data = Array.isArray(arr) && arr[0] === 'message' ? arr[1] : null;
    if (typeof data === 'string') data = JSON.parse(data);
  } catch { data = null; }
  return data && typeof data === 'object' && typeof data.type === 'string' ? { msg: data } : { unreadable: 'text' };
}

/** The editor <-> DS identity of a session, from its editor config: the DS host and the document key. */
function dsIdentity(config) {
  let dsHost = null;
  try { dsHost = new URL(config.documentServerUrl).host; } catch { dsHost = null; }
  return { dsHost, key: (config && config.document && config.document.key) || null };
}

/** Is this websocket URL the editor <-> DS connection of THIS session? (DS host, a /doc/ socket, our key). */
function isDsSocket(url, { dsHost, key }) {
  if (!dsHost || !key) return false;
  let u;
  try { u = new URL(url); } catch { return false; }
  const docSocket = /\/doc\//.test(u.pathname);
  return u.host === dsHost && docSocket && u.searchParams.get('shardkey') === key;
}

const anomaly = (obs, kind, extra = {}) => obs.anomalies.push({ kind, at: obs.now(), ...extra });

function sentLockRequest(obs, s, msg) {
  if (s.cycle || s.outstanding) anomaly(obs, 'save-cycle-while-busy', { socket: s.id });
  const from = msg.syncChangesIndex;
  s.cycle = { syncFrom: Number.isInteger(from) && from >= 0 ? from : null, granted: false };
}

function sentChanges(obs, s, msg) {
  obs.seq += 1;
  s.sends += 1;
  if (s.outstanding) anomaly(obs, 'save-sent-while-outstanding', { socket: s.id });
  if (!s.cycle || !s.cycle.granted) anomaly(obs, 'save-without-granted-lock', { socket: s.id });
  const send = { seq: obs.seq, socket: s.id, end: msg.endSaveChanges === true, at: obs.now(), ackedAt: null,
    syncFrom: s.cycle ? s.cycle.syncFrom : null, syncTo: null };
  s.outstanding = send;
  obs.lastSend = send;
}

function receivedAuth(obs, s, msg) { if (msg.result === 1) s.authed = true; }

function receivedLock(obs, s, msg) {
  if (!s.cycle) { anomaly(obs, 'save-lock-without-request', { socket: s.id }); return; }
  // not granted (another holder or an error): no save happens in this cycle
  if (msg.saveLock === false) s.cycle.granted = true; else s.cycle = null;
}

function receivedPartAck(obs, s) {
  if (!s.outstanding || s.outstanding.end) { anomaly(obs, 'part-ack-unpaired', { socket: s.id }); return; }
  s.outstanding.ackedAt = obs.now();
  s.outstanding = null;
}

function receivedFinalAck(obs, s, msg) {
  if (!s.outstanding) { s.cycle = null; return; } // the answer to a cycle that sent no changes
  const send = s.outstanding;
  if (!send.end) { anomaly(obs, 'final-ack-for-part', { socket: s.id }); return; }
  const sync = Number.isInteger(msg.syncChangesIndex) ? msg.syncChangesIndex : null;
  if (send.syncFrom === null) anomaly(obs, 'sync-from-unobserved', { socket: s.id });
  else if (sync === null || sync <= send.syncFrom) anomaly(obs, 'ack-sync-index-unpairable', { socket: s.id, sync, from: send.syncFrom });
  Object.assign(send, { ackedAt: obs.now(), syncTo: sync, saveTime: Number.isFinite(msg.time) ? msg.time : null });
  s.outstanding = null;
  s.cycle = null;
}

/** A forcesave request: bound to the last send at the moment it leaves, and to that send's ack save time. */
function sentForceSave(obs, s) {
  const last = obs.lastSend;
  const acked = !!(last && last.end && last.ackedAt !== null && last.socket === s.id);
  obs.forceSaves.push({ id: obs.forceSaves.length + 1, socket: s.id, seq: obs.seq, saveTime: acked ? last.saveTime : null,
    afterAck: acked, at: obs.now(), code: null, startTime: null, success: null, doneAt: null });
}

const fsMessages = msg => (msg.messages && typeof msg.messages === 'object' ? msg.messages : {});

function receivedForceSaveStart(obs, s, msg) {
  const fs = obs.forceSaves.find(f => f.socket === s.id && f.code === null); // the oldest unanswered request here
  if (!fs) { obs.strayForceSave += 1; return; }
  const m = fsMessages(msg);
  fs.code = Number.isInteger(m.code) ? m.code : -1;
  fs.startTime = Number.isFinite(m.time) ? m.time : null;
}

/** Does completion `m` (type 1 = button/API) answer request `f` on socket `s`: started, unfinished, the same save time? */
function completes(f, s, m) {
  const open = f.socket === s.id && f.code === 0 && f.success === null;
  return open && m.type === 1 && f.startTime !== null && m.time === f.startTime;
}

function receivedForceSaveDone(obs, s, msg) {
  const m = fsMessages(msg);
  const fs = obs.forceSaves.find(f => completes(f, s, m));
  if (!fs) { obs.strayForceSave += 1; return; }
  fs.success = m.success === true;
  fs.doneAt = obs.now();
}

const ON_SENT = { isSaveLock: sentLockRequest, saveChanges: sentChanges, forceSaveStart: sentForceSave };
const ON_RECEIVED = { auth: receivedAuth, saveLock: receivedLock, savePartChanges: receivedPartAck, unSaveLock: receivedFinalAck,
  forceSaveStart: receivedForceSaveStart, forceSave: receivedForceSaveDone };

const isOpenRequest = f => f.code === null || (f.code === 0 && f.success === null);

/** The verdict of a forceSave OPERATION: its window must hold exactly one request and nothing ambiguous. */
function forceSaveOpVerdictOf(obs, op) {
  const reqs = obs.forceSaves.slice(op.fromIndex);
  const open = obs.sockets.some(x => x.open);
  let out;
  if (op.openBefore > 0) out = { state: 'unconfirmed', reason: 'ambiguous:earlier-request-still-open' };
  else if (reqs.length > 1) out = { state: 'unconfirmed', reason: `ambiguous:${reqs.length}-requests-in-the-operation` };
  else if (obs.strayForceSave > op.strayBefore) out = { state: 'unconfirmed', reason: 'ambiguous:unmatched-forcesave-answer' };
  else if (reqs.length === 0) out = { state: open ? 'pending' : 'unconfirmed', reason: 'awaiting-request' };
  else out = forceSaveVerdictOf(obs, reqs[0].id);
  return out;
}

/** The state of forcesave request `id`: confirmed | pending | unconfirmed, with the reason. */
function forceSaveVerdictOf(obs, id) {
  const fs = obs.forceSaves[id - 1];
  const open = !!fs && obs.sockets.some(x => x.id === fs.socket && x.open);
  let out;
  if (!fs) out = { state: 'unconfirmed', reason: 'request-not-observed' };
  else if (obs.anomalies.length) out = { state: 'unconfirmed', reason: `anomaly:${obs.anomalies[0].kind}` };
  else if (!fs.afterAck || fs.saveTime === null) out = { state: 'unconfirmed', reason: 'request-before-acked-save' };
  else if (obs.seq > fs.seq) out = { state: 'unconfirmed', reason: 'superseded-by-a-newer-save' };
  else if (fs.code === null) out = { state: open ? 'pending' : 'unconfirmed', reason: 'awaiting-start' };
  else if (fs.code !== 0) out = { state: 'unconfirmed', reason: `start-code-${fs.code}` };
  else if (fs.startTime !== fs.saveTime) out = { state: 'unconfirmed', reason: 'start-time-not-the-acked-save' };
  else if (fs.success === null) out = { state: open ? 'pending' : 'unconfirmed', reason: 'awaiting-completion' };
  else out = fs.success ? { state: 'confirmed', reason: 'forcesave-confirmed' } : { state: 'unconfirmed', reason: 'completion-not-successful' };
  return out;
}

/**
 * The close-time verdict. `canSave` is the editor's own "saveable change remains" flag read now; `mutated` is the
 * driver's own record of whether any write command was dispatched in this session.
 *   verified   -- (a) the last send is a final chunk with its paired ack, or (b) no send at all in a session with no
 *                 mutation; in both: a DS socket handshake was observed, nothing in flight on any DS socket, the
 *                 editor has nothing to send, and no anomaly
 *   pending    -- a send or a cycle is still in flight on an open DS socket, or the editor still has a change
 *   unverified -- everything else
 */
function verdictOf(obs, canSave, { mutated }) {
  return blockingVerdict(obs, canSave) || outcomeVerdict(obs, mutated);
}

/** Anything that stops a verdict on the save itself: no DS observation, an anomaly, traffic in flight, unsent edits. */
function blockingVerdict(obs, canSave) {
  const ds = obs.sockets;
  const busy = ds.find(s => s.outstanding || s.cycle);
  if (!ds.some(s => s.authed)) return { state: 'unverified', reason: ds.length ? 'ds-handshake-not-observed' : 'no-ds-socket-observed' };
  if (obs.anomalies.length) return { state: 'unverified', reason: obs.anomalies[0].kind };
  if (busy) return { state: busy.open ? 'pending' : 'unverified', reason: busy.outstanding ? 'save-outstanding' : 'save-cycle-open' };
  if (canSave === false) return null;
  return { state: ds.some(s => s.open) ? 'pending' : 'unverified', reason: canSave === true ? 'editor-has-unsent-changes' : 'can-save-unknown' };
}

/** The save itself: the last send is an acked final chunk, or there was no send in a session with no mutation. */
function outcomeVerdict(obs, mutated) {
  const last = obs.lastSend;
  if (last) return last.end && last.ackedAt !== null ? { state: 'verified', reason: 'last-save-acked' } : { state: 'unverified', reason: 'last-send-not-an-acked-final' };
  if (mutated === false) return { state: 'verified', reason: 'no-save-traffic' };
  return { state: 'unverified', reason: mutated ? 'mutated-without-observed-save' : 'mutation-unknown' };
}

function createSaveAckObserver({ dsHost, key, now = () => Date.now() } = {}) {
  const obs = { now, seq: 0, sockets: [], lastSend: null, anomalies: [], foreignSockets: 0, forceSaves: [], strayForceSave: 0 };
  return {
    /** A new websocket: returns its frame sink. Only this session's DS socket feeds state; others are counted. */
    onSocketOpen(url) {
      if (!isDsSocket(url, { dsHost, key })) { obs.foreignSockets += 1; return { onFrame() {}, onClose() {} }; }
      const s = { id: obs.sockets.length + 1, open: true, authed: false, sends: 0, outstanding: null, cycle: null };
      obs.sockets.push(s);
      return {
        onFrame(dir, payload) {
          const f = readFrame(payload);
          if (f.unreadable) { anomaly(obs, `uninterpreted-ds-frame:${f.unreadable}`, { socket: s.id, dir }); return; }
          const handler = f.msg && (dir === 'sent' ? ON_SENT : ON_RECEIVED)[f.msg.type];
          if (handler) handler(obs, s, f.msg);
        },
        onClose() {
          s.open = false;
          if (s.outstanding || s.cycle) anomaly(obs, 'socket-closed-while-saving', { socket: s.id });
        },
      };
    },
    verdict: (canSave, opts = {}) => verdictOf(obs, canSave, opts),
    forceSaveVerdict: id => forceSaveVerdictOf(obs, id),
    /** Opens a forceSave operation window: what was already sent, still open, or stray before it. */
    beginForceSaveOp: () => ({ fromIndex: obs.forceSaves.length, openBefore: obs.forceSaves.filter(isOpenRequest).length, strayBefore: obs.strayForceSave }),
    forceSaveOpVerdict: op => forceSaveOpVerdictOf(obs, op),
    /** A point in the close: what had been sent and what anomalies existed when the pre-close verdict was taken. */
    mark: () => ({ seq: obs.seq, anomalies: obs.anomalies.length }),
    /** Anything since `mark` that voids a verified close: a late anomaly, a late send, a save still in flight. */
    lateSince: m => {
      const kinds = obs.anomalies.slice(m.anomalies).map(a => a.kind);
      if (kinds.length) return `late:${kinds[0]}`;
      if (obs.seq > m.seq) return 'late:save-sent-during-close';
      return obs.sockets.some(x => x.outstanding || x.cycle) ? 'late:save-in-flight-at-teardown' : null;
    },
    evidence: () => ({
      dsSockets: obs.sockets.map(s => ({ id: s.id, open: s.open, authed: s.authed, sends: s.sends, busy: !!(s.outstanding || s.cycle) })),
      foreignSockets: obs.foreignSockets,
      lastSend: obs.lastSend && { ...obs.lastSend },
      anomalies: obs.anomalies.map(a => a.kind),
      forceSaves: obs.forceSaves.map(f => ({ ...f, verdict: forceSaveVerdictOf(obs, f.id).reason })),
      strayForceSave: obs.strayForceSave,
    }),
  };
}

/** From the session's first frame: every websocket of the page goes through the observer. No events -> null. */
function observePage(page, ids) {
  if (!page || typeof page.on !== 'function') return null;
  const obs = createSaveAckObserver(ids);
  page.on('websocket', ws => {
    const sink = obs.onSocketOpen(ws.url());
    ws.on('framesent', f => sink.onFrame('sent', f.payload));
    ws.on('framereceived', f => sink.onFrame('recv', f.payload));
    ws.on('close', () => sink.onClose());
  });
  return obs;
}

/**
 * Waits, condition-based, until the observer's verdict leaves `pending` or the fixed deadline passes.
 * `readCanSave` returns true | false | null. `wait` is injectable so the offline controls do not sleep.
 */
async function awaitSaveAck(observer, { readCanSave, mutated, timeoutMs, pollMs, now = () => Date.now(), wait }) {
  const start = now();
  let v = observer.verdict(await readCanSave(), { mutated });
  while (v.state === 'pending' && now() - start < timeoutMs) {
    await wait(pollMs);
    v = observer.verdict(await readCanSave(), { mutated });
  }
  const timedOut = v.state === 'pending';
  return { state: timedOut ? 'unverified' : v.state, reason: timedOut ? `timeout:${v.reason}` : v.reason, waitedMs: now() - start };
}

/** Waits, condition-based and bounded, until forceSave operation `op` leaves `pending`. Timeout -> unconfirmed. */
async function awaitForceSave(observer, op, { timeoutMs, pollMs, now = () => Date.now(), wait }) {
  const start = now();
  let v = observer.forceSaveOpVerdict(op);
  while (v.state === 'pending' && now() - start < timeoutMs) {
    await wait(pollMs);
    v = observer.forceSaveOpVerdict(op);
  }
  const timedOut = v.state === 'pending';
  return { state: timedOut ? 'unconfirmed' : v.state, reason: timedOut ? `timeout:${v.reason}` : v.reason, waitedMs: now() - start };
}

/**
 * The forceSave sequence (owner GO, owner channel): (1) the precondition -- every change sent and the last send an
 * acked final chunk, nothing left to send; (2) only then ONE request via `requestForceSave()` -> { ok, sent, detail };
 * (3) its own confirmation. Returns { ok, confirmed, detail, forceSave }. ok:false = no request was made (the
 * precondition failed or the editor could not be asked). A session with nothing saved is never forced to save.
 */
async function confirmForceSave(observer, { readCanSave, mutated, requestForceSave, ackTimeoutMs, confirmTimeoutMs, pollMs, wait }) {
  const result = (ok, state, reason, extra = {}) => ({ ok, confirmed: state === 'confirmed', detail: `forceSave ${state}:${reason}`,
    forceSave: { state, reason, ...extra } });
  if (!observer) return result(false, 'unconfirmed', 'no-observer');
  const pre = await awaitSaveAck(observer, { readCanSave, mutated, timeoutMs: ackTimeoutMs, pollMs, wait });
  let out;
  if (pre.state !== 'verified') out = result(false, 'unconfirmed', `precondition:${pre.reason}`, { precondition: pre });
  else if (pre.reason === 'no-save-traffic') out = result(true, 'not-required', 'nothing-saved-in-session', { precondition: pre });
  else {
    const op = observer.beginForceSaveOp();
    const call = await requestForceSave().catch(err => ({ ok: false, detail: String(err && err.message || err).slice(0, 120) }));
    if (!call.ok) out = result(false, 'unconfirmed', call.detail || 'forceSave-call-failed', { precondition: pre });
    else if (!call.sent) out = result(true, 'unconfirmed', 'request-not-sent-by-editor', { precondition: pre });
    else {
      const fs = await awaitForceSave(observer, op, { timeoutMs: confirmTimeoutMs, pollMs, wait });
      const reqs = observer.evidence().forceSaves.slice(op.fromIndex);
      out = result(true, fs.state, fs.reason, { op, precondition: pre, waitedMs: fs.waitedMs, request: reqs.length === 1 ? reqs[0] : null });
    }
  }
  return out;
}

/**
 * close() end (owner review of 2ab3aca, owner channel): the FINAL close-ack and confirmation, taken AFTER requestClose
 * and the teardown on the same observer and the same forceSave operation window. Every event observed during the
 * close counts: a late send, a late anomaly or a save in flight at the teardown voids a verified pre-close verdict;
 * the confirmation is re-evaluated on that final state, and a failed close leaves no confirmation standing.
 */
function finishClose(observer, { pre, mark, op, closeOk }) {
  if (!observer) return { saveAck: pre, confirmation: null };
  const late = pre.state === 'verified' && mark ? observer.lateSince(mark) : null;
  const saveAck = { ...pre, ...observer.evidence(), ...(late ? { state: 'unverified', reason: late } : {}), lateDuringClose: late };
  let confirmation = op ? observer.forceSaveOpVerdict(op) : null;
  if (confirmation && confirmation.state === 'confirmed' && !closeOk) confirmation = { state: 'unconfirmed', reason: 'close-failed' };
  return { saveAck, confirmation };
}

/**
 * The close() result: ok only for a successful requestClose AND teardown AND a final verified close-ack; a failed
 * close (either step) leaves no save confirmation standing, and a teardown failure is named in the detail.
 */
function closeResult(observer, { pre, mark, op, lifecycle, torn }) {
  const closeOk = lifecycle.ok === true && torn.ok === true;
  const { saveAck, confirmation } = finishClose(observer, { pre, mark, op, closeOk });
  const tornNote = torn.ok ? '' : `; teardown-failed: ${String(torn.detail).replace(/^browser-close-failed: /, '')}`;
  return { ok: closeOk && saveAck.state === 'verified', detail: `${lifecycle.detail}; save-ack ${saveAck.state}:${saveAck.reason}${tornNote}`,
    saveAck, confirmation, teardown: torn };
}

/** close() helper: the bounded wait plus the evidence, never throwing; no observer -> a not-verified close. */
async function settleSaveAck(observer, opts) {
  const ack = observer
    ? await awaitSaveAck(observer, opts).catch(err => ({ state: 'unverified', reason: `ack-wait-threw:${String(err && err.message || err).slice(0, 80)}` }))
    : { state: 'unverified', reason: 'no-observer' };
  return { ...ack, ...(observer ? observer.evidence() : {}) };
}

module.exports = { readFrame, dsIdentity, isDsSocket, createSaveAckObserver, awaitSaveAck, awaitForceSave, confirmForceSave, observePage, settleSaveAck, finishClose, closeResult };
