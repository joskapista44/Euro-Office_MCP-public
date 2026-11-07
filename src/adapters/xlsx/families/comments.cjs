'use strict';

const { parseRange, rangeText, colName } = require('../a1.cjs');
const cells = require('./cells.cjs');

// GH #367 2nd unit: MUST match xlsx-driver.cjs's own MAX_COMMENT_SCAN_CELLS -- the adapter boundary
// (test/static/xlsx-adapter-boundary.test.cjs) forbids importing the driver, so the geometry is
// mirrored here rather than read live; see scanWindowOf() below.
const MAX_COMMENT_SCAN_CELLS = 20000;

/**
 * XLSX `comments` family: add_comment (GH #300 K2, public ApiRange.AddComment + ApiComment.SetUserId), edit_comment (GH #348, public
 * ApiComment.SetText) and delete_comment (GH #349, public ApiComment.Delete). The READ mirror and the verifier is read_comments (K1:
 * the driver's commentSheets read, the anchor found by a per-cell scan).
 *
 * THE AUTHOR (owner decision 2026-10-26): the authenticated Nextcloud user. It is NOT a step field (an `author` / `userId` /
 * any identity-bearing key is refused before dispatch); the driver adds it from the SESSION -- the editor config's
 * editorConfig.user, which Nextcloud signs for the requester's own credential and the editor session runs as -- in front of
 * the command (`__sessionAuthor`), and refuses to dispatch when the session has none. The stable uid (`id`, the
 * instance-prefixed Nextcloud uid) is the identity; the display name (`name`) is written as the author name. The read-back
 * must show BOTH, else the write is not verified.
 *
 * The K1 / K2 LIVE design probes (docs/evidence/gh300-k1-comments-probe-20261026.json, gh300-k2-comment-identity-probe-
 * 20261026.json) set the rules:
 *   - a second AddComment on a cell that already holds a comment RETURNS a comment but is NOT persisted (a silent no-op)
 *     -> a commented cell is refused before the write, unless it already holds exactly this comment by this author (the
 *     retry: SATISFIED, 0 write)
 *   - AddComment's third argument is ignored; SetUserId right after the AddComment persists (fresh reopen + the stored person)
 *   - a comment on a range lands on its top-left cell -> v1 takes ONE cell
 * Refused before the write (named): an unread / unobservable comment read, no session author, a commented cell, a missing
 * sheet, a protected sheet (protection-gate's default). L1: the cell's comment is exactly { text, author = session name,
 * userId = session uid, not resolved, no replies }. L3: every OTHER comment of the sheet unchanged (cell, text, author,
 * userId, time, resolved, replies) and every cell unchanged (cells.verifyPreservation).
 *
 * EDIT / DELETE (GH #348 / #349; the LIVE design probe docs/evidence/gh348-probe-live-20261030.json, prod DS 9.3.5, a comment loaded
 * from a SAVED file through the PUBLIC API, in-session AND after a fresh reopen):
 *   - SetText persists the new text and keeps the author, the uid, the solved flag, the replies and the cell value; the same text again
 *     is a no-op; a multi-line text and a 2000-character text persist
 *   - Delete removes the target comment (its replies go with it, a solved one too) and leaves every other comment and every cell value
 *     -- but ANOTHER author's reply is not lost silently: a thread that holds a reply whose uid is not the session user's is refused
 *     before the write (xlsx-comment-has-foreign-replies); the session user's own replies go with the comment;
 *     a SECOND Delete on the stale handle returns true and does not throw -> a retry cannot be told from the return value, so the
 *     decision is made from the FRESH read: a cell with no comment is SATISFIED (the end state holds), never re-dispatched
 *   - both calls also work on ANOTHER author's comment (the API does not stop it) -> the product's own rule: only the SESSION user's
 *     own comment (uid = the session uid) is edited or deleted; any other (a foreign author, a legacy note without a uid) is refused
 *     BEFORE the write, named (xlsx-comment-foreign-author). The dispatch re-checks the uid on the live comment itself.
 * Refused before the write (named): the same list as add_comment, plus xlsx-comment-cell-empty (edit: nothing to edit) and
 * xlsx-comment-foreign-author. L1 (edit): the cell's comment has exactly the new text and its author, uid, solved flag and replies are
 * the BEFORE comment's; (delete): the cell holds no comment. L3 as add_comment.
 *
 * GH #367 2nd unit (docs/GH367-COMMENT-SCAN-OUTSIDE-USED-RANGE-SOURCE-CONTRACT-20261105.md): `ApiComment` has no
 * address getter, so `read_comments`/`add_comment`'s shared `commentReads` scan (xlsx-driver.cjs:1101-1149) finds a
 * comment's cell by a BOUNDED, row-major scan from A1 -- a comment past that bound is named, never a silent partial
 * list. `add_comment`'s own `gate()` is trivially satisfied for a cell's FIRST comment (`expected=0` skips the scan
 * entirely), so a write whose target sits past that SAME bound dispatches with no protection at all, then becomes a
 * measured ghost: persisted, but invisible to every later read_comments/retry self-check on that sheet. `gate()`/
 * `classify()` now refuse such a target BEFORE dispatch, named `xlsx-comment-target-outside-scan` -- the same
 * geometry (`scanWindowOf()`, mirroring xlsx-driver.cjs:1125-1126), computed client-side from the sheet's own
 * used-range (`p.sheets[].used`, already carried on every read). A target INSIDE the window is unaffected.
 */

const INTENTS = Object.freeze(['add_comment', 'edit_comment', 'delete_comment', 'reply_to_comment', 'set_comment_resolved']);
const FIELDS = Object.freeze({
  add_comment: Object.freeze(['intent', 'id', 'sheet', 'cell', 'text']),
  edit_comment: Object.freeze(['intent', 'id', 'sheet', 'cell', 'text']),
  delete_comment: Object.freeze(['intent', 'id', 'sheet', 'cell']),
  reply_to_comment: Object.freeze(['intent', 'id', 'sheet', 'cell', 'text']),
  set_comment_resolved: Object.freeze(['intent', 'id', 'sheet', 'cell', 'resolved']),
});
// our implementation ceiling, not an engine limit
const MAX_TEXT = 2000;
// printable text and line breaks only: a control character in a comment is not something the probes measured
const CONTROL = /[\u0000-\u0009\u000B-\u001F\u007F]/;

function plan(step, index) {
  const intent = step.intent;
  const at = `step ${index}: ${intent}`;
  const fields = FIELDS[intent];
  if (!fields) return { ok: false, detail: `${at} is not a comments intent` };
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')} (the author is the authenticated session user, never a step field)` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.cell);
  if (!g || g.count !== 1) return { ok: false, detail: `${at} requires exactly one cell (A1)` };
  if (intent === 'set_comment_resolved' && typeof step.resolved !== 'boolean') return { ok: false, detail: `${at} requires resolved: true | false` };
  if (intent !== 'delete_comment' && intent !== 'set_comment_resolved') {
    if (typeof step.text !== 'string' || step.text.trim() === '') return { ok: false, detail: `${at} requires a non-empty text` };
    if (step.text.length > MAX_TEXT) return { ok: false, detail: `${at} text has ${step.text.length} characters, over the ${MAX_TEXT} ceiling` };
    if (CONTROL.test(step.text)) return { ok: false, detail: `${at} text holds a control character (only printable text and line breaks)` };
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent,
      family: 'comments',
      selector: Object.freeze({ sheet: step.sheet, cell: rangeText(g) }),
      ...(intent === 'delete_comment' || intent === 'set_comment_resolved' ? {} : { text: step.text }),
      ...(intent === 'set_comment_resolved' ? { resolved: step.resolved } : {}),
      // a comment owns no cell: every cell is preserved
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** GH #367 2nd unit: the SAME per-sheet scan geometry `commentReads` computes
 * (xlsx-driver.cjs:1111-1150) -- `ApiComment` has no address getter, so the comment anchor is
 * found by a bounded, row-major scan from A1; a comment PAST this window is a named
 * 'UNOBSERVABLE:comment-anchor-outside-scan:...' on the WHOLE sheet, never a silent partial list.
 * But `add_comment`'s own `gate()` (above) is TRIVIALLY satisfied for a target cell's FIRST
 * comment (nothing to scan for yet), so the write dispatches BEFORE that refusal can ever fire for
 * the cell being written -- the comment then becomes a measured "ghost write": persisted, but
 * invisible to every later `read_comments`/retry self-check on that sheet. This mirrors the SAME
 * formula client-side, from the sheet's own used-range (`p.sheets[].used`, the same field every
 * XLSX read already carries), so the write can be refused BEFORE dispatch instead. */
function scanWindowOf(used) {
  const g = used ? parseRange(used) : null;
  const usedCol = g ? g.c2 : 1;
  const usedRow = g ? g.r2 : 1;
  const maxCol = Math.min(16384, Math.max(26, usedCol + 10));
  const maxRow = Math.min(1048576, Math.max(200, usedRow + 100), Math.floor(MAX_COMMENT_SCAN_CELLS / maxCol));
  return { maxCol, maxRow };
}

/** The sheet's comment list and the session author, or the named reason they cannot be told. */
function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const read = p.commentReads ? p.commentReads[op.selector.sheet] : undefined;
  if (read === undefined) return { ok: false, reason: 'xlsx-comments-unobservable', why: 'not-read' };
  if (typeof read === 'string') return { ok: false, reason: 'xlsx-comments-unobservable', why: read };
  if (!read || !Array.isArray(read.comments)) return { ok: false, reason: 'xlsx-comments-unobservable', why: 'malformed-read' };
  const author = p.sessionUser;
  if (!author || typeof author !== 'object' || !author.id || !author.name) return { ok: false, reason: 'xlsx-comment-author-identity-unavailable', why: typeof author === 'string' ? author : 'not-read' };
  return { ok: true, comments: read.comments, author, at: read.comments.find(c => c.cell === op.selector.cell) || null, sheetUsed: s[0].used };
}

/** exactly the written comment: the text, the session author's name AND uid, open, no reply */
const isWritten = (c, op, author) => !!c && c.text === op.text && c.author === author.name && c.userId === author.id && c.solved === false && Array.isArray(c.replies) && c.replies.length === 0;

/** the comment is the SESSION user's own: the stable uid is the identity (a legacy note without one, or another author's, is not) */
const isOwn = (c, author) => !!c && c.userId === author.id;
/** a reply of the thread that is not the session user's (a legacy reply without a uid counts as foreign) */
const hasForeignReplies = (c, author) => !!c && Array.isArray(c.replies) && c.replies.some(r => r.userId !== author.id);

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, cell: op.selector.cell };
  if (op.intent === 'edit_comment') {
    if (!g.at) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-cell-empty' } };
    if (!isOwn(g.at, g.author)) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-foreign-author' } };
    if (g.at.text === op.text) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-the-cell-holds-this-text-by-the-session-author' } };
    return { state: 'APPLY', evidence: where };
  }
  if (op.intent === 'delete_comment') {
    // the observation is NAMED: the call cannot tell "deleted by an earlier run" from "never had a comment" (a mistyped cell), the caller can
    if (!g.at) return { state: 'SATISFIED', evidence: { ...where, observation: 'no-comment-at-cell', proof: 'fresh-live-the-cell-holds-no-comment' } };
    if (!isOwn(g.at, g.author)) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-foreign-author' } };
    if (hasForeignReplies(g.at, g.author)) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-has-foreign-replies' } };
    return { state: 'APPLY', evidence: where };
  }
  if (op.intent === 'reply_to_comment') {
    // GH #350 LIVE probe: a second AddReply with the same text makes a DUPLICATE -- the retry is decided from the fresh read
    if (!g.at) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-cell-empty' } };
    if (!Array.isArray(g.at.replies)) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-replies-unobservable' } };
    if (g.at.replies.some(r => r.userId === g.author.id && r.text === op.text)) return { state: 'SATISFIED', evidence: { ...where, observation: 'reply-already-present', proof: 'fresh-live-a-reply-by-the-session-user-holds-this-text' } };
    return { state: 'APPLY', evidence: where };
  }
  if (op.intent === 'set_comment_resolved') {
    // the resolved state is the THREAD's shared state, not the root author's content: any author's thread can be resolved / reopened
    if (!g.at) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-cell-empty' } };
    if (g.at.solved === op.resolved) return { state: 'SATISFIED', evidence: { ...where, observation: 'comment-already-in-that-state', proof: 'fresh-live-the-resolved-flag-holds-the-target' } };
    return { state: 'APPLY', evidence: where };
  }
  if (isWritten(g.at, op, g.author)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-the-cell-holds-this-comment-by-the-session-author' } };
  if (g.at) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-cell-occupied' } };
  // GH #367 2nd unit: `g.at` is null here because nothing is at this cell YET -- not because the
  // scan already proved it (the whole-sheet unobservable-read branch above would have caught an
  // EXISTING ghost already). A FIRST comment, though, dispatches with no scan at all (expected=0
  // skips it entirely); if the target itself sits past the SAME geometry `commentReads` would use
  // once the comment exists, the write would become a ghost the moment it lands. Refuse it now,
  // named, instead of letting it dispatch unprovably.
  const target = parseRange(op.selector.cell);
  const window = scanWindowOf(g.sheetUsed);
  if (target && (target.c2 > window.maxCol || target.r2 > window.maxRow)) {
    return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-target-outside-scan', window: `A1:${colName(window.maxCol)}${window.maxRow}` } };
  }
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, cell: selector.cell }) };
}

/**
 * edit / delete: one callCommand. The sheet identity, the live comment itself (GetComment), and ITS uid against the session uid
 * (a comment that is not the session user's own is not touched, even if the read said otherwise a moment ago), then SetText / Delete.
 */
function lowerEditOrDelete(op) {
  const { sheet, cell } = op.selector;
  if (op.intent === 'reply_to_comment' || op.intent === 'set_comment_resolved') return lowerReplyOrResolve(op);
  const act = op.intent === 'edit_comment' ? `c.SetText(${JSON.stringify(op.text)});` : `
  const nr = c.GetRepliesCount();
  for (let i = 0; i < nr; i += 1) { if (c.GetReply(i).GetUserId() !== __sessionAuthor.id) return { ok: false, stage: 'foreign-replies' }; }
  c.Delete();`;
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const c = ws.GetRange(${JSON.stringify(cell)}).GetComment();
  if (!c) return { ok: false, stage: 'cell-has-no-comment' };
  if (c.GetUserId() !== __sessionAuthor.id) return { ok: false, stage: 'foreign-author' };
  ${act}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'comments', intent: op.intent, sessionAuthor: true, sheet, cell, ...(op.intent === 'edit_comment' ? { text: op.text } : {}) }),
  };
}

/**
 * reply / resolve (GH #350 / #351): one callCommand. The sheet identity and the live comment, then AddReply with the SESSION author's
 * name and uid (the author is never a step field) / SetSolved(bool). Measured LIVE (docs/evidence/gh350-gh351-probe-live-20261030.json):
 * both persist through save + a fresh reopen, the thread, the root author and the other replies kept.
 */
function lowerReplyOrResolve(op) {
  const { sheet, cell } = op.selector;
  const act = op.intent === 'reply_to_comment' ? `c.AddReply(${JSON.stringify(op.text)}, __sessionAuthor.name, __sessionAuthor.id);` : `c.SetSolved(${op.resolved ? 'true' : 'false'});`;
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const c = ws.GetRange(${JSON.stringify(cell)}).GetComment();
  if (!c) return { ok: false, stage: 'cell-has-no-comment' };
  ${act}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'comments', intent: op.intent, sessionAuthor: true, sheet, cell, ...(op.intent === 'reply_to_comment' ? { text: op.text } : { resolved: op.resolved }) }),
  };
}

/**
 * add_comment: one callCommand: exact-name sheet check, then AddComment with the SESSION author's
 * name and SetUserId with its uid.
 *
 * GH #367 3rd unit v6 (LIVE finding): on a sheet that is NOT the document's active sheet, `AddComment`'s own return value is a
 * DETACHED copy -- `SetUserId` on it never reaches the actually-stored comment (the live measurement:
 * the HELD object reports the session user back, a FRESH `GetRange(cell).GetComment()` right after,
 * in the SAME script, reports an empty userId, and it stays empty after reopen). The real product
 * dispatches `create_sheet` and `add_comment` as separate tasks/sessions, so the target sheet is not
 * guaranteed to be active when this runs. `SetUserId` now runs on a FRESHLY re-fetched comment object
 * and is verified by a fresh re-read in THIS SAME script; if that still does not land, the sheet is
 * made active, the write is retried once on a fresh object, and the document's ORIGINAL active sheet
 * is restored afterwards (this write never leaves the user's active sheet changed).
 */
function lower(op) {
  if (op.intent !== 'add_comment') return lowerEditOrDelete(op);
  const { sheet, cell } = op.selector;
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  if (ws.GetRange(${JSON.stringify(cell)}).GetComment()) return { ok: false, stage: 'cell-occupied' };
  const c = ws.GetRange(${JSON.stringify(cell)}).AddComment(${JSON.stringify(op.text)}, __sessionAuthor.name);
  if (!c) return { ok: false, stage: 'add-returned-null' };
  const live = ws.GetRange(${JSON.stringify(cell)}).GetComment();
  if (!live) return { ok: false, stage: 'add-not-found-on-reread' };
  live.SetUserId(__sessionAuthor.id);
  let verify = ws.GetRange(${JSON.stringify(cell)}).GetComment();
  if (!verify || verify.GetUserId() !== __sessionAuthor.id) {
    const originalActive = Api.GetActiveSheet();
    ws.SetActive();
    const retryLive = ws.GetRange(${JSON.stringify(cell)}).GetComment();
    if (retryLive) retryLive.SetUserId(__sessionAuthor.id);
    if (originalActive && originalActive.GetName() !== name) originalActive.SetActive();
    verify = ws.GetRange(${JSON.stringify(cell)}).GetComment();
    if (!verify || verify.GetUserId() !== __sessionAuthor.id) return { ok: false, stage: 'userid-never-landed' };
  }
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'comments', sessionAuthor: true, sheet, cell, text: op.text }),
  };
}

/** edit L1: exactly the new text; the author, the uid, the solved flag and the replies are the BEFORE comment's (the probe: SetText keeps them) */
const isEdited = (c, op, was) => !!c && !!was && c.text === op.text && c.author === was.author && c.userId === was.userId && c.solved === was.solved
  && JSON.stringify(c.replies) === JSON.stringify(was.replies);

/** a reply list as a multiset (the read reverses the order after a reopen -- GH #350 LIVE probe -- so position is never compared) */
const replyKeys = list => (Array.isArray(list) ? list.map(r => JSON.stringify([r.text, r.author, r.userId])).sort() : null);

function verifyEditOrDelete(op, before, after) {
  const where = { sheet: op.selector.sheet, cell: op.selector.cell };
  if (op.intent === 'reply_to_comment' || op.intent === 'set_comment_resolved') {
    const a = gate(after && after.primary, op);
    const b = gate(before && before.primary, op);
    if (!a.ok || !b.ok) return { ok: false, evidence: { reason: !a.ok ? a.reason : 'xlsx-comment-before-unobservable' } };
    const c = a.at;
    const was = b.at;
    const rootKept = !!c && !!was && c.text === was.text && c.author === was.author && c.userId === was.userId;
    if (op.intent === 'set_comment_resolved') {
      const repliesKept = !!c && !!was && JSON.stringify(replyKeys(c.replies)) === JSON.stringify(replyKeys(was.replies));
      return { ok: rootKept && repliesKept && c.solved === op.resolved, evidence: { ...where, found: !!c, rootKept, repliesKept, resolved: c ? c.solved : null, proof: 'fresh-live-read_comments' } };
    }
    const want = replyKeys([...(was ? was.replies || [] : []), { text: op.text, author: b.author.name, userId: b.author.id }]);
    const repliesAsExpected = !!c && JSON.stringify(replyKeys(c.replies)) === JSON.stringify(want);
    const solvedKept = !!c && !!was && c.solved === was.solved;
    return { ok: rootKept && repliesAsExpected && solvedKept, evidence: { ...where, found: !!c, rootKept, repliesAsExpected, solvedKept, proof: 'fresh-live-read_comments' } };
  }
  const a = gate(after && after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: a.reason, why: a.why || null } };
  if (op.intent === 'delete_comment') {
    const bd = gate(before && before.primary, op);
    const lost = bd.ok && hasForeignReplies(bd.at, bd.author);
    return { ok: !a.at && !lost, evidence: { ...where, found: !!a.at, foreignReplyLost: lost, proof: 'fresh-live-read_comments' } };
  }
  const b = gate(before && before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: 'xlsx-comment-before-unobservable', before: b.reason } };
  const c = a.at;
  const was = b.at;
  return { ok: isEdited(c, op, was), evidence: { ...where, found: !!c, textMatches: !!c && c.text === op.text, authorKept: !!c && !!was && c.author === was.author,
    userIdKept: !!c && !!was && c.userId === was.userId, solvedKept: !!c && !!was && c.solved === was.solved,
    repliesKept: !!c && !!was && JSON.stringify(c.replies) === JSON.stringify(was.replies), proof: 'fresh-live-read_comments' } };
}

function verifyPrimitive(op, before, after) {
  if (op.intent !== 'add_comment') return verifyEditOrDelete(op, before, after);
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason, why: g.why || null } };
  const c = g.at;
  return { ok: isWritten(c, op, g.author), evidence: { sheet: op.selector.sheet, cell: op.selector.cell, found: !!c, textMatches: !!c && c.text === op.text,
    authorIsSession: !!c && c.author === g.author.name, userIdIsSession: !!c && c.userId === g.author.id, open: !!c && c.solved === false, proof: 'fresh-live-read_comments' } };
}

const sameComment = (x, y) => JSON.stringify(x) === JSON.stringify(y);

/** every OTHER comment of the sheet unchanged, and every cell (a comment owns none) */
function verifyPreservation(op, before, after) {
  const b = gate(before && before.primary, op);
  const a = gate(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-comment-preservation-unobservable', before: b.ok ? null : b.reason, after: a.ok ? null : a.reason } };
  const others = list => list.filter(c => c.cell !== op.selector.cell);
  const was = others(b.comments);
  const now = others(a.comments);
  const changed = was.filter(c => !now.some(x => sameComment(x, c))).map(c => c.cell);
  const added = now.filter(c => !was.some(x => x.cell === c.cell)).map(c => c.cell);
  const rest = cells.verifyPreservation(op, before, after);
  return { ok: changed.length === 0 && added.length === 0 && now.length === was.length && rest.ok, evidence: { ...rest.evidence, commentsCompared: was.length, commentsChanged: changed.slice(0, 20), commentsAdded: added.slice(0, 20) } };
}

function wholeTaskChecks(op, finalObservation) {
  const g = gate(finalObservation && finalObservation.primary, op);
  if (op.intent === 'delete_comment') {
    const gone = g.ok && !g.at;
    return [{ name: `${op.selector.sheet}!${op.selector.cell} holds no comment`, ok: gone, detail: gone ? null : { reason: g.reason || 'comment-still-there' } }];
  }
  if (op.intent === 'reply_to_comment') {
    const ok = g.ok && !!g.at && Array.isArray(g.at.replies) && g.at.replies.some(r => r.userId === g.author.id && r.text === op.text);
    return [{ name: `${op.selector.sheet}!${op.selector.cell} holds a reply by the session user with this text`, ok, detail: ok ? null : { reason: g.reason || 'reply-not-present' } }];
  }
  if (op.intent === 'set_comment_resolved') {
    const ok = g.ok && !!g.at && g.at.solved === op.resolved;
    return [{ name: `${op.selector.sheet}!${op.selector.cell} comment resolved = ${op.resolved}`, ok, detail: ok ? null : { reason: g.reason || 'resolved-flag-not-at-target' } }];
  }
  if (op.intent === 'edit_comment') {
    const ok = g.ok && !!g.at && g.at.text === op.text && isOwn(g.at, g.author);
    return [{ name: `${op.selector.sheet}!${op.selector.cell} comment text is the new one, by the session user`, ok, detail: ok ? null : { reason: g.reason || 'not-the-edited-comment' } }];
  }
  const ok = g.ok && isWritten(g.at, op, g.author);
  return [{ name: `${op.selector.sheet}!${op.selector.cell} comment by the session user`, ok, detail: ok ? null : { reason: g.reason || 'not-the-written-comment' } }];
}

const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.cell }];
const liveExtras = op => ({ commentSheets: [op.selector.sheet], sessionUser: true });

module.exports = { INTENTS, MAX_TEXT, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, scanWindowOf };
