'use strict';

const { parseRange, rangeText } = require('../a1.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `comments` family: add_comment (GH #300 K2, public ApiRange.AddComment + ApiComment.SetUserId). The READ mirror and
 * the verifier is read_comments (K1: the driver's commentSheets read, the anchor found by a per-cell scan).
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
 */

const INTENTS = Object.freeze(['add_comment']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'cell', 'text']);
// our implementation ceiling, not an engine limit
const MAX_TEXT = 2000;
// printable text and line breaks only: a control character in a comment is not something the probes measured
const CONTROL = /[\u0000-\u0009\u000B-\u001F\u007F]/;

function plan(step, index) {
  const at = `step ${index}: add_comment`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')} (the author is the authenticated session user, never a step field)` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.cell);
  if (!g || g.count !== 1) return { ok: false, detail: `${at} requires exactly one cell (A1)` };
  if (typeof step.text !== 'string' || step.text.trim() === '') return { ok: false, detail: `${at} requires a non-empty text` };
  if (step.text.length > MAX_TEXT) return { ok: false, detail: `${at} text has ${step.text.length} characters, over the ${MAX_TEXT} ceiling` };
  if (CONTROL.test(step.text)) return { ok: false, detail: `${at} text holds a control character (only printable text and line breaks)` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'add_comment',
      family: 'comments',
      selector: Object.freeze({ sheet: step.sheet, cell: rangeText(g) }),
      text: step.text,
      // a comment owns no cell: every cell is preserved
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
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
  return { ok: true, comments: read.comments, author, at: read.comments.find(c => c.cell === op.selector.cell) || null };
}

/** exactly the written comment: the text, the session author's name AND uid, open, no reply */
const isWritten = (c, op, author) => !!c && c.text === op.text && c.author === author.name && c.userId === author.id && c.solved === false && Array.isArray(c.replies) && c.replies.length === 0;

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, cell: op.selector.cell };
  if (isWritten(g.at, op, g.author)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-the-cell-holds-this-comment-by-the-session-author' } };
  if (g.at) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-comment-cell-occupied' } };
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, cell: selector.cell }) };
}

/** One callCommand: exact-name sheet check, then AddComment with the SESSION author's name and SetUserId with its uid. */
function lower(op) {
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
  c.SetUserId(__sessionAuthor.id);
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'comments', sessionAuthor: true, sheet, cell, text: op.text }),
  };
}

function verifyPrimitive(op, before, after) {
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
  const ok = g.ok && isWritten(g.at, op, g.author);
  return [{ name: `${op.selector.sheet}!${op.selector.cell} comment by the session user`, ok, detail: ok ? null : { reason: g.reason || 'not-the-written-comment' } }];
}

const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.cell }];
const liveExtras = op => ({ commentSheets: [op.selector.sheet], sessionUser: true });

module.exports = { INTENTS, MAX_TEXT, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
