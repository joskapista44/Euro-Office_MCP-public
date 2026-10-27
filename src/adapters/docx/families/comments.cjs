'use strict';

const { norm } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');
const { bodyKey } = require('./page-setup.cjs');

/**
 * DOCX `comments` family (GH #269c): resolve_comment (SetSolved true|false).
 *
 * resolve_comment is a NEW_INTENTS capability (capabilities.cjs), not one of the 28 migrated
 * intents -- like set_page_setup it is deliberately NOT in field-contract.cjs's STEP_KEYS, so the
 * unknown-field refusal is re-implemented HERE, locally.
 *
 * Evidence: docs/GH269-DOCX-COMMENTS-READ-WRITE-RESEARCH.md 3. (ApiComment.SetSolved(bSolved), both
 * directions, author-independent) + 7. 2.; the g269v2 LIVE ROUND R6 (SetSolved true then false) PASS.
 *
 * Identification: `text` + `author` + `quote` TOGETHER, exact string equality (the quote's TRAILING CR/LF aside, see `eol`), against the main
 * document's comments (the opt-in `includeComments` read, GH #269b). The `GetId` stability across a
 * reopen is NOT MEASURED, so the id is never used as the key: the command re-finds the comment by
 * the same three fields inside the editor. Zero or more than one match is refused BEFORE dispatch
 * (FAIL_CLOSED, a named reason), never resolved by position.
 *
 * Out of scope (stated): writing a new comment or a reply (the author identity is an OWNER decision),
 * Delete, and every SetAuthorName / SetUserId / SetTime / SetText / SetPosition setter -- the
 * command body below calls NONE of them (a static gate in the acceptance test pins that).
 */

const FIELDS = Object.freeze(['text', 'author', 'quote']);
const STEP_KEYS = Object.freeze(['intent', 'id', 'dependsOn', ...FIELDS, 'resolved']);
const isUnobservable = v => typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0;

function plan(step, index) {
  if (step.intent !== 'resolve_comment') return { ok: false, detail: `step ${index}: ${step.intent} is not a comments intent` };
  for (const key of Object.keys(step)) {
    if (!STEP_KEYS.includes(key)) return { ok: false, detail: `field-unsupported: step ${index} (resolve_comment).${key}` };
  }
  for (const f of ['text', 'author']) {
    if (typeof step[f] !== 'string' || !norm(step[f])) return { ok: false, detail: `step ${index}: resolve_comment requires a non-empty ${f} (text + author + quote identify the comment)` };
  }
  if (typeof step.quote !== 'string') return { ok: false, detail: `step ${index}: resolve_comment requires quote (a string, '' for a comment with no quoted range)` };
  if (step.resolved !== undefined && typeof step.resolved !== 'boolean') return { ok: false, detail: `step ${index}: resolve_comment resolved must be a boolean` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: 'resolve_comment', comment: { text: step.text, author: step.author, quote: step.quote },
      resolved: step.resolved !== false, preservationMode: 'PRIMITIVE_INTEGRAL',
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

/**
 * A quote's TRAILING CR/LF is not part of its identity (g269cv1): a paragraph comment reads 'GH269 gamma\r\n' in the editing session and 'GH269 gamma' after a save round-trip. Both
 * sides of the match drop it; an INNER CR/LF and every other character stay exact.
 */
const eol = v => (typeof v === 'string' ? v.replace(/[\r\n]+$/, '') : v);
const sameField = (a, b) => (a === null || a === undefined ? '' : a) === (b === null || b === undefined ? '' : b);
const matches = (row, c) => !!row && FIELDS.every(f => (typeof row[f] === 'string' || row[f] === null) && !isUnobservable(row[f]) && sameField(f === 'quote' ? eol(row[f]) : row[f], f === 'quote' ? eol(c[f]) : c[f]));

/** The state that decides classification and verification, from ONE observation. */
function commentState(primary, op) {
  const d = primary && primary.document;
  const list = d ? d.comments : null;
  if (!Array.isArray(list)) return { ok: false, reason: 'comments-unreadable' };
  const hits = list.filter(r => matches(r, op.comment));
  if (hits.length === 0) return { ok: false, reason: 'comment-not-found', total: list.length };
  if (hits.length > 1) return { ok: false, reason: 'comment-ambiguous', count: hits.length };
  const row = hits[0];
  if (typeof row.solved !== 'boolean') return { ok: false, reason: `comment-solved-state-unobservable${isUnobservable(row.solved) ? `:${row.solved}` : ''}` };
  return { ok: true, row, list, solved: row.solved, done: row.solved === op.resolved };
}

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  const s = commentState(primary, op);
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason, count: s.count } };
  if (s.done) {
    if (op.recovery === UNRESOLVED) return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } };
    return { state: 'SATISFIED', evidence: { resolved: op.resolved, proof: 'fresh-live-comment-solved-state-matches' } };
  }
  return { state: 'APPLY', evidence: { from: s.solved, to: op.resolved } };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  if (!op || op.intent !== 'resolve_comment') return { ok: false, target: null, selector, ambiguity: 'INVALID_SELECTOR' };
  const s = commentState(primary, op);
  if (!s.ok) return { ok: false, target: null, selector, ambiguity: s.reason.replace(/-/g, '_').toUpperCase() };
  return { ok: true, target: { kind: 'comment', comment: op.comment, resolved: op.resolved }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  const same = target && target.kind === 'comment' && FIELDS.every(f => target.comment[f] === op.comment[f]) && target.resolved === op.resolved;
  return same ? { ok: true } : { ok: false, dispatched: false, detail: 'comment-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({ family: 'comments', action: 'resolveComment', comment: target.comment, resolved: target.resolved });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand, public API only: re-finds the comment by text + author + quote (the id is not a
 * stable key), requires EXACTLY one match, then SetSolved(bool). No other setter is called.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  const str = v => (v === null || v === undefined ? '' : String(v));
  const eol = v => str(v).replace(/[\\r\\n]+$/, '');
  const hits = (doc.GetAllComments() || []).filter(c => str(c.GetText()) === m.comment.text && str(c.GetAuthorName()) === m.comment.author && eol(c.GetQuoteText()) === eol(m.comment.quote));
  if (hits.length !== 1) return;
  hits[0].SetSolved(m.resolved);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** L1: the one matching comment now reads the requested solved state; it was not there before. */
function verifyPrimitive(op, before, after) {
  const b = commentState(before && before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const a = commentState(after && after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  return { ok: !b.done && a.done, evidence: { solvedBefore: b.solved, solvedAfter: a.solved, resolved: op.resolved } };
}

const rowKey = r => JSON.stringify([r.id, r.text, r.author, r.userId, r.timeUtc, eol(r.quote), r.replies]);
const otherKeys = (s, op) => s.list.filter(r => !matches(r, op.comment)).map(r => JSON.stringify([rowKey(r), r.solved]));

/**
 * L3, PRIMITIVE_INTEGRAL: the body, the comment count, every OTHER comment (incl. its solved state),
 * and every field of the target comment except `solved` (text, author, userId, time, quote, replies --
 * the forgery surface the setters would open) are unchanged.
 */
function verifyPreservation(op, before, after) {
  const b = commentState(before && before.primary, op);
  const a = commentState(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` } };
  const sameCount = b.list.length === a.list.length;
  const sameOthers = JSON.stringify(otherKeys(b, op)) === JSON.stringify(otherKeys(a, op));
  const sameTarget = rowKey(b.row) === rowKey(a.row);
  const sameBody = bodyKey(before.primary) === bodyKey(after.primary);
  return { ok: sameCount && sameOthers && sameTarget && sameBody, evidence: { sameCount, otherCommentsUnchanged: sameOthers, targetFieldsUnchanged: sameTarget, bodyUnchanged: sameBody } };
}

function wholeTaskChecks(op, finalObservation) {
  const s = commentState(finalObservation.primary, op);
  return [{ name: `${op.id}: comment by ${op.comment.author} reads ${op.resolved ? 'resolved' : 'open'}`, ok: s.ok && s.done }];
}

// refs: none -- a comment state is not a body text unit a later step can name by reference.
function producedText() {
  return null;
}

module.exports = { FIELDS, eol, matches, rowKey, isUnobservable, plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, commentState };
