'use strict';

const { norm, textUnits } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');
const { bodyKey } = require('./page-setup.cjs');
const { eol, matches, rowKey, isUnobservable } = require('./comments.cjs');

/**
 * DOCX `comment-write` family (GH #269d): reply_comment (ApiComment.AddReply) and add_comment_on_match (ApiRange.AddComment).
 *
 * Both are NEW_INTENTS capabilities (capabilities.cjs), deliberately NOT in field-contract.cjs's STEP_KEYS (the unknown-field refusal is local, as in
 * comments.cjs / page-setup.cjs). Both are UNRELEASED.
 *
 * AUTHOR (OWNER DECISION: the author is the editing session's own identity, as for a real person):
 * the product NEVER names the author. The command calls AddReply(text) / AddComment(text) with ONE argument, so the editor stamps its own SESSION identity
 * (docs/GH269-DOCX-COMMENTS-READ-WRITE-RESEARCH.md 4. (C)). The task has NO author field (UNKNOWN_FIELD), and SetAuthorName / SetUserId / SetTime / SetTimeUTC /
 * SetText / SetPosition are never called (a static test pins the command bodies).
 * The session identity itself is NOT read by the product: reading it needs AscCommon.UserInfoParser / Asc.editor.documentUserId, which the NEW_DIRECT_INTERNAL=0 gate
 * (test/static/direct-internal-guard.test.cjs) forbids without an owner decision. So the offline contract checks the WRITE by what is observable through the public getters
 * (text, quote, a non-empty author, the replies/comment count, everything else unchanged); that the saved author EQUALS the session identity is measured by the LIVE
 * harness only (the g269 probe's own identity read), and the retry key is author-INDEPENDENT (stated limit below).
 *
 * reply_comment      { target: { text, author, quote }, text }  -- the target is identified as in resolve_comment (text + author + quote TOGETHER, exact, exactly one match);
 *                    `author` inside `target` IDENTIFIES the existing comment, it is not the author of the reply. Retry: the target already has a reply with this text (by ANY author)
 *                    -> SATISFIED (more than one -> FAIL_CLOSED reply-duplicated).
 * add_comment_on_match { matchText, text, occurrence? } -- a comment on a body text range. The occurrences of matchText are counted over the body text units (paragraph / heading /
 *                    list item; exact, case-sensitive, non-overlapping per unit, document order). More than one and no `occurrence` (1-based) -> FAIL_CLOSED before dispatch
 *                    (match-ambiguous:<n>). The command re-counts with doc.Search and writes ONLY when its count equals the planned one (a match inside a table makes the counts
 *                    differ: no write, the L1 then fails closed). Retry: a comment with this text on this quote exists (by ANY author) -> SATISFIED; the key is
 *                    text + quote, NOT the occurrence and NOT the author, so a second identical comment on ANOTHER occurrence of the same matchText, or the same text by another
 *                    person, is treated as already done (stated limit: no duplicate is ever written).
 */

const STEP_KEYS = Object.freeze({
  reply_comment: Object.freeze(['intent', 'id', 'dependsOn', 'target', 'text']),
  add_comment_on_match: Object.freeze(['intent', 'id', 'dependsOn', 'matchText', 'text', 'occurrence']),
});
const TARGET_KEYS = Object.freeze(['author', 'quote', 'text']);

const nonEmpty = v => typeof v === 'string' && norm(v) !== '';

function plan(step, index) {
  const keys = STEP_KEYS[step.intent];
  if (!keys) return { ok: false, detail: `step ${index}: ${step.intent} is not a comment-write intent` };
  for (const key of Object.keys(step)) {
    if (!keys.includes(key)) return { ok: false, detail: `field-unsupported: step ${index} (${step.intent}).${key}` };
  }
  if (!nonEmpty(step.text)) return { ok: false, detail: `step ${index}: ${step.intent} requires a non-empty text` };
  const common = { id: step.id || `s${index}`, intent: step.intent, text: step.text, preservationMode: 'PRIMITIVE_INTEGRAL', ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) };
  if (step.intent === 'reply_comment') {
    const t = step.target;
    if (!t || typeof t !== 'object' || Array.isArray(t) || Object.keys(t).sort().join(',') !== TARGET_KEYS.join(',')) {
      return { ok: false, detail: `step ${index}: reply_comment requires target: { text, author, quote } (exactly these keys: the existing comment's identity)` };
    }
    if (!nonEmpty(t.text) || !nonEmpty(t.author) || typeof t.quote !== 'string') return { ok: false, detail: `step ${index}: reply_comment target.text / target.author must be non-empty strings and target.quote a string ('' = no quoted range)` };
    return { ok: true, operation: { ...common, comment: { text: t.text, author: t.author, quote: t.quote } } };
  }
  if (!nonEmpty(step.matchText)) return { ok: false, detail: `step ${index}: add_comment_on_match requires a non-empty matchText` };
  if (step.occurrence !== undefined && !(Number.isInteger(step.occurrence) && step.occurrence >= 1)) return { ok: false, detail: `step ${index}: add_comment_on_match occurrence must be a positive integer (1-based)` };
  return { ok: true, operation: { ...common, matchText: step.matchText, ...(step.occurrence !== undefined ? { occurrence: step.occurrence } : {}) } };
}

// ---- observation --------------------------------------------------------------------------------

const commentsOf = primary => {
  const l = primary && primary.document ? primary.document.comments : null;
  return Array.isArray(l) ? l : null;
};

const occurrencesIn = (text, needle) => { let n = 0; for (let at = text.indexOf(needle); at >= 0; at = text.indexOf(needle, at + needle.length)) n += 1; return n; };

function replyState(primary, op) {
  const list = commentsOf(primary);
  if (!list) return { ok: false, reason: 'comments-unreadable' };
  const hits = list.filter(r => matches(r, op.comment));
  if (hits.length === 0) return { ok: false, reason: 'comment-not-found' };
  if (hits.length > 1) return { ok: false, reason: 'comment-ambiguous', count: hits.length };
  const row = hits[0];
  if (!Array.isArray(row.replies)) return { ok: false, reason: 'comment-replies-unobservable' };
  const mine = row.replies.filter(r => r && r.text === op.text).length;
  return { ok: true, list, row, mine };
}

function addState(primary, op) {
  const list = commentsOf(primary);
  if (!list) return { ok: false, reason: 'comments-unreadable' };
  const total = textUnits(primary).reduce((n, u) => n + occurrencesIn(String(u.text || ''), op.matchText), 0);
  if (total === 0) return { ok: false, reason: 'match-not-found' };
  if (op.occurrence === undefined && total > 1) return { ok: false, reason: `match-ambiguous:${total}`, count: total };
  if (op.occurrence !== undefined && op.occurrence > total) return { ok: false, reason: `occurrence-out-of-range:${op.occurrence}/${total}` };
  const mine = list.filter(r => r && r.text === op.text && eol(r.quote) === eol(op.matchText)).length;
  return { ok: true, list, total, mine };
}

const stateOf = (primary, op) => (op.intent === 'reply_comment' ? replyState(primary, op) : addState(primary, op));

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  const s = stateOf(primary, op);
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason, count: s.count } };
  if (s.mine > 1) return { state: 'FAIL_CLOSED', evidence: { reason: op.intent === 'reply_comment' ? 'reply-duplicated' : 'comment-duplicated', count: s.mine } };
  if (s.mine === 1) {
    if (op.recovery === UNRESOLVED) return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } };
    return { state: 'SATISFIED', evidence: { proof: 'fresh-live-text-already-present' } };
  }
  return { state: 'APPLY', evidence: {} };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  if (!op || !STEP_KEYS[op.intent]) return { ok: false, target: null, selector, ambiguity: 'INVALID_SELECTOR' };
  const s = stateOf(primary, op);
  if (!s.ok) return { ok: false, target: null, selector, ambiguity: s.reason.replace(/[-:/]/g, '_').toUpperCase() };
  return { ok: true, target: { kind: 'comment_write', intent: op.intent, comment: op.comment, matchText: op.matchText, occurrence: op.occurrence, total: s.total, text: op.text }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  const same = target && target.kind === 'comment_write' && target.intent === op.intent && target.text === op.text
    && (op.intent === 'reply_comment' ? target.comment && ['text', 'author', 'quote'].every(f => target.comment[f] === op.comment[f]) : target.matchText === op.matchText);
  return same ? { ok: true } : { ok: false, dispatched: false, detail: 'comment-write-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = op.intent === 'reply_comment'
    ? Object.freeze({ family: 'commentWrite', action: 'replyComment', comment: target.comment, text: target.text })
    : Object.freeze({ family: 'commentWrite', action: 'addComment', matchText: target.matchText, occurrence: target.occurrence === undefined ? 1 : target.occurrence, total: target.total, text: target.text });
  return { kind: 'callCommand', body: meta.action === 'replyComment' ? replyBody(meta) : addBody(meta), meta };
}

/** AddReply(text) with ONE argument: the editor stamps its session identity. The target is re-found by text + author + quote, exactly one. */
function replyBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  const str = v => (v === null || v === undefined ? '' : String(v));
  const eol = v => str(v).replace(/[\\r\\n]+$/, '');
  const hits = (doc.GetAllComments() || []).filter(c => str(c.GetText()) === m.comment.text && str(c.GetAuthorName()) === m.comment.author && eol(c.GetQuoteText()) === eol(m.comment.quote));
  if (hits.length !== 1) return;
  hits[0].AddReply(m.text);
} catch (_) {}
`;
}

/** AddComment(text) with ONE argument on the planned occurrence; writes ONLY when doc.Search finds exactly the planned number of occurrences. */
function addBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const doc = Api.GetDocument();
  const hits = (doc.Search(m.matchText, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || '') === m.matchText);
  if (hits.length !== m.total) return;
  const range = hits[m.occurrence - 1];
  if (!range) return;
  range.AddComment(m.text);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** the saved text carries SOME readable author (WHICH one is a LIVE measurement: it must equal the session identity) */
const authored = a => typeof a === 'string' && norm(a) !== '' && !isUnobservable(a);

const withSolved = r => JSON.stringify([rowKey(r), r.solved]);

function verifyPrimitive(op, before, after) {
  const b = stateOf(before && before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const a = stateOf(after && after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  if (op.intent === 'reply_comment') {
    const ok = a.mine === b.mine + 1 && a.row.replies.length === b.row.replies.length + 1;
    const last = a.row.replies[a.row.replies.length - 1];
    return { ok: ok && !!last && last.text === op.text && authored(last.author), evidence: { repliesBefore: b.row.replies.length, repliesAfter: a.row.replies.length, author: last ? last.author : null } };
  }
  const added = a.list.length === b.list.length + 1 && a.mine === b.mine + 1;
  const row = a.list.find(r => r && r.text === op.text && eol(r.quote) === eol(op.matchText));
  return { ok: added && !!row && authored(row.author) && row.solved === false && Array.isArray(row.replies) && row.replies.length === 0, evidence: { commentsBefore: b.list.length, commentsAfter: a.list.length, author: row ? row.author : null } };
}

/**
 * L3, PRIMITIVE_INTEGRAL: the body (a comment is not body text), every OTHER comment (incl. solved state and replies) unchanged. reply: the target keeps every field
 * and its OLD replies as a prefix; add: every comment of the before-state is still there, unchanged (the multiset grew by exactly the new one).
 */
function verifyPreservation(op, before, after) {
  const b = stateOf(before && before.primary, op);
  const a = stateOf(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` } };
  const sameBody = bodyKey(before.primary) === bodyKey(after.primary);
  if (op.intent === 'reply_comment') {
    const sameCount = b.list.length === a.list.length;
    const others = l => JSON.stringify(l.filter(r => !matches(r, op.comment)).map(withSolved));
    const head = r => JSON.stringify([r.id, r.text, r.author, r.userId, r.timeUtc, eol(r.quote), r.solved]);
    const prefix = JSON.stringify(a.row.replies.slice(0, b.row.replies.length)) === JSON.stringify(b.row.replies);
    const sameOthers = others(b.list) === others(a.list);
    const sameTarget = head(b.row) === head(a.row);
    return { ok: sameCount && sameOthers && sameTarget && prefix && sameBody, evidence: { sameCount, otherCommentsUnchanged: sameOthers, targetFieldsUnchanged: sameTarget, oldRepliesKept: prefix, bodyUnchanged: sameBody } };
  }
  const left = [...a.list.map(withSolved)];
  let keptAll = true;
  for (const k of b.list.map(withSolved)) { const i = left.indexOf(k); if (i < 0) { keptAll = false; break; } left.splice(i, 1); }
  return { ok: keptAll && left.length === 1 && sameBody, evidence: { everyOldCommentKeptUnchanged: keptAll, newRows: left.length, bodyUnchanged: sameBody } };
}

function wholeTaskChecks(op, finalObservation) {
  const s = stateOf(finalObservation.primary, op);
  return [{ name: `${op.id}: ${op.intent === 'reply_comment' ? 'the reply' : 'the comment'} "${op.text}" exists once`, ok: s.ok && s.mine === 1 }];
}

// refs: none -- a comment is not a body text unit a later step can name by reference.
function producedText() {
  return null;
}

module.exports = { STEP_KEYS, plan, classify, resolve, preDispatchCheck, lower, replyBody, addBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, replyState, addState };
