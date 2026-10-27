'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const comments = require('../../src/adapters/docx/families/comments.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * GH #269c offline acceptance: resolve_comment (ApiComment.SetSolved true|false) through the TARGET
 * path on the fake transport (modelled on docx-w11-page-setup.test.cjs). The intent is UNRELEASED
 * (not on any WRITE allowlist) -- these tests call the adapter directly; the LIVE acceptance is
 * test/live/docx-resolve-comment-product.live.cjs (zero runs it).
 */

const paragraphs = () => [{ text: 'Body heading', level: 1, props: {} }, { text: 'Body first.', props: {} }];
const row = (id, text, author, quote, solved = false, replies = []) => ({ id, text, author, userId: `u-${author}`, timeUtc: 1760000000000 + Number(id), quote, solved, replies });
const fixture = () => [row('1', 'Please check', 'Alice', 'Body first.'), row('2', 'Typo here', 'Bob', 'Body heading'), row('3', 'Typo here', 'Alice', 'Body heading', true, [{ text: 'done', author: 'Bob', userId: 'u-Bob' }])];
const store = (cs = fixture()) => createStore(paragraphs(), { doc: { comments: cs } });

async function run(s, t, script = {}) {
  let n = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(t));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'g269c', steps });
const rc = (text, author, quote, extra = {}) => ({ intent: 'resolve_comment', text, author, quote, ...extra });
const status = (x, i = 0) => x.appliedOperations[i].status;
const solvedOf = (s, id) => s.durableDoc.comments.find(c => c.id === id).solved;
const snapshot = s => JSON.stringify([s.durable.map(p => p.text), s.durableDoc.comments]);

test('RESOLVE: an open comment is resolved (APPLIED, 1 write), verified L1/L2/L3, then the SAME task again is SATISFIED with ZERO writes', async () => {
  const s = store();
  const t = task(rc('Please check', 'Alice', 'Body first.'));
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED);
  assert.deepEqual([first.m.writes, first.m.commands], [1, 1]);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.primitiveVerification));
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS');
  assert.equal(solvedOf(s, '1'), true);
  assert.equal(solvedOf(s, '2'), false, 'a sibling comment is untouched');
  const second = await run(s, t);
  assert.equal(second.result.ok, true);
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
});

test('REOPEN: resolved:false reopens a resolved comment; resolved:true on a resolved one is SATISFIED (no write)', async () => {
  const s = store();
  const open = await run(s, task(rc('Typo here', 'Alice', 'Body heading', { resolved: false })));
  assert.equal(status(open.x), OPERATION_STATUS.APPLIED);
  assert.equal(solvedOf(s, '3'), false);
  const s2 = store();
  const already = await run(s2, task(rc('Typo here', 'Alice', 'Body heading', { resolved: true })));
  assert.equal(status(already.x), OPERATION_STATUS.SATISFIED);
  assert.equal(already.m.writes, 0);
});

test('IDENTIFICATION: text alone is ambiguous; text + author + quote together select exactly one (the reply and the other author are untouched)', async () => {
  const s = store();
  const r = await run(s, task(rc('Typo here', 'Bob', 'Body heading')));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED);
  assert.equal(solvedOf(s, '2'), true);
  assert.equal(solvedOf(s, '3'), true, 'the same text by another author keeps its own state');
  assert.deepEqual(s.durableDoc.comments.find(c => c.id === '3').replies, [{ text: 'done', author: 'Bob', userId: 'u-Bob' }]);
});

test('FAIL-CLOSED before dispatch: no match, several matches, unreadable comments, unobservable solved state -- 0 writes, named reason', async () => {
  const dup = store([row('1', 'Same', 'Alice', 'q'), row('2', 'Same', 'Alice', 'q')]);
  const ambiguous = await run(dup, task(rc('Same', 'Alice', 'q')));
  assert.notEqual(status(ambiguous.x), OPERATION_STATUS.APPLIED);
  assert.equal(ambiguous.m.writes, 0);
  const none = store();
  const missing = await run(none, task(rc('Nope', 'Alice', 'Body first.')));
  assert.notEqual(status(missing.x), OPERATION_STATUS.APPLIED);
  assert.equal(missing.m.writes, 0);
  const wrongQuote = await run(store(), task(rc('Please check', 'Alice', 'Body heading')));
  assert.notEqual(status(wrongQuote.x), OPERATION_STATUS.APPLIED);
  assert.equal(wrongQuote.m.writes, 0);
  const unread = await run(store(), task(rc('Please check', 'Alice', 'Body first.')), { commentsUnreadable: true });
  assert.notEqual(status(unread.x), OPERATION_STATUS.APPLIED);
  assert.equal(unread.m.writes, 0);
  const unobs = store([{ ...row('1', 'Please check', 'Alice', 'Body first.'), solved: 'UNOBSERVABLE:no-such-getter' }]);
  const u = await run(unobs, task(rc('Please check', 'Alice', 'Body first.')));
  assert.notEqual(status(u.x), OPERATION_STATUS.APPLIED);
  assert.equal(u.m.writes, 0);
});

test('PLAN refusals: unknown field (selector, newText, author override), missing identity, non-boolean resolved -- before any dispatch', () => {
  const bad = [
    { ...rc('a', 'b', 'c'), selector: { text: 'x' } }, { ...rc('a', 'b', 'c'), newAuthor: 'x' }, { ...rc('a', 'b', 'c'), userId: 'x' },
    { intent: 'resolve_comment', author: 'b', quote: 'c' }, { intent: 'resolve_comment', text: 'a', quote: 'c' }, { intent: 'resolve_comment', text: 'a', author: 'b' },
    rc('a', 'b', 'c', { resolved: 'yes' }), rc('', 'b', 'c'),
  ];
  for (const step of bad) assert.equal(planOperations(task(step)).ok, false, JSON.stringify(step));
  assert.equal(planOperations(task(rc('a', 'b', ''))).ok, true, "quote '' = a comment with no quoted range");
});

test('NEGATIVE CONTROLS: the write not landing, a sibling comment changing, the author being forged, the wrong direction are each caught by L1/L3', async () => {
  for (const fault of ['ignore', 'rc-other-comment', 'rc-author-forged', 'rc-wrong-direction']) {
    const s = store();
    const r = await run(s, task(rc('Please check', 'Alice', 'Body first.')), { fault });
    assert.notEqual(r.result.ok, true, `${fault} must not end green`);
  }
});

test('NO FORGERY SURFACE: the command body calls SetSolved and NO author/time/text/position setter, Delete, AddComment or AddReply', () => {
  const body = comments.commandBody({ comment: { text: 'a', author: 'b', quote: 'c' }, resolved: true });
  assert.match(body, /\.SetSolved\(/);
  for (const banned of ['SetAuthorName', 'SetUserId', 'SetTime', 'SetTimeUTC', 'SetText', 'SetPosition', 'Delete', 'AddComment', 'AddReply']) assert.equal(body.includes(banned), false, banned);
});

test('RELEASED (g269ev1 PASS): resolve_comment routed in the adapter AND on both write allowlists (positive control of the former UNRELEASED check)', () => {
  const W = require('../../src/server/write-tools.cjs');
  const T = require('../../src/server/tools.cjs');
  for (const i of ['resolve_comment']) {
    assert.ok(W.DOCX_WRITE_ALLOWLIST.includes(i) && T.ACCEPTED_DOCX_WRITE_ALLOWLIST.includes(i), i);
  }
});

test('the pre-existing state is not mutated by a refused run (snapshot equal)', async () => {
  const s = store();
  const before = snapshot(s);
  await run(s, task(rc('Nope', 'Alice', 'Body first.')));
  assert.equal(snapshot(s), before);
});

test('QUOTE TRAILING CR/LF (g269cv1): a paragraph comment is found with OR without the quote\'s trailing CR/LF, in the adapter match and in the command body; an INNER CR/LF and any other difference stay exact', async () => {
  const crlf = store([row('1', 'Para note', 'Alice', 'Body first.\r\n'), row('2', 'Other', 'Bob', 'a\r\nb')]);
  const r1 = await run(crlf, task(rc('Para note', 'Alice', 'Body first.')));
  assert.equal(status(r1.x), OPERATION_STATUS.APPLIED, JSON.stringify(r1.x.appliedOperations.map(e => e.evidence)));
  assert.equal(solvedOf(crlf, '1'), true);
  const plain = store([row('1', 'Para note', 'Alice', 'Body first.')]);
  const r2 = await run(plain, task(rc('Para note', 'Alice', 'Body first.\r\n')));
  assert.equal(status(r2.x), OPERATION_STATUS.APPLIED);
  const inner = await run(store([row('1', 'Other', 'Bob', 'a\r\nb')]), task(rc('Other', 'Bob', 'a b')));
  assert.notEqual(status(inner.x), OPERATION_STATUS.APPLIED, 'an inner CR/LF is NOT trimmed');
  assert.equal(inner.m.writes, 0);
  const text = await run(store([row('1', 'Para note', 'Alice', 'Body first.\r\n')]), task(rc('Para note', 'Alice', 'Body second.')));
  assert.notEqual(status(text.x), OPERATION_STATUS.APPLIED);
  // the REAL command body, run against a mock Api: the CR/LF-ended quote of the editing session matches the CR/LF-less quote of the step
  const calls = [];
  const mk = (t, a, q) => ({ GetText: () => t, GetAuthorName: () => a, GetQuoteText: () => q, SetSolved: v => calls.push(v) });
  const run1 = body => new Function('Api', body)({ GetDocument: () => ({ GetAllComments: () => [mk('Para note', 'Alice', 'Body first.\r\n'), mk('Other', 'Bob', 'a\r\nb')] }) });
  run1(comments.commandBody({ comment: { text: 'Para note', author: 'Alice', quote: 'Body first.' }, resolved: true }));
  run1(comments.commandBody({ comment: { text: 'Other', author: 'Bob', quote: 'a b' }, resolved: true }));
  assert.deepEqual(calls, [true], 'only the trailing-CR/LF comment matched; the inner-CR/LF one did not');
});
