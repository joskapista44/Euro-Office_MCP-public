'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const cw = require('../../src/adapters/docx/families/comment-write.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * GH #269d offline acceptance: reply_comment (AddReply) and add_comment_on_match (ApiRange.AddComment) on the fake transport. UNRELEASED (no allowlist); the LIVE acceptance
 * is test/live/docx-reply-add-comment-product.live.cjs (zero runs it). AUTHOR = the editor SESSION identity (owner decision): the task has no author field. The product does not READ the identity (NEW_DIRECT_INTERNAL=0), so
 * 'saved author == session identity' is a LIVE measurement; offline the fake stamps its own constant and the contract checks a readable author + everything else.
 */

const SESSION = 'EO Session Agent';
const paragraphs = () => [{ text: 'Alpha ZQ one tail', props: {} }, { text: 'Beta ZQ two tail', props: {} }, { text: 'Gamma unique', props: {} }];
const row = (id, text, author, quote, solved = false, replies = []) => ({ id, text, author, userId: `u-${author}`, timeUtc: 1760000000000 + Number(id), quote, solved, replies });
const fixture = () => [row('1', 'Please check', 'Alice', 'Gamma unique'), row('2', 'Other', 'Bob', 'Alpha ZQ one tail', true, [{ text: 'ok', author: 'Alice', userId: 'u-Alice' }])];
const store = (cs = fixture()) => createStore(paragraphs(), { doc: { comments: cs } });

async function run(s, t, script = {}) {
  let n = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(t));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'g269d', steps });
const reply = (target, text, extra = {}) => ({ intent: 'reply_comment', target, text, ...extra });
const add = (matchText, text, extra = {}) => ({ intent: 'add_comment_on_match', matchText, text, ...extra });
const T1 = { text: 'Please check', author: 'Alice', quote: 'Gamma unique' };
const status = (x, i = 0) => x.appliedOperations[i].status;
const comments = s => s.durableDoc.comments;
const snapshot = s => JSON.stringify([s.durable.map(p => p.text), s.durableDoc.comments]);

test('REPLY: a reply lands (APPLIED, 1 write) by the SESSION author, verified L1/L2/L3; the SAME task again is SATISFIED with ZERO writes', async () => {
  const s = store();
  const t = task(reply(T1, 'Done, thanks'));
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED);
  assert.deepEqual([first.m.writes, first.m.commands], [1, 1]);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS') && first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify([first.x.primitiveVerification, first.x.preservationVerification]));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS');
  assert.deepEqual(comments(s)[0].replies, [{ text: 'Done, thanks', author: SESSION, userId: 'eo-session-1' }]);
  assert.deepEqual(comments(s)[1].replies, [{ text: 'ok', author: 'Alice', userId: 'u-Alice' }], 'the other comment keeps its reply');
  const second = await run(s, t);
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
});

test('ADD: a comment on a unique match lands by the SESSION author with the match as its quote; retry SATISFIED, zero writes', async () => {
  const s = store();
  const t = task(add('Gamma unique', 'New note'));
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED);
  assert.deepEqual([first.m.writes, first.m.commands], [1, 1]);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS') && first.x.preservationVerification.every(v => v.verdict === 'PASS'));
  const added = comments(s).find(c => c.text === 'New note');
  assert.deepEqual([added.author, added.quote, added.solved, added.replies], [SESSION, 'Gamma unique', false, []]);
  assert.equal(comments(s).length, 3);
  const second = await run(s, t);
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([second.m.writes, second.m.commands], [0, 0]);
});

test('ADD ambiguity: a matchText found twice is refused pre-dispatch without occurrence (named count); occurrence 2 selects the second; out of range is refused', async () => {
  const amb = await run(store(), task(add('ZQ', 'On a twin')));
  assert.notEqual(status(amb.x), OPERATION_STATUS.APPLIED);
  assert.equal(amb.m.writes, 0);
  const s = store();
  const ok = await run(s, task(add('ZQ', 'On the second', { occurrence: 2 })));
  assert.equal(status(ok.x), OPERATION_STATUS.APPLIED, JSON.stringify(ok.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(comments(s).find(c => c.text === 'On the second').quote, 'ZQ');
  const far = await run(store(), task(add('ZQ', 'Nope', { occurrence: 3 })));
  assert.notEqual(status(far.x), OPERATION_STATUS.APPLIED);
  assert.equal(far.m.writes, 0);
  const none = await run(store(), task(add('absent text', 'Nope')));
  assert.notEqual(status(none.x), OPERATION_STATUS.APPLIED);
  assert.equal(none.m.writes, 0);
});

test('FAIL-CLOSED before dispatch: reply target missing / ambiguous / wrong quote, unreadable comments - 0 writes', async () => {
  const missing = await run(store(), task(reply({ ...T1, text: 'Nope' }, 'x')));
  const wrongQuote = await run(store(), task(reply({ ...T1, quote: 'Alpha ZQ one tail' }, 'x')));
  const dup = await run(store([row('1', 'Same', 'Alice', 'q'), row('2', 'Same', 'Alice', 'q')]), task(reply({ text: 'Same', author: 'Alice', quote: 'q' }, 'x')));
  const unread = await run(store(), task(reply(T1, 'x')), { commentsUnreadable: true });
  for (const r of [missing, wrongQuote, dup, unread]) { assert.notEqual(status(r.x), OPERATION_STATUS.APPLIED); assert.equal(r.m.writes, 0); }
});

test('PLAN refusals: NO author field (the owner decision), unknown fields, malformed target, empty text, bad occurrence - before any dispatch', () => {
  const bad = [
    reply(T1, 'x', { author: 'Mallory' }), reply(T1, 'x', { userId: 'u' }), add('Gamma unique', 'x', { author: 'Mallory' }), add('Gamma unique', 'x', { selector: { text: 'y' } }),
    reply({ text: 'a', author: 'b' }, 'x'), reply({ ...T1, extra: 1 }, 'x'), reply('Please check', 'x'), reply(T1, ''), add('', 'x'), add('Gamma unique', ''),
    add('Gamma unique', 'x', { occurrence: 0 }), add('Gamma unique', 'x', { occurrence: 1.5 }),
  ];
  for (const step of bad) assert.equal(planOperations(task(step)).ok, false, JSON.stringify(step));
  assert.equal(planOperations(task(reply({ ...T1, quote: '' }, 'x'))).ok, true);
});

test('NEGATIVE CONTROLS: nothing landing, an EMPTY author, a duplicated write, a sibling comment changing are each caught', async () => {
  for (const [step, fault] of [[reply(T1, 'r'), 'ignore'], [reply(T1, 'r'), 'cw-empty-author'], [reply(T1, 'r'), 'cw-duplicate'], [reply(T1, 'r'), 'cw-other-comment'],
    [add('Gamma unique', 'a'), 'ignore'], [add('Gamma unique', 'a'), 'cw-empty-author'], [add('Gamma unique', 'a'), 'cw-duplicate'], [add('Gamma unique', 'a'), 'cw-other-comment']]) {
    const r = await run(store(), task(step), { fault });
    assert.notEqual(r.result.ok, true, `${step.intent} + ${fault} must not end green`);
  }
});

test('NO AUTHOR SURFACE: AddReply / AddComment get exactly ONE argument (the text); no author/time/text/position setter, no Delete, no SetSolved, in either command body', () => {
  const bodies = [cw.replyBody({ comment: { text: 'a', author: 'b', quote: 'c' }, text: 'x' }), cw.addBody({ matchText: 'm', occurrence: 1, total: 1, text: 'x' })];
  assert.match(bodies[0], /\.AddReply\(m\.text\)/);
  assert.match(bodies[1], /\.AddComment\(m\.text\)/);
  for (const b of bodies) for (const banned of ['SetAuthorName', 'SetUserId', 'SetTime', 'SetTimeUTC', 'SetText', 'SetPosition', 'Delete', 'SetSolved']) assert.equal(b.includes(banned), false, banned);
});

test('RELEASED (g269ev1 PASS): reply_comment, add_comment_on_match routed in the adapter AND on both write allowlists (positive control of the former UNRELEASED check)', () => {
  const W = require('../../src/server/write-tools.cjs');
  const T = require('../../src/server/tools.cjs');
  for (const i of ['reply_comment', 'add_comment_on_match']) {
    assert.ok(W.DOCX_WRITE_ALLOWLIST.includes(i) && T.ACCEPTED_DOCX_WRITE_ALLOWLIST.includes(i), i);
  }
});

test('a refused run does not mutate the document (snapshot equal)', async () => {
  const s = store();
  const before = snapshot(s);
  await run(s, task(add('ZQ', 'ambiguous')));
  await run(s, task(reply({ ...T1, text: 'Nope' }, 'x')));
  assert.equal(snapshot(s), before);
});

test('QUOTE TRAILING CR/LF (g269cv1): reply_comment finds the target with OR without the quote\'s trailing CR/LF (adapter + the real command body); add_comment_on_match retries across the same difference', async () => {
  const s = store([row('1', 'Para note', 'Alice', 'Gamma unique\r\n')]);
  const r = await run(s, task(reply({ text: 'Para note', author: 'Alice', quote: 'Gamma unique' }, 'Noted')));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => e.evidence)));
  assert.equal(comments(s)[0].replies.length, 1);
  const retry = await run(s, task(reply({ text: 'Para note', author: 'Alice', quote: 'Gamma unique\r\n' }, 'Noted')));
  assert.equal(status(retry.x), OPERATION_STATUS.SATISFIED);
  const inner = await run(store([row('1', 'Para note', 'Alice', 'a\r\nb')]), task(reply({ text: 'Para note', author: 'Alice', quote: 'a b' }, 'x')));
  assert.notEqual(status(inner.x), OPERATION_STATUS.APPLIED);
  const adds = [];
  const mk = (t, a, q) => ({ GetText: () => t, GetAuthorName: () => a, GetQuoteText: () => q, AddReply: x => adds.push(x) });
  new Function('Api', cw.replyBody({ comment: { text: 'Para note', author: 'Alice', quote: 'Gamma unique' }, text: 'R' }))({ GetDocument: () => ({ GetAllComments: () => [mk('Para note', 'Alice', 'Gamma unique\r\n'), mk('X', 'Bob', 'a\r\nb')] }) });
  assert.deepEqual(adds, ['R']);
  const crlfAdd = store([row('1', 'New note', 'EO Session Agent', 'Gamma unique\r\n')]);
  const again = await run(crlfAdd, task(add('Gamma unique', 'New note')));
  assert.equal(status(again.x), OPERATION_STATUS.SATISFIED, 'the same note on the same quote is already there, CR/LF or not');
});
