'use strict';

/**
 * GH #350 (reply_to_comment) + GH #351 (set_comment_resolved): a reply by the SESSION user (public ApiComment.AddReply) and the resolved
 * flag (public ApiComment.SetSolved) on a comment loaded from a saved file. The LIVE design probe (docs/evidence/gh350-gh351-probe-live-
 * 20261030.json) measured: both persist through save + a fresh reopen, the root author and the other replies kept; a SECOND AddReply with
 * the same text makes a DUPLICATE (so a retry is decided from the fresh read); the read lists the replies in REVERSED order after a reopen
 * (so the verifier compares them as a multiset, never by position); any author's comment can be replied to and resolved.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/comments.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const ME = { name: 'Fake Tester', userId: 'oc_fake_tester' };
const mk = (cell, over = {}) => ({ cell, text: `seed ${cell}`, author: ME.name, userId: ME.userId, timeUtc: 0, solved: false, replies: [], ...over });
const OTHER = { text: 'a reply', author: 'Excel Author', userId: 'oc_other', timeUtc: 0 };
const book = () => [{
  name: 'A',
  cells: { A1: { v: 'one' }, B2: { v: 'commented value' }, C3: { v: 7 }, D4: { v: 'foreign' } },
  comments: [
    mk('B2'),
    mk('C3', { solved: true, replies: [OTHER] }),
    mk('D4', { author: 'Excel Author', userId: 'oc_other' }),
    mk('F6'),
  ],
}, { name: 'Other', cells: { A1: { v: 'other' } }, comments: [mk('A1', { text: 'on another sheet' })] }];
const replyStep = (extra = {}) => ({ intent: 'reply_to_comment', sheet: 'A', cell: 'B2', text: 'my reply', ...extra });
const resolveStep = (extra = {}) => ({ intent: 'set_comment_resolved', sheet: 'A', cell: 'B2', resolved: true, ...extra });
const task = step => ({ id: 'gh350', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const commentAt = (store, cell, sheet = 0) => (store.durable[sheet].comments || []).find(c => c.cell === cell) || null;
const snapshot = store => JSON.parse(JSON.stringify(store.durable));
const mapped = (r, t) => W.mapExecutionResult(r.result, { fileId: 'x', task: t });

test('schema: reply takes sheet + cell + text, resolve takes sheet + cell + a boolean; an identity key, a range, a missing field refused', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(replyStep()).ok, true, JSON.stringify(v(replyStep())));
  assert.equal(v(resolveStep()).ok, true, JSON.stringify(v(resolveStep())));
  assert.equal(v(resolveStep({ resolved: false })).ok, true);
  for (const k of ['author', 'userId', 'uid']) assert.equal(v(replyStep({ [k]: 'Mallory' })).ok, false, `reply: ${k}`);
  assert.equal(v(resolveStep({ resolved: 'yes' })).ok, false, 'resolved must be a boolean');
  assert.equal(v({ intent: 'set_comment_resolved', sheet: 'A', cell: 'B2' }).ok, false, 'resolved is required');
  assert.equal(v(resolveStep({ text: 'x' })).ok, false, 'resolve carries no text');
  assert.equal(v(replyStep({ cell: 'B2:C3' })).ok, false);
  assert.equal(v(replyStep({ text: '' })).ok, false);
  assert.equal(fam.plan(replyStep({ text: 'bell \u0007' }), 0).ok, false);
  assert.equal(fam.plan(replyStep({ text: 'two\nlines' }), 0).ok, true);
});

test('REPLY: APPLIED -- one reply by the SESSION user appended; the root, the solved flag, the other comments and every cell unchanged; retry SATISFIED, named, 0 commands', async () => {
  const store = createStore(book());
  const before = snapshot(store);
  const { result, m } = await run(store, task(replyStep()));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.deepEqual(commentAt(store, 'B2'), mk('B2', { replies: [{ text: 'my reply', author: ME.name, userId: ME.userId, timeUtc: 0 }] }));
  const after = snapshot(store);
  assert.deepEqual(after[0].cells, before[0].cells);
  assert.deepEqual(after[1], before[1]);
  for (const c of ['C3', 'D4', 'F6']) assert.deepEqual(commentAt(store, c), before[0].comments.find(x => x.cell === c), c);
  const retry = await run(store, task(replyStep()));
  assert.equal(retry.result.ok, true);
  assert.deepEqual([retry.m.writes, retry.m.commands], [0, 0], 'a retry never makes a duplicate (the LIVE probe: AddReply twice = two replies)');
  assert.equal(mapped(retry, task(replyStep())).operations[0].observation, 'reply-already-present');
  assert.equal(commentAt(store, 'B2').replies.length, 1);
});

test('REPLY: on a solved comment with another author\'s reply (stays solved, both replies kept) and on ANOTHER author\'s comment (allowed)', async () => {
  for (const cell of ['C3', 'D4']) {
    const store = createStore(book());
    const was = commentAt(store, cell);
    const { result } = await run(store, task(replyStep({ cell })));
    assert.equal(result.ok, true, cell);
    const now = commentAt(store, cell);
    assert.deepEqual([now.text, now.author, now.userId, now.solved], [was.text, was.author, was.userId, was.solved], cell);
    assert.deepEqual(now.replies, [...was.replies, { text: 'my reply', author: ME.name, userId: ME.userId, timeUtc: 0 }], cell);
  }
});

test('RESOLVE / REOPEN: APPLIED both ways; text, author, replies kept; retry SATISFIED, named, 0 commands', async () => {
  const store = createStore(book());
  const r1 = await run(store, task(resolveStep()));
  assert.equal(r1.result.ok, true, JSON.stringify(r1.result.execution.primitiveVerification));
  assert.equal(commentAt(store, 'B2').solved, true);
  const r1b = await run(store, task(resolveStep()));
  assert.deepEqual([r1b.result.ok, r1b.m.writes, r1b.m.commands], [true, 0, 0]);
  assert.equal(mapped(r1b, task(resolveStep())).operations[0].observation, 'comment-already-in-that-state');
  const r2 = await run(store, task(resolveStep({ cell: 'C3', resolved: false })));
  assert.equal(r2.result.ok, true);
  assert.deepEqual(commentAt(store, 'C3'), mk('C3', { solved: false, replies: [OTHER] }));
  const r3 = await run(store, task(resolveStep({ cell: 'D4' })));
  assert.equal(r3.result.ok, true, 'another author\'s thread can be resolved (the flag is the thread\'s shared state)');
});

test('refused before the write, named: an empty cell (reply / resolve), a missing sheet, a session without an identity', async () => {
  for (const step of [replyStep({ cell: 'A1' }), resolveStep({ cell: 'A1' })]) {
    const store = createStore(book());
    const before = snapshot(store);
    const { result, m } = await run(store, task(step));
    assert.equal(result.ok, false, step.intent);
    assert.equal(m.commands, 0, step.intent);
    assert.match(JSON.stringify(result.execution.appliedOperations[0].evidence), /xlsx-comment-cell-empty/, step.intent);
    assert.deepEqual(snapshot(store), before, step.intent);
  }
  const ns = await run(createStore(book()), task(replyStep({ sheet: 'Nope' })));
  assert.equal(ns.result.ok, false);
  assert.equal(ns.m.commands, 0);
  const noId = await run(createStore(book()), task(replyStep()), { sessionUser: null });
  assert.equal(noId.result.ok, false, 'no reliable identity -> not dispatched');
  assert.equal(noId.m.commands, 0);
});

test('every broken engine behaviour fails the VERIFY (fake faults), never reported as applied', async () => {
  const cases = [
    [replyStep(), 'reply-wrong-author'], [replyStep({ cell: 'C3' }), 'reply-unsolves'], [replyStep({ cell: 'C3' }), 'reply-drops-thread'],
    [resolveStep({ cell: 'C3', resolved: false }), 'resolve-drops-replies'], [resolveStep(), 'resolve-changes-text'],
    [replyStep(), 'comment-touches-other'], [resolveStep(), 'comment-touches-other'], [replyStep(), 'ignore'], [resolveStep(), 'ignore'], [replyStep(), 'side-effect'],
  ];
  for (const [step, fault] of cases) {
    const { result } = await run(createStore(book()), task(step), { fault });
    assert.equal(result.ok, false, `${step.intent} ${step.cell} with fault ${fault} must fail`);
  }
});

test('the lowered command: the live comment checked, AddReply with the SESSION author (never a step value) / SetSolved(bool); the only step of its task', () => {
  const r = fam.lower(fam.plan(replyStep(), 0).operation);
  assert.equal(r.meta.sessionAuthor, true);
  assert.match(r.body, /GetComment\(\)/);
  assert.match(r.body, /c\.AddReply\("my reply", __sessionAuthor\.name, __sessionAuthor\.id\)/);
  const s = fam.lower(fam.plan(resolveStep({ resolved: false }), 0).operation);
  assert.match(s.body, /c\.SetSolved\(false\)/);
  const two = W.validateWriteRequest(req([replyStep(), { intent: 'write_range', sheet: 'A', range: 'A1', values: [['x']] }]));
  assert.equal(two.ok, false, 'the only step of its task');
});
