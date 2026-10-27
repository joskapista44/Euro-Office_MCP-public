'use strict';

/**
 * GH #300 K2: add_comment -- one comment on one cell, the author is the AUTHENTICATED session user (owner decision
 * 2026-10-26, rule A): the signed editor config's user, added by the driver, never a step field; no reliable identity ->
 * no comment; a commented cell is refused (a second AddComment is a silent no-op, K1 / K2 LIVE probes); L1 is the K1
 * read; L3 keeps every other comment and every cell.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/comments.cjs');
const { createEuroOfficeXlsxDriver, sessionUserOf } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const OTHER = { cell: 'D4', text: 'an older note', author: 'Excel Author', userId: '', timeUtc: 0, solved: false, replies: [] };
const book = () => [{ name: 'A', cells: { A1: { v: 'x' }, B2: { v: 'commented value' }, C3: { v: 7 } }, comments: [OTHER] }, { name: 'Other', cells: { A1: { v: 'y' } } }];
const stepOf = (extra = {}) => ({ intent: 'add_comment', sheet: 'A', cell: 'B2', text: 'Please check this', ...extra });
const task = step => ({ id: 'gh300-k2', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const commentAt = (store, cell) => (store.durable[0].comments || []).find(c => c.cell === cell) || null;

test('K2 schema: sheet + cell + text accepted; an author / userId / any identity key, a range, an extra field refused', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(stepOf()).ok, true, JSON.stringify(v(stepOf())));
  for (const k of ['author', 'userId', 'user', 'uid', 'authorName']) assert.equal(v(stepOf({ [k]: 'Mallory' })).ok, false, `${k} must be refused`);
  assert.equal(v(stepOf({ cell: 'B2:C3' })).ok, false, 'one cell only');
  assert.equal(v(stepOf({ text: '' })).ok, false, 'empty text');
  assert.equal(v(stepOf({ text: 'x'.repeat(fam.MAX_TEXT + 1) })).ok, false, 'over the ceiling');
  const ctl = fam.plan(stepOf({ text: 'bell \u0007' }), 0);
  assert.equal(ctl.ok, false, 'a control character');
  assert.equal(fam.plan(stepOf({ text: 'two\nlines' }), 0).ok, true, 'a line break is text');
});

test('K2: APPLIED -- the session user is the author (name AND uid), the older comment and every cell unchanged; retry SATISFIED 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(stepOf()));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.deepEqual(commentAt(store, 'B2'), { cell: 'B2', text: 'Please check this', author: 'Fake Tester', userId: 'oc_fake_tester', timeUtc: 0, solved: false, replies: [] });
  assert.deepEqual(commentAt(store, 'D4'), OTHER);
  assert.equal(store.durable[0].cells.B2.v, 'commented value');
  const again = await run(store, task(stepOf()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('K2: no session author -> refused before the write, 0 commands (no reliable identity -> no comment)', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(stepOf()), { sessionUser: null });
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-comment-author-identity-unavailable/);
  assert.equal(commentAt(store, 'B2'), null);
});

test('K2: a commented cell is refused before the write -- also the SAME text by ANOTHER author (never re-attributed)', async () => {
  for (const existing of [{ ...OTHER, cell: 'B2', text: 'someone else' }, { ...OTHER, cell: 'B2', text: 'Please check this' }]) {
    const b = book(); b[0].comments.push(existing);
    const store = createStore(b);
    const { result, m } = await run(store, task(stepOf()));
    assert.equal(result.ok, false);
    assert.equal(m.commands, 0);
    assert.match(JSON.stringify(result.execution), /xlsx-comment-cell-occupied/);
    assert.deepEqual(commentAt(store, 'B2'), existing, 'the existing comment and its author stay');
  }
});

test('K2: an unobservable comment read fails closed before the write, 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(stepOf()), { commentsUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-comments-unobservable/);
});

for (const fault of ['comment-wrong-author', 'comment-no-user-id', 'comment-touches-other', 'ignore', 'side-effect']) {
  test(`K2: a fake write with fault ${fault} fails the verify, not passes it`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task(stepOf()), { fault });
    assert.equal(m.commands, 1, 'the write was dispatched -- this is the VERIFY failing, not a refusal');
    assert.equal(result.ok, false);
  });
}

test('K2: add_comment must be the only step of its task; liveScope asks the comment read and the session user', () => {
  const p = planOperations({ id: 't', steps: [stepOf(), stepOf({ cell: 'C3' })] });
  assert.equal(p.ok, false);
  assert.match(p.detail, /only step of its task/);
  const s = liveScope(task(stepOf()));
  assert.deepEqual(s.commentSheets, ['A']);
  assert.equal(s.sessionUser, true);
});

test('K2: the lowered command takes the author ONLY from the driver\'s __sessionAuthor -- the step carries none', () => {
  const plan = fam.plan(stepOf(), 0);
  assert.equal(plan.ok, true, plan.detail);
  const cmd = fam.lower(plan.operation);
  assert.match(cmd.body, /AddComment\("Please check this", __sessionAuthor\.name\)/);
  assert.match(cmd.body, /SetUserId\(__sessionAuthor\.id\)/);
  assert.equal(cmd.meta.sessionAuthor, true);
  assert.doesNotMatch(cmd.body, /Fake Tester|oc_fake_tester/);
});

test('K2 driver: the session author is the signed config\'s user (id + name, trimmed), null when either is missing', () => {
  assert.deepEqual(sessionUserOf({ editorConfig: { user: { id: ' oc1_u ', name: ' U Name ', image: 'x' } } }), { id: 'oc1_u', name: 'U Name' });
  for (const bad of [null, {}, { editorConfig: {} }, { editorConfig: { user: { id: 'u' } } }, { editorConfig: { user: { id: '', name: 'n' } } }, { editorConfig: { user: { id: 7, name: 'n' } } }]) {
    assert.equal(sessionUserOf(bad), null, JSON.stringify(bad));
  }
});

test('K2 driver: dispatch refuses a sessionAuthor command without a session identity, before any editor call', async () => {
  const d = createEuroOfficeXlsxDriver({ url: 'http://127.0.0.1:1', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  const r = await d.dispatch(null, { kind: 'callCommand', body: 'return { ok: true };', meta: { family: 'comments', sessionAuthor: true } });
  assert.deepEqual(r, { ok: false, dispatched: false, outcome: 'comment-author-identity-unavailable' });
});
