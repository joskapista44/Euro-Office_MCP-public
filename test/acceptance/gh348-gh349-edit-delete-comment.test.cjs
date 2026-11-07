'use strict';

/**
 * GH #348 (edit_comment) + GH #349 (delete_comment): one comment of the SESSION user, edited (public ApiComment.SetText) or
 * deleted (public ApiComment.Delete). The LIVE design probe (docs/evidence/gh348-probe-live-20261030.json) measured: SetText keeps the
 * author / uid / solved flag / replies and persists; Delete removes the comment with its replies and nothing else; a second Delete on a
 * stale handle returns true (a retry is decided from the fresh read); both calls ALSO work on another author's comment -> the product
 * refuses a comment that is not the session user's own, before the write (xlsx-comment-foreign-author).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/comments.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const ME = { name: 'Fake Tester', userId: 'oc_fake_tester' };
const mk = (cell, over = {}) => ({ cell, text: `seed ${cell}`, author: ME.name, userId: ME.userId, timeUtc: 0, solved: false, replies: [], ...over });
const REPLY = { text: 'a reply', author: 'Excel Author', userId: 'oc_other', timeUtc: 0 };
const OWN_REPLY = { text: 'my own reply', author: ME.name, userId: ME.userId, timeUtc: 0 };
const LEGACY_REPLY = { text: 'a reply without a uid', author: 'Old Replier', userId: '', timeUtc: 0 };
const book = () => [{
  name: 'A',
  cells: { A1: { v: 'one' }, B2: { v: 'commented value' }, C3: { v: 7 }, D4: { v: 'foreign' }, E5: { v: 'legacy' } },
  comments: [
    mk('B2'),
    mk('C3', { solved: true, replies: [REPLY] }),
    mk('D4', { author: 'Excel Author', userId: 'oc_other' }),
    mk('E5', { author: 'Legacy Note', userId: '' }),
    mk('F6'),
    mk('G7', { solved: true, replies: [OWN_REPLY] }),
    mk('H8', { replies: [LEGACY_REPLY] }),
  ],
}, { name: 'Other', cells: { A1: { v: 'other' } }, comments: [mk('A1', { text: 'on another sheet' })] }];
const editStep = (extra = {}) => ({ intent: 'edit_comment', sheet: 'A', cell: 'B2', text: 'the edited text', ...extra });
const delStep = (extra = {}) => ({ intent: 'delete_comment', sheet: 'A', cell: 'B2', ...extra });
const task = step => ({ id: 'gh348', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const commentAt = (store, cell, sheet = 0) => (store.durable[sheet].comments || []).find(c => c.cell === cell) || null;
const snapshot = store => JSON.parse(JSON.stringify(store.durable));

test('schema: edit takes sheet + cell + text, delete takes sheet + cell only; an identity key, a range, a missing/extra field refused', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(editStep()).ok, true, JSON.stringify(v(editStep())));
  assert.equal(v(delStep()).ok, true, JSON.stringify(v(delStep())));
  for (const k of ['author', 'userId', 'user', 'uid', 'authorName']) {
    assert.equal(v(editStep({ [k]: 'Mallory' })).ok, false, `edit: ${k}`);
    assert.equal(v(delStep({ [k]: 'Mallory' })).ok, false, `delete: ${k}`);
  }
  assert.equal(v(delStep({ text: 'x' })).ok, false, 'delete carries no text');
  assert.equal(v(editStep({ cell: 'B2:C3' })).ok, false);
  assert.equal(v(delStep({ cell: 'B2:C3' })).ok, false);
  assert.equal(v(editStep({ text: '' })).ok, false);
  assert.equal(v(editStep({ text: 'x'.repeat(fam.MAX_TEXT + 1) })).ok, false);
  assert.equal(fam.plan(editStep({ text: 'bell \u0007' }), 0).ok, false);
  assert.equal(fam.plan(editStep({ text: 'two\nlines' }), 0).ok, true);
  assert.equal(fam.plan({ intent: 'edit_comment', sheet: 'A', cell: 'B2' }, 0).ok, false, 'edit needs the text');
  assert.equal(fam.plan({ intent: 'delete_comment', sheet: 'A', cell: 'B2', id: 'x' }, 0).ok, true);
});

test('EDIT: APPLIED -- only the text changes; author, uid, solved, replies, the other comments and every cell unchanged; retry SATISFIED 0 commands', async () => {
  const store = createStore(book());
  const before = snapshot(store);
  const { result, m } = await run(store, task(editStep()));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.deepEqual(commentAt(store, 'B2'), mk('B2', { text: 'the edited text' }));
  const after = snapshot(store);
  after[0].comments = after[0].comments.map(c => (c.cell === 'B2' ? { ...c, text: before[0].comments.find(x => x.cell === 'B2').text } : c));
  assert.deepEqual(after, before, 'nothing but the one text moved');
  const again = await run(store, task(editStep()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('EDIT: a SOLVED comment with a reply keeps its solved flag and its reply (the probe: E3 / E2)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(editStep({ cell: 'C3' })));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual(commentAt(store, 'C3'), mk('C3', { text: 'the edited text', solved: true, replies: [REPLY] }));
});

test('EDIT: an empty cell -> xlsx-comment-cell-empty before the write; ANOTHER author\'s and a legacy (no uid) comment -> xlsx-comment-foreign-author; 0 commands, nothing changes', async () => {
  for (const [cell, reason] of [['A1', 'xlsx-comment-cell-empty'], ['D4', 'xlsx-comment-foreign-author'], ['E5', 'xlsx-comment-foreign-author']]) {
    const store = createStore(book());
    const before = snapshot(store);
    const { result, m } = await run(store, task(editStep({ cell })));
    assert.equal(result.ok, false, cell);
    assert.equal(m.commands, 0, cell);
    assert.match(JSON.stringify(result.execution), new RegExp(reason), cell);
    assert.deepEqual(snapshot(store), before, `${cell}: the store is untouched`);
  }
});

test('EDIT: the same text on ANOTHER author\'s comment is NOT a satisfied retry -- still refused (never claims a comment that is not the session\'s)', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(editStep({ cell: 'D4', text: 'seed D4' })));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-comment-foreign-author/);
});

test('EDIT / DELETE: no session author -> refused before the write; an unobservable comment read fails closed; 0 commands', async () => {
  for (const step of [editStep(), delStep()]) {
    const a = await run(createStore(book()), task(step), { sessionUser: null });
    assert.equal(a.result.ok, false);
    assert.equal(a.m.commands, 0);
    assert.match(JSON.stringify(a.result.execution), /xlsx-comment-author-identity-unavailable/);
    const b = await run(createStore(book()), task(step), { commentsUnobservable: true });
    assert.equal(b.result.ok, false);
    assert.equal(b.m.commands, 0);
    assert.match(JSON.stringify(b.result.execution), /xlsx-comments-unobservable/);
  }
});

for (const fault of ['edit-changes-author', 'edit-drops-replies', 'edit-unsolves', 'comment-touches-other', 'ignore', 'side-effect']) {
  test(`EDIT: a fake write with fault ${fault} fails the verify, not passes it`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task(editStep({ cell: 'C3' })), { fault });
    if (fault === 'ignore') assert.equal(m.commands, 1);
    assert.equal(result.ok, false, fault);
  });
}

test('DELETE: APPLIED -- only the target comment is gone; the cell value, every other comment and the other sheet unchanged; retry SATISFIED 0 commands', async () => {
  const store = createStore(book());
  const before = snapshot(store);
  const { result, m } = await run(store, task(delStep()));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.equal(commentAt(store, 'B2'), null);
  const after = snapshot(store);
  before[0].comments = before[0].comments.filter(c => c.cell !== 'B2');
  assert.deepEqual(after, before, 'nothing but the one comment moved (cells included)');
  assert.equal(store.durable[0].cells.B2.v, 'commented value');
  const again = await run(store, task(delStep()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.match(JSON.stringify(again.result.execution.appliedOperations[0].evidence), /fresh-live-the-cell-holds-no-comment/);
});

test('DELETE: a SOLVED comment whose replies are all the session user\'s OWN is deleted with its thread; the neighbours stay', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(delStep({ cell: 'G7' })));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(commentAt(store, 'G7'), null);
  assert.ok(commentAt(store, 'B2') && commentAt(store, 'F6') && commentAt(store, 'C3'));
});

test('DELETE: a thread holding ANOTHER author\'s reply (or a reply without a uid) is refused before the write -- xlsx-comment-has-foreign-replies, 0 commands, the thread untouched', async () => {
  for (const cell of ['C3', 'H8']) {
    const store = createStore(book());
    const before = snapshot(store);
    const { result, m } = await run(store, task(delStep({ cell })));
    assert.equal(result.ok, false, cell);
    assert.equal(m.commands, 0, cell);
    assert.match(JSON.stringify(result.execution), /xlsx-comment-has-foreign-replies/, cell);
    assert.deepEqual(snapshot(store), before, cell);
  }
});

test('DELETE: a foreign reply that lands AFTER the read is stopped by the command body\'s own guard -- the comment stays, nothing is saved, the run fails', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(delStep({ cell: 'G7' })), { foreignReplyAtDispatch: 'G7' });
  assert.equal(m.commands, 1, 'dispatched: the read had no foreign reply');
  assert.equal(result.ok, false);
  const c = commentAt(store, 'G7');
  assert.ok(c, 'the comment is still there (the delete did not happen; the durable store is the saved file)');
  assert.deepEqual(c.replies, [OWN_REPLY], 'and nothing was saved over it');
});

test('LIMIT (documented, not wanted): with the body guard broken, a late foreign reply is lost WITH the thread and no verify can see it (it lives only in the deleted thread) -- the protection is the pre-write refusal + the body guard', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(delStep({ cell: 'G7' })), { foreignReplyAtDispatch: 'G7', fault: 'delete-drops-foreign-reply' });
  assert.equal(commentAt(store, 'G7'), null);
  assert.equal(result.ok, true, 'the thread is gone, so the verify has nothing left to compare -- this is why the guards sit BEFORE the write');
});

test('DELETE verify, defence in depth: a BEFORE comment that already held a foreign reply fails the primitive verify even though the comment is gone', () => {
  const obs = comments => ({ primary: { format: 'xlsx', complete: true, sheets: [{ name: 'A' }], sessionUser: { id: ME.userId, name: ME.name }, commentReads: { A: { comments } } } });
  const op = fam.plan(delStep(), 0).operation;
  const withForeign = obs([mk('B2', { replies: [REPLY] })]);
  const clean = obs([mk('B2', { replies: [OWN_REPLY] })]);
  const gone = obs([]);
  assert.equal(fam.verifyPrimitive(op, clean, gone).ok, true);
  const bad = fam.verifyPrimitive(op, withForeign, gone);
  assert.equal(bad.ok, false);
  assert.equal(bad.evidence.foreignReplyLost, true);
});

test('DELETE: a cell with no comment is SATISFIED and the receipt NAMES the observation (a mistyped cell is distinguishable from a delete)', async () => {
  const r = await run(createStore(book()), task(delStep({ cell: 'A1' })));
  assert.equal(r.result.ok, true);
  const ev = r.result.execution.appliedOperations[0].evidence;
  assert.equal(ev.observation, 'no-comment-at-cell');
  assert.equal(ev.proof, 'fresh-live-the-cell-holds-no-comment');
  // the CALLER sees it: the MCP receipt's operation carries the observation (not only the executor's internal evidence)
  const mapped = require('../../src/server/write-tools.cjs').mapExecutionResult(r.result, { fileId: 'x', task: task(delStep({ cell: 'A1' })) });
  assert.equal(mapped.operations[0].status, 'SATISFIED');
  assert.equal(mapped.operations[0].observation, 'no-comment-at-cell');
});

test('DELETE: ANOTHER author\'s and a legacy (no uid) comment -> xlsx-comment-foreign-author before the write; a cell with no comment is SATISFIED (0 commands)', async () => {
  for (const cell of ['D4', 'E5']) {
    const store = createStore(book());
    const before = snapshot(store);
    const { result, m } = await run(store, task(delStep({ cell })));
    assert.equal(result.ok, false, cell);
    assert.equal(m.commands, 0, cell);
    assert.match(JSON.stringify(result.execution), /xlsx-comment-foreign-author/, cell);
    assert.deepEqual(snapshot(store), before, cell);
  }
  const empty = await run(createStore(book()), task(delStep({ cell: 'A1' })));
  assert.equal(empty.result.ok, true);
  assert.equal(empty.m.commands, 0);
  assert.equal(empty.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

for (const fault of ['delete-takes-neighbour', 'delete-clears-cell', 'comment-touches-other', 'ignore', 'side-effect']) {
  test(`DELETE: a fake write with fault ${fault} fails the verify, not passes it`, async () => {
    const store = createStore(book());
    const { result } = await run(store, task(delStep()), { fault });
    assert.equal(result.ok, false, fault);
  });
}

test('edit_comment / delete_comment must each be the only step of their task; liveScope asks the comment read and the session user', () => {
  for (const [a, b] of [[editStep(), editStep({ cell: 'F6' })], [delStep(), delStep({ cell: 'F6' })], [editStep(), delStep({ cell: 'F6' })]]) {
    const p = planOperations({ id: 't', steps: [a, b] });
    assert.equal(p.ok, false);
    assert.match(p.detail, /(edit|delete)_comment must be the only step of its task/);
  }
  for (const step of [editStep(), delStep()]) {
    const s = liveScope(task(step));
    assert.deepEqual(s.commentSheets, ['A']);
    assert.equal(s.sessionUser, true);
  }
});

test('the lowered commands: the SESSION uid is checked on the live comment, the action is SetText / Delete, no identity is baked in', () => {
  const e = fam.lower(fam.plan(editStep(), 0).operation);
  assert.match(e.body, /GetComment\(\)/);
  assert.match(e.body, /c\.GetUserId\(\) !== __sessionAuthor\.id/);
  assert.match(e.body, /c\.SetText\("the edited text"\)/);
  assert.doesNotMatch(e.body, /AddComment|Delete\(\)/);
  assert.deepEqual({ family: e.meta.family, intent: e.meta.intent, sessionAuthor: e.meta.sessionAuthor, text: e.meta.text }, { family: 'comments', intent: 'edit_comment', sessionAuthor: true, text: 'the edited text' });
  const d = fam.lower(fam.plan(delStep(), 0).operation);
  assert.match(d.body, /c\.GetUserId\(\) !== __sessionAuthor\.id/);
  assert.match(d.body, /c\.GetReply\(i\)\.GetUserId\(\) !== __sessionAuthor\.id\) return \{ ok: false, stage: 'foreign-replies' \}/);
  assert.doesNotMatch(e.body, /GetReply/, 'an edit never touches the replies');
  assert.match(d.body, /c\.Delete\(\)/);
  assert.doesNotMatch(d.body, /SetText|AddComment/);
  assert.equal(d.meta.text, undefined);
  for (const c of [e, d]) assert.doesNotMatch(c.body, /Fake Tester|oc_fake_tester/);
  // add_comment is unchanged
  assert.match(fam.lower(fam.plan({ intent: 'add_comment', sheet: 'A', cell: 'B2', text: 'x' }, 0).operation).body, /AddComment\("x", __sessionAuthor\.name\)/);
});
