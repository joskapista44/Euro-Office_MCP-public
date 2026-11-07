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
const { colName } = require('../../src/adapters/xlsx/a1.cjs');
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

// ---- GH #367 2nd unit: a FIRST comment past the scan geometry dispatches no scan at all (expected=0 skips it) --
// the target cell is checked against the SAME window `commentReads` would use, BEFORE dispatch, instead.

test("GH #367 2nd unit: scanWindowOf() matches the driver's own commentReads formula (xlsx-driver.cjs:1125-1126) -- maxCol = min(16384, max(26, usedCol+10)), maxRow = min(1048576, max(200, usedRow+100), floor(20000/maxCol))", () => {
  assert.deepEqual(fam.scanWindowOf(null), { maxCol: 26, maxRow: 200 }); // no used range at all -> the two floors
  assert.deepEqual(fam.scanWindowOf('A1'), { maxCol: 26, maxRow: 200 });
  assert.deepEqual(fam.scanWindowOf('A1:C3'), { maxCol: 26, maxRow: 200 }); // usedCol=3+10=13 < floor 26; usedRow=3+100=103 < floor 200 -- both floors bind
  assert.deepEqual(fam.scanWindowOf('A1:Z500'), { maxCol: 36, maxRow: 555 }); // usedCol=26+10=36 > floor 26; floor(20000/36)=555 < usedRow+100=600 -- the cell-count ceiling binds
  assert.deepEqual(fam.scanWindowOf('A1:CZ1'), { maxCol: 114, maxRow: 175 }); // usedCol=104+10=114; floor(20000/114)=175 < the 200 row floor -- the cell-count ceiling binds on the row side too, from a WIDE sheet
});

test('GH #367 2nd unit: add_comment on a cell PAST the window (row beyond maxRow) is refused BEFORE dispatch, named, 0 commands -- the "csendes reszlista NEM" guarantee extended to the WRITE side', async () => {
  const store = createStore(book()); // sheet A used = A1:C3 -> window A1:Z200
  const { result, m } = await run(store, task(stepOf({ cell: 'Z201' })));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0, 'no dispatch -- never a ghost write');
  assert.match(JSON.stringify(result.execution), /xlsx-comment-target-outside-scan/);
  assert.match(JSON.stringify(result.execution), /A1:Z200/, 'the evidence names the ACTUAL window, not just the reason');
  assert.equal(commentAt(store, 'Z201'), null);
});

test('GH #367 2nd unit: add_comment on a cell PAST the window (column beyond maxCol) is also refused BEFORE dispatch, named', async () => {
  const store = createStore(book()); // window A1:Z200 -> column AA (27) is one past Z (26)
  const { result, m } = await run(store, task(stepOf({ cell: 'AA1' })));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-comment-target-outside-scan/);
});

test('GH #367 2nd unit: add_comment exactly AT the window boundary (Z200) is NOT refused by this check -- the family behaves exactly as before for every in-window cell', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(stepOf({ cell: 'Z200' })));
  assert.equal(result.ok, true, JSON.stringify(result.execution));
  assert.equal(m.commands, 1);
  assert.deepEqual(commentAt(store, 'Z200'), { cell: 'Z200', text: 'Please check this', author: 'Fake Tester', userId: 'oc_fake_tester', timeUtc: 0, solved: false, replies: [] });
});

test('GH #367 2nd unit: a widened used range widens the window the SAME way -- a far cell that used to be refused becomes reachable once the used range grows past it', async () => {
  const b = book();
  b[0].cells.Z600 = { v: 'widens the used range' }; // pushes used range to A1:Z600 -> maxRow floor(20000/36)=555... let's use a smaller, exact case
  const store = createStore(b);
  const window = fam.scanWindowOf('A1:Z600');
  const { result, m } = await run(store, task(stepOf({ cell: `${colName(window.maxCol)}${window.maxRow}` })));
  assert.equal(result.ok, true, JSON.stringify(result.execution));
  assert.equal(m.commands, 1);
});

test('GH #367 2nd unit: an edit/delete/reply/resolve on an EXISTING comment is unaffected -- the new check only gates add_comment\'s brand-new target, never an already-found cell', async () => {
  const b = book();
  b[0].comments.push({ cell: 'Z500', text: 'already there', author: 'Fake Tester', userId: 'oc_fake_tester', timeUtc: 0, solved: false, replies: [] });
  const store = createStore(b);
  const { result, m } = await run(store, { id: 'gh367', steps: [{ intent: 'edit_comment', sheet: 'A', cell: 'Z500', text: 'edited' }] });
  assert.equal(result.ok, true, JSON.stringify(result.execution));
  assert.equal(m.commands, 1);
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

test("GH #367 2nd unit: the family's local scan-cell ceiling stays in sync with the driver's MAX_COMMENT_SCAN_CELLS (the adapter boundary forbids importing it)", () => {
  const { MAX_COMMENT_SCAN_CELLS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
  const fam = require('../../src/adapters/xlsx/families/comments.cjs');
  // A1:Z500 -> maxCol 36, so the cell-count ceiling binds the row side: floor(ceiling / 36)
  assert.equal(fam.scanWindowOf('A1:Z500').maxRow, Math.floor(MAX_COMMENT_SCAN_CELLS / 36));
});

// GH #367 3rd unit v6 (zero LIVE finding, card an internal branch
// 20261105): on a NON-ACTIVE sheet, AddComment's own return value is a DETACHED copy -- a SetUserId
// on it never reaches the stored comment. `lower()`'s generated body is run DIRECTLY (new Function,
// same pattern as xlsx-b05w-bridge-write-cutover.test.cjs's own body tests) against a FAKE Api that
// models exactly the measured shape: SetUserId only lands while the TARGET sheet is the active one.
test('GH #367 3rd unit v6: add_comment lower() -- SetUserId runs on a fresh re-fetch, with a SetActive+restore fallback, three fake-engine scenarios', () => {
  const SHEET = 'Target';
  const OTHER_SHEET = 'Other';
  const CELL = 'B2';
  const AUTHOR = { id: 'oc1_u', name: 'U Name' };
  const op = { intent: 'add_comment', selector: { sheet: SHEET, cell: CELL }, text: 'hi' };

  function makeFakeApi(setUserIdLands) {
    let activeName = OTHER_SHEET;
    let stored = null; // becomes { userId } once a SetUserId call "lands"
    let commentExists = false;
    const view = () => ({
      GetUserId: () => (stored ? stored.userId : null),
      SetUserId: id => { if (setUserIdLands(activeName)) stored = { userId: id }; },
    });
    const sheets = {
      [SHEET]: { GetName: () => SHEET, SetActive: () => { activeName = SHEET; },
        GetRange: cell => { assert.equal(cell, CELL); return { GetComment: () => (commentExists ? view() : null), AddComment: () => { commentExists = true; return view(); } }; } },
      [OTHER_SHEET]: { GetName: () => OTHER_SHEET, SetActive: () => { activeName = OTHER_SHEET; } },
    };
    return { api: { GetSheet: name => sheets[name] || null, GetActiveSheet: () => sheets[activeName] }, activeNameOf: () => activeName };
  }

  const runLowered = api => {
    const body = `const __sessionAuthor = ${JSON.stringify(AUTHOR)};\n${fam.lower(op).body}`;
    return new Function('Api', body)(api);
  };

  // (1) a fresh re-fetch alone is enough (SetUserId lands regardless of active sheet) -- the
  // SetActive fallback is never needed, and the active sheet is never touched.
  {
    const { api, activeNameOf } = makeFakeApi(() => true);
    assert.deepEqual(runLowered(api), { ok: true });
    assert.equal(activeNameOf(), OTHER_SHEET);
  }

  // (2) the LIVE-measured shape: SetUserId only lands while the TARGET sheet is active -- the
  // fallback must SetActive, retry on a fresh object, then restore the ORIGINAL active sheet.
  {
    const { api, activeNameOf } = makeFakeApi(active => active === SHEET);
    assert.deepEqual(runLowered(api), { ok: true });
    assert.equal(activeNameOf(), OTHER_SHEET, 'the original active sheet is restored, not left on the target');
  }

  // (3) even the SetActive fallback does not make it land (a worse engine state than measured) --
  // a NAMED failure, never a silent { ok: true } over an unproven identity.
  {
    const { api, activeNameOf } = makeFakeApi(() => false);
    assert.deepEqual(runLowered(api), { ok: false, stage: 'userid-never-landed' });
    assert.equal(activeNameOf(), OTHER_SHEET, 'the active sheet is still restored even on this failure path');
  }
});
