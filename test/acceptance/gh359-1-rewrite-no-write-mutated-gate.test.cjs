'use strict';

const { HANDSHAKE_BODY } = require('../../src/core/runtime-identity.cjs');
const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');

const drv = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * GH #359 (1) 1/2: a `rewrite_paragraph` dispatch that never wrote (`commandBody`'s named no-write
 * reasons, docx-driver-rewrite-named-reason-vm.test.cjs) must not demand a save-ack at close() --
 * the MULTIRUN shape, measured live on a production Document Server, had the write-free dispatch
 * marked `state.mutated = true` unconditionally, so close() reported `mutated-without-observed-
 * save` and the cleanup FAILED even though nothing was ever written.
 *
 * This harness is `test/acceptance/m7-docx-save-ack-driver.test.cjs`'s own fake Playwright
 * environment, DUPLICATED here (that file's own header explains why: a module-local env is not
 * importable across a separate unit's test file) with ONE addition: the command branch of the fake
 * frame's `evaluate()` returns a scripted `result` instead of a hardcoded `null`, so a test can
 * choose what `commandBody` is claimed to have returned without running any real JS. m7's own write()
 * helper and its existing tests are untouched (they still dispatch with no `meta` at all, so
 * `rewroteNothing` in docx-driver.cjs's dispatch() is unreachable for them and `state.mutated` is set
 * exactly as before -- this file's first test is the explicit proof of that).
 */

const frame = msg => '42' + JSON.stringify(['message', msg]);
const auth = () => frame({ type: 'auth', result: 1 });
const isSaveLock = sync => frame({ type: 'isSaveLock', syncChangesIndex: sync });
const saveLock = granted => frame({ type: 'saveLock', saveLock: !granted });
const saveChanges = (start, end) => frame({ type: 'saveChanges', changes: '[]', startSaveChanges: start, endSaveChanges: end });

const DS_URL = 'wss://ds.example/9.3.4-x/doc/42/c/?shardkey=KEY1&EIO=4&transport=websocket';
const CONFIG = { documentServerUrl: 'https://ds.example/', document: { key: 'KEY1', url: 'https://ds.example/doc.docx' } };

function fakeWs(url) {
  return { url: () => url, handlers: {}, on(ev, cb) { this.handlers[ev] = cb; }, emit(ev, payload) { this.handlers[ev] && this.handlers[ev]({ payload }); } };
}

function fakeDriverEnv({ canSave = false } = {}) {
  const log = [];
  const ctl = { canSave, log, wsHandler: null, nextResult: null };
  const frameObj = {
    url: () => 'https://ds.example/9.3.4-x/web-apps/apps/documenteditor/main/index.html',
    evaluate: async (fn, arg) => {
      if (arg && arg.body === HANDSHAKE_BODY) { log.push('handshake'); return { ok: true, outcome: 'command-completed', result: null }; }
      // the ONLY change from m7-docx-save-ack-driver.test.cjs's own copy of this fake: a
      // test-scripted result (ctl.nextResult) instead of a hardcoded null, so the dispatch()
      // mutated-gate (keyed on command.meta.family/action and this exact string) can be driven.
      if (arg && arg.body) { log.push('command'); return { ok: true, outcome: 'command-completed', result: ctl.nextResult }; }
      const src = String(fn);
      if (src.includes('asc_isDocumentCanSave')) { log.push('canSave'); return { ok: true, canSave: ctl.canSave }; }
      return true; // the open condition: a non-empty live projection
    },
  };
  const page = {
    on(ev, cb) { if (ev === 'websocket') ctl.wsHandler = cb; },
    goto: async () => {}, waitForFunction: async () => {}, waitForTimeout: async () => {},
    frames: () => [frameObj],
    evaluate: async fn => {
      if (String(fn).includes('requestClose')) { log.push('requestClose'); if (ctl.onRequestClose) ctl.onRequestClose(); }
      return { ok: true, detail: 'requestClose-called' };
    },
  };
  const browser = { newPage: async () => page, close: async () => { log.push('browser-close'); if (ctl.ws && ctl.ws.handlers.close) ctl.ws.handlers.close(); } };
  return { ctl, loadPlaywright: () => ({ ok: true, pw: { chromium: { launch: async () => browser } } }) };
}

const opened = [];
after(async () => { for (const d of opened) await d.close().catch(() => {}); });

async function openFakeDriver(env, opts = {}) {
  const realFetch = globalThis.fetch;
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: CONFIG } }) });
  try {
    const d = drv.createEuroOfficeDocxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: env.loadPlaywright, pollMs: 5,
      closeAckTimeoutMs: 200, forceSaveAckTimeoutMs: 300, forceSaveConfirmTimeoutMs: 200, ...opts });
    opened.push(d);
    const r = await d.open({ fileId: '1' });
    assert.equal(r.ok, true, r.detail);
    if (env.ctl.wsHandler) { env.ctl.ws = fakeWs(DS_URL); env.ctl.wsHandler(env.ctl.ws); env.ctl.ws.emit('framereceived', auth()); }
    return d;
  } finally { globalThis.fetch = realFetch; }
}

/** A dispatch the m7 harness's own `write()` makes -- no `meta` at all. Kept byte-identical to m7's
 * helper (not imported, same reason m7 itself does not import from a sibling file) as the negative
 * control: this path must be UNCHANGED by the GH #359 fix. */
const writeNoMeta = d => d.dispatch(null, { kind: 'callCommand', body: 'return 1;' });

/** A rewrite_paragraph-shaped dispatch, with the family/action meta the real adapter's `lower()`
 * attaches (text.cjs:229-233) -- the only shape the new gate inspects. */
const writeRewrite = (d, env, result) => { env.ctl.nextResult = result; return d.dispatch(null, { kind: 'callCommand', body: 'commandBody-source-elided-in-this-fake', meta: { family: 'text', action: 'rewrite', oldText: 'x', newText: 'y' } }); };

// --- negative control: the pre-existing, meta-less write path is byte-for-byte unchanged ---------

test('unchanged control: a dispatch with no meta at all still unconditionally mutates (m7\'s own write()) -- close() still demands a save-ack', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  await writeNoMeta(d);
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [false, 'mutated-without-observed-save']);
});

// --- the fix: a rewrite_paragraph dispatch that reports a NAMED no-write reason does not mutate ---

test('GH #359 (1) FIX: a rewrite_paragraph dispatch whose commandBody reports a named no-write reason (the MULTIRUN shape) does NOT demand a save-ack -- close() succeeds with no observed save', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  const dispatchResult = await writeRewrite(d, env, 'rewrite-live-range-text-mismatch');
  assert.equal(dispatchResult.ok, true, 'the body ran to completion (command-completed); this is a REPORTED outcome, not a transport failure');
  // review #ref, point 1: the REAL driver's dispatch() return must carry the named reason itself
  // (not just suppress the mutated gate) -- a LIVE run needs this to know WHICH guard fired.
  assert.equal(dispatchResult.reason, 'rewrite-live-range-text-mismatch');
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.state, r.saveAck.reason], [true, 'verified', 'no-save-traffic']);
});

test('every one of REWRITE_NO_WRITE_REASONS suppresses the mutated gate, not just the MULTIRUN one', async () => {
  const { REWRITE_NO_WRITE_REASONS } = require('../../src/adapters/docx/families/text.cjs');
  for (const reason of REWRITE_NO_WRITE_REASONS) {
    const env = fakeDriverEnv();
    const d = await openFakeDriver(env);
    await writeRewrite(d, env, reason);
    const r = await d.close();
    assert.deepEqual([r.ok, r.saveAck.reason], [true, 'no-save-traffic'], `reason ${reason} should not have required a save-ack`);
  }
});

// --- negative controls on the fix itself: a genuine write, or a post-write rollback failure, must still mutate ---

test("NEGATIVE CONTROL: a rewrite_paragraph dispatch that reports 'ok' (a real write happened) still mutates -- close() still demands a save-ack", async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  await writeRewrite(d, env, 'ok');
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [false, 'mutated-without-observed-save']);
});

test("NEGATIVE CONTROL (review #ref point 2): a rewrite_paragraph dispatch that reports 'rewrite-live-addtext-failed' still mutates -- AddText IS the mutating call, a `!== true` return is not proof nothing was inserted", async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  await writeRewrite(d, env, 'rewrite-live-addtext-failed');
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [false, 'mutated-without-observed-save']);
});

test('NEGATIVE CONTROL: a rewrite_paragraph dispatch that reports a post-AddText rollback reason (text.cjs:273) still mutates -- a best-effort in-editor rollback is not proof nothing was ever written', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  await writeRewrite(d, env, 'rewrite-live-rollback-inserted-text-mismatch');
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [false, 'mutated-without-observed-save']);
});

test('NEGATIVE CONTROL: the no-write gate is scoped to family:text/action:rewrite -- the SAME result string on a different family/action still mutates', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  env.ctl.nextResult = 'rewrite-live-range-text-mismatch';
  await d.dispatch(null, { kind: 'callCommand', body: 'x', meta: { family: 'table', action: 'cell' } });
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [false, 'mutated-without-observed-save']);
});

test('NEGATIVE CONTROL: an unrecognised result string on a rewrite_paragraph dispatch still mutates -- the gate is an allowlist, not "anything non-ok"', async () => {
  const env = fakeDriverEnv();
  const d = await openFakeDriver(env);
  await writeRewrite(d, env, 'some-future-guard-not-yet-catalogued');
  const r = await d.close();
  assert.deepEqual([r.ok, r.saveAck.reason], [false, 'mutated-without-observed-save']);
});

// --- review #ref, point 1: the named reason must reach appliedOperations[0].uncertainty.dispatch-
//     Evidence.reason through the REAL executor pipeline (classifyDispatchAttempt/validateUncertainty),
//     not just the driver in isolation. The driver tested above has no commandBody string to run, so
//     this uses the FAKE adapter-level driver (test/fakes/fake-docx-driver.cjs), whose dispatch() now
//     scripts the same `reason` field the real driver's generic dispatch() branch adds (docx-driver.cjs,
//     this unit's own fix) -- the wiring under test is core/execution/{executor,uncertainty,result}.cjs,
//     identical for either driver, since handle.dispatch() (core/session/handle.cjs:46) is a raw
//     passthrough to whatever the driver returns. ------------------------------------------------------

{
  const { withSession } = require('../../src/core/session/lifecycle.cjs');
  const { createExecutor } = require('../../src/core/execution/executor.cjs');
  const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
  const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
  const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
  const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');

  const PROPS = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 11, language: 'hu-HU' };
  const OLD_TEXT = 'GH359-1 reason-propagation target';
  const paragraphs = () => [{ text: 'Intro.', props: PROPS }, { text: OLD_TEXT, props: PROPS }, { text: 'Closing.', props: PROPS }];
  const selectorFor = ps => {
    const primary = normalizeInventory({ html: htmlOf(ps), paragraphs: ps });
    const b = primary.blocks.find(x => x.text === OLD_TEXT);
    return { text: OLD_TEXT, kind: b.kind, before: b.locator.context.before, after: b.locator.context.after };
  };

  async function runRewrite(reason) {
    let t = 0;
    const ps = paragraphs();
    const driver = createFakeDocxDriver(createStore(ps), { fault: 'ignore', rewriteDispatchReason: reason });
    const task = { id: 'gh359-1-reason', steps: [{ intent: 'rewrite_paragraph', selector: selectorFor(ps), text: 'replacement' }] };
    const options = { driver, fileId: 'disposable', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
    return withSession(options, docxAdapter, createExecutor(task));
  }

  test('GH #359 (1) review #ref point 1: a no-write reason on dispatch reaches appliedOperations[0].uncertainty.dispatchEvidence.reason through withSession+executor', async () => {
    const result = await runRewrite('rewrite-live-range-text-mismatch');
    const e = result.execution.appliedOperations[0];
    assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
    assert.equal(e.uncertainty.dispatchEvidence.reason, 'rewrite-live-range-text-mismatch');
  });

  test('NEGATIVE CONTROL: a different named reason propagates as itself, not a fixed string -- the field carries the actual value, not a boolean flag', async () => {
    const result = await runRewrite('rewrite-live-paragraph-text-not-unique');
    const e = result.execution.appliedOperations[0];
    assert.equal(e.uncertainty.dispatchEvidence.reason, 'rewrite-live-paragraph-text-not-unique');
  });

  test('NEGATIVE CONTROL: no scripted reason at all (the pre-fix shape) still reaches dispatchEvidence as undefined/absent, not a stale leftover value', async () => {
    let t = 0;
    const ps = paragraphs();
    const driver = createFakeDocxDriver(createStore(ps), { fault: 'ignore' });
    const task = { id: 'gh359-1-no-reason', steps: [{ intent: 'rewrite_paragraph', selector: selectorFor(ps), text: 'replacement' }] };
    const options = { driver, fileId: 'disposable', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
    const result = await withSession(options, docxAdapter, createExecutor(task));
    const e = result.execution.appliedOperations[0];
    assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
    assert.equal(e.uncertainty.dispatchEvidence.reason, undefined);
  });
}
