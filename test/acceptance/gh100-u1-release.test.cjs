'use strict';

const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');

const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { fakeDriverEnv, openFakeDriver, saveCycle, answerOk, write } = require('../fakes/fake-xlsx-session-env.cjs');

/**
 * GH #100 U1: release() / resetForNextCall() / alive() on the XLSX driver -- the call end of a session that STAYS open. Nothing uses them yet.
 * LIVE (probe, 3 of 3): every forceSave of ONE session gets its own confirmation, so the observer is not reset between calls.
 * release() = close() minus requestClose and teardown; close() is untouched (its own suite, xlsx-m7-p1-close-ack, stays green).
 */

const opened = [];
after(async () => { for (const d of opened) await d.close(); }); // a driver left open keeps its loopback host alive
const open = async (env, o) => { const d = await openFakeDriver(drv, env, o); opened.push(d); return d; };

const forceSaveCall = async (env, d, from, to) => {
  await write(d);
  saveCycle(env.ctl.ws, from, to, { time: 700 + to }); // the confirmation must carry THIS ack's save time
  env.ctl.onForceSave = () => answerOk(env.ctl.ws, 700 + to);
  return d.forceSave();
};

test('release after an acked write: ok, verified, the confirmation stands -- and NOTHING was closed', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  const f = await forceSaveCall(env, d, 0, 4004);
  assert.equal(f.confirmed, true);
  const r = await d.release();
  assert.equal(r.ok, true, r.detail);
  assert.equal(r.saveAck.state, 'verified');
  assert.equal(r.confirmation.state, 'confirmed');
  assert.match(r.detail, /session-kept; save-ack verified/);
  assert.equal(env.ctl.log.includes('requestClose'), false);
  assert.equal(env.ctl.log.includes('browser-close'), false);
  assert.equal(await d.alive(), true, 'the editor still answers');
});

test('release with an unacked write: not ok (a failed call end), still nothing closed, the reason names what was pending', async () => {
  const env = fakeDriverEnv();
  const d = await open(env, { closeAckTimeoutMs: 60 });
  await write(d);
  saveCycle(env.ctl.ws, 0, 4004, { ack: false });
  const r = await d.release();
  assert.equal(r.ok, false);
  assert.match(r.saveAck.reason, /^timeout:save-outstanding/);
  assert.equal(env.ctl.log.includes('browser-close'), false, 'the caller decides to close; release never tears down');
});

test('release vs close on the SAME events: the same save claim (ok, ack state and reason, confirmation) -- release only skips the teardown', async () => {
  const claims = [];
  for (const how of ['release', 'close']) {
    const env = fakeDriverEnv();
    const d = await open(env);
    await forceSaveCall(env, d, 0, 4004);
    const r = await d[how]();
    claims.push({ ok: r.ok, ack: [r.saveAck.state, r.saveAck.reason], confirmation: r.confirmation && r.confirmation.state });
    assert.equal(env.ctl.log.includes('browser-close'), how === 'close');
  }
  assert.deepEqual(claims[0], claims[1]);
});

test('ONE session, THREE calls: write + forceSave + release + reset, three times -- every call has its own confirmation', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  let sync = 0;
  for (let call = 1; call <= 3; call += 1) {
    const f = await forceSaveCall(env, d, sync, sync + 4004);
    sync += 4004;
    assert.deepEqual([f.ok, f.confirmed], [true, true], `call ${call}`);
    const r = await d.release();
    assert.deepEqual([r.ok, r.confirmation && r.confirmation.state], [true, 'confirmed'], `release ${call}`);
    assert.equal(d.resetForNextCall().ok, true);
  }
  assert.equal(env.ctl.log.filter(x => x === 'launch').length, 1, 'one browser for all three calls');
  assert.equal(env.ctl.log.filter(x => x === 'forceSave').length, 3);
});

test('a write-free call after an acked write call ends clean, and owns no forceSave operation', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  await forceSaveCall(env, d, 0, 4004);
  assert.equal((await d.release()).ok, true);
  d.resetForNextCall();
  const quiet = await d.release();                       // a read-only second call
  assert.equal(quiet.ok, true, quiet.detail);
  assert.equal(quiet.confirmation, null, 'no forceSave operation belongs to the write-free call');
});

test('release on a dead frame: not ok (the call end is not trusted); alive() is false and never throws', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  env.ctl.frameDead = true;
  assert.equal(await d.alive(), false);
  const r = await d.release();
  assert.equal(r.ok, false);
  assert.match(r.detail, /no live editor/);
  env.ctl.frameDead = false;
});

test('before open: release / alive / resetForNextCall are safe and say "no session"', async () => {
  const d = drv.createEuroOfficeXlsxDriver({ url: 'https://nc.example', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false, detail: 'none' }) });
  assert.equal(await d.alive(), false);
  assert.equal(d.resetForNextCall().ok, false);
  const r = await d.release();
  assert.equal(r.ok, false);
});

test('resetForNextCall clears `mutated`: a dispatch with no observed save is a failed call end; after the reset the NEXT (write-free) call ends clean', async () => {
  const env = fakeDriverEnv();
  const d = await open(env);
  await write(d);                                   // a write attempt that produced no save traffic
  const bad = await d.release();
  assert.deepEqual([bad.ok, bad.saveAck.reason], [false, 'mutated-without-observed-save']);
  assert.equal(d.resetForNextCall().ok, true);
  const next = await d.release();
  assert.deepEqual([next.ok, next.saveAck.reason], [true, 'no-save-traffic']);
});
