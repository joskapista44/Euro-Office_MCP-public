'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * M11.5d.4 (owner: M11.5 continuous mandate, internal tracker) -- the B18 cutover END TO END, offline: the REAL product
 * command (list.lower -> commandBody, with the driver's own b18BridgeWritePrelude) runs against a fake document whose
 * bridge is the REAL word /2 candidate block (or the /1 block). The removal reaches the editor routine ONLY through the
 * Bridge write(); on any other runtime the gate refuses with zero removals; there is no direct internal path.
 */

const list = require('../../src/adapters/docx/families/list.cjs');
const driver = require('../../src/drivers/eurooffice/docx-driver.cjs');
const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-word-v1');
const V1 = fs.readFileSync(path.join(DIR, 'bridge-block.js'), 'utf8');
const V2 = fs.readFileSync(path.join(DIR, 'rev2-candidate', 'bridge-block.js'), 'utf8');

/** fake Word document + the REAL bridge block installed on the same Api; every RemoveNumPr and bridge write() recorded. */
function editor(block, paras) {
  const calls = { remove: [], write: 0 };
  const P = paras.map(x => ({ ...x, num: x.num ? { ...x.num } : null }));
  const api = p => ({
    GetNumbering: () => (p.num ? { GetLevelIndex: () => p.num.level, GetNumbering: () => ({ GetLevel: () => ({}) }) } : null),
    Paragraph: { Pr: { NumPr: p.num ? { NumId: p.num.numId } : null }, RemoveNumPr: () => { calls.remove.push(p.text); p.num = null; } },
  });
  const doc = { Search: t => P.filter(p => p.text.includes(t)).map(p => ({ GetText: () => p.text, GetParagraph: () => api(p) })) };
  const Api = { GetDocument: () => doc };
  new Function('Api', block.replace(/,\s*$/, ''))(Api);
  const raw = Api.GetEuroOfficeBridge;
  Api.GetEuroOfficeBridge = function () { const b = raw.apply(this, arguments); if (typeof b.write === 'function') { const w = b.write; b.write = function () { calls.write += 1; return w.apply(this, arguments); }; } return b; };
  return { Api, calls, P };
}
const PARAS = [{ text: 'Elso elem', num: { level: 0, numId: 3 } }, { text: 'Masodik elem', num: { level: 1, numId: 3 } }, { text: 'Sima', num: null }];
const command = (listAction, text, beforeLevel, extra = {}) => {
  const meta = { family: 'list', action: 'list', listAction, targetText: text, beforeType: 'bullet', beforeLevel, bridgeNumberingRemove: true, ...extra };
  return { kind: 'callCommand', body: list.commandBody(meta), meta };
};
const runIn = (env, cmd) => new Function('Api', `${driver.b18BridgeWritePrelude()}${cmd.body}`)(env.Api);

test('word /2: remove_list -> ONE bridge write, the editor routine runs exactly once INSIDE it; receipt ok; others untouched', () => {
  const env = editor(V2, PARAS);
  const r = runIn(env, command('remove_list', 'Masodik elem', 1));
  assert.deepEqual(r, { ok: true });
  assert.deepEqual([env.calls.write, env.calls.remove], [1, ['Masodik elem']]);
  assert.deepEqual(env.P.map(p => p.num && p.num.level), [0, null, null]);
  assert.deepEqual(driver.b18DispatchResult({ ok: true, outcome: 'command-completed', result: r }), { ok: true, dispatched: true, outcome: 'command-completed', error: null });
});

test('word /2: outdent_list_item at level 0 is the same removal (one bridge write)', () => {
  const env = editor(V2, PARAS);
  assert.deepEqual(runIn(env, command('outdent_list_item', 'Elso elem', 0)), { ok: true });
  assert.deepEqual([env.calls.write, env.calls.remove], [1, ['Elso elem']]);
});

test('word /1 (no Bridge WRITE): the gate refuses BEFORE any write -- zero removals, mutation NONE, dispatch failed', () => {
  const env = editor(V1, PARAS);
  const r = runIn(env, command('remove_list', 'Elso elem', 0));
  assert.deepEqual([r.ok, r.stage, r.reason, r.mutation], [false, 'bridge-write-gate', 'PATCH_VERSION_MISMATCH', 'NONE']);
  assert.deepEqual([env.calls.write, env.calls.remove], [0, []]);
  const d = driver.b18DispatchResult({ ok: true, outcome: 'command-completed', result: r });
  assert.deepEqual([d.ok, d.dispatched, d.outcome], [false, true, 'b18-bridge-write-failed']);
});

test('a moved live level: the product body refuses before the bridge (no write); the bridge refuses a stale level itself too', () => {
  const env = editor(V2, PARAS);
  assert.equal(runIn(env, command('remove_list', 'Masodik elem', 0)), undefined, 'the body\'s own live-level re-check');
  assert.deepEqual([env.calls.write, env.calls.remove], [0, []]);
  const b = env.Api.GetEuroOfficeBridge();
  assert.deepEqual(b.write('docx.paragraph.numbering.remove/v1', { text: 'Masodik elem', level: 0 }), { ok: false, reason: 'level-mismatch', level: 1 });
  assert.deepEqual(env.calls.remove, []);
});

test('the B18 receipt rule: a pre-call refusal is mutation NONE; an after-call failure or a malformed receipt is uncertain / UNKNOWN', () => {
  const W = 'docx.paragraph.numbering.remove/v1';
  const client = r => new Function('Api', 'R', `${driver.b18BridgeWritePrelude()} return __eoBridgeNumberingRemove({ write: () => R }, ${JSON.stringify(W)}, { text: 'x', level: 0 });`)({}, r);
  assert.deepEqual(client({ ok: false, reason: 'not-numbered' }), { ok: false, reason: 'BRIDGE:not-numbered', uncertain: false, mutation: 'NONE' });
  assert.deepEqual(client({ ok: false, reason: 'RemoveNumPr-missing' }).mutation, 'NONE');
  for (const reason of ['remove-threw', 'remove-not-observed', 'something-new']) assert.deepEqual(client({ ok: false, reason }), { ok: false, reason: `BRIDGE:${reason}`, uncertain: true, mutation: 'UNKNOWN' }, reason);
  for (const bad of [{ ok: true, value: { applied: ['numbering-removed'] } }, { ok: true, value: { applied: ['x'], before: { level: 0 } } }, { ok: true, value: { applied: ['numbering-removed'], before: { level: 3 } } }, null, { ok: 'yes' }])
    assert.deepEqual(client(bad).reason, 'UNEXPECTED_SHAPE', JSON.stringify(bad));
  assert.deepEqual(client({ ok: true, value: { applied: ['numbering-removed'], before: { level: 0, numId: '1' } } }), { ok: true });
});

test('every other command stays byte-identical: the prelude is NOT in a non-removal list body', () => {
  const body = list.commandBody({ family: 'list', action: 'list', listAction: 'set_list_level', targetText: 'Elso elem', beforeType: 'bullet', beforeLevel: 0, level: 1 });
  assert.doesNotMatch(body, /__eoBridgeWriteGate|__eoBridgeNumberingRemove/);
});
