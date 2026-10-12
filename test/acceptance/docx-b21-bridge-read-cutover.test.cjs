'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * M11.5e.4 (owner: M11.5 continuous mandate, internal tracker) -- the B21 cutover END TO END, offline: the REAL
 * product command (header-footer.lower -> commandBody, with the driver's own b21BridgeReadPrelude) runs against a fake
 * document whose bridge is the REAL word /3 candidate block (or the /2 block, which lacks the capability). The
 * title/even re-check reaches the editor's section-variant state ONLY through the Bridge read(); on word /2 the gate
 * refuses with zero Bridge reads and the direct internal getters are never called from product code.
 */

const hf = require('../../src/adapters/docx/families/header-footer.cjs');
const driver = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { DOCX_B21_BRIDGE_READ_RUNTIME_ENTRIES, DOCX_B18_BRIDGE_WRITE_RUNTIME_ENTRIES, runtimeIdOf } = require('../../src/core/runtime-identity.cjs');
const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-word-v1');
const V2 = fs.readFileSync(path.join(DIR, 'rev2-candidate', 'bridge-block.js'), 'utf8');
const V3 = fs.readFileSync(path.join(DIR, 'rev3-candidate', 'bridge-block.js'), 'utf8');

/** fake Word document + the REAL bridge block installed on the same Api; every Get_TitlePage/IsEvenAndOdd and bridge
 * read() recorded. ONE section, header-footer part already exists (title/even are flipped, not created). */
function editor(block, { titlePage = false, evenAndOdd = false } = {}) {
  const calls = { getters: [], read: 0 };
  const state = { titlePage, evenAndOdd, added: null };
  const para = { GetClassType: () => 'paragraph', GetElementsCount: () => 0, GetElement: () => null, AddText: t => { state.added = t; } };
  const part = { GetElementsCount: () => 1, GetElement: () => para, RemoveAllElements: () => { state.added = null; } };
  const inner = {
    Get_TitlePage: () => { calls.getters.push('Get_TitlePage'); return state.titlePage; },
    IsEvenAndOdd: () => { calls.getters.push('IsEvenAndOdd'); return state.evenAndOdd; },
  };
  const section = { Section: inner, GetHeader: v => (v === 'default' || v === 'title' ? part : null), GetFooter: () => null,
    SetTitlePage: v => { state.titlePage = v; } };
  const doc = { GetSections: () => [section] };
  const Api = { GetDocument: () => doc };
  new Function('Api', block.replace(/,\s*$/, ''))(Api);
  const raw = Api.GetEuroOfficeBridge;
  Api.GetEuroOfficeBridge = function () { const b = raw.apply(this, arguments); if (typeof b.read === 'function') { const r = b.read; b.read = function () { calls.read += 1; return r.apply(this, arguments); }; } return b; };
  return { Api, calls, state };
}

const target = v => ({ kind: 'section_part', sectionIndex: 0, sectionCount: 1, part: 'header', variant: v, create: false, activate: true });
const command = v => hf.lower({ intent: 'edit_header', text: 'New' }, target(v));
const runIn = (env, cmd) => new Function('Api', `${driver.b21BridgeReadPrelude()}${cmd.body}`)(env.Api);

test('word /3: a title command re-checks through ONE Bridge read; the direct internal getters are never called from product code', () => {
  const env = editor(V3, { titlePage: false, evenAndOdd: false });
  const cmd = command('title');
  assert.equal(cmd.meta.bridgeSectionVariantRead, true);
  runIn(env, cmd);
  assert.equal(env.calls.read, 1, 'ONE Bridge read() call');
  assert.equal(env.state.titlePage, true, 'activated: the flag flips');
  assert.equal(env.state.added, 'New', 'the content is written too');
  // the Get_TitlePage/IsEvenAndOdd calls recorded here are the REAL /3 block's OWN internal implementation of the
  // capability -- not a call from product code (product only ever calls __eoB21Variant / __eoBridgeRead)
  assert.deepEqual(env.calls.getters, ['Get_TitlePage', 'IsEvenAndOdd']);
});

test('word /3: already-active flag -> no activation side effect, still ONE Bridge read', () => {
  const env = editor(V3, { titlePage: true, evenAndOdd: false });
  runIn(env, command('title'));
  assert.equal(env.calls.read, 1);
  assert.equal(env.state.added, null, 'already active: the command returns before writing');
});

test('word /2 (no B21 Bridge READ capability): the product re-check gets null, zero mutation -- the driver\'s second line refuses it before this body would even run', () => {
  const env = editor(V2, { titlePage: false, evenAndOdd: false });
  const cmd = command('title');
  runIn(env, cmd);
  assert.equal(env.calls.read, 0, 'PATCH_VERSION_MISMATCH refuses INSIDE __eoBridgeRead, before bridge.read() is even called');
  assert.equal(env.state.titlePage, false, 'not activated: __eoB21Variant returned null, the body returned early');
  assert.equal(env.state.added, null);
  // the driver's second line (the actual production gate, exercised before ANY dispatch) refuses the SAME command
  const refused = driver.internalWriteRefusal(cmd, '9.3.4|euro-office-bridge-word-v1/9.3.4.60/2');
  assert.deepEqual([refused.ok, refused.dispatched, refused.outcome], [false, false, 'b21-bridge-read-runtime-not-proven']);
  assert.equal(driver.internalWriteRefusal(cmd, '9.3.4|euro-office-bridge-word-v1/9.3.4.60/3'), null, 'word /3: let through');
});

test('b21BridgeReadPrelude: the editor-side runtime gate itself is narrowed to word /3 ONLY -- this is a SECOND, independent layer behind the driver\'s dispatch-side second line, and a widened gate here must be caught even if the dispatch gate is also somehow bypassed', () => {
  // (a) structural: the generated source embeds the word-/3-only runtime list, not the wider B18 (/2+/3) one
  const src = driver.b21BridgeReadPrelude();
  assert.ok(src.includes(JSON.stringify(DOCX_B21_BRIDGE_READ_RUNTIME_ENTRIES)), 'embeds DOCX_B21_BRIDGE_READ_RUNTIME_ENTRIES ([/3] only)');
  assert.equal(src.includes(JSON.stringify(DOCX_B18_BRIDGE_WRITE_RUNTIME_ENTRIES)), false, 'does NOT embed the wider B18 (/2+/3) list');
  // (b) behavioural, isolating the runtime check from capability-presence masking: a manifest on word /2 that
  // (unrealistically) ALSO advertises the capability must still be refused -- by the RUNTIME check alone, not by
  // capability absence. If the prelude's own runtime-list parameter were ever widened to include /2 (e.g. swapped for
  // DOCX_B18_BRIDGE_WRITE_RUNTIME_ENTRIES), this fake would let it through; a realistic /2 fixture (capability
  // genuinely absent) could not tell the difference, which is exactly the gap this test closes.
  const manifest = patchVersion => ({ bridgeVersion: '1.0.0', documentServerVersion: '9.3.4', patchVersion,
    capabilities: [{ id: 'docx.section.variantState.read/v1', version: 1, kind: 'READ', format: 'DOCX', mutation: 'NO' }] });
  const client = patchVersion => {
    const Api = { GetEuroOfficeBridge: () => ({ manifest: () => manifest(patchVersion), read: () => ({ ok: true, value: { titlePage: true, evenAndOdd: false } }) }) };
    return new Function('Api', `${driver.b21BridgeReadPrelude()} return __eoB21Variant(0);`)(Api);
  };
  assert.equal(client('euro-office-bridge-word-v1/9.3.4.60/2'), null, 'word /2, even with the capability (unrealistically) advertised: the runtime gate itself refuses');
  assert.deepEqual(client('euro-office-bridge-word-v1/9.3.4.60/3'), { titlePage: true, evenAndOdd: false }, 'word /3: accepted');
  assert.equal(runtimeIdOf({ documentServerVersion: '9.3.4', patchVersion: 'euro-office-bridge-word-v1/9.3.4.60/3' }), runtimeIdOf(DOCX_B21_BRIDGE_READ_RUNTIME_ENTRIES[0]), 'sanity: /3 is the entry under test');
});

test('__eoB21Variant: the prelude\'s OWN validator rejects any malformed/incomplete bridge value as null, never trusts it as-is', () => {
  const manifest = () => ({ bridgeVersion: '1.0.0', documentServerVersion: '9.3.4', patchVersion: 'euro-office-bridge-word-v1/9.3.4.60/3',
    capabilities: [{ id: 'docx.section.variantState.read/v1', version: 1, kind: 'READ', format: 'DOCX', mutation: 'NO' }] });
  const client = readResult => {
    const Api = { GetEuroOfficeBridge: () => ({ manifest, read: () => readResult }) };
    return new Function('Api', `${driver.b21BridgeReadPrelude()} return __eoB21Variant(0);`)(Api);
  };
  for (const bad of [
    { ok: true, value: { titlePage: 'yes', evenAndOdd: false } }, // non-boolean
    { ok: true, value: { titlePage: true } }, // missing evenAndOdd
    { ok: true, value: { titlePage: true, evenAndOdd: false, extra: 1 } }, // extra key
    { ok: true, value: null },
    { ok: true, value: 'nope' },
    { ok: false, reason: 'section-not-found' },
    null,
  ]) assert.equal(client(bad), null, JSON.stringify(bad));
  assert.deepEqual(client({ ok: true, value: { titlePage: true, evenAndOdd: false } }), { titlePage: true, evenAndOdd: false }, 'a well-shaped value IS trusted');
});

test('a default-variant command never gets the driver-injected prelude (__eoBridgeRead is undefined there, so the dead __eoB21Variant branch is never reached)', () => {
  const cmd = hf.lower({ intent: 'edit_header', text: 'New' }, target('default'));
  assert.equal(cmd.meta.bridgeSectionVariantRead, undefined);
  assert.doesNotMatch(cmd.body, /__eoBridgeRead/, 'the prelude-only helper name never comes from commandBody itself');
  const env = editor(V3, { titlePage: false, evenAndOdd: false });
  assert.doesNotThrow(() => new Function('Api', cmd.body)(env.Api), 'runs fine WITHOUT the prelude: the m.variant===\'default\' branch never calls __eoB21Variant');
  assert.equal(env.calls.read, 0);
});
