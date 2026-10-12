'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { BRIDGE_CLIENT_SRC, BRIDGE_HYDRATIONS, BRIDGE_HYDRATE_CLIENT_SRC } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
const { RUNTIMES } = require('../../src/core/runtime-identity.cjs');
const { OTHER_PATCH } = require('../fakes/ds-bridge-runtime.cjs');

/**
 * B15 /5 -- sheet.model.ensureReady/v1, KIND=RUNTIME_HYDRATION (owner GO the owner, owner channel 2026-10-09 msg 33257,
 * internal tracker), OFFLINE. The ACTUAL shipped bridge-block.js text, evaluated with `new Function('Api', source)` against
 * a fake Api (the Wave B test pattern): the manifest classification, the separate hydrate() entry (read() refuses it,
 * so the READ contract is untouched), every fail-closed reason, nothing but the view creator touched, and the MCP-side
 * client's gates -- including the previous revision (/4) and a READ-classified entry refused, never a fallback.
 */

const BLOCK_PATH = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1', 'bridge-block.js');
const SOURCE = fs.readFileSync(BLOCK_PATH, 'utf8').replace(/,\s*$/, '');
const bridgeOf = api => new Function('Api', `${SOURCE}; return Api.GetEuroOfficeBridge();`)(api);
const CAP = BRIDGE_HYDRATIONS.SHEET_MODEL_READY;

/** A fake workbook: sheet `name` at `index`; the view creator returns a view only for `viewIndex`; every call logged. */
function hydrationApi({ name = 'N', index = 1, viewIndex = 1, creator, model } = {}) {
  const calls = [];
  const wb = creator === undefined ? { getWorksheet: i => { calls.push(['getWorksheet', i]); return i === viewIndex ? { view: i } : null; } } : creator;
  const ws = { worksheet: model === undefined ? { getIndex: () => index, workbook: { oApi: { wb } } } : model };
  return { api: { GetSheet: n => (n === name ? ws : null) }, calls };
}

const hydrate = (api, cap = CAP, runtime = RUNTIMES.xlsx, args = { sheet: 'N' }) => new Function('Api',
  `${BRIDGE_HYDRATE_CLIENT_SRC}; return __eoBridgeHydrate(${JSON.stringify(cap.id)}, ${cap.version}, ${JSON.stringify(runtime)}, ${JSON.stringify(args)});`)(api);

test('manifest /5: exactly one sheet.model.ensureReady/v1, kind RUNTIME_HYDRATION, mutation EPHEMERAL_VIEW_ONLY -- never READ', () => {
  const m = bridgeOf({ GetSheet: () => null }).manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.4.60/5');
  const c = m.capabilities.filter(x => x.id === CAP.id);
  assert.equal(c.length, 1);
  assert.deepEqual([c[0].version, c[0].kind, c[0].format, c[0].mutation], [1, 'RUNTIME_HYDRATION', 'XLSX', 'EPHEMERAL_VIEW_ONLY']);
  assert.equal(m.capabilities.filter(x => x.kind !== 'READ').length, 1, 'the hydration entry is the ONLY non-READ capability');
  assert.equal(m.capabilities.filter(x => x.kind === 'READ' && x.mutation === 'NO').length, 15, 'the fifteen /4 READ entries are unchanged');
});

test('POSITIVE: hydrate() runs the view creator once for that sheet\'s index and returns only { modelReady: true }', () => {
  const { api, calls } = hydrationApi();
  assert.deepEqual(bridgeOf(api).hydrate(CAP.id, { sheet: 'N' }), { ok: true, value: { modelReady: true } });
  assert.deepEqual(calls, [['getWorksheet', 1]]);
});

test('the READ contract is untouched: read() refuses the hydration id, hydrate() refuses every READ id', () => {
  const { api, calls } = hydrationApi();
  const b = bridgeOf(api);
  assert.deepEqual(b.read(CAP.id, { sheet: 'N' }), { ok: false, reason: 'capability-not-advertised' });
  assert.deepEqual(b.hydrate('drawing.anchor.read/v1', { sheet: 'N', drawing: 'X' }), { ok: false, reason: 'capability-not-advertised' });
  assert.deepEqual(calls, [], 'NEGATIVE CONTROL: neither call reached the view creator');
});

test('FAIL CLOSED: bad-args, sheet-not-found, model-threw, view-creator-missing, view-creator-threw, no-view', () => {
  const r = (opts, args = { sheet: 'N' }) => bridgeOf(hydrationApi(opts).api).hydrate(CAP.id, args);
  assert.deepEqual(r({}, {}), { ok: false, reason: 'bad-args' });
  assert.deepEqual(r({}, { sheet: 7 }), { ok: false, reason: 'bad-args' });
  assert.deepEqual(r({}, { sheet: 'X' }), { ok: false, reason: 'sheet-not-found' });
  assert.deepEqual(r({ model: { get workbook() { throw new Error('boom'); } } }), { ok: false, reason: 'model-threw' });
  assert.deepEqual(r({ creator: {} }), { ok: false, reason: 'view-creator-missing' });
  assert.deepEqual(r({ model: { workbook: { oApi: { wb: { getWorksheet: () => ({}) } } } } }), { ok: false, reason: 'view-creator-missing' }, 'no getIndex');
  assert.deepEqual(r({ creator: { getWorksheet: () => { throw new Error('x'); } } }), { ok: false, reason: 'view-creator-threw' });
  assert.deepEqual(r({ viewIndex: 9 }), { ok: false, reason: 'no-view' });
});

test('MCP client: PASS on the accepted runtime; the previous revision (/4) refuses before the bridge is asked', () => {
  const { api, calls } = hydrationApi();
  bridgeOf(api); // installs the shipped block's Api.GetEuroOfficeBridge on this fake
  assert.deepEqual(hydrate(api), { ok: true }, 'POSITIVE CONTROL on /5');
  assert.equal(OTHER_PATCH, 'euro-office-bridge-v1/9.3.4.60/4');
  assert.deepEqual(hydrate(api, CAP, { ...RUNTIMES.xlsx, patchVersion: OTHER_PATCH }), { ok: false, reason: 'PATCH_VERSION_MISMATCH' });
  assert.deepEqual(hydrate(api, CAP, { ...RUNTIMES.xlsx, documentServerVersion: '9.3.5' }), { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' });
  assert.equal(calls.length, 1, 'only the positive control reached the view creator');
});

test('MCP client: missing / wrong-version / READ-classified / duplicated entry or no hydrate() -> CAPABILITY_*; no fallback', () => {
  const { api, calls } = hydrationApi();
  const real = bridgeOf(api);
  const m = real.manifest();
  const variant = (caps, extra = {}) => ({ GetEuroOfficeBridge: () => ({ manifest: () => ({ ...m, capabilities: caps(m.capabilities) }), read: real.read, hydrate: real.hydrate, ...extra }) });
  const mine = c => c.id === CAP.id;
  assert.deepEqual(hydrate(variant(cs => cs.filter(c => !mine(c)))), { ok: false, reason: 'CAPABILITY_MISSING' });
  assert.deepEqual(hydrate(variant(cs => cs.map(c => (mine(c) ? { ...c, version: 2 } : c)))), { ok: false, reason: 'CAPABILITY_VERSION_MISMATCH' });
  assert.deepEqual(hydrate(variant(cs => cs.map(c => (mine(c) ? { ...c, kind: 'READ', mutation: 'NO' } : c)))), { ok: false, reason: 'CAPABILITY_MISSING' }, 'a READ-classified entry is not a hydration');
  assert.deepEqual(hydrate(variant(cs => cs.map(c => (mine(c) ? { ...c, mutation: 'YES' } : c)))), { ok: false, reason: 'CAPABILITY_MISSING' });
  assert.deepEqual(hydrate(variant(cs => [...cs, cs.find(mine)])), { ok: false, reason: 'CAPABILITY_MISSING' }, 'advertised twice');
  assert.deepEqual(hydrate(variant(cs => cs, { hydrate: undefined })), { ok: false, reason: 'CAPABILITY_MISSING' });
  assert.deepEqual(hydrate({}), { ok: false, reason: 'BRIDGE_UNAVAILABLE' });
  assert.deepEqual(calls, [], 'NEGATIVE CONTROL: no refused case reached the view creator');
  assert.doesNotMatch(BRIDGE_HYDRATE_CLIENT_SRC, /getWorksheet|\.worksheet\b|oApi/, 'the client names no internal -- there is nothing to fall back to');
  assert.doesNotMatch(BRIDGE_CLIENT_SRC, /__eoBridgeHydrate/, 'the READ client stays byte-identical (no hydration in a read-only body)');
});

test('MCP client: only exactly { ok: true, value: { modelReady: true } } passes; the bridge\'s own refusal is namespaced', () => {
  const { api } = hydrationApi();
  const real = bridgeOf(api);
  const answering = res => ({ GetEuroOfficeBridge: () => ({ manifest: real.manifest, read: real.read, hydrate: () => res }) });
  for (const bad of [{ ok: true }, { ok: true, value: true }, { ok: true, value: { modelReady: false } }, { ok: true, value: { modelReady: true, data: 1 } },
    { ok: true, value: { modelReady: true }, extra: 1 }, { ok: false }, { ok: false, reason: '' }, null, 'x']) {
    assert.deepEqual(hydrate(answering(bad)), { ok: false, reason: 'UNEXPECTED_SHAPE' }, JSON.stringify(bad));
  }
  assert.deepEqual(hydrate(answering({ ok: false, reason: 'no-view' })), { ok: false, reason: 'BRIDGE:no-view' });
  const throwing = { GetEuroOfficeBridge: () => ({ manifest: real.manifest, read: real.read, hydrate: () => { throw new Error('x'); } }) };
  assert.deepEqual(hydrate(throwing), { ok: false, reason: 'BRIDGE_THREW' });
});
