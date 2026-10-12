'use strict';

// DS Bridge v1 READ pilot (internal tracker): the in-editor client's fail-closed gates, run against fake `Api` objects.
// The client source is evaluated exactly as the callCommand body embeds it. Since owner decision (B) (internal tracker
// #ref) the identity gate is the manifest's stable (documentServerVersion, patchVersion) -- never a build/cache key.
const test = require('node:test');
const assert = require('node:assert/strict');
const { BRIDGE_CLIENT_SRC, BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { RUNTIMES } = require('../../src/core/runtime-identity.cjs');

const EXPECTED = RUNTIMES.xlsx;
const CAP = BRIDGE_CAPABILITIES.BOLD_READ;

function clientWith(Api) {
  // eslint-disable-next-line no-new-func
  return new Function('Api', `${BRIDGE_CLIENT_SRC} return __eoBridgeRead;`)(Api);
}

const manifest = (over = {}) => ({
  bridgeVersion: '1.0.0',
  documentServerVersion: '9.3.4',
  documentServerBuild: '9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b',
  patchVersion: 'euro-office-bridge-v1/9.3.4.60/5',
  capabilities: [{ id: CAP.id, version: 1, kind: 'READ', format: 'XLSX', mutation: 'NO', semanticContract: 'x' }],
  ...over,
});

function apiWith({ m = manifest(), read = () => ({ ok: true, value: true }) } = {}) {
  return { GetEuroOfficeBridge: () => ({ manifest: () => m, read }) };
}

const call = (Api, expected = EXPECTED) => clientWith(Api)(CAP.id, CAP.version, expected, { sheet: 'S', cell: 'A1' }, 'boolean');

test('PASS: an advertised capability on the expected runtime returns the bridge value (true and false)', () => {
  assert.deepEqual(call(apiWith()), { ok: true, value: true });
  assert.deepEqual(call(apiWith({ read: () => ({ ok: true, value: false }) })), { ok: true, value: false });
});

test('the bridge receives exactly the capability ID and the semantic args -- no method name, no path', () => {
  const seen = [];
  call(apiWith({ read: (id, args) => { seen.push([id, args]); return { ok: true, value: true }; } }));
  assert.deepEqual(seen, [[CAP.id, { sheet: 'S', cell: 'A1' }]]);
});

test('gate 1 BRIDGE_UNAVAILABLE: no entry point, entry point throws, incomplete object, manifest throws', () => {
  assert.deepEqual(call({}), { ok: false, reason: 'BRIDGE_UNAVAILABLE' });
  assert.deepEqual(call({ GetEuroOfficeBridge: () => { throw new Error('x'); } }), { ok: false, reason: 'BRIDGE_UNAVAILABLE' });
  assert.deepEqual(call({ GetEuroOfficeBridge: () => ({ manifest: () => manifest() }) }), { ok: false, reason: 'BRIDGE_UNAVAILABLE' });
  assert.deepEqual(call({ GetEuroOfficeBridge: () => ({ manifest: () => { throw new Error('x'); }, read: () => ({}) }) }), { ok: false, reason: 'BRIDGE_UNAVAILABLE' });
});

test('MANIFEST_MALFORMED: a non-object manifest or a missing / non-string identity field or capability list', () => {
  for (const m of [null, 'x', manifest({ bridgeVersion: undefined }), manifest({ bridgeVersion: 1 }), manifest({ documentServerVersion: undefined }),
    manifest({ patchVersion: 3 }), manifest({ capabilities: undefined }), manifest({ capabilities: 'all' })]) {
    assert.deepEqual(call(apiWith({ m })), { ok: false, reason: 'MANIFEST_MALFORMED' }, JSON.stringify(m));
  }
});

test('BRIDGE_VERSION_MISMATCH: another bridge major', () => {
  for (const v of ['2.0.0', '0.9.0']) {
    assert.deepEqual(call(apiWith({ m: manifest({ bridgeVersion: v }) })), { ok: false, reason: 'BRIDGE_VERSION_MISMATCH' }, v);
  }
});

test('RUNTIME_VERSION_MISMATCH: another DS runtime version, or the caller expects nothing', () => {
  assert.deepEqual(call(apiWith({ m: manifest({ documentServerVersion: '9.3.5' }) })), { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' });
  assert.deepEqual(call(apiWith(), null), { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' });
  assert.deepEqual(call(apiWith(), '9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b'), { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' }, 'a cache-key string is not an expected runtime');
});

test('PATCH_VERSION_MISMATCH: another patch family, DS patch runtime or revision', () => {
  for (const v of ['euro-office-bridge-v1/9.3.4.60/2', 'euro-office-bridge-v1/9.3.4.61/3', 'euro-office-bridge-word-v1/9.3.4.60/3']) {
    assert.deepEqual(call(apiWith({ m: manifest({ patchVersion: v }) })), { ok: false, reason: 'PATCH_VERSION_MISMATCH' }, v);
  }
});

test('CACHE-KEY CHANGE: the manifest\'s own documentServerBuild (the patch-time cache key) is NOT compared -- any value PASSES', () => {
  for (const b of ['9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b', '9.3.4-94bc577ecccd9583be1b130bd5a6b407', undefined]) {
    assert.deepEqual(call(apiWith({ m: manifest({ documentServerBuild: b }) })), { ok: true, value: true }, String(b));
  }
});

test('CAPABILITY_MISSING / CAPABILITY_VERSION_MISMATCH: absent, other version, not READ, mutating, or advertised twice', () => {
  const capWith = c => manifest({ capabilities: c });
  const base = { id: CAP.id, version: 1, kind: 'READ', format: 'XLSX', mutation: 'NO' };
  for (const caps of [[], [{ ...base, id: 'other.read/v1' }], [{ ...base, kind: 'WRITE' }], [{ ...base, mutation: 'YES' }], [base, base]]) {
    assert.deepEqual(call(apiWith({ m: capWith(caps) })), { ok: false, reason: 'CAPABILITY_MISSING' }, JSON.stringify(caps));
  }
  assert.deepEqual(call(apiWith({ m: capWith([{ ...base, version: 2 }]) })), { ok: false, reason: 'CAPABILITY_VERSION_MISMATCH' });
});

test('gate 5 UNEXPECTED_SHAPE: a result outside { ok, value } / { ok, reason } is never accepted', () => {
  for (const r of [null, 'true', { ok: 'yes' }, { ok: true }, { ok: true, value: 'true' }, { ok: true, value: true, extra: 1 }, { ok: false }, { ok: false, reason: '' }, { ok: false, reason: 'x', value: true }]) {
    assert.deepEqual(call(apiWith({ read: () => r })), { ok: false, reason: 'UNEXPECTED_SHAPE' }, JSON.stringify(r));
  }
});

test('the bridge\'s own refusal passes through, namespaced; a throwing read is BRIDGE_THREW', () => {
  assert.deepEqual(call(apiWith({ read: () => ({ ok: false, reason: 'cell-not-found' }) })), { ok: false, reason: 'BRIDGE:cell-not-found' });
  assert.deepEqual(call(apiWith({ read: () => { throw new Error('x'); } })), { ok: false, reason: 'BRIDGE_THREW' });
});

test('readQuery (B01 / B02 cutover): bold and italic go ONLY through the bridge client -- no direct internal font read', () => {
  const w = [{ sheet: 'S', range: 'A1:A1', format: true }];
  const q = readQuery(w, {});
  assert.equal(/getBold|getItalic|getFont/.test(q), false, 'no direct internal font read, and so no fallback to one');
  assert.equal((q.match(/const __eoBridgeRead = /g) || []).length, 1, 'the client is defined once');
  assert.ok(q.includes(JSON.stringify(BRIDGE_CAPABILITIES.BOLD_READ)) && q.includes(JSON.stringify(BRIDGE_CAPABILITIES.ITALIC_READ)));
  assert.ok(q.includes(JSON.stringify(EXPECTED)), 'the expected runtime is the stable manifest identity');
  assert.equal(/9\.3\.4-[0-9a-f]{8,}/.test(q), false, 'no cache key in the read body');
});
