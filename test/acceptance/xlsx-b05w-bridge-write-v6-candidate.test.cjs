'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { RUNTIMES, classifyHandshake } = require('../../src/core/runtime-identity.cjs');

/**
 * M11.5a.1 (owner GO: the owner, owner channel) -- OFFLINE acceptance of the /6 CANDIDATE block
 * (documentserver-patches/9.3.4.60/euro-office-bridge-v1/rev6-candidate/bridge-block.js): the FIRST Bridge WRITE
 * capability, range.alignmentExt.write/v1. The ACTUAL candidate text is evaluated with `new Function('Api', source)`
 * against a fake Api whose objects are recording spies (the Wave B / B15 test pattern). Product code is NOT touched by
 * this milestone (cutover HOLD): the current product still accepts only /5 -- pinned below as a compatibility fact.
 */

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const read = p => fs.readFileSync(p, 'utf8');
const V5 = read(path.join(DIR, 'bridge-block.js'));
const V6 = read(path.join(DIR, 'rev6-candidate', 'bridge-block.js'));
const bridgeOf = (src, api) => new Function('Api', `${src.replace(/,\s*$/, '')}; return Api.GetEuroOfficeBridge();`)(api);
const W = 'range.alignmentExt.write/v1';

/** fake workbook: sheet `S`, any single cell resolves to ONE shared inner range whose every touch is recorded. */
function fakeApi({ setIndent = true, setShrink = true, throwIndent = false, throwShrink = false, getRangeThrows = false, noInner = false } = {}) {
  const calls = []; const state = { indent: 0, shrink: false };
  const target = {
    getAlign: () => ({ getReadingOrder: () => 0, getIndent: () => state.indent, getShrinkToFit: () => state.shrink }),
  };
  if (setIndent) target.setIndent = v => { if (throwIndent) throw new Error('boom-indent'); state.indent = v; };
  if (setShrink) target.setShrinkToFit = v => { if (throwShrink) throw new Error('boom-shrink'); state.shrink = v; };
  const inner = new Proxy(target, { get(t, k) { if (typeof k === 'string') calls.push(['inner.get', k]); const v = t[k]; return typeof v === 'function' ? (...a) => { calls.push(['inner.call', k, ...a]); return v(...a); } : v; } });
  const ws = { GetRange: a => { calls.push(['ws.GetRange', a]); if (getRangeThrows) throw new Error('bad'); return noInner ? {} : { range: inner }; } };
  const api = { GetSheet: n => { calls.push(['Api.GetSheet', n]); return n === 'S' ? ws : null; } };
  return { api, calls, state, mutations: () => calls.filter(c => c[0] === 'inner.call' && /^set/.test(c[1])) };
}
const write = (env, args, id = W) => bridgeOf(V6, env.api).write(id, args);

test('/6 manifest: patchVersion /6, /5 capabilities byte-identical, exactly one new WRITE entry', () => {
  const m5 = bridgeOf(V5, { GetSheet: () => null }).manifest(); const m6 = bridgeOf(V6, { GetSheet: () => null }).manifest();
  assert.equal(m5.patchVersion, 'euro-office-bridge-v1/9.3.4.60/5');
  assert.equal(m6.patchVersion, 'euro-office-bridge-v1/9.3.4.60/6');
  assert.equal(m6.capabilities.length, m5.capabilities.length + 1);
  assert.deepEqual(m6.capabilities.slice(0, m5.capabilities.length), m5.capabilities);
  const w = m6.capabilities.filter(c => c.id === W);
  assert.equal(w.length, 1);
  assert.deepEqual([w[0].version, w[0].kind, w[0].format, w[0].mutation], [1, 'WRITE', 'XLSX', 'TARGETED']);
  assert.equal(m5.capabilities.some(c => c.kind === 'WRITE'), false, '/5 has no WRITE class');
  for (const k of ['bridgeVersion', 'documentServerVersion', 'documentServerBuild']) assert.equal(m6[k], m5[k]);
});

test('/6 = /5 + three inserted segments + the patchVersion bump (every /5 implementation byte-identical)', () => {
  let t = V6.replace('patchVersion: "euro-office-bridge-v1/9.3.4.60/6"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/5"');
  const cut = (s, startMark, endMark) => { const i = s.indexOf(startMark); const j = s.indexOf(endMark, i); assert.ok(i > 0 && j > i, startMark); return s.slice(0, i) + s.slice(j); };
  t = cut(t, '{ id: "range.alignmentExt.write/v1"', '], };');
  t = cut(t, 'var EO_BRIDGE_WRITES = ', 'var EO_BRIDGE_HYDRATIONS = ');
  t = cut(t, 'write: function (capabilityId, args)', '}; },');
  assert.equal(t, V5);
});

test('class separation: read()/hydrate() refuse the WRITE id; write() refuses READ / HYDRATION / unknown version ids', () => {
  const env = fakeApi(); const b = bridgeOf(V6, env.api);
  assert.deepEqual(b.read(W, { sheet: 'S', cell: 'A1', indent: 1 }), { ok: false, reason: 'capability-not-advertised' });
  assert.deepEqual(b.hydrate(W, { sheet: 'S' }), { ok: false, reason: 'capability-not-advertised' });
  for (const id of ['range.alignmentExt.read/v1', 'sheet.model.ensureReady/v1', 'range.alignmentExt.write/v2', 'nope'])
    assert.deepEqual(b.write(id, { sheet: 'S', cell: 'A1', indent: 1 }), { ok: false, reason: 'capability-not-advertised' }, id);
  assert.deepEqual(env.mutations(), []);
  assert.equal(typeof bridgeOf(V5, env.api).write, 'undefined', '/5: MISSING_CAPABILITY is detectable (no write entry)');
});

test('minimal mutation: indent only / shrink only / both -- only the requested setter(s), in order, nothing else touched', () => {
  const only = (env) => env.calls.filter(c => c[0] === 'inner.call').map(c => c[1]);
  let e = fakeApi(); assert.deepEqual(write(e, { sheet: 'S', cell: 'B2', indent: 3 }), { ok: true, value: { applied: ['indent'] } });
  assert.deepEqual(only(e), ['setIndent']); assert.deepEqual(e.state, { indent: 3, shrink: false });
  e = fakeApi(); assert.deepEqual(write(e, { sheet: 'S', cell: 'B2', shrink: true }), { ok: true, value: { applied: ['shrink'] } });
  assert.deepEqual(only(e), ['setShrinkToFit']); assert.deepEqual(e.state, { indent: 0, shrink: true });
  e = fakeApi(); assert.deepEqual(write(e, { sheet: 'S', cell: 'B2', indent: 15, shrink: false }), { ok: true, value: { applied: ['indent', 'shrink'] } });
  assert.deepEqual(only(e), ['setIndent', 'setShrinkToFit']);
  // non-creating: only GetSheet + GetRange on the Api/sheet, no getter on the inner range during a write
  assert.deepEqual(e.calls.filter(c => c[0] !== 'inner.get' && c[0] !== 'inner.call').map(c => c[0]), ['Api.GetSheet', 'ws.GetRange']);
  assert.deepEqual(e.calls.filter(c => c[0] === 'inner.get').map(c => c[1]).sort(), ['setIndent', 'setIndent', 'setShrinkToFit', 'setShrinkToFit']);
});

test('write -> fresh range.alignmentExt.read/v1 is the verification (the receipt carries no observed value)', () => {
  const e = fakeApi(); const b = bridgeOf(V6, e.api);
  const r = b.write(W, { sheet: 'S', cell: 'C3', indent: 2, shrink: true });
  assert.deepEqual(Object.keys(r.value), ['applied']);
  assert.deepEqual(b.read('range.alignmentExt.read/v1', { sheet: 'S', cell: 'C3' }), { ok: true, value: { readingOrder: 0, indent: 2, shrink: true } });
});

test('invalid inputs refuse BEFORE any mutation (zero setter calls)', () => {
  const cases = [
    [undefined, 'bad-args'], [{ cell: 'A1', indent: 1 }, 'bad-args'], [{ sheet: 'S', indent: 1 }, 'bad-args'], [{ sheet: 1, cell: 'A1', indent: 1 }, 'bad-args'],
    [{ sheet: 'S', cell: 'A1' }, 'no-property'], [{ sheet: 'S', cell: 'A1', indent: 1, wrap: true }, 'unsupported-property'],
    [{ sheet: 'S', cell: 'A1', indent: 1, requester: 'x' }, 'unsupported-property'],
    ...[-1, 16, 1.5, '2', NaN, Infinity, null, true].map(v => [{ sheet: 'S', cell: 'A1', indent: v }, 'invalid-indent']),
    ...['true', 1, 0, null].map(v => [{ sheet: 'S', cell: 'A1', shrink: v }, 'invalid-shrink']),
    ...['A1:B2', 'A:A', '1:1', 'a1', 'A0', '', 'ZZZZ1', 'A1 '].map(c => [{ sheet: 'S', cell: c, indent: 1 }, 'invalid-target']),
    [{ sheet: 'X', cell: 'A1', indent: 1 }, 'sheet-not-found'],
  ];
  for (const [args, reason] of cases) { const e = fakeApi(); assert.deepEqual(write(e, args), { ok: false, reason }, JSON.stringify(args)); assert.deepEqual(e.mutations(), [], JSON.stringify(args)); }
  let e = fakeApi({ getRangeThrows: true }); assert.deepEqual(write(e, { sheet: 'S', cell: 'A1', indent: 1 }), { ok: false, reason: 'cell-not-found' });
  e = fakeApi({ noInner: true }); assert.deepEqual(write(e, { sheet: 'S', cell: 'A1', indent: 1 }), { ok: false, reason: 'range-missing' });
});

test('missing setter: refused before the FIRST mutation, also when the other requested setter exists', () => {
  let e = fakeApi({ setIndent: false }); assert.deepEqual(write(e, { sheet: 'S', cell: 'A1', indent: 1 }), { ok: false, reason: 'setIndent-missing' }); assert.deepEqual(e.mutations(), []);
  e = fakeApi({ setShrink: false }); assert.deepEqual(write(e, { sheet: 'S', cell: 'A1', indent: 1, shrink: true }), { ok: false, reason: 'setShrinkToFit-missing' });
  assert.deepEqual(e.mutations(), [], 'indent NOT applied when the shrink setter is missing');
  e = fakeApi({ setShrink: false }); assert.deepEqual(write(e, { sheet: 'S', cell: 'A1', indent: 4 }), { ok: true, value: { applied: ['indent'] } }, 'an unrequested missing setter does not block');
});

test('setter throws: controlled failure; PARTIAL_WRITE_POSSIBLE=YES is reported, never rolled back, no fallback', () => {
  let e = fakeApi({ throwIndent: true });
  assert.deepEqual(write(e, { sheet: 'S', cell: 'A1', indent: 1, shrink: true }), { ok: false, reason: 'setIndent-threw', partial: false, applied: [] });
  assert.deepEqual(e.state, { indent: 0, shrink: false });
  e = fakeApi({ throwShrink: true });
  assert.deepEqual(write(e, { sheet: 'S', cell: 'A1', indent: 5, shrink: true }), { ok: false, reason: 'setShrinkToFit-threw', partial: true, applied: ['indent'] });
  assert.deepEqual(e.state, { indent: 5, shrink: false }, 'the partial write stays -- the caller must re-observe');
  assert.deepEqual(e.mutations().map(c => c[1]), ['setIndent', 'setShrinkToFit'], 'no compensating / rollback call');
});

test('COMPAT MATRIX fact (product untouched): the current product accepts only /5 -- a /6 manifest -> PATCH_VERSION_MISMATCH', () => {
  const m6 = bridgeOf(V6, { GetSheet: () => null }).manifest();
  const caps = m6.capabilities.map(c => ({ id: c.id, version: c.version, kind: c.kind, mutation: c.mutation }));
  assert.deepEqual(classifyHandshake({ ok: true, manifest: { ...m6, capabilities: caps } }, RUNTIMES.xlsx), { ok: false, reason: 'PATCH_VERSION_MISMATCH' });
  const m5 = bridgeOf(V5, { GetSheet: () => null }).manifest();
  assert.equal(classifyHandshake({ ok: true, manifest: m5 }, RUNTIMES.xlsx).ok, true, 'positive control: /5 accepted');
});
