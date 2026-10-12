'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { wireBody, decodeWire, encodeWire, ENCODER_SRC, WIRE_VERSION } = require('../../src/drivers/eurooffice/wire-codec.cjs');
const { readQuery, digestQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

/**
 * M9-2 O1 (owner GO, owner channel, 2026-10-06): the XLSX live read crosses the editor boundary as ONE string through a
 * lossless codec. Contract under test: decode(encode(x)) is deepStrictEqual to x for every value the read can carry
 * (the special numbers and undefined included), the codec's own tags cannot collide with document data, everything it
 * cannot represent exactly fails the ENCODE, and every malformed / foreign / version-mismatched input fails the DECODE.
 */

const roundTrip = v => {
  const d = decodeWire(encodeWire(v));
  assert.equal(d.ok, true, d.reason);
  return d.value;
};

test('round trip: plain JSON shapes are unchanged', () => {
  const v = { ok: true, sheets: [{ name: 'S', index: 0, used: 'A1:D20' }], cells: { 'S\u0000A1': { v: 'x', f: '', n: 1.5 } }, nested: [[1, [2, [3]]], {}], s: '', z: 0, neg: -3, big: 1e300, tiny: 5e-324 };
  assert.deepStrictEqual(roundTrip(v), v);
});

test('special values: undefined, NaN, +Infinity, -Infinity and -0 survive exactly (also as keys’ values and in arrays)', () => {
  const v = { u: undefined, nan: NaN, pinf: Infinity, ninf: -Infinity, nz: -0, arr: [undefined, NaN, Infinity, -Infinity, -0, 0], deep: { x: { y: [undefined] } } };
  const back = roundTrip(v);
  assert.deepStrictEqual(back, v);
  assert.ok(Object.prototype.hasOwnProperty.call(back, 'u'), 'an undefined-valued key is kept, not dropped');
  assert.ok(Object.is(back.nz, -0) && Object.is(back.arr[4], -0) && Object.is(back.arr[5], 0), '-0 and 0 stay distinct');
  assert.ok(Number.isNaN(back.nan) && back.pinf === Infinity && back.ninf === -Infinity);
  assert.equal(back.arr.length, 6);
  assert.ok(1 in back.arr && 0 in back.arr, 'no hole: an undefined element is a real element');
});

test('NEGATIVE CONTROL: plain JSON would lose exactly these values (the reason the codec exists)', () => {
  const v = { u: undefined, nan: NaN, nz: -0, arr: [undefined] };
  assert.notDeepStrictEqual(JSON.parse(JSON.stringify(v)), v);
});

test('tag collision: document keys and values that LOOK like tags or escapes come back verbatim', () => {
  const v = { $: 'u', $$: 'nan', $x: 1, $$y: { $: '-0' }, '$eo-wire': 1, '$eo-wire-error': 'x', data: '$', list: [{ $: '+inf' }, '$', { $$$: 2 }], 'S\u0000$A$1': '$' };
  assert.deepStrictEqual(roundTrip(v), v);
});

test('object shape: __proto__ / constructor as DATA keys stay own data properties, never a prototype change', () => {
  const v = JSON.parse('{"__proto__": {"x": 1}, "constructor": "c", "toString": 2}');
  const back = roundTrip(v);
  assert.deepStrictEqual(back, v);
  assert.equal(Object.getPrototypeOf(back), Object.prototype);
  assert.ok(Object.prototype.hasOwnProperty.call(back, '__proto__'));
});

test('array vs object stays unambiguous ({0:..} is not [..], [] is not {})', () => {
  const v = { a: [], o: {}, ao: { 0: 'a', 1: 'b', length: 2 }, arr: ['a', 'b'] };
  const back = roundTrip(v);
  assert.deepStrictEqual(back, v);
  assert.ok(Array.isArray(back.arr) && !Array.isArray(back.ao) && Array.isArray(back.a) && !Array.isArray(back.o));
});

test('ENCODE fails closed on what it cannot represent exactly -- never dropped or approximated', () => {
  const sparse = [1]; sparse[3] = 2;
  const extra = [1]; extra.k = 2;
  const cyc = {}; cyc.self = cyc;
  class K { constructor() { this.a = 1; } }
  for (const [label, v] of [['function', { f() {} }], ['symbol', { s: Symbol('x') }], ['bigint', { b: 1n }], ['Date', { d: new Date(0) }], ['Map', { m: new Map() }], ['class instance', { k: new K() }], ['sparse array', { sparse }], ['array extra key', { extra }], ['cycle', cyc], ['symbol key', { [Symbol('k')]: 1 }]]) {
    const d = decodeWire(encodeWire(v));
    assert.equal(d.ok, false, label);
    assert.match(d.reason, /^wire-encode-error:/, label);
  }
});

test('a shared (non-cyclic) reference is encoded twice -- the same values, not a cycle error', () => {
  const shared = { a: [1, NaN] };
  const back = roundTrip({ x: shared, y: shared });
  assert.deepStrictEqual(back, { x: { a: [1, NaN] }, y: { a: [1, NaN] } });
});

test('DECODE fails closed: non-string, parse error, envelope, version, unknown tag, tag with extra keys, unescaped $ key', () => {
  const env = data => JSON.stringify({ '$eo-wire': WIRE_VERSION, data });
  const cases = [
    [{ ok: true }, 'wire-not-a-string'],
    [undefined, 'wire-not-a-string'],
    ['{not json', 'wire-parse-error'],
    ['[1,2]', 'wire-envelope-shape'],
    ['"x"', 'wire-envelope-shape'],
    [JSON.stringify({ data: 1 }), 'wire-envelope-shape'],
    [JSON.stringify({ '$eo-wire': 1, data: 1, extra: 1 }), 'wire-envelope-shape'],
    [JSON.stringify({ '$eo-wire': 2, data: 1 }), 'wire-version-mismatch'],
    [JSON.stringify({ '$eo-wire': '1', data: 1 }), 'wire-version-mismatch'],
    [env({ $: 'bogus' }), /^wire-decode:unknown tag/],
    [env({ $: 1 }), /^wire-decode:unknown tag/],
    [env({ $: 'toString' }), /^wire-decode:unknown tag/],
    [env({ $: 'u', x: 1 }), /^wire-decode:tag with extra keys/],
    [env({ $x: 1 }), /^wire-decode:unescaped \$ key/],
    [env([{ ok: { $: 'nope' } }]), /^wire-decode:unknown tag/],
    [JSON.stringify({ '$eo-wire-error': 'cycle' }), /^wire-encode-error:cycle/],
  ];
  for (const [input, want] of cases) {
    const d = decodeWire(input);
    assert.equal(d.ok, false, String(input));
    assert.equal(d.value, undefined, 'no partial value on failure');
    if (typeof want === 'string') assert.equal(d.reason, want, String(input)); else assert.match(d.reason, want, String(input));
  }
});

test('wireBody: the read body is embedded UNCHANGED and the result is the encoder over it (evaluated like the editor does)', () => {
  const body = "try { return { ok: true, v: undefined, n: NaN, z: -0 }; } catch (err) { return { ok: false, error: String(err) }; }";
  const w = wireBody(body);
  assert.ok(w.includes(body));
  const out = new Function(w)();
  assert.equal(typeof out, 'string');
  assert.deepStrictEqual(decodeWire(out).value, { ok: true, v: undefined, n: NaN, z: -0 });
});

test('the wire wrapper adds no editor Api call (the read-only gate of the bodies is unaffected)', () => {
  const calls = [...ENCODER_SRC.matchAll(/\.([A-Z]\w*)\s*\(/g)].map(m => m[1]);
  assert.deepEqual(calls, []);
  assert.doesNotMatch(ENCODER_SRC, /\bApi\b/);
  // positive control: the extractor sees capitalised calls when there are any
  assert.deepEqual([...'a.GetValue(); b.SetX(1)'.matchAll(/\.([A-Z]\w*)\s*\(/g)].map(m => m[1]), ['GetValue', 'SetX']);
});

test('the readQuery / digestQuery bodies themselves are NOT changed by O1 (the wire wraps, it does not rewrite)', () => {
  for (const b of [readQuery([{ sheet: 'S', range: 'A1:B2' }]), digestQuery([{ sheet: 'S', range: 'A1:B2' }])]) {
    assert.doesNotMatch(b, /__eoWire/);
    assert.ok(wireBody(b).includes(b));
  }
});

/** The REAL driver on a stand-in editor frame: `answer(body)` is what the editor returns for a command body. */
async function readThrough(answer, opts = {}) {
  const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
  const build = '9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b';
  const frameObj = {
    url: () => `https://ds/${build}/web-apps/apps/spreadsheeteditor/main/index.html`,
    evaluate: async (fn, arg) => {
      if (arg && typeof arg.body === 'string' && /GetActiveSheet/.test(arg.body) && !/__eoWire/.test(arg.body)) return { ok: true, result: 'S' };
      if (arg && typeof arg.body === 'string') return { ok: true, result: answer(arg.body) };
      return { ok: true };
    },
  };
  const page = { evaluate: async () => ({ ok: true, detail: 'stub' }), goto: async () => {}, waitForFunction: async () => {}, waitForTimeout: async () => {}, frames: () => [frameObj] };
  const pw = { chromium: { launch: async () => ({ newPage: async () => page, close: async () => {} }) } };
  const realFetch = global.fetch;
  global.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: { documentServerUrl: 'https://ds/', document: { key: 'k' } } } }) });
  try {
    const d = drv.createEuroOfficeXlsxDriver({ url: 'https://nc', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: true, pw }), timeoutMs: 1000, ...opts });
    assert.equal((await d.open({ fileId: '1' })).ok, true);
    const r = await d.readLive(null, { windows: [] });
    await d.close();
    return r;
  } finally { global.fetch = realFetch; }
}

const CANNED = { ok: true, sheets: [{ name: 'S', index: 0, used: null }], cells: [], truncated: false, unsupported: [{ field: 'x', reason: undefined }], objects: {} };
const probe = { ok: true, sheets: [{ name: 'S', used: null }] };
const isProbe = b => !/__eoWire/.test(b) && /GetUsedRange/.test(b) && !/LIMIT/.test(b);

test('readLive (wire, the default): the decoded payload equals the object payload, undefined kept', async () => {
  const r = await readThrough(b => (isProbe(b) ? probe : encodeWire(CANNED)));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepStrictEqual(r.payload.unsupported, CANNED.unsupported);
  assert.ok(Object.prototype.hasOwnProperty.call(r.payload.unsupported[0], 'reason'));
});

test('readLive FAILS CLOSED on a bad wire answer -- a non-string, a parse error, a version mismatch -- never a partial payload', async () => {
  for (const bad of [CANNED, '{oops', JSON.stringify({ '$eo-wire': 99, data: CANNED })]) {
    const r = await readThrough(b => (isProbe(b) ? probe : bad));
    assert.equal(r.ok, false);
    assert.equal(r.payload, undefined);
    assert.match(String(r.detail), /^wire-/);
  }
});

test('the TEST SEAM readTransport "object" is the pre-O1 path (the parity harness compares the two); an unknown mode fails', async () => {
  const seen = [];
  const r = await readThrough(b => { seen.push(/__eoWire/.test(b)); return isProbe(b) ? probe : CANNED; }, { readTransport: () => 'object' });
  assert.equal(r.ok, true);
  assert.ok(!seen.includes(true), 'no wire body sent on the object path');
  const u = await readThrough(b => (isProbe(b) ? probe : CANNED), { readTransport: 'smoke-signals' });
  assert.equal(u.ok, false);
  assert.match(u.detail, /^read-transport-unknown/);
});
