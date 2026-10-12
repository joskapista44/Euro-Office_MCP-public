'use strict';

/**
 * M9-2 O1 (owner GO, owner channel, 2026-10-06): the lossless WIRE codec of the XLSX live read.
 *
 * MEASURED (docs/evidence/M9-2-READ-ANALYSIS-20261006.md): 85-90 % of an XLSX live read was the transfer of the
 * payload as an object from the editor to node (469 ms vs 63 ms for the same payload as one JSON string). The read
 * bodies (readQuery / digestQuery) are UNCHANGED; only their result crosses the editor/driver boundary as one string
 * and is decoded here into the same value. Representation only -- no verification axis, scope or read changes.
 *
 * Plain JSON would lose information, so these values are TAGGED, never coerced:
 *   undefined -> {"$":"u"}   NaN -> {"$":"nan"}   +Infinity -> {"$":"+inf"}   -Infinity -> {"$":"-inf"}   -0 -> {"$":"-0"}
 * A tag is an object whose ONLY key is "$". A legitimate object key that starts with "$" is escaped by one more "$"
 * ("$" -> "$$", "$x" -> "$$x"), so no document value can look like a tag, and a decoded "$..." key without the escape
 * is malformed. Everything the codec cannot represent exactly -- a function, symbol, bigint, a non-plain object (Date,
 * Map, a class instance), a sparse array or an array with extra keys, a symbol key, a cycle -- makes the ENCODER fail
 * (the read fails); it is never dropped or approximated.
 *
 * FAIL-CLOSED decode: a non-string, a parse error, a wrong envelope / version, an unknown tag, a tag object with other
 * keys, an unescaped "$" key, or an encoder-side error all give { ok: false, reason } -- the live read FAILS. There is
 * no partial or default interpretation of a bad payload.
 */

const WIRE_VERSION = 1;
const ENVELOPE_KEY = '$eo-wire';
const ERROR_KEY = '$eo-wire-error';

/**
 * The editor-side encoder, as SOURCE: it runs inside the editor (callCommand), the unit tests evaluate the same text.
 * Only lowercase builtins (JSON.stringify, Object.keys, Array.isArray, ...): no editor Api method is called.
 */
const ENCODER_SRC = `
function __eoWireEncode(root) {
  const stack = [];
  const enc = v => {
    if (v === undefined) return { $: 'u' };
    if (v === null || typeof v === 'boolean' || typeof v === 'string') return v;
    if (typeof v === 'number') {
      if (v !== v) return { $: 'nan' };
      if (v === Infinity) return { $: '+inf' };
      if (v === -Infinity) return { $: '-inf' };
      if (v === 0 && 1 / v < 0) return { $: '-0' };
      return v;
    }
    if (typeof v !== 'object') throw new Error('unsupported type ' + typeof v);
    if (stack.indexOf(v) !== -1) throw new Error('cycle');
    if (Object.getOwnPropertySymbols(v).length) throw new Error('symbol key');
    stack.push(v);
    let out;
    if (Array.isArray(v)) {
      if (Object.keys(v).length !== v.length) throw new Error('sparse array or extra array keys');
      out = [];
      for (let i = 0; i < v.length; i += 1) out.push(enc(v[i]));
    } else {
      const p = Object.getPrototypeOf(v);
      if (p !== null && Object.getPrototypeOf(p) !== null) throw new Error('non-plain object');
      out = {};
      for (const k of Object.keys(v)) Object.defineProperty(out, k.charAt(0) === '$' ? '$' + k : k, { value: enc(v[k]), enumerable: true, writable: true, configurable: true });
    }
    stack.pop();
    return out;
  };
  try { return JSON.stringify({ '${ENVELOPE_KEY}': ${WIRE_VERSION}, data: enc(root) }); }
  catch (err) { return JSON.stringify({ '${ERROR_KEY}': String(err && err.message || err) }); }
}
`;

/** A read body whose result crosses the boundary as one wire string (the body itself is unchanged). */
function wireBody(body) {
  return `const __eoWireQuery = function () {\n${body}\n};\n${ENCODER_SRC}\nreturn __eoWireEncode(__eoWireQuery());`;
}

const TAGS = Object.freeze({ u: undefined, nan: NaN, '+inf': Infinity, '-inf': -Infinity, '-0': -0 });

function isPlainObject(v) {
  return v !== null && typeof v === 'object' && !Array.isArray(v);
}

function dec(v, path) {
  if (v === null || typeof v === 'boolean' || typeof v === 'string' || typeof v === 'number') return v;
  if (Array.isArray(v)) return v.map((x, i) => dec(x, `${path}[${i}]`));
  if (!isPlainObject(v)) throw new Error(`unexpected value at ${path}`);
  const keys = Object.keys(v);
  if (keys.includes('$')) {
    if (keys.length !== 1) throw new Error(`tag with extra keys at ${path}`);
    if (typeof v.$ !== 'string' || !Object.prototype.hasOwnProperty.call(TAGS, v.$)) throw new Error(`unknown tag at ${path}`);
    return TAGS[v.$];
  }
  const out = {};
  for (const k of keys) {
    let name = k;
    if (k.charAt(0) === '$') {
      if (k.charAt(1) !== '$') throw new Error(`unescaped $ key at ${path}`);
      name = k.slice(1);
    }
    Object.defineProperty(out, name, { value: dec(v[k], `${path}.${name}`), enumerable: true, writable: true, configurable: true });
  }
  return out;
}

/** The node side: one wire string -> { ok: true, value } | { ok: false, reason } (fail-closed, never partial). */
function decodeWire(text) {
  if (typeof text !== 'string') return { ok: false, reason: 'wire-not-a-string' };
  let env;
  try { env = JSON.parse(text); } catch (err) { return { ok: false, reason: 'wire-parse-error' }; }
  if (!isPlainObject(env)) return { ok: false, reason: 'wire-envelope-shape' };
  const keys = Object.keys(env).sort();
  if (keys.length === 1 && keys[0] === ERROR_KEY) return { ok: false, reason: `wire-encode-error:${String(env[ERROR_KEY]).slice(0, 120)}` };
  if (keys.length !== 2 || keys[0] !== ENVELOPE_KEY || keys[1] !== 'data') return { ok: false, reason: 'wire-envelope-shape' };
  if (env[ENVELOPE_KEY] !== WIRE_VERSION) return { ok: false, reason: 'wire-version-mismatch' };
  try { return { ok: true, value: dec(env.data, 'data') }; } catch (err) { return { ok: false, reason: `wire-decode:${err.message}` }; }
}

/** The SAME encoder text, evaluated in node: for the unit tests and the fakes that stand in for the editor. */
const encodeWire = new Function(`${ENCODER_SRC}\nreturn __eoWireEncode;`)();

module.exports = { wireBody, decodeWire, encodeWire, ENCODER_SRC, WIRE_VERSION };
