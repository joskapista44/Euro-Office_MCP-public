'use strict';

/**
 * GH #358: the wire form of a tool's inputSchema on `tools/list`. The step schemas repeat identical sub-schemas
 * inline (a DOCX paragraph format 4x, border objects 16x, ...), which made `tools/list` ~117 KB -- a fixed cost on
 * every request of a context-bound client. `hoistDefs` lifts every sub-schema that occurs at least twice and is
 * larger than `minBytes` into a root `$defs` table and replaces each occurrence with a local `{ "$ref": "#/$defs/dN" }`.
 *
 * Only SCHEMA positions are considered (a property's schema, `items`, the members of anyOf / oneOf / allOf, `not`,
 * an object-valued `additionalProperties`, the `$defs` bodies themselves) -- never a `properties` map as a whole, an
 * `enum` / `const` value or any other keyword's data, where a `$ref` would change the meaning.
 * The result is semantically identical: `derefDefs(hoistDefs(s))` equals `s` (test/static/schema-defs.test.cjs, every
 * shipped tool). Measured before the change (#358): the llama.cpp grammar conversion accepts the `$ref` form for all six
 * tools; other clients' `$ref` support is NOT_MEASURED. The in-process schemas (`inputSchema()`) stay inline -- only the
 * `tools/list` reply carries this form.
 */

const SCHEMA_ARRAY_KEYS = Object.freeze(['anyOf', 'oneOf', 'allOf', 'prefixItems']);
const SCHEMA_KEYS = Object.freeze(['items', 'not', 'additionalProperties', 'if', 'then', 'else', 'contains']);

/** Key-order-independent JSON text: the identity of a sub-schema. */
function canon(x) {
  if (Array.isArray(x)) return `[${x.map(canon).join(',')}]`;
  if (x && typeof x === 'object') return `{${Object.keys(x).sort().map(k => `${JSON.stringify(k)}:${canon(x[k])}`).join(',')}}`;
  return JSON.stringify(x);
}

/** Calls `fn(subSchema)` for every schema in a schema position below `s` (not `s` itself). */
function eachChild(s, fn) {
  if (!s || typeof s !== 'object' || Array.isArray(s)) return;
  if (s.properties && typeof s.properties === 'object') Object.values(s.properties).forEach(fn);
  for (const k of SCHEMA_ARRAY_KEYS) if (Array.isArray(s[k])) s[k].forEach(fn);
  for (const k of SCHEMA_KEYS) if (s[k] && typeof s[k] === 'object' && !Array.isArray(s[k])) fn(s[k]);
}

/** A copy of `s` with every schema-position child mapped through `fn` (the node's own keywords kept as they are). */
function mapChildren(s, fn) {
  if (!s || typeof s !== 'object' || Array.isArray(s)) return s;
  const out = { ...s };
  if (s.properties && typeof s.properties === 'object') out.properties = Object.fromEntries(Object.entries(s.properties).map(([k, v]) => [k, fn(v)]));
  for (const k of SCHEMA_ARRAY_KEYS) if (Array.isArray(s[k])) out[k] = s[k].map(fn);
  for (const k of SCHEMA_KEYS) if (s[k] && typeof s[k] === 'object' && !Array.isArray(s[k])) out[k] = fn(s[k]);
  return out;
}

function hoistDefs(schema, { minBytes = 150 } = {}) {
  let root = schema;
  const defs = {};
  for (;;) {
    const seen = new Map();
    const visit = node => {
      if (!node || typeof node !== 'object' || Array.isArray(node) || node.$ref !== undefined) return;
      const k = canon(node);
      const e = seen.get(k);
      if (e) e.n += 1; else seen.set(k, { n: 1, node });
      eachChild(node, visit);
    };
    eachChild(root, visit);
    Object.values(defs).forEach(d => eachChild(d, visit));
    let best = null;
    for (const [k, e] of seen) {
      if (e.n < 2 || k.length <= minBytes) continue;
      const gain = k.length * (e.n - 1);
      if (!best || gain > best.gain || (gain === best.gain && k < best.k)) best = { k, gain, node: e.node };
    }
    if (!best) break;
    const name = `d${Object.keys(defs).length}`;
    const replace = node => (node && typeof node === 'object' && !Array.isArray(node) && node.$ref === undefined && canon(node) === best.k ? { $ref: `#/$defs/${name}` } : mapChildren(node, replace));
    root = mapChildren(root, replace);
    for (const d of Object.keys(defs)) defs[d] = mapChildren(defs[d], replace);
    defs[name] = mapChildren(best.node, replace);
  }
  return Object.keys(defs).length ? { ...root, $defs: defs } : root;
}

/** The inline schema back: every local `{ "$ref": "#/$defs/x" }` replaced by its body, `$defs` dropped. */
function derefDefs(schema) {
  const defs = (schema && schema.$defs) || {};
  const walk = x => {
    if (Array.isArray(x)) return x.map(walk);
    if (!x || typeof x !== 'object') return x;
    if (typeof x.$ref === 'string' && Object.keys(x).length === 1 && x.$ref.startsWith('#/$defs/')) return walk(defs[x.$ref.slice('#/$defs/'.length)]);
    return Object.fromEntries(Object.entries(x).filter(([k]) => k !== '$defs' || x !== schema).map(([k, v]) => [k, walk(v)]));
  };
  const out = walk(schema);
  if (out && typeof out === 'object') delete out.$defs;
  return out;
}

const cache = new WeakMap();
/** The `tools/list` form of a tool's inputSchema (computed once per schema object). */
function wireSchema(schema) {
  if (!schema || typeof schema !== 'object') return schema;
  if (!cache.has(schema)) cache.set(schema, hoistDefs(schema));
  return cache.get(schema);
}

module.exports = { hoistDefs, derefDefs, wireSchema, canon };
