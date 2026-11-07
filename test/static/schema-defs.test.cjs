'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { PassThrough } = require('node:stream');

const { hoistDefs, derefDefs, wireSchema, canon } = require('../../src/server/schema-defs.cjs');
const W = require('../../src/server/write-tools.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const R = require('../../src/server/read-tools.cjs');
const { createMcpServer } = require('../../src/server/mcp.cjs');

/**
 * GH #358 acceptance 1 + 3: the `$defs` / `$ref` wire form is byte-for-byte the same schema once dereferenced, for
 * every shipped tool schema, and it is smaller. Acceptance 2 (the llama.cpp grammar probe) is a LIVE measurement on
 * the client side, recorded on the issue; it is not repeatable here.
 */

const bytes = x => JSON.stringify(x).length;
// the full DOCX step union as well (the exposed tool's allowlist may be narrower than the family set)
const docxAll = () => ({ ...D.inputSchema(), properties: { ...D.inputSchema().properties, task: { ...D.inputSchema().properties.task, properties: { ...D.inputSchema().properties.task.properties, steps: { type: 'array', items: { oneOf: Object.values(D.DOCX_STEP_SCHEMAS_ALL) } } } } } });
const SCHEMAS = () => ({ xlsx_execute_task: W.inputSchema(), docx_execute_task: D.inputSchema(), docx_all_steps: docxAll(), xlsx_read: R.inputSchema('xlsx'), docx_read: R.inputSchema('docx') });

test('every shipped tool schema: deref(hoist(s)) is the same schema (key-order-independent), and the big ones shrink', () => {
  for (const [name, s] of Object.entries(SCHEMAS())) {
    const h = hoistDefs(s);
    assert.equal(canon(derefDefs(h)), canon(s), `${name}: the dereferenced wire schema differs from the inline one`);
    assert.ok(bytes(h) <= bytes(s), `${name}: ${bytes(h)} > ${bytes(s)}`);
  }
  const x = SCHEMAS().xlsx_execute_task;
  assert.ok(bytes(hoistDefs(x)) < bytes(x) * 0.95, 'xlsx_execute_task: at least 5% smaller');
  const d = SCHEMAS().docx_all_steps;
  assert.ok(bytes(hoistDefs(d)) < bytes(d) * 0.8, 'docx (all steps): at least 20% smaller');
});

test('every $ref is local, resolvable and sits in a SCHEMA position -- never inside properties-as-a-map, enum or const', () => {
  for (const [name, s] of Object.entries(SCHEMAS())) {
    const h = hoistDefs(s);
    const defs = h.$defs || {};
    const bad = [];
    const walk = (x, path) => {
      if (Array.isArray(x)) { x.forEach((v, i) => walk(v, `${path}[${i}]`)); return; }
      if (!x || typeof x !== 'object') return;
      if (x.$ref !== undefined) {
        const parent = path.split('.').slice(-2, -1)[0] || '';
        if (Object.keys(x).length !== 1 || !String(x.$ref).startsWith('#/$defs/') || !defs[x.$ref.slice(8)]) bad.push(`${path}: ${JSON.stringify(x)}`);
        if (/(^|\])(enum|const|default|examples)$/.test(parent) || path.endsWith('.properties')) bad.push(`${path}: $ref in a data position`);
      }
      for (const [k, v] of Object.entries(x)) walk(v, `${path}.${k}`);
    };
    walk(h, name);
    assert.deepEqual(bad, [], name);
  }
});

test('RED/GREEN on a hand-built schema: a repeated object property schema is hoisted once; small or single ones stay inline', () => {
  const big = { type: 'object', properties: { a: { type: 'string', pattern: '^[A-Z]{1,3}$', description: 'x'.repeat(200) } }, additionalProperties: false };
  const s = { type: 'object', properties: { p: big, q: { ...big }, r: { type: 'integer' }, e: { enum: [{ type: 'string', pattern: '^[A-Z]{1,3}$', description: 'x'.repeat(200) }] } } };
  const h = hoistDefs(s);
  assert.deepEqual(Object.keys(h.$defs), ['d0']);
  assert.deepEqual(h.properties.p, { $ref: '#/$defs/d0' });
  assert.deepEqual(h.properties.q, { $ref: '#/$defs/d0' });
  assert.deepEqual(h.properties.r, { type: 'integer' });
  assert.deepEqual(h.properties.e, s.properties.e, 'an enum value is data, never replaced');
  assert.equal(canon(derefDefs(h)), canon(s));
  assert.deepEqual(hoistDefs({ type: 'object', properties: { r: { type: 'integer' } } }), { type: 'object', properties: { r: { type: 'integer' } } }, 'nothing to hoist -> no $defs at all');
});

test('the tools/list reply carries the wire form; the tool object keeps its inline schema', async () => {
  const big = { type: 'object', properties: { a: { type: 'string', description: 'y'.repeat(300) } } };
  const inline = { type: 'object', properties: { one: big, two: { ...big } } };
  const input = new PassThrough();
  const output = new PassThrough();
  const lines = [];
  output.on('data', d => lines.push(...String(d).split('\n').filter(Boolean)));
  const log = { info() {}, warn() {}, error() {}, debug() {} };
  const srv = createMcpServer({ tools: { t: { description: 'd', inputSchema: inline, handler: async () => ({ ok: true }) } }, info: { version: '0' }, log, input, output });
  const done = srv.start();
  input.write(`${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18' } })}\n`);
  input.write(`${JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/list' })}\n`);
  for (let i = 0; i < 50 && lines.length < 2; i += 1) await new Promise(r => setTimeout(r, 10));
  input.end();
  await done;
  const list = JSON.parse(lines.find(l => JSON.parse(l).id === 2)).result.tools[0];
  assert.deepEqual(list.inputSchema, wireSchema(inline));
  assert.deepEqual(list.inputSchema.properties.one, { $ref: '#/$defs/d0' });
  assert.deepEqual(inline.properties.one, big, 'the in-process schema is not mutated');
});
