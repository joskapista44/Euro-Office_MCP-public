'use strict';

/**
 * GH #76 follow-up: the reporter's 4-step NAD write took 118.9 s while the same shape on a disposable copy took 7.2 s,
 * and the product log (starting / stopping only) could not say where the time went. Every xlsx_execute_task now logs its
 * phase timings and every tools/call its wall time to the per-UID log -- numbers and codes only, never content.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { PassThrough } = require('node:stream');

const { harness, req } = require('../fakes/write-wiring.cjs');
const { createMcpServer } = require('../../src/server/mcp.cjs');

const STEP = { intent: 'write_range', sheet: 'Budget', range: 'F1:G2', values: [['secret-cell-text', 1], ['y', 2]] };

test('a write reports its phase timings once: numbers and codes, no cell content', async () => {
  const seen = [];
  const h = harness({ onTiming: t => seen.push(t) });
  const r = await h.tool.handler(req([STEP, { ...STEP, id: 'b', range: 'H1:I2' }]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(seen.length, 1);
  const t = seen[0];
  assert.deepEqual([t.steps, t.ok, t.persistence], [2, true, 'FORCESAVE_CONFIRMED']);
  for (const k of ['sessionOpenMs', 'readyMs', 'taskElapsedMs', 'totalElapsedMs', 'freshReads', 'commands', 'persistencePolls']) assert.equal(typeof t[k], 'number', k);
  assert.equal(t.commands, 2);
  assert.doesNotMatch(JSON.stringify(t), /secret-cell-text/);
});

test('a failing timing hook never changes the answer', async () => {
  const h = harness({ onTiming: () => { throw new Error('log disk full'); } });
  const r = await h.tool.handler(req([STEP]));
  assert.equal(r.ok, true, JSON.stringify(r));
});

test('tools/call logs the tool, its wall time and outcome code -- not its arguments', async () => {
  const lines = [];
  const log = { info: (msg, f) => lines.push([msg, f]), error() {}, warn() {}, debug() {} };
  const input = new PassThrough();
  const output = new PassThrough();
  let buf = '';
  const got = new Promise(res => output.on('data', d => { buf += d; if (buf.includes('\n')) res(); }));
  const tools = { echo: { description: 'e', inputSchema: { type: 'object' }, handler: async () => ({ ok: false, code: 'REFUSED_PRE_MUTATION' }) } };
  const done = createMcpServer({ tools, info: { version: 't' }, log, input, output }).start();
  input.write(`${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'echo', arguments: { secret: 'secret-arg' } } })}\n`);
  await got;
  input.end();
  await done;
  const call = lines.find(([m]) => m === 'tool call');
  assert.ok(call, JSON.stringify(lines));
  assert.equal(call[1].tool, 'echo');
  assert.equal(typeof call[1].ms, 'number');
  assert.deepEqual([call[1].ok, call[1].code], [false, 'REFUSED_PRE_MUTATION']);
  assert.doesNotMatch(JSON.stringify(lines), /secret-arg/);
});
