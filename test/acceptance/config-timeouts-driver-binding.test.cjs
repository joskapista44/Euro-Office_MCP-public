'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { buildServerTools } = require('../../src/server/server-tools.cjs');
// docx_execute_task's PUBLIC handler always refuses PRE_DISPATCH (DOCX_WRITE_ACCEPTED is empty -- no DOCX write
// intent is released yet, an orthogonal gate to this card). createDocxWriteExecution is exported for exactly
// this: direct, narrow testing of the execution wiring separately from the intent-allowlist gate (the same
// reason write-tools.cjs exports its own xlsx equivalents).
const { createDocxWriteExecution } = require('../../src/server/docx-write-tools.cjs');

/**
 * config.cjs validates `timeouts.openMs` / `timeouts.commandMs` (positive integer) but, before this change,
 * nothing downstream ever read `config.timeouts` back -- `buildServerTools` only read `config.broker` /
 * `config.playwright.path`, and `productDriver(format, { configSource, loadPlaywright })` (read-tools.cjs)
 * destructured away anything else, so the driver always ran on its own hardcoded defaults (90000/60000/15000)
 * regardless of what the config file said.
 *
 * Proves, through the EXISTING `readTool`/`writeTool.createDriver` test seams (no new seam needed): a config
 * carrying `timeouts.openMs`/`timeouts.commandMs` now reaches the driver constructor as `timeoutMs`/
 * `commandTimeoutMs`, on BOTH formats (xlsx, docx) and BOTH paths (read: xlsx_read/docx_read; write:
 * xlsx_execute_task/docx_execute_task) -- and a config WITHOUT the key leaves the driver's own default in
 * place (no `timeoutMs` key reaches the constructor call at all).
 */

const BROKER_CONFIG = (timeouts) => ({
  schemaVersion: 2,
  nextcloud: { url: 'https://nc.example' },
  broker: { url: 'http://127.0.0.1:3420/api/euro-broker' },
  ...(timeouts ? { timeouts } : {}),
});

/** A stub driver that stops the call right after construction (open() fails on purpose) -- all these tests
 * care about is the OPTIONS createDriver was called with, never whether a task completes. */
function spyCreateDriver(seen) {
  return (format, opts) => {
    seen.push({ format, timeoutMs: opts.timeoutMs, commandTimeoutMs: opts.commandTimeoutMs });
    return { open: async () => ({ ok: false, detail: 'stub-stop-here' }), close: async () => ({ ok: true }) };
  };
}

function toolsWith(config, seen) {
  return buildServerTools({
    config,
    configIdentity: 'test',
    readTool: { createDriver: spyCreateDriver(seen), loadPlaywright: () => ({}) },
    writeTool: { createDriver: spyCreateDriver(seen), loadPlaywright: () => ({}) },
  });
}

test('xlsx_read: config.timeouts reaches the driver constructor as timeoutMs/commandTimeoutMs', async () => {
  const seen = [];
  const { tools } = toolsWith(BROKER_CONFIG({ openMs: 12345, commandMs: 6789 }), seen);
  await tools.xlsx_read.handler({ fileId: '1', operation: 'workbook_overview' });
  assert.equal(seen.length, 1);
  assert.deepEqual(seen[0], { format: 'xlsx', timeoutMs: 12345, commandTimeoutMs: 6789 });
});

test('docx_read: config.timeouts reaches the driver constructor as timeoutMs/commandTimeoutMs', async () => {
  const seen = [];
  const { tools } = toolsWith(BROKER_CONFIG({ openMs: 12345, commandMs: 6789 }), seen);
  await tools.docx_read.handler({ fileId: '1', operation: 'read_text' });
  assert.equal(seen.length, 1);
  assert.deepEqual(seen[0], { format: 'docx', timeoutMs: 12345, commandTimeoutMs: 6789 });
});

test('xlsx_execute_task: config.timeouts reaches the driver constructor as timeoutMs/commandTimeoutMs', async () => {
  const seen = [];
  const { tools } = toolsWith(BROKER_CONFIG({ openMs: 12345, commandMs: 6789 }), seen);
  await tools.xlsx_execute_task.handler({ fileId: '1', task: { id: 't', steps: [{ intent: 'write_range', id: 'w1', sheet: 'Budget', range: 'F2', values: [[1]] }] } }).catch(() => null);
  assert.equal(seen.length, 1);
  assert.deepEqual(seen[0], { format: 'xlsx', timeoutMs: 12345, commandTimeoutMs: 6789 });
});

test('docx_execute_task: config.timeouts reaches the driver constructor as timeoutMs/commandTimeoutMs (createDocxWriteExecution directly -- the public handler always refuses PRE_DISPATCH, DOCX_WRITE_ACCEPTED is empty, an orthogonal gate)', async () => {
  const seen = [];
  const execute = createDocxWriteExecution({ configSource: { kind: 'BROKER' }, createDriver: spyCreateDriver(seen), loadPlaywright: () => ({}), timeoutMs: 12345, commandTimeoutMs: 6789 });
  await execute({ fileId: '1', task: { id: 't', steps: [] } }).catch(() => null);
  assert.equal(seen.length, 1);
  assert.deepEqual(seen[0], { format: 'docx', timeoutMs: 12345, commandTimeoutMs: 6789 });
});

test('NEG: without a timeouts key in the config, no timeoutMs/commandTimeoutMs reaches the driver constructor (the driver keeps its own default)', async () => {
  const seen = [];
  const { tools } = toolsWith(BROKER_CONFIG(null), seen);
  await tools.xlsx_read.handler({ fileId: '1', operation: 'workbook_overview' });
  await tools.docx_read.handler({ fileId: '1', operation: 'read_text' });
  await tools.xlsx_execute_task.handler({ fileId: '1', task: { id: 't', steps: [{ intent: 'write_range', id: 'w1', sheet: 'Budget', range: 'F2', values: [[1]] }] } }).catch(() => null);
  // docx_execute_task's public handler always refuses PRE_DISPATCH (DOCX_WRITE_ACCEPTED is empty) -- same reason
  // the positive test above uses createDocxWriteExecution directly. Here, undefined (not forwarded at all,
  // mirroring an absent config.timeouts) is the NEG case for that same layer.
  await createDocxWriteExecution({ configSource: { kind: 'BROKER' }, createDriver: spyCreateDriver(seen), loadPlaywright: () => ({}) })({ fileId: '1', task: { id: 't', steps: [] } }).catch(() => null);
  assert.equal(seen.length, 4);
  for (const s of seen) {
    assert.equal(s.timeoutMs, undefined, `${s.format}: timeoutMs must be absent, not a stale/default number, so the DRIVER's own default parameter applies`);
    assert.equal(s.commandTimeoutMs, undefined, `${s.format}: commandTimeoutMs must be absent`);
  }
});

test('NEG: a config.timeouts that sets only openMs leaves commandTimeoutMs absent (independent keys, not all-or-nothing)', async () => {
  const seen = [];
  const { tools } = toolsWith(BROKER_CONFIG({ openMs: 12345 }), seen);
  await tools.xlsx_read.handler({ fileId: '1', operation: 'workbook_overview' });
  assert.deepEqual(seen[0], { format: 'xlsx', timeoutMs: 12345, commandTimeoutMs: undefined });
});
