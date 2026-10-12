'use strict';

/**
 * M11.4a.1 acceptance (offline): the WITHHELD product WRITE tool `xlsx_execute_task` -- handler + strict schema +
 * allowlist / refusal + pre-dispatch security + canonical task adapter (owner decision, owner channel 2026-10-10).
 *
 * The executor is NOT wired in this step: a refused task must never reach the `execute` seam (the M11.4a.2 wiring
 * point), and a valid task reaches it exactly once, carrying the canonical task verbatim. The seam used here counts
 * what the M11.4a.2 path would touch (broker, browser, driver, executor, write); global fetch is counted too.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { PassThrough } = require('node:stream');

const W = require('../../src/server/write-tools.cjs');
const { buildServerTools } = require('../../src/server/server-tools.cjs');
const { createMcpServer } = require('../../src/server/mcp.cjs');
const { exposedTools, EXPOSED_CAPABILITY_TOOLS, TOOL_ROWS } = require('../../src/server/tools.cjs');
const { xlsxAdapter } = require('../../src/adapters/xlsx/adapter.cjs');

const CONFIG = { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: 'http://127.0.0.1:9/api/euro-broker', instance: 'production', timeoutMs: 1000 } };
const { FILE } = require('../fakes/write-refusals.cjs');

const { VALID, REFUSALS, req, wr } = require('../fakes/write-refusals.cjs');

/** The registry tool with a counting M11.4a.2-shaped execute seam, plus a counted global fetch. */
function harness() {
  const n = { broker: 0, browser: 0, driver: 0, executor: 0, write: 0, fetch: 0, tasks: [] };
  const execute = async v => { n.broker += 1; n.browser += 1; n.driver += 1; n.executor += 1; n.write += v.task.steps.length; n.tasks.push(v.task); return { ok: true, stage: 'SEAM' }; };
  const { registry } = buildServerTools({ config: CONFIG, configIdentity: 'fixture', writeTool: { execute } });
  return { n, tool: registry.xlsx_execute_task };
}

async function withCountedFetch(n, fn) {
  const real = globalThis.fetch;
  globalThis.fetch = async () => { n.fetch += 1; throw new Error('no network in this test'); };
  try { return await fn(); } finally { globalThis.fetch = real; }
}

const ZERO = n => ({ BROKER_CALL: n.broker + n.fetch, BROWSER_LAUNCH: n.browser, DRIVER_OPEN: n.driver, EXECUTOR_CALL: n.executor, WRITE: n.write });
const NONE = { BROKER_CALL: 0, BROWSER_LAUNCH: 0, DRIVER_OPEN: 0, EXECUTOR_CALL: 0, WRITE: 0 };

test('ALLOWLIST: exactly the six XLSX intents; DOCX_WRITE_ALLOWLIST empty; no DOCX handler', () => {
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST], ['write_range', 'insert_rows', 'delete_rows', 'rename_sheet', 'unmerge_range', 'set_chart_title']);
  assert.equal(Object.isFrozen(W.XLSX_WRITE_ALLOWLIST), true);
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], []);
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.throws(() => W.createWriteTool('docx'), /no WRITE handler for docx/);
  const { registry } = buildServerTools({ config: CONFIG, configIdentity: 'fixture' });
  assert.equal(registry.docx_execute_task, undefined, 'no docx_execute_task handler is built');
  assert.equal(typeof registry.xlsx_execute_task.handler, 'function');
});

test('TEMPORARY_BOUNDARY: the legacy set equals the M11.4a.0 census; no allowlisted intent depends on a boundary', () => {
  const census = JSON.parse(fs.readFileSync(path.join(__dirname, '../../docs/evidence/m11-4a0/census.json'), 'utf8'));
  const fromCensus = Object.values(census.legacy_dependents).flat().map(x => x.split(':')[1]).sort();
  assert.deepEqual(Object.keys(W.TEMPORARY_BOUNDARY).sort(), fromCensus);
  for (const intent of W.XLSX_WRITE_ALLOWLIST) {
    assert.equal(Object.prototype.hasOwnProperty.call(W.TEMPORARY_BOUNDARY, intent), false, intent);
    const row = census.rows.find(r => r.FORMAT === 'XLSX' && r.INTENT === intent);
    assert.equal(row.PRODUCT_RELEASE_CLASS, 'ACCEPTED_WRITE', intent);
    assert.equal(row.DIRECT_INTERNAL_DEPENDENCY, 'none', intent);
  }
});

test('WRITE_SCHEMA: top level { fileId, task }, one explicit step schema per intent, additionalProperties false everywhere', () => {
  const s = W.inputSchema();
  assert.deepEqual(Object.keys(s.properties), ['fileId', 'task']);
  assert.deepEqual(s.required, ['fileId', 'task']);
  assert.deepEqual(Object.keys(s.properties.task.properties), ['id', 'steps']);
  const steps = s.properties.task.properties.steps.items.oneOf;
  assert.deepEqual(steps.map(x => x.properties.intent.const), [...W.XLSX_WRITE_ALLOWLIST]);
  const objects = [];
  const walk = x => { if (!x || typeof x !== 'object') return; if (x.type === 'object') objects.push(x); Object.values(x).forEach(walk); };
  walk(s);
  assert.ok(objects.length >= 9);
  for (const o of objects) assert.equal(o.additionalProperties, false, JSON.stringify(Object.keys(o.properties || {})));
  // no identity / credential / route field is offered anywhere in the schema
  const keys = [];
  const keysOf = x => { if (!x || typeof x !== 'object') return; if (x.properties) keys.push(...Object.keys(x.properties)); Object.values(x).forEach(keysOf); };
  keysOf(s);
  for (const k of ['requester', 'user', 'uid', 'fileOwner', 'credential', 'password', 'token', 'keyName', 'brokerUrl', 'url', 'identity', 'strategy', 'preconditions']) {
    assert.equal(keys.includes(k), false, k);
  }
});

test('VALID_SINGLE_STEP: each of the six -> accepted, planned by the canonical planner, task passed verbatim, execute once', async () => {
  for (const [intent, st] of Object.entries(VALID)) {
    const h = harness();
    const args = req([st]);
    const v = W.validateWriteRequest(args);
    assert.equal(v.ok, true, `${intent}: ${JSON.stringify(v)}`);
    assert.equal(v.plannedOperations, 1);
    assert.deepEqual(v.task, args.task, `${intent}: the canonical task is passed on verbatim`);
    assert.equal(xlsxAdapter.planOperations(v.task).ok, true, `${intent}: the executor's own planner accepts the same task`);
    const r = await withCountedFetch(h.n, () => h.tool.handler(args));
    assert.equal(r.stage, 'SEAM');
    assert.equal(h.n.executor, 1);
    assert.deepEqual(h.n.tasks[0], args.task);
    assert.equal(h.n.fetch, 0);
  }
});

test('VALID_MULTI_STEP: allowlisted combinations, every step schema-valid -> accepted', () => {
  const combos = {
    two_disjoint_writes: [wr('A1:B1', [['a', 1]], 'w1'), wr('D5', [['x']], 'w2')],
    write_and_unmerge: [wr('A1', [['a']], 'w1'), { ...VALID.unmerge_range, id: 'u1' }],
    write_and_chart_title: [wr('A1', [['a']], 'w1'), { ...VALID.set_chart_title, id: 'c1' }],
    unmerge_chart_write: [{ ...VALID.unmerge_range, id: 'u1' }, { ...VALID.set_chart_title, id: 'c1' }, wr('F1:F2', [[1], [2]], 'w1')],
  };
  for (const [name, steps] of Object.entries(combos)) {
    const v = W.validateWriteRequest(req(steps));
    assert.equal(v.ok, true, `${name}: ${JSON.stringify(v)}`);
    assert.equal(v.plannedOperations, steps.length, name);
  }
});

test('DEFAULT WIRING (M11.4a.2): with no browser runtime configured a valid task stops at RUNTIME -- no broker, no driver', async () => {
  // M11.4a.1 asserted EXECUTION_NOT_WIRED here; since M11.4a.2 the handler runs the canonical executor path, and an
  // unconfigured runtime (playwright.path absent) is refused before the broker or a browser is touched
  const { registry } = buildServerTools({ config: CONFIG, configIdentity: 'fixture' });
  const n = { fetch: 0 };
  const r = await withCountedFetch(n, () => registry.xlsx_execute_task.handler(req([VALID.write_range])));
  assert.equal(r.ok, false);
  assert.equal(r.stage, 'RUNTIME');
  assert.equal(r.code, 'BROWSER_RUNTIME_UNAVAILABLE');
  assert.equal(r.dispatched, false);
  assert.equal(n.fetch, 0);
});

test('PRE_DISPATCH_REFUSALS: the refusal matrix -- whole task refused, BROKER/BROWSER/DRIVER/EXECUTOR/WRITE = 0', async () => {
  const seen = new Set();
  for (const [label, args, code] of REFUSALS) {
    const h = harness();
    const r = await withCountedFetch(h.n, () => h.tool.handler(JSON.parse(JSON.stringify(args))));
    assert.equal(r.ok, false, `${label}: ${JSON.stringify(r)}`);
    assert.equal(r.code, code, `${label}: ${r.code} -- ${r.reason}`);
    assert.equal(r.stage, 'PRE_DISPATCH', label);
    assert.equal(r.dispatched, false, label);
    assert.equal(r.wholeTask, true, label);
    assert.deepEqual(ZERO(h.n), NONE, label);
    seen.add(code);
  }
  for (const c of ['REFUSED_UNKNOWN_INTENT', 'REFUSED_NOT_ALLOWLISTED', 'REFUSED_TEMPORARY_BOUNDARY', 'REFUSED_CROSS_FORMAT', 'REFUSED_READ_OPERATION',
    'REFUSED_SHAPE', 'REFUSED_COMPOSITION', 'MALFORMED_TASK', 'EMPTY_TASK', 'LIMIT', 'CREDENTIAL_FIELD', 'REQUESTER_OVERRIDE', 'ROUTE_OVERRIDE', 'UNKNOWN_FIELD']) {
    assert.ok(seen.has(c), `refusal class covered: ${c}`);
  }
});

test('SECURITY_FIELDS: a credential / requester name is reported as such, before the unknown-field check, at any depth', () => {
  const r1 = W.validateWriteRequest(req([wr('A1', [[{ password: 'x' }]])]));
  assert.equal(r1.code, 'CREDENTIAL_FIELD');
  assert.match(r1.path, /steps\[0\]\.values\[0\]\[0\]\.password$/);
  const r2 = W.validateWriteRequest({ fileId: FILE, task: { steps: [{ ...VALID.insert_rows, anchor: { column: 'A', value: 'x', onBehalfOf: 'maintainer' } }] } });
  assert.equal(r2.code, 'REQUESTER_OVERRIDE');
  const r3 = W.validateWriteRequest({ fileId: FILE, task: { steps: [VALID.write_range] }, __proto__: { token: 'x' } });
  assert.equal(r3.ok, true, 'an inherited (non-own) key is not an argument');
  assert.equal(W.validateWriteRequest(JSON.parse(`{"fileId":"${FILE}","task":{"steps":[${JSON.stringify(VALID.write_range)}]},"__proto__":{"token":"x"}}`)).code, 'CREDENTIAL_FIELD', 'an OWN __proto__ key (JSON) is walked like any other');
});

test('SCHEMA_PLANNER_CONSISTENCY: every schema-accepted single step is accepted by the canonical planner; the planner-only value rules refuse as REFUSED_SHAPE', () => {
  for (const st of Object.values(VALID)) assert.equal(xlsxAdapter.planOperations({ steps: [st] }).ok, true, st.intent);
  for (const text of ['=x', '+1', '-1', '@x', "'x", '12', '1.5', '1e3', '10%', 'TRUE', 'false', '2026-10-10', '10/10', '12:30', '9:15 pm']) {
    const r = W.validateWriteRequest(req([wr('A1', [[text]])]));
    assert.equal(r.code, 'REFUSED_SHAPE', text);
    assert.equal(xlsxAdapter.planOperations({ steps: [wr('A1', [[text]])] }).ok, false, `the planner refuses ${text} too`);
  }
});

test('RESULT_MAPPING: FinalTaskResult -> MCP result with the canonical statuses only', () => {
  const v = W.validateWriteRequest(req([VALID.write_range]));
  const final = (over = {}) => ({ kind: 'FinalTaskResult', ok: true, editOk: true, outcome: 'task-complete', documentOutcome: 'SUCCESS', mutationCertainty: 'CERTAIN',
    authority: 'LIVE_VERIFY', noOp: false, persistenceOutcome: 'SUCCESS', persistenceState: 'FORCESAVE_CONFIRMED', saveEvidence: { reason: 'forcesave-confirmed' },
    cleanupOutcome: 'SUCCESS', openFailure: null,
    execution: { appliedOperations: [{ operationId: 's0', status: 'APPLIED', verification: 'LIVE_VERIFY' }], operationAccounting: { planned: 1 } }, ...over });
  const applied = W.mapExecutionResult(final(), v);
  assert.equal(applied.ok, true);
  assert.deepEqual(applied.operations, [{ id: 's0', status: 'APPLIED', verification: 'LIVE_VERIFY', l1: null, l3: null }]); // M11.4a.2: + the L1 / L3 verdicts
  assert.equal(applied.persistence.state, 'FORCESAVE_CONFIRMED');
  assert.equal(applied.persistence.storedContentVerified, false);
  const retry = W.mapExecutionResult(final({ noOp: true, persistenceState: 'NOT_REQUIRED', persistenceOutcome: 'NOT_REQUIRED',
    execution: { appliedOperations: [{ operationId: 's0', status: 'SATISFIED', verification: 'LIVE_VERIFY' }] } }), v);
  assert.equal(retry.noOp, true);
  assert.equal(retry.operations[0].status, 'SATISFIED');
  const unknown = W.mapExecutionResult(final({ ok: false, editOk: false, mutationCertainty: 'UNCERTAIN', persistenceState: 'UNCERTAIN_NOT_PERSISTED',
    execution: { appliedOperations: [{ operationId: 's0', status: 'DISPATCHED_OUTCOME_UNKNOWN', verification: 'DISPATCH_ONLY' }] } }), v);
  assert.equal(unknown.ok, false);
  assert.equal(unknown.operations[0].status, 'DISPATCHED_OUTCOME_UNKNOWN');
  const rejected = W.mapExecutionResult(final({ ok: false, outcome: 'task-save-not-confirmed', persistenceState: 'EDITOR_CLEAN', saveEvidence: { reason: 'barrier:completion-not-successful' } }), v);
  assert.equal(rejected.persistence.error, 'STORAGE_WRITE_REJECTED');
  assert.equal(rejected.code, 'STORAGE_WRITE_REJECTED');
  assert.equal(rejected.ok, false);
  const statuses = new Set(['APPLIED', 'SATISFIED', 'FAILED', 'DISPATCHED_OUTCOME_UNKNOWN', 'SKIPPED_DUE_TO_PRIOR_FAILURE']);
  for (const r of [applied, retry, unknown, rejected]) for (const o of r.operations) assert.ok(statuses.has(o.status));
});

test('ERROR_MAPPING: broker denial passes through (ACCESS_DENIED), read-only config -> TARGET_READ_ONLY, never APPLIED', () => {
  const v = W.validateWriteRequest(req([VALID.write_range]));
  const denied = W.mapExecutionResult({ openFailure: { ok: false, code: 'ACCESS_DENIED', detail: 'PRODUCT_OCS_CONFIG_DENY_LEVEL' } }, v);
  assert.deepEqual([denied.ok, denied.stage, denied.code, denied.dispatched], [false, 'OPEN', 'ACCESS_DENIED', false]);
  assert.equal(denied.operations, undefined);
  assert.equal(W.mapExecutionResult({ openFailure: { ok: false, code: 'ACCESS_DENIED', detail: 'broker PRODUCT_OCS_CONFIG_DENY_CONFIG_BELOW_REQUEST' } }, v).code, 'TARGET_READ_ONLY');
  assert.equal(W.mapExecutionResult({ openFailure: { ok: false, detail: 'x' } }, v).code, 'OPEN_FAILED');
  assert.deepEqual(Object.keys(W.WRITE_ERRORS), ['ACCESS_DENIED', 'TARGET_READ_ONLY', 'STORAGE_WRITE_REJECTED']);
});

/** An MCP session over in-memory streams. */
function mcp(tools) {
  const input = new PassThrough();
  const output = new PassThrough();
  const waiting = new Map();
  let buf = '';
  output.on('data', d => {
    buf += d.toString();
    let i;
    while ((i = buf.indexOf('\n')) >= 0) { const m = JSON.parse(buf.slice(0, i)); buf = buf.slice(i + 1); const w = waiting.get(m.id); if (w) { waiting.delete(m.id); w(m); } }
  });
  const log = { info() {}, error() {}, warn() {}, debug() {} };
  const done = createMcpServer({ tools, info: { version: 'test' }, log, input, output }).start();
  let id = 0;
  const send = (method, params) => { id += 1; const my = id; const p = new Promise(res => waiting.set(my, res)); input.write(`${JSON.stringify({ jsonrpc: '2.0', id: my, method, params })}\n`); return p; };
  return { send, end: () => { input.end(); return done; } };
}

test('M11.4b.1 surface: xlsx_execute_task listed and its tools/call reaches this handler; docx_execute_task stays Unknown tool', async () => {
  const h = harness();
  const { tools } = buildServerTools({ config: CONFIG, configIdentity: 'fixture', writeTool: { execute: async () => { h.n.executor += 1; return { ok: true }; } } });
  assert.deepEqual(exposedTools(), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  // the DOCX WRITE tool keeps BOTH release switches off (the capability list AND the tool row)
  assert.deepEqual([...EXPOSED_CAPABILITY_TOOLS], ['xlsx_read', 'docx_read', 'xlsx_execute_task']);
  assert.equal(TOOL_ROWS.find(r => r.TOOL === 'docx_execute_task').EXPOSE_IN_PRODUCT_MCP, 'NO');
  assert.equal(TOOL_ROWS.find(r => r.TOOL === 'xlsx_execute_task').EXPOSE_IN_PRODUCT_MCP, 'YES');
  assert.deepEqual(Object.keys(tools), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  const s = mcp(tools);
  await s.send('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 't', version: '0' } });
  const list = await s.send('tools/list', {});
  assert.deepEqual(list.result.tools.map(t => t.name), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  const ok = await s.send('tools/call', { name: 'xlsx_execute_task', arguments: req([VALID.write_range]) });
  assert.equal(ok.error, undefined, JSON.stringify(ok));
  const refused = await s.send('tools/call', { name: 'xlsx_execute_task', arguments: req([{ intent: 'merge_range', sheet: 'S', range: 'A1:B1' }]) });
  assert.equal(refused.result.structuredContent.stage, 'PRE_DISPATCH');
  const r = await s.send('tools/call', { name: 'docx_execute_task', arguments: req([VALID.write_range]) });
  assert.match(r.error.message, /Unknown tool: docx_execute_task/);
  await s.end();
  assert.equal(h.n.executor, 1, 'exactly the valid call reached the execution; the refusal and the DOCX tool did not');
});
