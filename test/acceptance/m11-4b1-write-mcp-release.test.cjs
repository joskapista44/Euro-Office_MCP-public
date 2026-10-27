'use strict';

/**
 * M11.4b.1 -- XLSX WRITE MCP release, OFFLINE acceptance (owner decision, owner channel).
 *
 * Through the product MCP server (initialize / tools/list / tools/call) over the registry `buildServerTools` serves:
 * the exact four-tool surface, `xlsx_execute_task` reaching the LANDED handler path (validate -> session gate -> runtime
 * check -> product WRITE driver wrapper -> withSession -> canonical executor -> persistence -> result mapping) with the
 * fake XLSX transport behind the createDriver seam, the six intents APPLIED then SATISFIED, the three-step task, the
 * refusal matrix (pre-dispatch, side-effect counters 0), BUSY shared with the READ tools, and `docx_execute_task`
 * as Unknown tool. No handler / executor / adapter / driver code is changed by M11.4b.1; LIVE acceptance is M11.4b.2.
 * Set M11_4B1_EVIDENCE=<file> to also write the measured matrices as JSON.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const { PassThrough } = require('node:stream');

const { buildServerTools } = require('../../src/server/server-tools.cjs');
const { createMcpServer } = require('../../src/server/mcp.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { book, seedChart, req } = require('../fakes/write-wiring.cjs');
const { REFUSALS } = require('../fakes/write-refusals.cjs');
const { BROKER } = require('../fakes/read-editors.cjs');

const EVIDENCE = {};
test.after(() => { if (process.env.M11_4B1_EVIDENCE) fs.writeFileSync(process.env.M11_4B1_EVIDENCE, `${JSON.stringify(EVIDENCE, null, 1)}\n`); });

const EXACT = ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task'];  // GH #228 (PROPOSED): + docx_execute_task
const STEP = Object.freeze({
  write_range: { intent: 'write_range', sheet: 'Budget', range: 'F1:G2', values: [['x', 1], ['y', 2]] },
  insert_rows: { intent: 'insert_rows', sheet: 'Budget', at: 3, count: 2, anchor: { column: 'A', value: 'pear' } },
  delete_rows: { intent: 'delete_rows', sheet: 'Budget', at: 2, count: 1, anchor: { column: 'A', value: 'pear' } },
  rename_sheet: { intent: 'rename_sheet', sheet: 'Budget', name: 'Budget_2026' },
  unmerge_range: { intent: 'unmerge_range', sheet: 'Budget', range: 'C1:D1' },
  set_chart_title: { intent: 'set_chart_title', sheet: 'K', name: 'R1', title: 'Revenue' },
});
const fresh = () => seedChart(createStore(book()));
const statuses = r => (r.operations || []).map(o => o.status);

/**
 * The SERVED product registry (buildServerTools -> exposedTools) with only the transport faked: the WRITE tool gets the
 * fake XLSX driver (createDriver seam), a counted runtime check and the lifecycle's test clocks; the READ tools share the
 * same process gate. `hold` delays every WRITE open until released (an editor still loading).
 */
function served(store, { gate = createSessionGate(), hold = null } = {}) {
  const configSource = { kind: 'BROKER' };
  const n = { runtime: 0, drivers: 0, open: 0, dispatch: 0, forceSave: 0, close: 0 };
  const drivers = [];
  const createDriver = (format, opts) => {
    n.drivers += 1;
    if (format !== 'xlsx' || !opts || opts.configSource !== configSource) throw new Error('unexpected WRITE driver request');
    const d = createFakeXlsxDriver(store, {});
    if (hold) { const real = d.open.bind(d); d.open = async a => { await hold; return real(a); }; }
    drivers.push(d);
    return d;
  };
  const runtime = { check: () => { n.runtime += 1; return { ok: true }; }, load: () => ({ ok: false, detail: 'unused by the fake' }) };
  let t = 0;
  const session = { ready: { wait: async () => {} }, persistence: { wait: async () => {} }, now: () => (t += 1) };
  const built = buildServerTools({ config: { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: BROKER } }, configIdentity: 'fixture',
    readTool: { gate, runtime: { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'no READ editor offline here' }) } },
    writeTool: { configSource, runtime, createDriver, session } });
  const calls = () => {
    const sum = k => drivers.reduce((a, d) => a + d.calls[k], 0);
    return { runtime: n.runtime, drivers: n.drivers, open: sum('open'), dispatch: sum('dispatch'), forceSave: sum('forceSave'), close: sum('close') };
  };
  return { ...built, calls };
}

/** An MCP session over in-memory streams; call() resolves with the JSON-RPC answer of that id. */
function mcp(tools) {
  const input = new PassThrough();
  const output = new PassThrough();
  const waiting = new Map();
  let buf = '';
  output.on('data', d => {
    buf += String(d);
    let i;
    while ((i = buf.indexOf('\n')) >= 0) { const m = JSON.parse(buf.slice(0, i)); buf = buf.slice(i + 1); const w = waiting.get(m.id); if (w) { waiting.delete(m.id); w(m); } }
  });
  const log = { info() {}, error() {}, warn() {}, debug() {} };
  const server = createMcpServer({ tools, info: { version: 'test' }, log, input, output });
  const done = server.start();
  let id = 0;
  const send = (method, params) => { id += 1; const my = id; const p = new Promise(res => waiting.set(my, res)); input.write(`${JSON.stringify({ jsonrpc: '2.0', id: my, method, params })}\n`); return p; };
  return { send, call: (name, args) => send('tools/call', { name, arguments: args }), end: () => { input.end(); return done; } };
}
const sc = m => { assert.equal(m.error, undefined, JSON.stringify(m.error)); return m.result.structuredContent; };

test('initialize + tools/list: EXACTLY product_status, xlsx_read, docx_read, xlsx_execute_task, docx_execute_task (GH #228 proposed state)', async () => {
  const p = served(await fresh());
  const s = mcp(p.tools);
  const init = await s.send('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 't', version: '0' } });
  assert.equal(init.result.serverInfo.name, 'euro-office-product');
  const list = await s.send('tools/list', {});
  assert.deepEqual(list.result.tools.map(t => t.name), EXACT);
  const w = list.result.tools.find(t => t.name === 'xlsx_execute_task');
  assert.equal(w.inputSchema.additionalProperties, false);
  assert.deepEqual(w.inputSchema.required, ['fileId', 'task']);
  // the listed metadata is factual and lifecycle-free (owner, owner channel): no WITHHELD / RELEASED / milestone words
  assert.doesNotMatch(w.description, /WITHHELD|RELEASED|M11/i);
  for (const i of Object.keys(STEP)) assert.match(w.description, new RegExp(`\\b${i}\\b`), i);
  const docx = await s.call('docx_execute_task', req([STEP.write_range]));
  assert.equal(docx.error, undefined);
  assert.equal(docx.result.structuredContent.code, 'REFUSED_CROSS_FORMAT');
  assert.equal(docx.result.structuredContent.dispatched, false);
  await s.end();
  assert.deepEqual(p.calls(), { runtime: 0, drivers: 0, open: 0, dispatch: 0, forceSave: 0, close: 0 });
  EVIDENCE.surface = { INITIALIZE: 'PASS', TOOLS_LIST: list.result.tools.map(t => t.name), DOCX_EXECUTE_TASK: 'served; xlsx step refused REFUSED_CROSS_FORMAT pre-dispatch' };
});

test('WRITE_6_OF_6 + RETRY_6_OF_6 via tools/call: APPLIED / FORCESAVE_CONFIRMED, then a new call SATISFIED with dispatch and forceSave 0', async () => {
  EVIDENCE.intents = {};
  for (const [intent, step] of Object.entries(STEP)) {
    const store = await fresh();
    const p1 = served(store);
    const s1 = mcp(p1.tools);
    const r1 = sc(await s1.call('xlsx_execute_task', req([step])));
    await s1.end();
    const c1 = p1.calls();
    assert.equal(r1.ok, true, `${intent}: ${JSON.stringify(r1)}`);
    assert.deepEqual(statuses(r1), ['APPLIED'], intent);
    assert.equal(r1.operations[0].l1, 'PASS', intent);
    assert.equal(r1.wholeTask, 'PASS', intent);
    assert.notEqual(r1.operations[0].l3, 'FAIL', intent);
    assert.equal(r1.persistence.state, 'FORCESAVE_CONFIRMED', intent);
    assert.deepEqual([c1.runtime, c1.drivers, c1.open, c1.dispatch, c1.forceSave, c1.close], [1, 1, 1, 1, 1, 1], intent);
    const p2 = served(store);
    const s2 = mcp(p2.tools);
    const r2 = sc(await s2.call('xlsx_execute_task', req([step])));
    await s2.end();
    const c2 = p2.calls();
    assert.equal(r2.ok, true, `${intent} retry: ${JSON.stringify(r2)}`);
    assert.deepEqual(statuses(r2), ['SATISFIED'], intent);
    assert.equal(r2.persistence.state, 'NOT_REQUIRED', intent);
    assert.deepEqual([c2.open, c2.dispatch, c2.forceSave], [1, 0, 0], `${intent}: the retry is a new session with no write and no save`);
    EVIDENCE.intents[intent] = { FIRST: statuses(r1)[0], PERSISTENCE: r1.persistence.state, FIRST_DISPATCH: c1.dispatch, FIRST_FORCESAVE: c1.forceSave,
      SECOND: statuses(r2)[0], SECOND_DISPATCH: c2.dispatch, SECOND_FORCESAVE: c2.forceSave };
  }
});

test('MULTISTEP via one tools/call: write_range + unmerge_range + set_chart_title -> 3x APPLIED, one persistence cycle; retry SATISFIED', async () => {
  const steps = [{ ...STEP.write_range, id: 'w' }, { ...STEP.unmerge_range, id: 'u' }, { ...STEP.set_chart_title, id: 'c' }];
  const store = await fresh();
  const p1 = served(store);
  const s1 = mcp(p1.tools);
  const r1 = sc(await s1.call('xlsx_execute_task', req(steps)));
  await s1.end();
  assert.equal(r1.ok, true, JSON.stringify(r1));
  assert.deepEqual(statuses(r1), ['APPLIED', 'APPLIED', 'APPLIED']);
  assert.equal(r1.wholeTask, 'PASS');
  assert.equal(r1.persistence.state, 'FORCESAVE_CONFIRMED');
  assert.equal(p1.calls().forceSave, 1, 'one persistence barrier for the whole task');
  const p2 = served(store);
  const s2 = mcp(p2.tools);
  const r2 = sc(await s2.call('xlsx_execute_task', req(steps)));
  await s2.end();
  assert.deepEqual(statuses(r2), ['SATISFIED', 'SATISFIED', 'SATISFIED']);
  assert.deepEqual([p2.calls().dispatch, p2.calls().forceSave], [0, 0]);
  EVIDENCE.multistep = { FIRST: statuses(r1), PERSISTENCE: r1.persistence.state, FORCESAVE: p1.calls().forceSave, SECOND: statuses(r2), SECOND_DISPATCH: p2.calls().dispatch };
});

test(`REFUSAL_REGRESSION via tools/call: the ${REFUSALS.length}-case matrix is pre-dispatch; runtime, driver, open, dispatch, forceSave 0`, async () => {
  const p = served(await fresh());
  const s = mcp(p.tools);
  for (const [label, args, code] of REFUSALS) {
    const m = await s.call('xlsx_execute_task', args);
    const r = m.error ? null : m.result.structuredContent;
    assert.ok(r, `${label}: answered as a tool result, not a protocol error`);
    assert.equal(r.ok, false, label);
    assert.equal(r.code, code, label);
    assert.equal(r.stage, 'PRE_DISPATCH', label);
  }
  await s.end();
  assert.deepEqual(p.calls(), { runtime: 0, drivers: 0, open: 0, dispatch: 0, forceSave: 0, close: 0 });
  EVIDENCE.refusals = { COUNT: REFUSALS.length, ALL_PRE_DISPATCH: true, SIDE_EFFECTS: p.calls() };
});

test('BUSY via tools/call: an active WRITE -> a second WRITE and a READ are BUSY, no second driver; a new call works after release', async () => {
  let release;
  const hold = new Promise(res => { release = res; });
  const gate = createSessionGate();
  const p = served(await fresh(), { gate, hold });
  const s = mcp(p.tools);
  const first = s.call('xlsx_execute_task', req([STEP.write_range]));
  for (let i = 0; i < 20 && gate.active() === 0; i += 1) await new Promise(res => setImmediate(res));
  assert.equal(gate.active(), 1);
  const w2 = sc(await s.call('xlsx_execute_task', req([STEP.unmerge_range])));
  const rd = sc(await s.call('xlsx_read', { fileId: '1303626', operation: 'workbook_overview' }));
  assert.deepEqual([w2.code, w2.stage], ['BUSY', 'ADMISSION']);
  assert.equal(rd.code, 'BUSY');
  assert.equal(p.calls().drivers, 1, 'no second driver / browser while busy');
  release();
  const r1 = sc(await first);
  assert.equal(r1.ok, true, JSON.stringify(r1));
  assert.equal(gate.active(), 0);
  const next = sc(await s.call('xlsx_execute_task', req([STEP.unmerge_range])));
  assert.equal(next.ok, true, 'admitted again after the session closed');
  await s.end();
  EVIDENCE.busy = { SECOND_WRITE: w2.code, SECOND_READ: rd.code, SECOND_BROWSER_LAUNCH: 0, AFTER_RELEASE: next.ok };
});
