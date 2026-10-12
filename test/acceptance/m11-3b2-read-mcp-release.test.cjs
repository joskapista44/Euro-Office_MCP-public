'use strict';

/**
 * M11.3b.2 -- READ MCP release, OFFLINE acceptance (owner decision, owner channel 2026-10-10).
 *
 * Through the product MCP server (initialize / tools/list / tools/call) with the product tool registry
 * (buildServerTools): the exact exposed set, the accepted READ operations on the real drivers + adapters (fake editor,
 * fake broker), the pre-dispatch refusal matrix with its side-effect counters, the browser runtime check, the
 * one-session admission gate (BUSY) and the bounded shutdown. LIVE acceptance is a later, separate owner gate.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { PassThrough } = require('node:stream');

const docx = require('../../src/drivers/eurooffice/docx-driver.cjs');
const xlsx = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createBrokerConfigSource } = require('../../src/server/broker-config-source.cjs');
const { buildServerTools } = require('../../src/server/server-tools.cjs');
const { createMcpServer } = require('../../src/server/mcp.cjs');
const { OPERATIONS, writeIntents, inputSchema } = require('../../src/server/read-tools.cjs');
const { createBrowserRuntime, PINNED, EXECUTABLE } = require('../../src/server/browser-runtime.cjs');
const { createSessionGate, MAX_ACTIVE_SESSIONS_PER_PROCESS } = require('../../src/server/session-gate.cjs');
const { BROKER, TOKEN, FILE, editConfig, xlsxEditor, docxEditor, NON_MUTATION } = require('../fakes/read-editors.cjs');

// the fake editors' DS socket carries the fixture fileId's document key, so both formats use that fileId offline
const DOCX_FILE = FILE;
const XLSX_FILE = FILE;
const CONFIG = { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: BROKER, instance: 'production', timeoutMs: 1000 } };

/** A recording fake broker for both formats (the fileId names the format). */
function broker() {
  const calls = [];
  const state = { format: 'docx' };
  const fetchImpl = async (url, init) => {
    const body = JSON.parse(init.body);
    calls.push(body);
    return { status: 200, json: async () => ({ ok: true, status: 200, config: editConfig(body.fileId, state.format) }) };
  };
  return { calls, state, source: createBrokerConfigSource({ brokerUrl: BROKER, readToken: () => TOKEN, fetchImpl }) };
}

/** The product registry with only the editor, the broker and the browser runtime faked; real drivers + adapters. */
function product({ blockFirstOpen = false } = {}) {
  const b = broker();
  const eds = { xlsx: xlsxEditor(), docx: docxEditor() };
  const counts = { runtimeCheck: 0, runtimeLoad: 0, drivers: 0, blockedLaunches: 0 };
  const runtime = { check: () => { counts.runtimeCheck += 1; return { ok: true }; }, load: () => { counts.runtimeLoad += 1; return { ok: false, detail: 'unused' }; } };
  let release = null;
  const createDriver = fmt => {
    counts.drivers += 1;
    b.state.format = fmt; // the fake broker answers the config of the format whose driver is being opened
    if (blockFirstOpen && counts.drivers === 1) {
      // session 1: a driver whose open stays pending (an editor still loading) until the test releases it
      return { open: () => { counts.blockedLaunches += 1; return new Promise(res => { release = () => res({ ok: false, detail: 'released by test' }); }); },
        close: async () => ({ ok: true, detail: 'closed' }) };
    }
    const ed = eds[fmt];
    const ctor = fmt === 'xlsx' ? xlsx.createEuroOfficeXlsxDriver : docx.createEuroOfficeDocxDriver;
    const d = ctor({ configSource: b.source, loadPlaywright: ed.loadPlaywright, pollMs: 5, closeAckTimeoutMs: 200, timeoutMs: 2000 });
    for (const m of ['dispatch', 'forceSave']) { const real = d[m].bind(d); d[m] = (...a) => { ed.api[m] += 1; return real(...a); }; }
    return d;
  };
  const built = buildServerTools({ config: CONFIG, configIdentity: 'fixture', readTool: { configSource: b.source, runtime, createDriver } });
  const launches = () => eds.xlsx.launches() + eds.docx.launches() + counts.blockedLaunches;
  return { ...built, b, eds, counts, launches, release: () => release && release() };
}

/** An MCP session over in-memory streams. call() resolves with the JSON-RPC answer of that id. */
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

const result = m => m.result.structuredContent;

test('MCP initialize + tools/list: the READ tools keep their schema next to xlsx_execute_task (M11.4b.1); docx_execute_task is Unknown tool', async () => {
  const p = product();
  const s = mcp(p.tools);
  const init = await s.send('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 't', version: '0' } });
  assert.equal(init.result.serverInfo.name, 'euro-office-product');
  const list = await s.send('tools/list', {});
  assert.deepEqual(list.result.tools.map(t => t.name), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  for (const t of list.result.tools.filter(x => x.name === 'xlsx_read' || x.name === 'docx_read')) {
    assert.equal(t.inputSchema.additionalProperties, false);
    assert.deepEqual(t.inputSchema.required, ['fileId', 'operation']);
  }
  for (const name of ['docx_execute_task']) {
    const r = await s.call(name, { fileId: DOCX_FILE, operation: 'write_range' });
    assert.match(r.error.message, new RegExp(`Unknown tool: ${name}`));
  }
  await s.end();
  assert.equal(p.b.calls.length, 0);
  assert.equal(p.launches(), 0);
});

test('ACCEPTED_INTENTS: the allowlist is exactly xlsx 3 + docx 5 and the MCP schema offers only those', () => {
  const acc = f => Object.entries(OPERATIONS[f]).filter(([, o]) => o.status === 'ACCEPTED').map(([k]) => k);
  assert.deepEqual(acc('xlsx'), ['workbook_overview', 'read_range', 'read_validations']);
  assert.deepEqual(acc('docx'), ['read_text', 'read_formatting', 'read_tables', 'read_headers_footers', 'read_bookmarks']);
  assert.deepEqual(inputSchema('xlsx').properties.operation.enum, acc('xlsx'));
  assert.deepEqual(inputSchema('docx').properties.operation.enum, acc('docx'));
});

test('XLSX_ACCEPTED_OFFLINE: the three accepted operations through tools/call -- real driver/adapter, non-mutating', async () => {
  const p = product();
  const s = mcp(p.tools);
  const cases = [['workbook_overview', {}], ['read_range', { sheet: 'Budget', range: 'A1:B3' }], ['read_validations', { sheet: 'Budget', range: 'B2:B3' }]];
  for (const [operation, extra] of cases) {
    const r = result(await s.call('xlsx_read', { fileId: XLSX_FILE, operation, ...extra }));
    assert.equal(r.ok, true, `${operation}: ${JSON.stringify(r)}`);
    assert.equal(r.tool, 'xlsx_read');
    NON_MUTATION(p.eds.xlsx, r);
  }
  await s.end();
  assert.equal(p.b.calls.length, 3, 'one broker request per call');
  assert.ok(p.b.calls.every(c => c.fileId === XLSX_FILE));
  assert.equal(p.eds.xlsx.launches(), 3, 'one browser per call, never two at once');
});

test('DOCX_ACCEPTED_OFFLINE: the five accepted operations through tools/call -- real driver/adapter, non-mutating', async () => {
  const p = product();
  const s = mcp(p.tools);
  for (const operation of ['read_text', 'read_formatting', 'read_tables', 'read_headers_footers', 'read_bookmarks']) {
    const r = result(await s.call('docx_read', { fileId: DOCX_FILE, operation }));
    assert.equal(r.ok, true, `${operation}: ${JSON.stringify(r)}`);
    NON_MUTATION(p.eds.docx, r);
  }
  await s.end();
  assert.equal(p.b.calls.length, 5);
  assert.equal(p.eds.docx.launches(), 5);
});

// ---- the refusal matrix: every refusal is PRE-DISPATCH with BROKER_CALL=0 BROWSER_LAUNCH=0 DRIVER_OPEN=0 -------------

const REFUSALS = [
  ['xlsx_read', { operation: 'read_charts' }, 'REFUSED_PARTIAL'],
  ['xlsx_read', { operation: 'read_pivots' }, 'REFUSED_PARTIAL'],
  ['xlsx_read', { operation: 'read_drawings' }, 'REFUSED_PARTIAL'],
  ['xlsx_read', { operation: 'read_conditional_formats' }, 'REFUSED_PARTIAL'],
  ['xlsx_read', { operation: 'read_print_rendering' }, 'REFUSED_NOT_MEASURED'],
  ['docx_read', { operation: 'read_numbering' }, 'REFUSED_PARTIAL'],
  ['docx_read', { operation: 'read_table_grid' }, 'REFUSED_PARTIAL'],
  ['docx_read', { operation: 'read_notes' }, 'REFUSED_PARTIAL'],
  ['docx_read', { operation: 'read_images' }, 'REFUSED_NOT_MEASURED'],
  ['docx_read', { operation: 'read_comments' }, 'REFUSED_NOT_MEASURED'],
  ['docx_read', { operation: 'read_section_flags' }, 'REFUSED_INTERNAL'],
  ['xlsx_read', { operation: 'write_range' }, 'REFUSED_WRITE'],
  ['docx_read', { operation: 'drop_everything' }, 'REFUSED_UNKNOWN'],
  ['xlsx_read', { operation: 'drop_everything' }, 'REFUSED_UNKNOWN'],
  // schema injection: identity / credential fields refused by name, before anything else
  ['docx_read', { operation: 'read_text', requester: 'maintainer' }, 'REQUESTER_OVERRIDE'],
  ['docx_read', { operation: 'read_text', user: 'maintainer' }, 'CREDENTIAL_FIELD'],
  ['docx_read', { operation: 'read_text', uid: 'maintainer' }, 'REQUESTER_OVERRIDE'],
  ['docx_read', { operation: 'read_text', fileOwner: 'maintainer' }, 'REQUESTER_OVERRIDE'],
  ['xlsx_read', { operation: 'workbook_overview', credential: 'x' }, 'CREDENTIAL_FIELD'],
  ['xlsx_read', { operation: 'workbook_overview', password: 'x' }, 'CREDENTIAL_FIELD'],
  ['xlsx_read', { operation: 'workbook_overview', token: 'x' }, 'CREDENTIAL_FIELD'],
  ['xlsx_read', { operation: 'workbook_overview', keyName: 'x' }, 'CREDENTIAL_FIELD'],
  ['xlsx_read', { operation: 'workbook_overview', extra: 1 }, 'UNKNOWN_FIELD'],
];

test('REFUSAL_MATRIX: PARTIAL / NOT_MEASURED / INTERNAL / WRITE / UNKNOWN / schema injection -- BROKER_CALL=0 BROWSER_LAUNCH=0 DRIVER_OPEN=0', async () => {
  const docxWrite = [...writeIntents()].find(i => !OPERATIONS.xlsx[i] && !OPERATIONS.docx[i] && /^[a-z_]+$/.test(i) && i !== 'write_range');
  const rows = [...REFUSALS, ['docx_read', { operation: docxWrite }, 'REFUSED_WRITE']];
  const p = product();
  const s = mcp(p.tools);
  for (const [name, args, code] of rows) {
    const fileId = name === 'xlsx_read' ? XLSX_FILE : DOCX_FILE;
    const r = result(await s.call(name, { fileId, ...args }));
    assert.equal(r.ok, false, `${name} ${JSON.stringify(args)}`);
    assert.equal(r.code, code, `${name} ${JSON.stringify(args)} -> ${r.code}`);
    assert.equal(r.stage, 'PRE_DISPATCH');
    assert.equal(r.dispatched, false);
  }
  await s.end();
  assert.equal(p.b.calls.length, 0, 'BROKER_CALL=0');
  assert.equal(p.launches(), 0, 'BROWSER_LAUNCH=0');
  assert.equal(p.counts.drivers, 0, 'DRIVER_OPEN=0 (no driver even constructed)');
  assert.equal(p.counts.runtimeCheck, 0, 'the runtime is not even checked');
  assert.deepEqual([p.eds.xlsx.api, p.eds.docx.api], [{ dispatch: 0, forceSave: 0 }, { dispatch: 0, forceSave: 0 }], 'DOCUMENT_MUTATION=0');
});

// ---- concurrency: MAX_ACTIVE_SESSIONS_PER_PROCESS=1 ------------------------------------------------------------------

test('MAX_ACTIVE_SESSIONS=1: session 1 active -> session 2 BUSY, browser count does not grow; session 1 closed -> session 2 succeeds', async () => {
  assert.equal(MAX_ACTIVE_SESSIONS_PER_PROCESS, 1);
  const p = product({ blockFirstOpen: true });
  assert.equal(p.gate.max, 1);
  const s = mcp(p.tools);
  const first = s.call('xlsx_read', { fileId: XLSX_FILE, operation: 'workbook_overview' });
  // let session 1 reach its pending open
  for (let i = 0; i < 20 && p.counts.blockedLaunches === 0; i += 1) await new Promise(r => setImmediate(r));
  assert.equal(p.gate.active(), 1);
  const launchesBefore = p.launches();
  const busyDocx = result(await s.call('docx_read', { fileId: DOCX_FILE, operation: 'read_text' }));
  const busyXlsx = result(await s.call('xlsx_read', { fileId: XLSX_FILE, operation: 'workbook_overview' }));
  for (const r of [busyDocx, busyXlsx]) {
    assert.equal(r.ok, false);
    assert.equal(r.code, 'BUSY');
    assert.equal(r.stage, 'ADMISSION');
  }
  assert.equal(p.launches(), launchesBefore, 'SECOND_BROWSER_LAUNCH=0');
  assert.equal(p.counts.drivers, 1, 'no second driver: no shared and no second editor session');
  assert.equal(p.b.calls.length, 0, 'BUSY never reaches the broker');
  p.release();
  const r1 = result(await first);
  assert.equal(r1.code, 'OPEN_FAILED');
  assert.equal(p.gate.active(), 0, 'SESSION_RELEASE: the lease is released when session 1 ends');
  const next = result(await s.call('docx_read', { fileId: DOCX_FILE, operation: 'read_text' }));
  assert.equal(next.ok, true, `NEXT_REQUEST=ALLOWED: ${JSON.stringify(next)}`);
  await s.end();
});

test('SESSION_RELEASE on every exit: a refused open, a runtime refusal and a thrown driver all free the slot', async () => {
  const gate = createSessionGate();
  const { createReadTool } = require('../../src/server/read-tools.cjs');
  const b = broker();
  const mk = opts => createReadTool('docx', { configSource: b.source, gate, ...opts });
  const runtimeOk = { check: () => ({ ok: true }), load: () => ({ ok: false }) };
  await mk({ runtime: runtimeOk, createDriver: () => ({ open: async () => ({ ok: false, detail: 'x' }), close: async () => ({}) }) }).handler({ fileId: DOCX_FILE, operation: 'read_text' });
  assert.equal(gate.active(), 0);
  await mk({ runtime: { check: () => ({ ok: false, code: 'BROWSER_RUNTIME_UNAVAILABLE', detail: 'x' }), load: () => null } }).handler({ fileId: DOCX_FILE, operation: 'read_text' });
  assert.equal(gate.active(), 0);
  await assert.rejects(mk({ runtime: runtimeOk, createDriver: () => { throw new Error('boom'); } }).handler({ fileId: DOCX_FILE, operation: 'read_text' }));
  assert.equal(gate.active(), 0);
});

// ---- browser runtime -------------------------------------------------------------------------------------------------

test('BROWSER_RUNTIME: pinned playwright-core 1.62.1 (exact dependency + lockfile) and headless shell revision 1234', () => {
  const root = path.join(__dirname, '..', '..');
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.deepEqual(pkg.dependencies, { 'playwright-core': '1.62.1' }, 'exact version, no range, the only runtime dependency');
  const lock = JSON.parse(fs.readFileSync(path.join(root, 'package-lock.json'), 'utf8'));
  const entry = lock.packages['node_modules/playwright-core'];
  assert.equal(entry.version, '1.62.1');
  assert.match(entry.integrity, /^sha512-/);
  assert.equal(entry.hasInstallScript, undefined, 'no install script: nothing downloads a browser (BROWSER_DOWNLOAD=NO)');
  assert.deepEqual(PINNED, { playwrightCore: '1.62.1', browser: 'chromium-headless-shell', revision: '1234' });
  // positive control: the real installed playwright-core passes when the pinned executable is present
  const fsOk = { accessSync: p => { if (p !== path.join('/opt/ms-playwright', ...EXECUTABLE)) throw new Error('ENOENT'); } };
  assert.equal(createBrowserRuntime({ browsersPath: '/opt/ms-playwright', fsImpl: fsOk }).check().ok, true);
});

test('BROWSER_MISSING=CONTROLLED: unconfigured / empty dir / wrong playwright-core / wrong revision -> BROWSER_RUNTIME_UNAVAILABLE, no broker, no crash', async ctx => {
  const empty = fs.mkdtempSync(path.join(os.tmpdir(), 'm113b2-nobrowser-'));
  ctx.after(() => fs.rmSync(empty, { recursive: true, force: true }));
  const pkgFiles = over => name => (over[name] !== undefined ? over[name]
    : { 'package.json': { version: '1.62.1' }, 'browsers.json': { browsers: [{ name: 'chromium-headless-shell', revision: '1234' }] } }[name]);
  const notFound = () => { const e = new Error('nf'); e.code = 'MODULE_NOT_FOUND'; throw e; };
  const cases = [
    createBrowserRuntime({}),
    createBrowserRuntime({ browsersPath: 'relative/dir' }),
    createBrowserRuntime({ browsersPath: empty }),
    createBrowserRuntime({ browsersPath: '/opt/ms-playwright', readPackageFile: pkgFiles({ 'package.json': { version: '1.61.0' } }), fsImpl: { accessSync() {} } }),
    createBrowserRuntime({ browsersPath: '/opt/ms-playwright', readPackageFile: pkgFiles({ 'browsers.json': { browsers: [{ name: 'chromium-headless-shell', revision: '1300' }] } }), fsImpl: { accessSync() {} } }),
    createBrowserRuntime({ browsersPath: '/opt/ms-playwright', readPackageFile: notFound, fsImpl: { accessSync() {} } }),
  ];
  // NEG control for the cases above: the same fakes with the pinned values pass the check
  assert.equal(createBrowserRuntime({ browsersPath: '/opt/ms-playwright', readPackageFile: pkgFiles({}), fsImpl: { accessSync() {} } }).check().ok, true);
  for (const rt of cases) {
    const c = rt.check();
    assert.equal(c.ok, false);
    assert.equal(c.code, 'BROWSER_RUNTIME_UNAVAILABLE', c.detail);
    assert.equal(rt.load().code, 'BROWSER_RUNTIME_UNAVAILABLE');
  }
  // through the product tool: the check runs before the driver, so the broker is never asked
  const b = broker();
  const { createReadTool } = require('../../src/server/read-tools.cjs');
  let drivers = 0;
  const t = createReadTool('xlsx', { configSource: b.source, browsersPath: empty, createDriver: () => { drivers += 1; throw new Error('no'); } });
  const r = await t.handler({ fileId: XLSX_FILE, operation: 'workbook_overview' });
  assert.deepEqual([r.ok, r.stage, r.code, b.calls.length, drivers], [false, 'RUNTIME', 'BROWSER_RUNTIME_UNAVAILABLE', 0, 0]);
});

// ---- bounded shutdown ------------------------------------------------------------------------------------------------

test('BOUNDED_SHUTDOWN: an active session is closed through its driver close(); a close that never ends is cut at the timeout', async () => {
  const gate = createSessionGate();
  assert.deepEqual(await gate.shutdown({ timeoutMs: 50 }), { sessions: 0, cleanup: 'NONE', ms: 0 });

  const g1 = createSessionGate();
  let closed = 0;
  g1.tryEnter().track({ close: async () => { closed += 1; return { ok: true }; } });
  const r1 = await g1.shutdown({ timeoutMs: 1000 });
  assert.deepEqual([r1.sessions, r1.cleanup, closed], [1, 'CLOSED', 1]);
  assert.equal(g1.tryEnter(), null, 'no new session once shutdown started');

  const g2 = createSessionGate();
  g2.tryEnter().track({ close: () => new Promise(() => {}) }); // a close that never settles
  const t0 = Date.now();
  const r2 = await g2.shutdown({ timeoutMs: 150 });
  assert.equal(r2.cleanup, 'TIMEOUT');
  assert.ok(Date.now() - t0 < 1000, `bounded: ${Date.now() - t0} ms`);
});

test('BOUNDED_SHUTDOWN through the product tool: shutdown during an active session closes it via close(); the request returns, the slot is free', async () => {
  const p = product({ blockFirstOpen: true });
  const s = mcp(p.tools);
  const pending = s.call('xlsx_read', { fileId: XLSX_FILE, operation: 'workbook_overview' });
  for (let i = 0; i < 50 && p.counts.blockedLaunches === 0; i += 1) await new Promise(r => setImmediate(r));
  assert.equal(p.gate.active(), 1, 'a session is active when shutdown starts');
  // the stub driver models a real teardown: its close() ends the pending open
  assert.equal(p.gate.tryEnter(), null, 'still BUSY while the session is active');
  const sd = await p.gate.shutdown({ timeoutMs: 2000 }).then(r => { p.release(); return r; });
  assert.deepEqual([sd.sessions, sd.cleanup], [1, 'CLOSED']);
  const out = result(await pending);
  assert.equal(out.ok, false);
  assert.equal(p.gate.active(), 0, 'SESSION_CLEANUP: the slot is free after the session ended');
  assert.equal(p.gate.tryEnter(), null, 'no new session after shutdown started');
  await s.end();
});
