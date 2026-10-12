'use strict';

/**
 * M11.3a -- READ tools xlsx_read / docx_read: handlers + schemas, WITHHELD. OFFLINE acceptance (owner GO, owner channel
 * 2026-10-09, after the M11.2 closure).
 *
 * The REAL product handler (src/server/read-tools.cjs) runs end to end against the REAL drivers and adapters:
 *   MCP args -> schema -> the M11.2 broker configSource (fake broker, own token) -> the canonical driver open(READ) +
 *   readLive + close -> normalizeInventory -> projection.
 * The editor is fake: XLSX command bodies run against an in-memory Api that has GETTERS ONLY (a write call would throw
 * and fail the read; every missing API name the body touches is logged), DOCX answers the real property query.
 * Proven: the accepted reads return structured results with DOCUMENT_MUTATION=0 / HISTORY_MUTATION=0 / WRITE=0 /
 * SAVE_REQUIRED=NO; every refused request (unknown / PARTIAL / NOT_MEASURED / INTERNAL / WRITE intent, malformed
 * arguments, credential injection, requester spoof, wrong fileId type) stops PRE-DISPATCH -- no broker call, no driver;
 * the tools stayed WITHHELD in M11.3a (neither listed nor callable); M11.3b.2 released exactly these two (see below).
 */

const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');
const { PassThrough } = require('node:stream');

const docx = require('../../src/drivers/eurooffice/docx-driver.cjs');
const xlsx = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { OPERATIONS, STATUS, createReadTool, validateReadRequest, writeIntents, inputSchema } = require('../../src/server/read-tools.cjs');
const { buildServerTools } = require('../../src/server/server-tools.cjs');
const { createMcpServer } = require('../../src/server/mcp.cjs');
const { exposedTools, TOOL_ROWS } = require('../../src/server/tools.cjs');

const { BROKER, TOKEN, FILE, fakeBroker, xlsxEditor, docxEditor, NON_MUTATION } = require('../fakes/read-editors.cjs');

const drivers = [];
after(async () => { for (const d of drivers) await d.close().catch(() => null); });

/** The real read tool on the real driver; only the editor, the broker and the timings are fakes. */
function tool(format, ed, broker) {
  const createDriver = (fmt, deps) => {
    assert.equal(deps.configSource.kind, 'BROKER', 'the handler hands the driver the broker configSource');
    const ctor = fmt === 'xlsx' ? xlsx.createEuroOfficeXlsxDriver : docx.createEuroOfficeDocxDriver;
    const d = ctor({ ...deps, pollMs: 5, closeAckTimeoutMs: 200, timeoutMs: 2000 });
    drivers.push(d);
    // the driver-API level proof: a READ never calls the write/save surface of the canonical driver
    for (const m of ['dispatch', 'forceSave']) { const real = d[m].bind(d); d[m] = (...a) => { ed.api[m] += 1; return real(...a); }; }
    return d;
  };
  return createReadTool(format, { configSource: broker.source, loadPlaywright: ed.loadPlaywright, createDriver });
}

async function noDirectFetch(fn) {
  const real = globalThis.fetch;
  let direct = 0;
  globalThis.fetch = async () => { direct += 1; throw new Error('the product READ path must not fetch Nextcloud directly'); };
  try { const out = await fn(); assert.equal(direct, 0, 'no direct Nextcloud / Basic-auth fetch'); return out; } finally { globalThis.fetch = real; }
}

// ---- inventory ----------------------------------------------------------------------------------------------------

const accepted = f => Object.entries(OPERATIONS[f]).filter(([, o]) => o.status === STATUS.ACCEPTED).map(([k]) => k);
const byStatus = s => ['xlsx', 'docx'].flatMap(f => Object.entries(OPERATIONS[f]).filter(([, o]) => o.status === s).map(([k]) => `${f}:${k}`));

test('INVENTORY: the accepted READ allowlist, every catalogue row carries evidence, the 76 routed intents are WRITE', () => {
  assert.deepEqual(accepted('xlsx'), ['workbook_overview', 'read_range', 'read_validations']);
  assert.deepEqual(accepted('docx'), ['read_text', 'read_formatting', 'read_tables', 'read_headers_footers', 'read_bookmarks']);
  assert.equal(byStatus(STATUS.PARTIAL).length, 7);
  assert.equal(byStatus(STATUS.NOT_MEASURED).length, 3);
  assert.equal(byStatus(STATUS.INTERNAL).length, 1);
  for (const f of ['xlsx', 'docx']) for (const [k, o] of Object.entries(OPERATIONS[f])) assert.ok(o.evidence && o.evidence.length > 3, `${f}:${k} evidence`);
  assert.equal(writeIntents().size, 76);
  for (const f of ['xlsx', 'docx']) for (const k of Object.keys(OPERATIONS[f])) assert.equal(writeIntents().has(k), false, `${k} is not a write intent`);
});

test('SCHEMA: fileId + operation required, only accepted operations offered, no credential or requester field, additionalProperties false', () => {
  for (const f of ['xlsx', 'docx']) {
    const s = inputSchema(f);
    assert.deepEqual(s.required, ['fileId', 'operation']);
    assert.equal(s.additionalProperties, false);
    assert.equal(s.properties.fileId.type, 'string');
    assert.deepEqual(s.properties.operation.enum, accepted(f));
    for (const k of Object.keys(s.properties)) assert.equal(/user|pass|secret|token|cred|vault|key|requester|uid|owner|actor/i.test(k), false, `${f}.${k}`);
  }
});

// ---- valid reads through the real stack -----------------------------------------------------------------------------

test('VALID_XLSX_READ: read_range + workbook_overview + read_validations -- broker READ config, real driver/adapter, non-mutating', async () => {
  await noDirectFetch(async () => {
    const b = fakeBroker('xlsx');
    const ed = xlsxEditor();
    const t = tool('xlsx', ed, b);
    const r = await t.handler({ fileId: FILE, operation: 'read_range', sheet: 'Budget', range: 'A1:B3' });
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(r.complete, true);
    assert.deepEqual(r.result.cells.map(c => [c.address, c.v]), [['A1', 'Item'], ['B1', 'Amount'], ['A2', 'Rent'], ['B2', '1200'], ['A3', 'Food'], ['B3', '340']]);
    assert.equal(b.calls.length, 1);
    assert.equal(b.calls[0].bearer, TOKEN, 'the own token only');
    assert.deepEqual(Object.keys(b.calls[0].body).sort(), ['fileId', 'instance', 'op', 'requestedMode']);
    assert.equal(b.calls[0].body.requestedMode, 'READ');
    assert.ok(ed.log.includes('read-body'), 'the canonical readLive query ran (core dispatch of the READ)');
    NON_MUTATION(ed, r);

    const o = await t.handler({ fileId: FILE, operation: 'workbook_overview' });
    assert.equal(o.ok, true, JSON.stringify(o));
    assert.deepEqual(o.result.sheets, [{ name: 'Budget', index: 0, used: 'A1:B3' }]);
    NON_MUTATION(ed, o);

    const v = await t.handler({ fileId: FILE, operation: 'read_validations', sheet: 'Budget', range: 'B2:B3' });
    assert.equal(v.ok, true, JSON.stringify(v));
    assert.deepEqual(v.result.validation, { type: 'xlValidateWholeNumber', alertStyle: 'xlValidAlertStop', operator: 'xlBetween', formula1: '0', formula2: '10000' });
    NON_MUTATION(ed, v);
    assert.equal(ed.launches(), 3);
  });
});

test('VALID_DOCX_READ: the five accepted operations -- broker READ config, real driver/adapter, non-mutating, no PARTIAL/INTERNAL field', async () => {
  await noDirectFetch(async () => {
    const b = fakeBroker('docx');
    const ed = docxEditor();
    const t = tool('docx', ed, b);
    const results = {};
    for (const op of accepted('docx')) {
      const r = await t.handler({ fileId: FILE, operation: op });
      assert.equal(r.ok, true, `${op}: ${JSON.stringify(r)}`);
      assert.equal(r.complete, true);
      NON_MUTATION(ed, r);
      results[op] = r.result;
      // numbering (PARTIAL), gridCalc (PARTIAL), title-page/even-odd flags (INTERNAL), the internal gate: never projected
      assert.equal(/numbering|removeNumPr|gridCalc|titlePage|evenAndOdd|internalGate/.test(JSON.stringify(r.result)), false, `${op} leaks a refused field`);
    }
    assert.deepEqual(results.read_text.blocks.map(x => x.kind), ['heading', 'paragraph', 'list', 'table']);
    assert.equal(results.read_text.blocks[1].text, 'Hello world');
    assert.deepEqual(results.read_text.blocks[2].items, ['one', 'two']);
    assert.equal(results.read_formatting.units[1].text, 'Hello world');
    assert.deepEqual(results.read_formatting.units[1].format, { alignment: 'left' });
    assert.deepEqual(results.read_formatting.defaultText, { fontFamily: 'Arial', fontSize: 11 });
    assert.deepEqual(results.read_tables.tables[0].rows[0].map(c => c.text), ['a', 'b']);
    assert.equal(results.read_headers_footers.sections[0].header.default.paras[0].text, 'Head');
    assert.deepEqual(results.read_bookmarks.bookmarks, [{ name: 'bm1', text: 'Hello' }]);
    assert.equal(b.calls.length, 5);
    assert.ok(b.calls.every(c => c.bearer === TOKEN && c.body.requestedMode === 'READ'));
    assert.equal(ed.log.filter(l => l === 'read-query').length, 5, 'the canonical property query ran once per read');
    assert.equal(ed.log.filter(l => l === 'read-query:opt-out').length, 5, 'M11.3a.1: every product READ asks without the B21 section variant state');
  });
});

// ---- pre-dispatch refusals --------------------------------------------------------------------------------------------

const REFUSALS = [
  ['UNKNOWN_INTENT', 'xlsx', { fileId: FILE, operation: 'read_everything' }, 'REFUSED_UNKNOWN'],
  ['PARTIAL intent (xlsx)', 'xlsx', { fileId: FILE, operation: 'read_charts' }, 'REFUSED_PARTIAL'],
  ['PARTIAL intent (docx)', 'docx', { fileId: FILE, operation: 'read_numbering' }, 'REFUSED_PARTIAL'],
  ['NOT_MEASURED intent', 'docx', { fileId: FILE, operation: 'read_images' }, 'REFUSED_NOT_MEASURED'],
  ['INTERNAL read', 'docx', { fileId: FILE, operation: 'read_section_flags' }, 'REFUSED_INTERNAL'],
  ['WRITE_INTENT (xlsx)', 'xlsx', { fileId: FILE, operation: 'write_range', sheet: 'Budget', range: 'A1' }, 'REFUSED_WRITE'],
  ['WRITE_INTENT (docx intent on the xlsx tool)', 'xlsx', { fileId: FILE, operation: [...writeIntents()].find(i => !OPERATIONS.xlsx[i] && !/_range|_sheet|_rows|_columns/.test(i)) }, 'REFUSED_WRITE'],
  ['malformed args: not an object', 'docx', ['1302148'], 'INVALID_ARGS'],
  ['malformed args: missing range', 'xlsx', { fileId: FILE, operation: 'read_range', sheet: 'Budget' }, 'INVALID_ARGS'],
  ['malformed args: bad range', 'xlsx', { fileId: FILE, operation: 'read_range', sheet: 'Budget', range: 'Budget!A1' }, 'INVALID_ARGS'],
  ['malformed args: flag type', 'xlsx', { fileId: FILE, operation: 'read_range', sheet: 'Budget', range: 'A1', format: 'yes' }, 'INVALID_ARGS'],
  ['malformed args: unknown field', 'docx', { fileId: FILE, operation: 'read_text', depth: 3 }, 'UNKNOWN_FIELD'],
  ['malformed args: field of another operation', 'xlsx', { fileId: FILE, operation: 'workbook_overview', sheet: 'Budget' }, 'UNKNOWN_FIELD'],
  ['range over the read ceiling', 'xlsx', { fileId: FILE, operation: 'read_range', sheet: 'Budget', range: 'A1:Z1000' }, 'LIMIT'],
  ['CREDENTIAL_INJECTION: password', 'docx', { fileId: FILE, operation: 'read_text', password: 'x' }, 'CREDENTIAL_FIELD'],
  ['CREDENTIAL_INJECTION: user', 'xlsx', { fileId: FILE, operation: 'workbook_overview', user: 'maintainer' }, 'CREDENTIAL_FIELD'],
  ['CREDENTIAL_INJECTION: token', 'docx', { fileId: FILE, operation: 'read_text', token: 'abc' }, 'CREDENTIAL_FIELD'],
  ['CREDENTIAL_INJECTION: key_name', 'docx', { fileId: FILE, operation: 'read_text', key_name: 'nc/maintainer' }, 'CREDENTIAL_FIELD'],
  ['REQUESTER_SPOOF: requester', 'docx', { fileId: FILE, operation: 'read_text', requester: 'maintainer' }, 'REQUESTER_OVERRIDE'],
  ['REQUESTER_SPOOF: uid', 'xlsx', { fileId: FILE, operation: 'workbook_overview', uid: 'maintainer' }, 'REQUESTER_OVERRIDE'],
  ['REQUESTER_SPOOF: onBehalfOf', 'docx', { fileId: FILE, operation: 'read_text', onBehalfOf: 'maintainer' }, 'REQUESTER_OVERRIDE'],
  ['wrong fileId type: number', 'docx', { fileId: 1302148, operation: 'read_text' }, 'INVALID_ARGS'],
  ['wrong fileId type: path', 'xlsx', { fileId: '/Budget.xlsx', operation: 'workbook_overview' }, 'INVALID_ARGS'],
  ['wrong fileId type: missing', 'docx', { operation: 'read_text' }, 'INVALID_ARGS'],
];

for (const [name, format, args, code] of REFUSALS) {
  test(`PRE-DISPATCH refusal -- ${name} -> ${code}: no broker call, no driver, no browser`, async () => {
    const b = fakeBroker(format);
    let created = 0;
    const t = createReadTool(format, { configSource: b.source, loadPlaywright: () => { throw new Error('no browser'); }, createDriver: () => { created += 1; throw new Error('no driver'); } });
    const r = await t.handler(args);
    assert.equal(r.ok, false);
    assert.equal(r.stage, 'PRE_DISPATCH');
    assert.equal(r.code, code, JSON.stringify(r));
    assert.equal(r.dispatched, false);
    assert.equal(b.calls.length, 0);
    assert.equal(created, 0);
    assert.equal(JSON.stringify(r).includes(TOKEN), false, 'SECRET_LEAK=NO');
  });
}

test('the refusal classification is the validator\'s, independent of the handler (same codes)', () => {
  for (const [, format, args, code] of REFUSALS) assert.equal(validateReadRequest(format, args).code, code);
});

// ---- exposure (M11.3a: withheld; M11.3b.2, owner decision 2026-10-10: the two READ tools released) ------------------

test('M11.4b.1 exposure: tools/list = product_status + xlsx_read + docx_read + xlsx_execute_task; docx_execute_task stays Unknown tool', async () => {
  const config = { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: BROKER, instance: 'production', timeoutMs: 1000 } };
  const { registry, tools } = buildServerTools({ config, configIdentity: 'fixture' });
  assert.deepEqual(Object.keys(registry).sort(), ['docx_read', 'product_status', 'xlsx_execute_task', 'xlsx_read']); // M11.4a.1 built; M11.4b.1 released
  assert.deepEqual(Object.keys(tools), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  assert.deepEqual(exposedTools(), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  assert.deepEqual(TOOL_ROWS.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'YES').map(r => r.TOOL).sort(), ['docx_execute_task']);

  const input = new PassThrough();
  const output = new PassThrough();
  const lines = [];
  output.on('data', d => lines.push(...String(d).split('\n').filter(Boolean).map(l => JSON.parse(l))));
  const log = { info() {}, error() {}, warn() {}, debug() {} };
  const server = createMcpServer({ tools, info: { version: 'test' }, log, input, output });
  const done = server.start();
  input.write(`${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list' })}\n`);
  input.write(`${JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'xlsx_execute_task', arguments: {} } })}\n`);
  input.write(`${JSON.stringify({ jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'docx_execute_task', arguments: {} } })}\n`);
  input.end();
  await done;
  const byId = Object.fromEntries(lines.map(m => [m.id, m]));
  assert.deepEqual(byId[1].result.tools.map(x => x.name), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  assert.equal(byId[2].error, undefined, 'xlsx_execute_task is callable (an empty task is refused by its own schema)');
  assert.equal(byId[2].result.structuredContent.ok, false);
  assert.match(byId[3].error.message, /Unknown tool: docx_execute_task/);
});

test('the product default with no browsers directory configured: an accepted READ stops at BROWSER_RUNTIME_UNAVAILABLE, no broker call', async () => {
  const b = fakeBroker('docx');
  const { registry } = buildServerTools({ config: { broker: { url: BROKER } } });
  const t = createReadTool('docx', { configSource: b.source });
  const r = await t.handler({ fileId: FILE, operation: 'read_text' });
  assert.equal(r.ok, false);
  assert.equal(r.stage, 'RUNTIME');
  assert.equal(r.code, 'BROWSER_RUNTIME_UNAVAILABLE');
  assert.equal(b.calls.length, 0, 'the browser runtime check refuses before the config is fetched');
  assert.equal(typeof registry.docx_read.handler, 'function');
});
