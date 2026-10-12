'use strict';

/**
 * M11.3a -- the product READ tools `xlsx_read` / `docx_read`: handlers + schemas (owner decision, owner channel
 * 2026-10-09, after the M11.2 closure); M11.3b.2 -- released into the product MCP with the browser runtime, the
 * one-session admission gate and bounded shutdown (owner decision, owner channel 2026-10-10).
 *
 *   MCP arguments -> validated schema (PRE-DISPATCH refusal) -> the M11.2 requester-own broker configSource ->
 *   the existing canonical driver (open READ + readLive + close) -> the existing adapter normalizeInventory ->
 *   a read-only projection of the named operation -> a structured result.
 *
 * No new read engine: the read is the drivers' own readLive and the adapters' own normalizeInventory; this module only
 * validates, routes and projects. Nothing here dispatches a command, saves or force-saves.
 *
 * OPERATIONS is the allowlist with its evidence status. Only ACCEPTED operations run; PARTIAL / NOT_MEASURED /
 * INTERNAL operations, the routed WRITE intents and anything unknown are refused before a driver exists.
 *
 * The requester is NOT a tool argument: the identity is the calling agent's own token, resolved by the broker
 * (M11.2). Credential-like and requester-override fields are refused by name, before the unknown-field check.
 *
 * Exposure: the exposure gate (tools.cjs EXPOSED_CAPABILITY_TOOLS) releases exactly these two capability tools since
 * M11.3b.2; the entrypoint registers only `exposedTools()`. The WRITE tools (`*_execute_task`) stay WITHHELD.
 */

const { createEuroOfficeXlsxDriver, MAX_READ_CELLS } = require('../drivers/eurooffice/xlsx-driver.cjs');
const { createEuroOfficeDocxDriver } = require('../drivers/eurooffice/docx-driver.cjs');
const xlsxInventory = require('../adapters/xlsx/inventory.cjs');
const docxInventory = require('../adapters/docx/inventory.cjs');
const { parseRange, cellsOf } = require('../adapters/xlsx/a1.cjs');
const docxIntents = require('../adapters/docx/intents.cjs');
const { createBrowserRuntime } = require('./browser-runtime.cjs');
const { createSessionGate } = require('./session-gate.cjs');

const STATUS = Object.freeze({ ACCEPTED: 'ACCEPTED', PARTIAL: 'PARTIAL', NOT_MEASURED: 'NOT_MEASURED', INTERNAL: 'INTERNAL' });

const SHEET = Object.freeze({ type: 'string', minLength: 1, maxLength: 31, description: 'worksheet name' });
const RANGE = Object.freeze({ type: 'string', pattern: '^\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6}(:\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6})?$', description: 'A1 range on that sheet, no sheet prefix' });
const FLAG = Object.freeze({ type: 'boolean' });

/**
 * The operation catalogue. `args`: the operation's own arguments (required unless listed in `optional`).
 * Evidence is the accepted acceptance record of the field set the projection returns (docs/ references).
 */
const OPERATIONS = Object.freeze({
  xlsx: Object.freeze({
    workbook_overview: { status: STATUS.ACCEPTED, args: {}, evidence: 'sheets + used range: M6-W0 LIVE probe F2-F6; M6-W1/W2' },
    read_range: { status: STATUS.ACCEPTED, args: { sheet: SHEET, range: RANGE, format: FLAG, style: FLAG }, optional: ['format', 'style'],
      evidence: 'v/f/count/merge: M6-W0 F2-F6; nf/fill: M6-W1B, X-F08/F10; style: M6 matrix W4 26/26, X-F02/F05/F07a; bold/italic/align/border/readingOrder/indent/shrink via DS Bridge B01/B02/B04/B05: DS-BRIDGE-INTERNAL-MIGRATION parity + cutover' },
    read_validations: { status: STATUS.ACCEPTED, args: { sheet: SHEET, range: RANGE }, evidence: 'range-level validation reader: M6 W6 (set/clear_validation)' },
    read_charts: { status: STATUS.PARTIAL, evidence: 'name/type/title/series W7, but dataLabels X-O05 PARTIAL (bridge data-label read pending)' },
    read_pivots: { status: STATUS.PARTIAL, evidence: 'X-O10; after reopen the source is content-proven only (M7-H)' },
    read_drawings: { status: STATUS.PARTIAL, evidence: 'X-O09' },
    read_conditional_formats: { status: STATUS.PARTIAL, evidence: 'X-O01: CF types beyond xlCellValue/xlExpression/3 scalar types partial; rendered effect X-O02 NOT_MEASURED' },
    read_print_rendering: { status: STATUS.NOT_MEASURED, evidence: 'X-P15, X-P04' },
  }),
  docx: Object.freeze({
    read_text: { status: STATUS.ACCEPTED, args: {}, evidence: 'blocks (heading/paragraph/list/table text) + locator: M4-WAVE1 (M8 124/124), M10 matrix W-08/W-09-3' },
    read_formatting: { status: STATUS.ACCEPTED, args: {}, evidence: 'run/paragraph props + format, hyperlinks, note refs: M4-WAVE1, M4-W11, M4-W3-RICH; defaultText M5 gate 6' },
    read_tables: { status: STATUS.ACCEPTED, args: {}, evidence: 'public table read (cell text/format/paraFormats): M4-W4A-D' },
    read_headers_footers: { status: STATUS.ACCEPTED, args: {}, evidence: 'header/footer default/even/title parts: M4-W5 14/14, M5 hf variants 16/16' },
    read_bookmarks: { status: STATUS.ACCEPTED, args: {}, evidence: 'M4-W3-RICH' },
    read_numbering: { status: STATUS.PARTIAL, evidence: 'B19 bridge parity 8/8, M10 matrix MIGRATED + PARTIALLY COVERED (PRODUCTION-DRIVER-BRIDGE-ROUTE)' },
    read_table_grid: { status: STATUS.PARTIAL, evidence: 'gridCalc B20 bridge parity 5/5, M10 matrix PARTIALLY COVERED' },
    read_notes: { status: STATUS.PARTIAL, evidence: 'footnote/endnote text: first paragraph only' },
    read_images: { status: STATUS.NOT_MEASURED, evidence: 'images only a preservation multiset; no positive LIVE image fixture' },
    read_comments: { status: STATUS.NOT_MEASURED, evidence: 'comments are not read by the driver' },
    read_section_flags: { status: STATUS.INTERNAL, evidence: 'B21: titlePage/evenAndOdd come from the internal ApiSection getters (B21, runtime-gated legacy), not an accepted public read' },
  }),
});

const XLSX_WRITE_CANDIDATES = () => require('./tools.cjs').xlsxRoutedIntents();
const DOCX_WRITE_INTENTS = () => (Array.isArray(docxIntents.INTENTS) ? docxIntents.INTENTS.map(r => r.intent) : Object.keys(docxIntents.INTENTS));
/** The routed WRITE intents of both formats (the 76 of the M11.0 inventory): never a READ operation. */
const writeIntents = () => new Set([...XLSX_WRITE_CANDIDATES(), ...DOCX_WRITE_INTENTS()]);

// Field names refused by NAME (normalized: lower case, no '_' / '-'), before the unknown-field check.
const CREDENTIAL_FIELDS = new Set(['user', 'username', 'password', 'pass', 'passwd', 'pwd', 'secret', 'token', 'bearer', 'auth',
  'authorization', 'credential', 'credentials', 'keyname', 'key', 'vault', 'vaultkey', 'apppassword', 'ncuser', 'ncpassword',
  'basic', 'basicauth', 'apikey', 'cookie', 'session']);
const REQUESTER_FIELDS = new Set(['requester', 'requesterid', 'requesteruid', 'uid', 'userid', 'ncuid', 'as', 'asuser', 'onbehalfof',
  'impersonate', 'actor', 'principal', 'agent', 'agentid', 'owner', 'fileowner', 'identity', 'runas']);
const normName = k => String(k).toLowerCase().replace(/[_\-\s]/g, '');

const FILE_ID = /^[0-9]{1,15}$/;
const COMMON = ['fileId', 'operation'];

const refuse = (code, reason) => Object.freeze({ ok: false, stage: 'PRE_DISPATCH', code, reason, dispatched: false });

// The validation stages: each returns a refusal or null (pass). The order is the contract: identity-bearing names first
// (so a credential is never reported as a mere unknown field), then target, operation class, the operation's own fields.
function checkNames(args) {
  for (const k of Object.keys(args)) {
    if (CREDENTIAL_FIELDS.has(normName(k))) return refuse('CREDENTIAL_FIELD', `credential-like field refused: ${k}`);
    if (REQUESTER_FIELDS.has(normName(k))) return refuse('REQUESTER_OVERRIDE', `requester/identity override refused: ${k} (the requester is the caller's own token)`);
  }
  return null;
}

function checkTarget({ fileId, operation }) {
  if (typeof fileId !== 'string' || !FILE_ID.test(fileId)) return refuse('INVALID_ARGS', 'fileId must be a numeric string');
  if (typeof operation !== 'string' || operation === '') return refuse('INVALID_ARGS', 'operation must be a non-empty string');
  return null;
}

function checkOperation(ops, operation) {
  const op = Object.prototype.hasOwnProperty.call(ops, operation) ? ops[operation] : null;
  if (op && op.status === STATUS.ACCEPTED) return null;
  if (op) return refuse(`REFUSED_${op.status}`, `${operation} is ${op.status} (${op.evidence})`);
  if (writeIntents().has(operation)) return refuse('REFUSED_WRITE', `write intent refused by a READ tool: ${operation}`);
  return refuse('REFUSED_UNKNOWN', `unknown read operation: ${operation}`);
}

const inLength = (v, spec) => v.length >= (spec.minLength || 1) && v.length <= (spec.maxLength || 64);
const validString = (v, spec) => typeof v === 'string' && inLength(v, spec) && !/[\u0000-\u001f]/.test(v);
const typeOk = (v, spec) => (spec.type === 'boolean' ? typeof v === 'boolean' : validString(v, spec));

function checkFields(op, operation, args) {
  const allowed = new Set([...COMMON, ...Object.keys(op.args)]);
  const extra = Object.keys(args).find(k => !allowed.has(k));
  if (extra) return refuse('UNKNOWN_FIELD', `unknown field for ${operation}: ${extra}`);
  const missing = Object.keys(op.args).find(k => args[k] === undefined && !(op.optional || []).includes(k));
  if (missing) return refuse('INVALID_ARGS', `${operation} requires ${missing}`);
  const bad = Object.entries(op.args).find(([k, spec]) => args[k] !== undefined && !typeOk(args[k], spec));
  return bad ? refuse('INVALID_ARGS', `${bad[0]} must be a valid ${bad[1].type}`) : null;
}

function checkRange(range) {
  if (range === undefined) return null;
  const g = parseRange(range);
  if (!g) return refuse('INVALID_ARGS', 'range must be an A1 range');
  return (g.c2 - g.c1 + 1) * (g.r2 - g.r1 + 1) > MAX_READ_CELLS ? refuse('LIMIT', `range larger than ${MAX_READ_CELLS} cells`) : null;
}

/**
 * validateReadRequest(format, args) -> { ok:true, fileId, operation, args } | refusal (never throws).
 * Refusal codes: INVALID_ARGS | CREDENTIAL_FIELD | REQUESTER_OVERRIDE | UNKNOWN_FIELD | REFUSED_PARTIAL |
 * REFUSED_NOT_MEASURED | REFUSED_INTERNAL | REFUSED_WRITE | REFUSED_UNKNOWN | LIMIT.
 */
function validateReadRequest(format, args) {
  const ops = OPERATIONS[format];
  if (!ops) return refuse('INVALID_ARGS', `unknown format: ${String(format)}`);
  if (!args || typeof args !== 'object' || Array.isArray(args)) return refuse('INVALID_ARGS', 'arguments must be an object');
  const op = ops[args.operation];
  const stages = [() => checkNames(args), () => checkTarget(args), () => checkOperation(ops, args.operation),
    () => checkFields(op, args.operation, args), () => checkRange(args.range)];
  for (const stage of stages) { const refusal = stage(); if (refusal) return refusal; }
  const out = Object.fromEntries(Object.keys(op.args).filter(k => args[k] !== undefined).map(k => [k, args[k]]));
  if (out.range !== undefined) out.range = out.range.replace(/\$/g, '').toUpperCase();
  return { ok: true, fileId: args.fileId, operation: args.operation, args: Object.freeze(out) };
}

/** The readLive request an XLSX operation needs (the driver's own request fields; never `tiered`). */
function xlsxRequest(operation, a) {
  if (operation === 'read_range') return { windows: [{ sheet: a.sheet, range: a.range, format: a.format === true, style: a.style === true }] };
  if (operation === 'read_validations') return { windows: [], validationRanges: [{ sheet: a.sheet, range: a.range }] };
  return { windows: [] };
}

const CELL_FIELDS = ['v', 'f', 'merge', 'nf', 'fill', 'bold', 'italic', 'wrap', 'rot', 'readingOrder', 'indent', 'shrink', 'font', 'align', 'border'];

function readRange(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, range: a.range, error: 'SHEET_NOT_FOUND' };
  // only the requested window's cells, row-major; an unread cell is said so, never defaulted
  const cells = cellsOf(parseRange(a.range)).map(address => {
    const c = inv.cells[xlsxInventory.key(a.sheet, address)];
    return c ? { address, ...Object.fromEntries(CELL_FIELDS.filter(k => k in c).map(k => [k, c[k]])) } : { address, read: false };
  });
  return { sheet: a.sheet, range: a.range, cells };
}

const XLSX_PROJECTIONS = Object.freeze({
  workbook_overview: inv => ({ sheets: inv.sheets.map(s => ({ name: s.name, index: s.index, used: s.used })) }),
  read_range: readRange,
  read_validations: (inv, a) => {
    const v = inv.validationReads ? inv.validationReads[`${a.sheet}!${a.range}`] : undefined;
    return { sheet: a.sheet, range: a.range, validation: v === undefined ? 'UNOBSERVABLE:not-read' : v };
  },
});

// The DOCX projections. Never returned: numbering (PARTIAL), gridCalc (PARTIAL), titlePage/evenAndOdd (INTERNAL),
// notes text (PARTIAL), images (NOT_MEASURED), the internal gate.
const orNull = v => (v === undefined ? null : v);
const textOf = b => (b.kind === 'table' ? { rows: b.rows.map(r => r.map(c => c.text)) } : { text: b.text });
const formatOf = n => ({ text: n.text, props: orNull(n.props), format: orNull(n.format), runs: orNull(n.runs), hyperlinks: orNull(n.hyperlinks), noteRefs: orNull(n.noteRefs) });
const unitsOf = blocks => blocks.flatMap(b => (b.kind === 'list' ? b.items.map(it => ({ kind: 'list_item', ...formatOf(it) }))
  : (b.kind === 'paragraph' || b.kind === 'heading' ? [{ kind: b.kind, ...formatOf(b) }] : [])));
const cellOf = c => ({ text: c.text, simple: c.simple, format: c.format, paraFormats: c.paraFormats });

const DOCX_PROJECTIONS = Object.freeze({
  read_text: inv => ({ blocks: inv.blocks.map(b => ({ kind: b.kind, ...(b.kind === 'heading' ? { level: b.level } : {}),
    ...(b.kind === 'list' ? { ordered: b.ordered, items: b.items.map(i => i.text) } : {}), ...textOf(b), fingerprint: b.locator.fingerprint })) }),
  read_formatting: inv => ({ propertyRead: inv.propertyRead, defaultText: inv.document ? inv.document.defaultText : null, units: unitsOf(inv.blocks) }),
  read_tables: inv => ({ tables: inv.document && inv.document.tables ? inv.document.tables.map(t => ({ rows: t.rows.map(r => r.map(cellOf)) })) : null }),
  read_headers_footers: inv => ({ sections: inv.document && inv.document.sections ? inv.document.sections.map(x => ({ index: x.index, header: x.header, footer: x.footer })) : null }),
  read_bookmarks: inv => ({ bookmarks: inv.document ? inv.document.bookmarks : null }),
});


/** The ONE product-side driver constructor: configSource only (the M11.2 broker source), never url/user/pass. */
function productDriver(format, { configSource, loadPlaywright }) {
  if (!configSource || configSource.kind !== 'BROKER') throw new TypeError('a product READ driver needs the broker configSource');
  return format === 'xlsx' ? createEuroOfficeXlsxDriver({ configSource, loadPlaywright }) : createEuroOfficeDocxDriver({ configSource, loadPlaywright });
}

const opsOf = (format, status) => Object.entries(OPERATIONS[format]).filter(([, o]) => o.status === status).map(([k]) => k);

/** The MCP inputSchema: only ACCEPTED operations are offered; additionalProperties false. */
function inputSchema(format) {
  const accepted = opsOf(format, STATUS.ACCEPTED);
  const props = { fileId: { type: 'string', pattern: FILE_ID.source, description: 'Nextcloud fileId of the target file' },
    operation: { type: 'string', enum: accepted, description: 'the read operation (accepted read operations only)' } };
  for (const op of accepted) for (const [k, spec] of Object.entries(OPERATIONS[format][op].args)) props[k] = spec;
  return { type: 'object', properties: props, required: ['fileId', 'operation'], additionalProperties: false };
}

const readFailed = reason => ({ ok: false, stage: 'READ', code: 'READ_FAILED', reason, dispatched: false });
const FORMATS = Object.freeze({
  xlsx: { inventory: xlsxInventory, read: (d, v) => d.readLive(null, xlsxRequest(v.operation, v.args)), project: (p, v) => XLSX_PROJECTIONS[v.operation](p, v.args), complete: p => p.complete === true },
  // M11.3a.1: the product READ never needs the B21 section variant state -> the internal getters are not called
  docx: { inventory: docxInventory, read: d => d.readLive(null, { includeSectionVariantState: false }), project: (p, v) => DOCX_PROJECTIONS[v.operation](p), complete: p => p.propertyRead === true },
});

/** One live read on an OPEN driver: readLive -> normalizeInventory -> projection. Never throws. */
async function readOnce(f, driver, v, format) {
  try {
    const r = await f.read(driver, v);
    if (!r || !r.ok) return readFailed((r && r.detail) || 'read failed');
    const primary = f.inventory.normalizeInventory(r.payload);
    if (!primary) return readFailed('the live read produced no inventory');
    return { ok: true, tool: `${format}_read`, fileId: v.fileId, operation: v.operation, complete: f.complete(primary), result: f.project(primary, v) };
  } catch (e) {
    return readFailed(String((e && e.message) || e));
  }
}

const closeDetail = closed => (closed && closed.detail ? String(closed.detail) : null);

/**
 * createReadTool(format, { configSource, browsersPath, gate, runtime, loadPlaywright, createDriver }) -> { description, inputSchema, handler(args) }
 *
 * M11.3b.2 request path (each step refuses BEFORE the next one starts):
 *   validate (pre-dispatch: schema, credential/requester fields, operation class)
 *   -> admission: the process-wide session gate (MAX_ACTIVE_SESSIONS_PER_PROCESS=1) -> BUSY, no queue
 *   -> browser runtime check (pinned playwright-core / revision under playwright.path) -> BROWSER_RUNTIME_UNAVAILABLE
 *   -> driver open (the broker call happens here) -> readLive -> projection -> close -> lease released.
 * `gate`: the product passes ONE gate shared by every tool of the process. `runtime` / `loadPlaywright` / `createDriver`
 * are test seams; the product default is the configured browser runtime and productDriver (broker configSource only).
 */
function createReadTool(format, { configSource, browsersPath, gate = createSessionGate(), runtime = null, loadPlaywright = null, createDriver = productDriver } = {}) {
  const f = FORMATS[format];
  const rt = runtime || (loadPlaywright ? { check: () => ({ ok: true }), load: loadPlaywright } : createBrowserRuntime({ browsersPath }));
  return {
    description: `Read-only ${format.toUpperCase()} read of a Nextcloud file as the calling requester (own credential via the broker). Accepted read operations only; never writes or saves.`,
    inputSchema: inputSchema(format),
    async handler(args) {
      const v = validateReadRequest(format, args);
      if (!v.ok) return v;
      const lease = gate.tryEnter();
      if (!lease) return Object.freeze({ ok: false, stage: 'ADMISSION', code: 'BUSY', reason: `an editor session is active (MAX_ACTIVE_SESSIONS_PER_PROCESS=${gate.max}); retry after it closes`, dispatched: false });
      try {
        const ready = rt.check();
        if (!ready.ok) return { ok: false, stage: 'RUNTIME', code: ready.code || 'BROWSER_RUNTIME_UNAVAILABLE', reason: ready.detail, dispatched: false };
        const driver = createDriver(format, { configSource, loadPlaywright: rt.load });
        lease.track(driver);
        const opened = (await driver.open({ fileId: v.fileId, operation: 'READ' })) || {};
        if (!opened.ok) return { ok: false, stage: 'OPEN', code: opened.code || 'OPEN_FAILED', reason: opened.detail || 'open failed', dispatched: false };
        const out = await readOnce(f, driver, v, format);
        // the read-only proof the driver itself gives: a session that never dispatched settles with no save traffic
        out.close = closeDetail(await driver.close().catch(e => ({ detail: String((e && e.message) || e) })));
        return out;
      } finally {
        lease.release();
      }
    },
  };
}

module.exports = { OPERATIONS, STATUS, validateReadRequest, inputSchema, createReadTool, productDriver, xlsxRequest, writeIntents, CREDENTIAL_FIELDS, REQUESTER_FIELDS };
