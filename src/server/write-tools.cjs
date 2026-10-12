'use strict';

/**
 * M11.4a.1 -- the product WRITE tool `xlsx_execute_task`: handler + strict schema, WITHHELD (owner decision, Talk
 * owner channel 2026-10-10, after M11.4a.0 DESIGN_ACCEPTED). Design: docs/M11-4A0-WRITE-DESIGN.md; supported shapes and
 * refusal matrix: docs/M11-4A1-WRITE-HANDLER.md.
 *
 *   MCP arguments { fileId, task: { id?, steps: [{ intent, ... }] } }
 *     -> recursive security validation (credential / requester / broker / identity override names, any depth)
 *     -> strict schema (additionalProperties false at every level, one explicit step schema per allowlisted intent)
 *     -> intent class (allowlist | temporary boundary | not allowlisted | cross-format | read operation | unknown)
 *     -> supported-shape constraints + the canonical planner on each step alone (REFUSED_SHAPE)
 *     -> composition: sole-step intents + the canonical planner on the whole task + validatePlan (REFUSED_COMPOSITION)
 *     -> execute(validated)   (M11.4a.2: the process-wide session gate -> browser runtime -> the product driver opened
 *        in WRITE mode -> withSession(xlsxAdapter, createExecutor(task, sequential)) -> mapExecutionResult)
 *
 * The task is the canonical executor task `{ id, steps: [{ intent, ... }] }`: this module validates and passes it on,
 * it never translates it into a second operation language. Any refused step refuses the WHOLE task before anything
 * starts: no broker call, no browser, no driver, no executor, no write.
 *
 * Exposure: tools.cjs EXPOSED_CAPABILITY_TOOLS does not name this tool -> it is neither listed nor callable
 * (static exact-set guard). There is no DOCX handler: DOCX_WRITE_ACCEPTED is empty and docx_execute_task stays withheld.
 */

const { xlsxAdapter, liveScope } = require('../adapters/xlsx/adapter.cjs');
const { parseRange } = require('../adapters/xlsx/a1.cjs');
const { validatePlan } = require('../core/execution/plan.cjs');
const docxIntents = require('../adapters/docx/intents.cjs');
const { OPERATIONS: READ_OPERATIONS, CREDENTIAL_FIELDS, REQUESTER_FIELDS, productDriver } = require('./read-tools.cjs');
const { withSession } = require('../core/session/lifecycle.cjs');
const { createExecutor } = require('../core/execution/executor.cjs');
const { createBrowserRuntime } = require('./browser-runtime.cjs');
const { createSessionGate } = require('./session-gate.cjs');

// The first WRITE release allowlist -- EXACTLY these six (owner 2026-10-10). Widening it is an owner decision.
const XLSX_WRITE_ALLOWLIST = Object.freeze(['write_range', 'insert_rows', 'delete_rows', 'rename_sheet', 'unmerge_range', 'set_chart_title']);
const DOCX_WRITE_ALLOWLIST = Object.freeze([]);

// The temporary direct-internal boundaries (M10 legacy set, M11.5 migrates them): refused by name, never routed here.
const TEMPORARY_BOUNDARY = Object.freeze({
  copy_sheet: 'B03', format_range: 'B05-WRITE',
  set_print_fit: 'B10', set_header_footer: 'B10', set_page_break: 'B10', set_first_page_number: 'B10',
  set_print_centering: 'B10', set_print_scale: 'B10', set_paper_size: 'B10', set_header_footer_flag: 'B10',
  remove_list: 'B18', outdent_list_item: 'B18', edit_header: 'B21', edit_footer: 'B21',
});

const MAX_STEPS = 50; // the M11.4a.1 release bound (the core plan ceiling is 1000)
const MAX_CELLS = 1000; // = the write_range / unmerge_range planner ceiling (cells.cjs / merge.cjs)
const MAX_TEXT = 32767; // = the write_range planner text ceiling (cells.cjs)

// ---- the step schemas (JSON Schema subset; the SAME objects validate and are published as inputSchema) ----

const ID = Object.freeze({ type: 'string', pattern: '^[A-Za-z0-9_.-]{1,64}$' });
// a sheet named only to ADDRESS cells (write_range / unmerge_range): any valid worksheet name
const SHEET_ADDRESS = Object.freeze({ type: 'string', pattern: '^[^\\u0000-\\u001f\\[\\]:*?/\\\\]{1,31}$' });
// a sheet that formulas / references may name (structure shift, rename, chart): the measured no-quote class only
const SHEET_BARE = Object.freeze({ type: 'string', pattern: '^[A-Za-z_][A-Za-z0-9_]{0,30}$' });
const RANGE = Object.freeze({ type: 'string', pattern: '^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$' });
const CELL_VALUE = Object.freeze({ anyOf: [
  { type: 'string', minLength: 1, maxLength: MAX_TEXT, pattern: '^\\S(?:[\\s\\S]*\\S)?$' },
  { type: 'number' },
] });

const step = (intent, properties, required) => Object.freeze({
  type: 'object',
  properties: Object.freeze({ intent: { const: intent }, id: ID, ...properties }),
  required: Object.freeze(['intent', ...required]),
  additionalProperties: false,
});

const ROWS = (intent) => step(intent, {
  sheet: SHEET_BARE,
  at: { type: 'integer', minimum: 1, maximum: 1048576 },
  count: { type: 'integer', minimum: 1, maximum: 100 },
  anchor: { type: 'object', properties: { column: { type: 'string', pattern: '^[A-Z]{1,3}$' },
    value: { anyOf: [{ type: 'string', minLength: 1, maxLength: 255 }, { type: 'number' }] } },
  required: ['column', 'value'], additionalProperties: false },
}, ['sheet', 'at', 'count', 'anchor']);

const STEP_SCHEMAS = Object.freeze({
  write_range: step('write_range', { sheet: SHEET_ADDRESS, range: RANGE,
    values: { type: 'array', minItems: 1, maxItems: MAX_CELLS, items: { type: 'array', minItems: 1, maxItems: MAX_CELLS, items: CELL_VALUE } } },
  ['sheet', 'range', 'values']),
  insert_rows: ROWS('insert_rows'),
  delete_rows: ROWS('delete_rows'),
  rename_sheet: step('rename_sheet', { sheet: SHEET_BARE, name: SHEET_BARE }, ['sheet', 'name']),
  unmerge_range: step('unmerge_range', { sheet: SHEET_ADDRESS, range: RANGE }, ['sheet', 'range']),
  set_chart_title: step('set_chart_title', { sheet: SHEET_BARE, name: { type: 'string', pattern: '^[A-Za-z_][A-Za-z0-9_. -]{0,99}$' },
    title: { type: 'string', minLength: 1, maxLength: 255, pattern: '^[^\\r\\n]+$' } }, ['sheet', 'name', 'title']),
});

// ---- the supported-shape constraints a JSON Schema cannot state (each returns null or why) ----

const CELL_REF_SHAPED = n => /^[A-Za-z]{1,3}\d+$/.test(n) || /^([RrCc]\d*|[Rr]\d*[Cc]\d*)$/.test(n);
const RESERVED = n => /^(history|true|false)$/i.test(n);
const bareNameProblem = n => (CELL_REF_SHAPED(n) ? 'shaped like a cell reference' : RESERVED(n) ? 'a reserved word' : null);

function rangeProblem(range) {
  const g = parseRange(range);
  if (!g) return 'range must be a plain A1 range, top-left first';
  return g.count > MAX_CELLS ? `range has ${g.count} cells, over the ${MAX_CELLS} ceiling` : null;
}

const CONSTRAINTS = Object.freeze({
  write_range(s) {
    const why = rangeProblem(s.range);
    if (why) return why;
    const g = parseRange(s.range);
    const fits = s.values.length === g.rows && s.values.every(row => row.length === g.cols);
    return fits ? null : `values must be a ${g.rows}x${g.cols} matrix matching ${s.range}`;
  },
  insert_rows: s => bareNameProblem(s.sheet),
  delete_rows: s => bareNameProblem(s.sheet),
  rename_sheet: s => bareNameProblem(s.sheet) || bareNameProblem(s.name)
    || (s.name.toLowerCase() === s.sheet.toLowerCase() ? 'the new name equals the old one, case-insensitively' : null),
  unmerge_range: s => rangeProblem(s.range),
  set_chart_title: s => bareNameProblem(s.sheet),
});

// a step that changes the coordinates / sheet graph the other steps would be resolved against: the only step
const SOLE_STEP = Object.freeze(['insert_rows', 'delete_rows', 'rename_sheet']);

// ---- a minimal validator for the schema subset above ----

const typeOf = v => (Array.isArray(v) ? 'array' : v === null ? 'null' : typeof v);
const isType = (v, t) => (t === 'integer' ? Number.isInteger(v) : t === 'number' ? typeof v === 'number' && Number.isFinite(v) : typeOf(v) === t);

/** -> null | { code, path, reason }. */
function check(schema, v, path) {
  if (schema.anyOf) return schema.anyOf.some(s => check(s, v, path) === null) ? null : { code: 'REFUSED_SHAPE', path, reason: `${path}: unsupported value` };
  if ('const' in schema) return v === schema.const ? null : { code: 'REFUSED_SHAPE', path, reason: `${path} must be ${schema.const}` };
  if (!isType(v, schema.type)) return { code: 'REFUSED_SHAPE', path, reason: `${path} must be ${schema.type === 'number' ? 'a finite number' : `a ${schema.type}`}` };
  if (schema.type === 'string') {
    if (v.length < (schema.minLength || 0) || (schema.maxLength && v.length > schema.maxLength)) return { code: 'REFUSED_SHAPE', path, reason: `${path} length out of range` };
    if (schema.pattern && !new RegExp(schema.pattern, 'u').test(v)) return { code: 'REFUSED_SHAPE', path, reason: `${path} is outside the supported shape` };
  }
  if ((schema.type === 'integer' || schema.type === 'number') && ((schema.minimum !== undefined && v < schema.minimum) || (schema.maximum !== undefined && v > schema.maximum))) {
    return { code: 'REFUSED_SHAPE', path, reason: `${path} out of range` };
  }
  if (schema.type === 'array') {
    if (v.length < (schema.minItems || 0) || (schema.maxItems && v.length > schema.maxItems)) return { code: 'REFUSED_SHAPE', path, reason: `${path} has ${v.length} items` };
    for (const [i, item] of v.entries()) { const bad = check(schema.items, item, `${path}[${i}]`); if (bad) return bad; }
  }
  if (schema.type === 'object') {
    const extra = Object.keys(v).find(k => !Object.prototype.hasOwnProperty.call(schema.properties, k));
    if (extra !== undefined) return { code: 'UNKNOWN_FIELD', path: `${path}.${extra}`, reason: `unknown field ${path}.${extra}` };
    const missing = (schema.required || []).find(k => v[k] === undefined);
    if (missing) return { code: 'REFUSED_SHAPE', path: `${path}.${missing}`, reason: `${path}.${missing} is required` };
    for (const [k, s] of Object.entries(schema.properties)) {
      if (v[k] === undefined) continue;
      const bad = check(s, v[k], `${path}.${k}`);
      if (bad) return bad;
    }
  }
  return null;
}

// ---- the refusal stages ----

const refuse = (code, reason, extra = {}) => Object.freeze({ ok: false, stage: 'PRE_DISPATCH', code, reason, dispatched: false, wholeTask: true, ...extra });

// identity / route override names refused at ANY depth (normalized: lower case, no '_' / '-' / space)
const ROUTE_FIELDS = new Set(['brokerurl', 'broker', 'brokerendpoint', 'url', 'baseurl', 'endpoint', 'ncurl', 'nextcloudurl',
  'nextcloud', 'ncidentity', 'ncinstance', 'instance', 'host', 'server']);
const normName = k => String(k).toLowerCase().replace(/[_\-\s]/g, '');

/** The recursive security walk: every object key at every depth (arrays included). */
function checkNamesDeep(value, path) {
  if (Array.isArray(value)) {
    for (const [i, item] of value.entries()) { const r = checkNamesDeep(item, `${path}[${i}]`); if (r) return r; }
    return null;
  }
  if (!value || typeof value !== 'object') return null;
  for (const k of Object.keys(value)) {
    const n = normName(k);
    if (CREDENTIAL_FIELDS.has(n)) return refuse('CREDENTIAL_FIELD', `credential-like field refused: ${path}.${k}`, { path: `${path}.${k}` });
    if (REQUESTER_FIELDS.has(n)) return refuse('REQUESTER_OVERRIDE', `requester/identity override refused: ${path}.${k} (the requester is the caller's own token)`, { path: `${path}.${k}` });
    if (ROUTE_FIELDS.has(n)) return refuse('ROUTE_OVERRIDE', `broker / Nextcloud route override refused: ${path}.${k}`, { path: `${path}.${k}` });
    const r = checkNamesDeep(value[k], `${path}.${k}`);
    if (r) return r;
  }
  return null;
}

const xlsxRouted = () => new Set(require('./tools.cjs').xlsxRoutedIntents());
const docxRouted = () => new Set(Array.isArray(docxIntents.INTENTS) ? docxIntents.INTENTS.map(r => r.intent) : Object.keys(docxIntents.INTENTS));
const readOperations = () => new Set([...Object.keys(READ_OPERATIONS.xlsx), ...Object.keys(READ_OPERATIONS.docx)]);

/** The class of one step's intent in the XLSX WRITE tool: null = allowlisted, else the refusal. */
function intentRefusal(intent, index) {
  const at = { stepIndex: index, intent };
  if (typeof intent !== 'string' || intent === '') return refuse('MALFORMED_TASK', `step ${index}: intent must be a non-empty string`, { stepIndex: index });
  if (XLSX_WRITE_ALLOWLIST.includes(intent)) return null;
  if (Object.prototype.hasOwnProperty.call(TEMPORARY_BOUNDARY, intent)) return refuse('REFUSED_TEMPORARY_BOUNDARY', `step ${index}: ${intent} depends on the temporary direct-internal boundary ${TEMPORARY_BOUNDARY[intent]} (migrated in M11.5, not released)`, at);
  if (xlsxRouted().has(intent)) return refuse('REFUSED_NOT_ALLOWLISTED', `step ${index}: ${intent} is a routed XLSX intent outside the first WRITE release allowlist`, at);
  if (docxRouted().has(intent)) return refuse('REFUSED_CROSS_FORMAT', `step ${index}: ${intent} is a DOCX intent (DOCX_WRITE_ACCEPTED is empty; docx_execute_task is withheld)`, at);
  if (readOperations().has(intent)) return refuse('REFUSED_READ_OPERATION', `step ${index}: ${intent} is a READ operation (use the READ tool)`, at);
  return refuse('REFUSED_UNKNOWN_INTENT', `step ${index}: unknown intent ${intent}`, at);
}

const TASK_SCHEMA_KEYS = Object.freeze(['id', 'steps']);
const FILE_ID = /^[0-9]{1,15}$/;

/**
 * validateWriteRequest(args) -> { ok:true, fileId, task, plannedOperations } | refusal (never throws).
 * Every refusal is PRE_DISPATCH and refuses the WHOLE task. Codes: INVALID_ARGS | CREDENTIAL_FIELD |
 * REQUESTER_OVERRIDE | ROUTE_OVERRIDE | UNKNOWN_FIELD | MALFORMED_TASK | EMPTY_TASK | LIMIT | REFUSED_UNKNOWN_INTENT |
 * REFUSED_NOT_ALLOWLISTED | REFUSED_TEMPORARY_BOUNDARY | REFUSED_CROSS_FORMAT | REFUSED_READ_OPERATION | REFUSED_SHAPE |
 * REFUSED_COMPOSITION.
 */
function validateWriteRequest(args) {
  if (!args || typeof args !== 'object' || Array.isArray(args)) return refuse('INVALID_ARGS', 'arguments must be an object');
  const names = checkNamesDeep(args, 'args');
  if (names) return names;
  const extra = Object.keys(args).find(k => !['fileId', 'task'].includes(k));
  if (extra !== undefined) return refuse('UNKNOWN_FIELD', `unknown field: ${extra}`, { path: `args.${extra}` });
  if (typeof args.fileId !== 'string' || !FILE_ID.test(args.fileId)) return refuse('INVALID_ARGS', 'fileId must be a numeric string');
  const task = args.task;
  if (!task || typeof task !== 'object' || Array.isArray(task)) return refuse('MALFORMED_TASK', 'task must be an object { id?, steps }');
  const taskExtra = Object.keys(task).find(k => !TASK_SCHEMA_KEYS.includes(k));
  if (taskExtra !== undefined) return refuse('UNKNOWN_FIELD', `unknown task field: ${taskExtra}`, { path: `args.task.${taskExtra}` });
  if (task.id !== undefined && check(ID, task.id, 'args.task.id')) return refuse('MALFORMED_TASK', 'task.id must match ^[A-Za-z0-9_.-]{1,64}$');
  if (!Array.isArray(task.steps)) return refuse('MALFORMED_TASK', 'task.steps must be an array');
  if (task.steps.length === 0) return refuse('EMPTY_TASK', 'task.steps is empty');
  if (task.steps.length > MAX_STEPS) return refuse('LIMIT', `task has ${task.steps.length} steps, over the ${MAX_STEPS} release bound`);
  for (const [i, s] of task.steps.entries()) {
    if (!s || typeof s !== 'object' || Array.isArray(s)) return refuse('MALFORMED_TASK', `step ${i} is not an object`, { stepIndex: i });
    const cls = intentRefusal(s.intent, i);
    if (cls) return cls;
    const bad = check(STEP_SCHEMAS[s.intent], s, `args.task.steps[${i}]`);
    if (bad) return refuse(bad.code, `step ${i}: ${s.intent} -- ${bad.reason}`, { stepIndex: i, intent: s.intent, path: bad.path });
    const why = CONSTRAINTS[s.intent](s);
    if (why) return refuse('REFUSED_SHAPE', `step ${i}: ${s.intent} -- ${why}`, { stepIndex: i, intent: s.intent });
    // the canonical planner on the step ALONE (pure): its per-value refusals (e.g. formula-like / number-like text)
    // are shape refusals too -- the single source of those rules stays the adapter
    const alone = xlsxAdapter.planOperations({ steps: [s] });
    if (!alone || !alone.ok) return refuse('REFUSED_SHAPE', `step ${i}: ${s.intent} -- ${(alone && alone.detail) || 'the planner refused the step'}`, { stepIndex: i, intent: s.intent });
  }
  const sole = task.steps.findIndex(s => SOLE_STEP.includes(s.intent));
  if (sole >= 0 && task.steps.length > 1) return refuse('REFUSED_COMPOSITION', `step ${sole}: ${task.steps[sole].intent} must be the only step of its task`, { stepIndex: sole, intent: task.steps[sole].intent });
  // the canonical planner, pre-dispatch (pure: no session) -- the same call the executor makes in-session
  const planned = xlsxAdapter.planOperations(task);
  if (!planned || !planned.ok) return refuse('REFUSED_COMPOSITION', (planned && planned.detail) || 'the planner refused the task', planned && planned.index !== undefined ? { stepIndex: planned.index } : {});
  const validity = validatePlan(planned.operations);
  if (!validity.ok) return refuse('REFUSED_COMPOSITION', `plan invalid: ${validity.problems.join('; ')}`);
  return Object.freeze({ ok: true, fileId: args.fileId, task: toExecutorTask(task), plannedOperations: planned.operations.length });
}

/** The canonical executor task, verbatim (a deep copy, frozen): no second task language. */
function toExecutorTask(task) {
  const copy = JSON.parse(JSON.stringify({ ...(task.id !== undefined ? { id: task.id } : {}), steps: task.steps }));
  const freeze = o => { if (o && typeof o === 'object') { Object.values(o).forEach(freeze); Object.freeze(o); } return o; };
  return freeze(copy);
}

/** The MCP inputSchema: fileId + the canonical task, each step one of the six explicit step schemas. */
function inputSchema() {
  return {
    type: 'object',
    properties: {
      fileId: { type: 'string', pattern: FILE_ID.source, description: 'Nextcloud fileId of the target XLSX file' },
      task: { type: 'object', properties: { id: ID, steps: { type: 'array', minItems: 1, maxItems: MAX_STEPS,
        items: { oneOf: XLSX_WRITE_ALLOWLIST.map(i => STEP_SCHEMAS[i]) } } }, required: ['steps'], additionalProperties: false },
    },
    required: ['fileId', 'task'],
    additionalProperties: false,
  };
}

// ---- result and error mapping (FinalTaskResult -> MCP result; no new outcome taxonomy) ----

// the write-denial names and WHERE they are decided (M11.4a.1: mapping only, no storage probing)
const WRITE_ERRORS = Object.freeze({
  ACCESS_DENIED: 'ACCESS_DENIED', // broker 403 PRODUCT_OCS_CONFIG_DENY_LEVEL/AUTHZ/... (broker-config-source.cjs), stage OPEN
  TARGET_READ_ONLY: 'TARGET_READ_ONLY', // an open WRITE config without permissions.edit, stage OPEN (before the first dispatch)
  STORAGE_WRITE_REJECTED: 'STORAGE_WRITE_REJECTED', // save completion success:false ('completion-not-successful'), stage PERSIST
});

// the broker's answer to a WRITE request Nextcloud can serve only with a VIEW config (broker-config-source.cjs maps it to
// ACCESS_DENIED and keeps the broker code in the detail): for the WRITE tool this is the read-only target
const BROKER_READ_ONLY = 'PRODUCT_OCS_CONFIG_DENY_CONFIG_BELOW_REQUEST';

/** The open-stage error class: the driver/broker code as-is; a VIEW-only answer to a WRITE request is TARGET_READ_ONLY. */
function mapOpenFailure(openFailure) {
  const f = openFailure || {};
  if (f.code === WRITE_ERRORS.TARGET_READ_ONLY || String(f.detail || '').includes(BROKER_READ_ONLY)) return WRITE_ERRORS.TARGET_READ_ONLY;
  return typeof f.code === 'string' && f.code !== '' ? f.code : 'OPEN_FAILED';
}

// the canonical save reason carries its stage ('barrier:<reason>' | 'close:<reason>'): the completion is the same signal
const storageRejected = reason => /(^|:)completion-not-successful$/.test(String(reason || ''));

/** mapExecutionResult(final, v) -> the MCP result of an executed task (canonical statuses only). */
function mapExecutionResult(final, v) {
  const base = { tool: 'xlsx_execute_task', fileId: v.fileId, taskId: v.task.id === undefined ? null : v.task.id };
  if (final.openFailure) {
    return { ok: false, stage: 'OPEN', code: mapOpenFailure(final.openFailure), reason: final.openFailure.detail || 'open failed', dispatched: false, ...base };
  }
  const exec = final.execution || {};
  const reason = final.saveEvidence ? final.saveEvidence.reason : null;
  const rejected = storageRejected(reason);
  const layer = (list, id) => { const v = (list || []).find(x => x.evidence && x.evidence.operationId === id); return v ? v.verdict : null; };
  return {
    ok: final.ok === true && !rejected,
    stage: 'EXECUTED',
    ...(rejected ? { code: WRITE_ERRORS.STORAGE_WRITE_REJECTED } : {}),
    ...base,
    outcome: final.outcome,
    editOk: final.editOk === true,
    documentOutcome: final.documentOutcome,
    mutationCertainty: final.mutationCertainty,
    authority: final.authority,
    noOp: final.noOp === true,
    operations: (exec.appliedOperations || []).map(e => ({ id: e.operationId, status: e.status, verification: e.verification || null,
      l1: layer(exec.primitiveVerification, e.operationId), l3: layer(exec.preservationVerification, e.operationId),
      ...(e.uncertainty && e.uncertainty.reason ? { uncertainty: e.uncertainty.reason } : {}) })),
    wholeTask: exec.wholeTaskVerification ? exec.wholeTaskVerification.verdict : null, // L2
    operationAccounting: exec.operationAccounting || null,
    persistence: {
      outcome: final.persistenceOutcome,
      state: final.persistenceState,
      reason,
      storedContentVerified: false,
      // EDITOR_CLEAN is the barrier's own fresh canSave poll after the forceSave (the editor reports nothing left to
      // save) -- never a forceSave request or timeout; a timeout is a FAILED barrier
      editorCleanSource: final.persistenceState === 'EDITOR_CLEAN' && final.persistence && final.persistence.editorClean === true ? 'fresh-canSave-poll-after-forceSave' : null,
      error: rejected ? WRITE_ERRORS.STORAGE_WRITE_REJECTED : null,
    },
    cleanupOutcome: final.cleanupOutcome,
  };
}

/**
 * The WRITE open mode (M11.4a.2): the shared session lifecycle opens a driver with { url, user, pass, fileId } and no
 * operation; the product WRITE path names it explicitly -- the broker is asked for requestedMode WRITE. Every other
 * driver method is the canonical driver's own (bound); the shared driver's default semantics are not changed.
 */
function productWriteDriver(driver) {
  return new Proxy(driver, {
    get(target, prop) {
      if (prop === 'open') return args => target.open({ fileId: args && args.fileId, operation: 'WRITE' });
      const v = target[prop];
      return typeof v === 'function' ? v.bind(target) : v;
    },
  });
}

/**
 * createWriteExecution({ configSource, browsersPath, gate, runtime, loadPlaywright, createDriver, session }) -> execute(v)
 *   the process-wide session gate (MAX_ACTIVE_SESSIONS_PER_PROCESS=1, shared with the READ tools) -> BUSY, no queue
 *   -> browser runtime check -> BROWSER_RUNTIME_UNAVAILABLE
 *   -> the product driver (broker configSource only) opened in WRITE mode
 *   -> withSession(xlsxAdapter, createExecutor(task, sequential)): the canonical lifecycle, executor, verification and
 *      persistence barrier -- nothing here decides APPLIED, saves or retries
 *   -> mapExecutionResult; the lease is released on every exit.
 * `runtime` / `loadPlaywright` / `createDriver` / `session` (the lifecycle's own ready / persistence / now options) are
 * test seams; the product default is the configured browser runtime and productDriver.
 */
function createWriteExecution({ configSource, browsersPath, gate = createSessionGate(), runtime = null, loadPlaywright = null, createDriver = productDriver, session = {} } = {}) {
  const rt = runtime || (loadPlaywright ? { check: () => ({ ok: true }), load: loadPlaywright } : createBrowserRuntime({ browsersPath }));
  return async function execute(v) {
    const lease = gate.tryEnter();
    if (!lease) return Object.freeze({ ok: false, stage: 'ADMISSION', code: 'BUSY', reason: `an editor session is active (MAX_ACTIVE_SESSIONS_PER_PROCESS=${gate.max}); retry after it closes`, dispatched: false });
    try {
      const ready = rt.check();
      if (!ready.ok) return { ok: false, stage: 'RUNTIME', code: ready.code || 'BROWSER_RUNTIME_UNAVAILABLE', reason: ready.detail, dispatched: false };
      const driver = productWriteDriver(createDriver('xlsx', { configSource, loadPlaywright: rt.load }));
      lease.track(driver);
      const options = { driver, fileId: v.fileId, ...(session.ready ? { ready: session.ready } : {}), ...(session.persistence ? { persistence: session.persistence } : {}), ...(session.now ? { now: session.now } : {}) };
      const final = await withSession(options, xlsxAdapter, createExecutor(v.task, { strategy: 'sequential', observation: { live: liveScope(v.task) } }));
      return mapExecutionResult(final, v);
    } finally {
      lease.release();
    }
  };
}

/** createWriteTool('xlsx', { execute, ...createWriteExecution options }) -> { description, inputSchema, handler(args) }. */
function createWriteTool(format, { execute = null, ...wiring } = {}) {
  if (format !== 'xlsx') throw new TypeError(`no WRITE handler for ${String(format)} (DOCX_WRITE_ACCEPTED is empty)`);
  const run = execute || createWriteExecution(wiring);
  return {
    description: `XLSX WRITE task on a Nextcloud file as the calling requester (own credential via the broker): the canonical task { id?, steps }, run by the canonical executor and persisted with a confirmed forcesave. Supported intents: ${XLSX_WRITE_ALLOWLIST.join(', ')}; any other intent or shape is refused before dispatch.`,
    inputSchema: inputSchema(),
    async handler(args) {
      const v = validateWriteRequest(args);
      if (!v.ok) return v;
      return run(v);
    },
  };
}

module.exports = {
  XLSX_WRITE_ALLOWLIST, DOCX_WRITE_ALLOWLIST, TEMPORARY_BOUNDARY, STEP_SCHEMAS, CONSTRAINTS, SOLE_STEP, MAX_STEPS,
  ROUTE_FIELDS, WRITE_ERRORS, BROKER_READ_ONLY, validateWriteRequest, toExecutorTask, inputSchema, mapExecutionResult,
  mapOpenFailure, createWriteTool, createWriteExecution, productWriteDriver, checkNamesDeep,
};
