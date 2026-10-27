'use strict';

/**
 * GH #219 2/2 acceptance (offline): the WITHHELD, UNREGISTERED product WRITE tool `docx_execute_task` --
 * handler + strict schema + allowlist/refusal + pre-dispatch security + canonical task adapter, docx
 * mirror of test/acceptance/m11-4a1-write-handler.test.cjs. Design: docs/GH219-DOCX-WRITE-CONTRACT-DESIGN.md.
 *
 * DOCX_WRITE_ALLOWLIST was EMPTY in GH #219 (decision 2 and 3 of the design doc's section 9): every docx intent was
 * REFUSED_NOT_ALLOWLISTED, never dispatched. GH #228 (PROPOSED STATE, owner decision pending) allowlists rewrite_paragraph only;
 * every OTHER docx intent is still REFUSED_NOT_ALLOWLISTED (the cases below pick a non-allowlisted intent). "ambiguous target" (one of the GH #219
 * required-test categories) is NOT re-tested here -- it is already covered by the docx adapter's own,
 * family-level tests (e.g. the header-footer family's unsafe-inheritance-refused-before-dispatch cases),
 * which this module does not duplicate or re-express.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const { EXPOSED_CAPABILITY_TOOLS, TOOL_ROWS } = require('../../src/server/tools.cjs');
const { INTENTS } = require('../../src/adapters/docx/intents.cjs');
const { NEW_INTENTS } = require('../../src/adapters/docx/capabilities.cjs');

const FILE = '1303626';
const req = steps => ({ fileId: FILE, task: { steps } });

/** A harness with a counting execute seam, same discipline as m11-4a1's harness() -- but built directly
 * from createDocxWriteTool (this tool is not wired into buildServerTools/server-tools.cjs in this unit). */
function harness() {
  const n = { executor: 0, tasks: [] };
  const execute = async v => { n.executor += 1; n.tasks.push(v.task); return { ok: true, stage: 'SEAM' }; };
  return { n, tool: D.createDocxWriteTool({ execute }) };
}

test('INTENT_INVENTORY (GH #219 2/2, decision 1): 28 M4-migrated + 10 post-M4 = 38, resolving the "33 vs 28" question', () => {
  assert.equal(INTENTS.length, 28);
  // the 10 post-M4 additions adapter.cjs's own BY_INTENT merges in alongside intents.cjs's INTENTS
  // (GH #268c: + set_image_size, set_image_wrap -- adapter-complete, PREPARATION, not yet in the
  // product schema/allowlist, same state format_table/set_numbering were in before their own release)
  assert.deepEqual(NEW_INTENTS.map(i => i.intent).sort(), ['add_comment_on_match', 'define_style', 'format_table', 'modify_style', 'reply_comment', 'resolve_comment', 'set_image_size', 'set_image_wrap', 'set_numbering', 'set_page_setup']);
  assert.equal(INTENTS.length + NEW_INTENTS.length, 38, 'matches the card\'s "33 adapter semantics" + GH #268c\'s 2 additions');
  assert.equal(D.docxRoutedIntents().size, 38);
  for (const i of [...INTENTS, ...NEW_INTENTS]) assert.ok(D.docxRoutedIntents().has(i.intent), i.intent);
});

test('ALLOWLIST: DOCX_WRITE_ALLOWLIST = [rewrite_paragraph, insert_paragraph_before, insert_paragraph_after, delete_paragraph, move_paragraph_before, move_paragraph_after, apply_list, set_list_level, set_list_type, indent_list_item, outdent_list_item, remove_list] (GH #228 + #233 + #234 + #235 + #236 + #237); registered; createDocxWriteTool works as a plain function', () => {
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], ['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'apply_list', 'set_list_level', 'indent_list_item', 'outdent_list_item', 'remove_list', 'replace_table_cell', 'format_text', 'insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table', 'set_page_setup', 'edit_header', 'edit_footer', 'define_style', 'modify_style', 'set_numbering', 'set_list_type', 'add_footnote', 'add_endnote', 'add_link', 'add_bookmark', 'resolve_comment', 'reply_comment', 'add_comment_on_match', 'set_image_size', 'set_image_wrap']);
  assert.equal(Object.isFrozen(W.DOCX_WRITE_ALLOWLIST), true);
  assert.deepEqual(Object.keys(D.DOCX_STEP_SCHEMAS), [...W.DOCX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(D.DOCX_CONSTRAINTS), [...W.DOCX_WRITE_ALLOWLIST]);
  // GH #228 (PROPOSED): the tool IS registered -- EXPOSED_CAPABILITY_TOOLS / TOOL_ROWS name it
  assert.deepEqual([...EXPOSED_CAPABILITY_TOOLS], ['xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  assert.equal(TOOL_ROWS.find(r => r.TOOL === 'docx_execute_task').EXPOSE_IN_PRODUCT_MCP, 'YES');
  // the write-tools.cjs xlsx path is untouched: createWriteTool('docx') still throws (this is a SEPARATE module/function)
  assert.throws(() => W.createWriteTool('docx'), /no WRITE handler for docx/);
  assert.equal(typeof D.createDocxWriteTool({}).handler, 'function');
});

test('DOCX_SCHEMA: top level { fileId, task }, additionalProperties false everywhere, one step schema per allowlisted intent', () => {
  const s = D.inputSchema();
  assert.deepEqual(Object.keys(s.properties), ['fileId', 'task']);
  assert.deepEqual(s.required, ['fileId', 'task']);
  // GH #76 MODEL 1 DOCX port (owner decision "A", 2026-10-19 -- supersedes the former "decision 4"):
  // verifyStoredContent IS now a docx task field, the SAME wire shape as the xlsx schema.
  assert.deepEqual(Object.keys(s.properties.task.properties), ['id', 'steps', 'verifyStoredContent']);
  assert.deepEqual(s.properties.task.properties.steps.items.oneOf, [D.DOCX_STEP_SCHEMAS.rewrite_paragraph, D.DOCX_STEP_SCHEMAS.insert_paragraph_before, D.DOCX_STEP_SCHEMAS.insert_paragraph_after, D.DOCX_STEP_SCHEMAS.delete_paragraph, D.DOCX_STEP_SCHEMAS.move_paragraph_before, D.DOCX_STEP_SCHEMAS.move_paragraph_after, D.DOCX_STEP_SCHEMAS.apply_list, D.DOCX_STEP_SCHEMAS.set_list_level, D.DOCX_STEP_SCHEMAS.indent_list_item, D.DOCX_STEP_SCHEMAS.outdent_list_item, D.DOCX_STEP_SCHEMAS.remove_list, D.DOCX_STEP_SCHEMAS.replace_table_cell, D.DOCX_STEP_SCHEMAS.format_text, D.DOCX_STEP_SCHEMAS.insert_table_row_before, D.DOCX_STEP_SCHEMAS.insert_table_row_after, D.DOCX_STEP_SCHEMAS.delete_table_row, D.DOCX_STEP_SCHEMAS.insert_table_column_before, D.DOCX_STEP_SCHEMAS.insert_table_column_after, D.DOCX_STEP_SCHEMAS.delete_table_column, D.DOCX_STEP_SCHEMAS.merge_table_cells, D.DOCX_STEP_SCHEMAS.split_table_cell, D.DOCX_STEP_SCHEMAS.format_table, D.DOCX_STEP_SCHEMAS.set_page_setup, D.DOCX_STEP_SCHEMAS.edit_header, D.DOCX_STEP_SCHEMAS.edit_footer, D.DOCX_STEP_SCHEMAS.define_style, D.DOCX_STEP_SCHEMAS.modify_style, D.DOCX_STEP_SCHEMAS.set_numbering, D.DOCX_STEP_SCHEMAS.set_list_type, D.DOCX_STEP_SCHEMAS.add_footnote, D.DOCX_STEP_SCHEMAS.add_endnote, D.DOCX_STEP_SCHEMAS.add_link, D.DOCX_STEP_SCHEMAS.add_bookmark, D.DOCX_STEP_SCHEMAS.resolve_comment, D.DOCX_STEP_SCHEMAS.reply_comment, D.DOCX_STEP_SCHEMAS.add_comment_on_match, D.DOCX_STEP_SCHEMAS.set_image_size, D.DOCX_STEP_SCHEMAS.set_image_wrap], 'exactly the allowlisted intents\' step schemas');
  assert.equal(s.properties.task.properties.verifyStoredContent.type, 'boolean');
  assert.equal(s.properties.task.additionalProperties, false);
  assert.equal(s.additionalProperties, false);
  // RED: an unknown top-level field is refused UNKNOWN_FIELD, not silently accepted
  const r = D.validateDocxWriteRequest({ fileId: FILE, task: { steps: [{ intent: 'rewrite_paragraph' }] }, extra: 1 });
  assert.equal(r.code, 'UNKNOWN_FIELD');
  // verifyStoredContent is now ACCEPTED as a task field (the xlsx schema always accepted it; this one now matches)
  const ok = D.validateDocxWriteRequest({ fileId: FILE, task: { steps: [{ intent: 'rewrite_paragraph', selector: { text: 'x' }, text: 'y' }], verifyStoredContent: true } });
  assert.equal(ok.ok, true, JSON.stringify(ok));
  assert.equal(ok.task.verifyStoredContent, true);
  // RED: the TYPE is still checked -- a non-boolean is MALFORMED_TASK, not silently coerced
  const bad = D.validateDocxWriteRequest({ fileId: FILE, task: { steps: [{ intent: 'rewrite_paragraph' }], verifyStoredContent: 'yes' } });
  assert.equal(bad.code, 'MALFORMED_TASK');
});

test('SECURITY_FIELDS: credential / requester / route names refused at any depth, before the unknown-field check -- unchanged from the xlsx path (checkNamesDeep is shared, unmodified)', () => {
  const r1 = D.validateDocxWriteRequest(req([{ intent: 'rewrite_paragraph', selector: { password: 'x' } }]));
  assert.equal(r1.code, 'CREDENTIAL_FIELD');
  assert.match(r1.path, /steps\[0\]\.selector\.password$/);
  const r2 = D.validateDocxWriteRequest(req([{ intent: 'rewrite_paragraph', onBehalfOf: 'someone-else' }]));
  assert.equal(r2.code, 'REQUESTER_OVERRIDE');
  // RED: a credential-like field in a DOCX-shaped step is caught the SAME way the xlsx test proves for an xlsx step
  const r3 = D.validateDocxWriteRequest(req([{ intent: 'rewrite_paragraph', brokerUrl: 'https://evil.example' }]));
  assert.equal(r3.code, 'ROUTE_OVERRIDE');
  const r4 = D.validateDocxWriteRequest({ fileId: FILE, task: { steps: [{ intent: 'rewrite_paragraph' }] }, __proto__: { token: 'x' } });
  assert.equal(r4.code, 'REFUSED_SHAPE', 'an inherited (non-own) key is not an argument -- the walk reaches the step-schema check next (the step carries no selector/text)');
});

test('DOCX_REFUSED_NOT_ALLOWLISTED: a known DOCX intent is refused "not allowlisted", NOT "cross-format" -- proving the docx-ag already consults DOCX_WRITE_ALLOWLIST', () => {
  for (const intent of [...INTENTS, ...NEW_INTENTS].map(i => i.intent).filter(i => !W.DOCX_WRITE_ALLOWLIST.includes(i)).slice(0, 3)) {
    const r = D.validateDocxWriteRequest(req([{ intent }]));
    assert.equal(r.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${r.code} -- ${r.reason}`);
    assert.equal(r.stage, 'PRE_DISPATCH');
    assert.equal(r.dispatched, false);
    assert.equal(r.wholeTask, true);
  }
  // RED: the SAME intent offered to the xlsx tool classifies differently (REFUSED_CROSS_FORMAT) -- this is
  // the asymmetry this unit's docxIntentRefusal exists to produce: xlsx sees a docx intent as cross-format,
  // docx sees its OWN intent as not-yet-allowlisted. Both refuse; the CODE differs and that difference is the point.
  const xlsxView = W.validateWriteRequest(req([{ intent: INTENTS[0].intent }]));
  assert.equal(xlsxView.code, 'REFUSED_CROSS_FORMAT');
  // RED: an xlsx intent offered to the DOCX tool is REFUSED_CROSS_FORMAT from the other side
  const docxView = D.validateDocxWriteRequest(req([{ intent: 'write_range', sheet: 'S', range: 'A1', values: [['x']] }]));
  assert.equal(docxView.code, 'REFUSED_CROSS_FORMAT', docxView.reason);
  // RED: a truly unknown intent on either side is REFUSED_UNKNOWN_INTENT, not swallowed into either allowlist class
  assert.equal(D.validateDocxWriteRequest(req([{ intent: 'zz_never_registered' }])).code, 'REFUSED_UNKNOWN_INTENT');
  // a routed READ operation offered to the WRITE tool is REFUSED_READ_OPERATION
  assert.equal(D.validateDocxWriteRequest(req([{ intent: 'read_text' }])).code, 'REFUSED_READ_OPERATION');
});

// GH #267 B: in the final union EVERY routed DOCX intent is released, so a still-unreleased one is picked at run time (undefined = none left; the release trimming makes these cases live again).
const notReleasedIntent = () => [...INTENTS, ...NEW_INTENTS].map(i => i.intent).find(i => !W.DOCX_WRITE_ALLOWLIST.includes(i));

test('MIXED_PREFLIGHT: a multi-step task with an earlier bad step refuses the WHOLE task on that step, before any later step is even inspected', () => {
  // a still-unallowlisted routed intent (insert_paragraph_* and format_text are released; add_footnote is not yet released)
  const nr = notReleasedIntent();
  if (nr === undefined) return;
  const steps = [{ intent: nr }, { notAnObject: true }];
  const r = D.validateDocxWriteRequest(req(steps));
  assert.equal(r.code, 'REFUSED_NOT_ALLOWLISTED', 'the FIRST bad step wins -- step 1 is never reached');
  assert.equal(r.stepIndex, 0);
  assert.equal(r.wholeTask, true);
  // RED: swap the order -- now step 0 is malformed and THAT is what refuses, proving the loop is positional, not a scan-for-worst
  const swapped = D.validateDocxWriteRequest(req([{ notAnObject: true }, { intent: nr }]));
  assert.equal(swapped.code, 'MALFORMED_TASK');
  assert.equal(swapped.stepIndex, 0);
});

test('PRE_DISPATCH_REFUSALS: every refusal stops at PRE_DISPATCH with dispatched:false, wholeTask:true, and the execute seam is never called', async () => {
  const cases = [
    ['empty task', req([]), 'EMPTY_TASK'],
    ['unknown intent', req([{ intent: 'zz_probe' }]), 'REFUSED_UNKNOWN_INTENT'],
    ...(notReleasedIntent() ? [['not allowlisted', req([{ intent: notReleasedIntent() }]), 'REFUSED_NOT_ALLOWLISTED']] : []),
    ['cross format', req([{ intent: 'write_range' }]), 'REFUSED_CROSS_FORMAT'],
    ['read operation', req([{ intent: 'read_text' }]), 'REFUSED_READ_OPERATION'],
    ['credential field', req([{ intent: 'rewrite_paragraph', password: 'x' }]), 'CREDENTIAL_FIELD'],
    ['malformed task', { fileId: FILE, task: 'not-an-object' }, 'MALFORMED_TASK'],
    ['invalid fileId', { fileId: 'not-numeric', task: { steps: [{ intent: 'rewrite_paragraph' }] } }, 'INVALID_ARGS'],
  ];
  const seen = new Set();
  for (const [label, args, code] of cases) {
    const h = harness();
    const r = await h.tool.handler(args);
    assert.equal(r.ok, false, label);
    assert.equal(r.code, code, `${label}: ${r.code} -- ${r.reason}`);
    assert.equal(r.stage, 'PRE_DISPATCH', label);
    assert.equal(r.dispatched, false, label);
    assert.equal(r.wholeTask, true, label);
    assert.equal(h.n.executor, 0, label);
    seen.add(code);
  }
  for (const c of ['EMPTY_TASK', 'REFUSED_UNKNOWN_INTENT', ...(notReleasedIntent() ? ['REFUSED_NOT_ALLOWLISTED'] : []), 'REFUSED_CROSS_FORMAT', 'REFUSED_READ_OPERATION', 'CREDENTIAL_FIELD', 'MALFORMED_TASK', 'INVALID_ARGS']) {
    assert.ok(seen.has(c), c);
  }
});

test('DOCX_RESULT_MAPPING: mapExecutionResult projects a docx FinalTaskResult with tool:"docx_execute_task" -- proving the receipt layer needed zero docx-specific change (RETRY / no-op SATISFIED included)', () => {
  const v = { fileId: FILE, task: { id: undefined, steps: [{ intent: 'rewrite_paragraph' }] } };
  const final = (over = {}) => ({ kind: 'FinalTaskResult', ok: true, editOk: true, outcome: 'task-complete', documentOutcome: 'SUCCESS',
    mutationCertainty: 'CERTAIN', authority: 'LIVE_VERIFY', noOp: false, persistenceOutcome: 'SUCCESS', persistenceState: 'FORCESAVE_CONFIRMED',
    saveEvidence: { reason: 'forcesave-confirmed' }, cleanupOutcome: 'SUCCESS', openFailure: null,
    execution: { appliedOperations: [{ operationId: 's0', status: 'APPLIED', verification: 'LIVE_VERIFY' }], operationAccounting: { planned: 1 } }, ...over });
  const applied = W.mapExecutionResult(final(), v, 'docx_execute_task');
  assert.equal(applied.tool, 'docx_execute_task');
  assert.equal(applied.ok, true);
  assert.deepEqual(applied.operations, [{ id: 's0', status: 'APPLIED', verification: 'LIVE_VERIFY', l1: null, l3: null }]);
  // RETRY / no-op SATISFIED: the SAME second call, SATISFIED + noOp -- the format-neutral executor's own
  // idempotency mechanism (core/execution/) produces this status; the receipt layer here only projects it.
  const retry = W.mapExecutionResult(final({ noOp: true, persistenceState: 'NOT_REQUIRED', persistenceOutcome: 'NOT_REQUIRED',
    execution: { appliedOperations: [{ operationId: 's0', status: 'SATISFIED', verification: 'LIVE_VERIFY' }] } }), v, 'docx_execute_task');
  assert.equal(retry.tool, 'docx_execute_task');
  assert.equal(retry.noOp, true);
  assert.equal(retry.operations[0].status, 'SATISFIED');
  // RED: omitting the third argument falls back to the xlsx literal -- proving the xlsx call sites (which
  // never pass a third argument) keep their EXACT, unchanged output.
  assert.equal(W.mapExecutionResult(final(), v).tool, 'xlsx_execute_task');
});

test('SURFACE_EXACTNESS: docx_execute_task is BUILT by server-tools.cjs (GH #219 3) and, in the GH #228 proposed state, SERVED by the exposure gate', () => {
  const { exposedTools } = require('../../src/server/tools.cjs');
  assert.deepEqual(exposedTools(), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task'], 'GH #228 proposed state -- docx_execute_task is exposed next to xlsx_execute_task');
  const { buildServerTools } = require('../../src/server/server-tools.cjs');
  const { registry, tools } = buildServerTools({ config: { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: 'http://127.0.0.1:9/api/euro-broker', instance: 'production', timeoutMs: 1000 } }, configIdentity: 'fixture' });
  // GH #219 3: registry.docx_execute_task exists (built, fully wired); GH #228 (proposed state) also lets `tools`
  // (the exposed subset, filtered by exposedTools()) carry it, exactly as xlsx_execute_task since M11.4b.1.
  assert.equal(typeof registry.docx_execute_task, 'object', 'server-tools.cjs now builds this tool (GH #219 3)');
  assert.equal(typeof registry.docx_execute_task.handler, 'function');
  assert.equal(typeof tools.docx_execute_task, 'object', 'the exposed subset carries it -- exposedTools() names it (GH #228 proposed state)');
  // RED: the module itself remains fully usable standalone, unchanged by this unit's wiring
  assert.equal(typeof D.createDocxWriteTool({}).handler, 'function');
  assert.ok(D.inputSchema().properties.fileId, 'the schema exists and is well-formed even though nothing serves it yet');
});
