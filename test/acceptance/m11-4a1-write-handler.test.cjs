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

const fmt0 = () => require('../../src/adapters/xlsx/families/format.cjs');
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

test('ALLOWLIST: the accepted XLSX intents; DOCX_WRITE_ALLOWLIST = [rewrite_paragraph, insert_paragraph_before, insert_paragraph_after, delete_paragraph, move_paragraph_before, move_paragraph_after, apply_list, set_list_level, set_list_type, indent_list_item, outdent_list_item, remove_list] (GH #228 + #233 + #234 + #235 + #236 + #237); no XLSX-side DOCX handler', () => {
  // X-N01 (owner decision 2026-10-12): + format_range, numberFormat-only (the TEMPORARY_BOUNDARY test below proves the shape)
  // W1a (owner decision 2026-10-13): + the 5 census-ACCEPTED_WRITE intents
  // W1d (owner decision 2026-10-13): + 10 more (W1c LIVE fail-closed evidence); set_chart_legend is the one that
  // stays held back (write-tools.cjs's own XLSX_WRITE_ALLOWLIST comment names why: PRESERVATION_EVIDENCE NONE_FOUND)
  // X-O04/11 (owner decision "B", internal tracker): + set_chart_style, restricted at the schema/CONSTRAINTS
  // level to the two targets the X-O04/10 DS Bridge /11 READ makes verifiable (marker, gridlines) -- the
  // other six internal targets stay unreleased, named.
  // X-O07/3 (internal tracker; X-O07/2 LIVE MEASURED and reviewed): + delete_chart, a NEW intent
  // (not a set_chart_style target) -- the DS Bridge /13 WRITE round-trip makes it verifiable the
  // same way.
  // GH #139 U2 (same #141 roadmap GO as GH #138): + move_chart -- the GH #139 U1 matrix's own C16
  // gap, already adapter-complete and LIVE-measured (X-O06-6); verified by read_drawings's pos fields.
  // GH #139 U3 (same #141 roadmap GO): + insert_shape / insert_image -- the drawings-insert family's
  // v1 scope, already adapter-complete and LIVE-measured (X-O08); verified by read_drawings.
  // GH #137 U2 (same #141 roadmap GO; READ mirror read_runs released in U1): + format_run -- the
  // format-run.cjs family's 5-axis per-run font write, already adapter-complete and LIVE-measured (X-F04).
  // GH #138 U4 (same #141 roadmap GO): + the remaining 8 print.cjs intents -- fit, header/footer text,
  // manual page breaks, first page number, print centering, scale, paper size, header/footer flags. Already
  // adapter-complete and LIVE-measured (print.cjs's W8b/M7-F2 probes); the write path migrated to the DS
  // Bridge WRITE sheet.printSettings.write/v1 (/7+) in M11.5b.4, the SAME migration set_page_layout (U3)
  // already rode -- TEMPORARY_BOUNDARY's 8 B10 entries are removed with this release.
  // GH #140 U4 (same #141 roadmap GO): + refresh_pivot -- the pivot family's SetSource + RefreshTable path (M7 P3), verified by read_pivots (GH #140 U2).
  // GH #140 U3 (same #141 roadmap GO; docs/GH140-PIVOT-CONTRACT.md section 4): + create_pivot -- census
  // ACCEPTED_WRITE, already adapter-complete and LIVE-measured (M7-H); the READ mirror is read_pivots (GH #140 U2).
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST], ['write_range', 'insert_rows', 'delete_rows', 'rename_sheet', 'unmerge_range', 'set_chart_title', 'format_range',
    'merge_range', 'sort_range', 'add_conditional_format', 'set_print_area', 'layout_range',
    'freeze_panes', 'unfreeze_panes', 'set_validation', 'clear_validation', 'delete_conditional_format',
    'clear_filter', 'clear_print_area', 'create_sheet', 'rename_chart', 'set_chart_axis_title', 'clear_range', 'set_chart_style',
    'insert_columns', 'delete_columns', 'delete_sheet', 'move_sheet', 'set_chart', 'set_chart_legend', 'copy_range', 'copy_chart', 'copy_sheet', 'move_range', 'delete_chart', 'filter_range', 'set_defined_name', 'delete_defined_name', 'delete_cells', 'set_print_titles', 'clear_print_titles', 'set_page_layout', 'move_chart', 'insert_shape', 'insert_image', 'format_run',
    'set_print_fit', 'set_header_footer', 'set_page_break', 'set_first_page_number', 'set_print_centering', 'set_print_scale', 'set_paper_size', 'set_header_footer_flag', 'refresh_pivot', 'create_pivot', 'write_array_formula', 'add_comment', 'set_shape_text', 'set_chart_point_data_label', 'set_shape_rotation', 'set_sheet_visibility']);
  assert.equal(Object.isFrozen(W.XLSX_WRITE_ALLOWLIST), true);
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], ['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'apply_list', 'set_list_level', 'indent_list_item', 'outdent_list_item', 'remove_list', 'replace_table_cell', 'format_text', 'insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table', 'set_page_setup', 'edit_header', 'edit_footer', 'define_style', 'modify_style', 'set_numbering', 'set_list_type', 'add_footnote', 'add_endnote', 'add_link', 'add_bookmark', 'resolve_comment', 'reply_comment', 'add_comment_on_match', 'set_image_size', 'set_image_wrap']);
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.throws(() => W.createWriteTool('docx'), /no WRITE handler for docx/);
  const { registry } = buildServerTools({ config: CONFIG, configIdentity: 'fixture' });
  // GH #219 3: docx_execute_task IS built now (server-tools.cjs) -- withheld from `tools` (the
  // exposed subset, test/static/tool-exposure-exact-set-guard.test.cjs's own job), not absent here.
  assert.equal(typeof registry.docx_execute_task.handler, 'function', 'GH #219 3: a docx_execute_task handler now exists in the registry');
  assert.equal(typeof registry.xlsx_execute_task.handler, 'function');
});

test('TEMPORARY_BOUNDARY: the legacy set equals the M11.4a.0 census; no allowlisted intent depends on a boundary', () => {
  const census = JSON.parse(fs.readFileSync(path.join(__dirname, '../../docs/evidence/m11-4a0/census.json'), 'utf8'));
  const fromCensus = Object.values(census.legacy_dependents).flat().map(x => x.split(':')[1]).sort();
  assert.deepEqual(Object.keys(W.TEMPORARY_BOUNDARY).sort(), fromCensus);
  // X-N01: format_range is released with a KEY-RESTRICTED shape -- the intent stays in the census' boundary set (its other
  // keys), the released key (numberFormat) lowers to a PUBLIC setter only, with no DS Bridge write prelude
  // GitHub #81: + border (edge modes) -- also a PUBLIC setter (SetBorders), no DS Bridge write prelude either
  // GitHub #88: + fill '#RRGGBB' -- the PUBLIC SetFillColor, no DS Bridge write prelude
  // GitHub #89: + font ({ color?, bold?, italic? } -- PUBLIC SetFontColor/SetBold/SetItalic), same shape
  // GitHub #93: + align ({ horizontal?, vertical? } -- PUBLIC SetAlignHorizontal/SetAlignVertical), same shape
  assert.deepEqual([...W.FORMAT_RANGE_RELEASED_KEYS], ['numberFormat', 'border', 'fill', 'font', 'align', 'wrap']);
  const pf = fmt0().plan({ intent: 'format_range', sheet: 'Budget', range: 'A1:B2', format: { fill: '#FFC000' } }, 0);
  const lf = fmt0().lower(pf.operation, {});
  assert.notEqual(lf.meta && lf.meta.bridgeAlignExtWrite, true);
  assert.match(lf.body, /SetFillColor\(Api\.CreateColorFromRGB\(255, 192, 0\)\)/);
  assert.doesNotMatch(lf.body, /__eo|Bridge/);
  const pn = fmt0().plan({ intent: 'format_range', sheet: 'Budget', range: 'A1:C3', format: { font: { color: '#3A7BC8', bold: true, italic: false } } }, 0);
  const ln = fmt0().lower(pn.operation, {});
  assert.notEqual(ln.meta && ln.meta.bridgeAlignExtWrite, true);
  assert.match(ln.body, /SetFontColor\(Api\.CreateColorFromRGB\(58, 123, 200\)\)/);
  assert.match(ln.body, /SetBold\(true\)/);
  assert.match(ln.body, /SetItalic\(false\)/);
  assert.doesNotMatch(ln.body, /__eo|Bridge/);
  const pa = fmt0().plan({ intent: 'format_range', sheet: 'Budget', range: 'F80:F84', format: { align: { horizontal: 'center', vertical: 'top' } } }, 0);
  const la = fmt0().lower(pa.operation, {});
  assert.notEqual(la.meta && la.meta.bridgeAlignExtWrite, true);
  assert.match(la.body, /SetAlignHorizontal\("center"\)/);
  assert.match(la.body, /SetAlignVertical\("top"\)/);
  assert.doesNotMatch(la.body, /__eo|Bridge/);
  const fmt = require('../../src/adapters/xlsx/families/format.cjs');
  const planned = fmt.plan({ intent: 'format_range', sheet: 'Budget', range: 'B2', format: { numberFormat: '0%' } }, 0);
  const lowered = fmt.lower(planned.operation, {});
  assert.notEqual(lowered.meta && lowered.meta.bridgeAlignExtWrite, true);
  assert.match(lowered.body, /SetNumberFormat\("0%"\)/);
  assert.doesNotMatch(lowered.body, /__eo|Bridge/);
  for (const mode of ['outline', 'inside', 'all']) {
    const pb = fmt.plan({ intent: 'format_range', sheet: 'Budget', range: 'A1:C3', format: { border: { mode, style: 'thin', color: '#000000' } } }, 0);
    const lb = fmt.lower(pb.operation, {});
    assert.notEqual(lb.meta && lb.meta.bridgeAlignExtWrite, true, mode);
    assert.match(lb.body, /SetBorders\(/, mode);
    assert.doesNotMatch(lb.body, /__eo|Bridge/, mode);
  }
  // X-O04/11: set_chart_style is released through a DIFFERENT mechanism (the X-O04/10 DS Bridge
  // /11 READ makes marker/gridlines verifiable -- classifyStyle reads-and-compares, see
  // charts.cjs) -- it was never part of the M11.4a.0/W1a/W1d census this loop checks against, so
  // it carries no census row at all; excluded the same way format_range is (a different, but
  // equally out-of-census, release path).
  // X-O07/3: delete_chart is the SAME out-of-census release path (the X-O07/2 DS Bridge /13
  // WRITE round-trip, classifyDelete reads-and-compares) -- excluded for the same reason.
  // X-O07/4b: copy_chart carries no census row either -- same out-of-census release path.
  // L-18/1: delete_cells is a NEW intent the M11.4a.0/W1a/W1d census never enumerated (it did not
  // exist yet) -- same out-of-census release path, acceptance evidence is the LIVE probe (5 rounds)
  // docs/L18-DELETE-CELLS-SHIFT-PLAN.md cites instead.
  // GH #301 D4: write_array_formula is a NEW intent the census never enumerated -- same out-of-census release path,
  // acceptance evidence is the LIVE design probe + release harness (docs/evidence/gh301-d4-*).
  // GH #138 U4: the 8 print.cjs intents DO carry a census row, but it classifies them
  // TEMPORARY_BOUNDARY/B10 -- STALE since M11.5b.4 migrated the family's write path off that boundary onto
  // the DS Bridge WRITE sheet.printSettings.write/v1 (/7+), the SAME migration set_page_layout (U3) already
  // rode. Excluded here for that reason, not because the row is missing.
  const PRINT_U4 = ['set_print_fit', 'set_header_footer', 'set_page_break', 'set_first_page_number', 'set_print_centering', 'set_print_scale', 'set_paper_size', 'set_header_footer_flag'];
  // GH #300 K2: add_comment is a NEW intent the census never enumerated either (same out-of-census release path; evidence:
  // the K1 / K2 LIVE design probes + the release harness, docs/evidence/gh300-*).
  for (const intent of W.XLSX_WRITE_ALLOWLIST.filter(i => !['format_range', 'set_chart_style', 'delete_chart', 'copy_chart', 'delete_cells', 'write_array_formula', 'add_comment', 'set_shape_text', 'set_chart_point_data_label', 'set_shape_rotation', 'set_sheet_visibility', ...PRINT_U4].includes(i))) {
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
  // GH #76 MODEL 1 (owner decision 2026-10-18): opt-in verifyStoredContent, default false
  assert.deepEqual(Object.keys(s.properties.task.properties), ['id', 'steps', 'verifyStoredContent']);
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

test('W1a HATOKOR +1: write_range formulas -- a =SUM(...) matrix is accepted, planned by the canonical planner', () => {
  const step = { intent: 'write_range', sheet: 'Budget', range: 'C1:C2', formulas: [['=SUM(A1:A2)'], ['=1+1']] };
  const v = W.validateWriteRequest(req([step]));
  assert.equal(v.ok, true, JSON.stringify(v));
  assert.equal(v.plannedOperations, 1);
  assert.equal(xlsxAdapter.planOperations({ steps: [step] }).ok, true);
});

test('GH79a: write_range replaceFormulas -- accepted by the schema, type-checked boolean', () => {
  const step = { intent: 'write_range', sheet: 'Budget', range: 'C1', formulas: [['=SUM(A1:A2)']], replaceFormulas: true };
  const v = W.validateWriteRequest(req([step]));
  assert.equal(v.ok, true, JSON.stringify(v));
  assert.equal(xlsxAdapter.planOperations({ steps: [step] }).ok, true);
  const notBoolean = W.validateWriteRequest(req([{ ...step, replaceFormulas: 'yes' }]));
  assert.equal(notBoolean.ok, false);
  assert.equal(notBoolean.code, 'REFUSED_SHAPE');
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
  // REFUSED_NOT_ALLOWLISTED is not in this list since GH #140 U3: no routed XLSX intent is outside the allowlist any more
  // (xlsxRoutedIntents() minus XLSX_WRITE_ALLOWLIST is empty), so the class has no live example -- its branch is covered by
  // the dedicated test below with a probe intent added to the routed set.
  for (const c of ['REFUSED_UNKNOWN_INTENT', 'REFUSED_TEMPORARY_BOUNDARY', 'REFUSED_CROSS_FORMAT', 'REFUSED_READ_OPERATION',
    'REFUSED_SHAPE', 'REFUSED_COMPOSITION', 'MALFORMED_TASK', 'EMPTY_TASK', 'LIMIT', 'CREDENTIAL_FIELD', 'REQUESTER_OVERRIDE', 'ROUTE_OVERRIDE', 'UNKNOWN_FIELD']) {
    assert.ok(seen.has(c), `refusal class covered: ${c}`);
  }
});

test('REFUSED_NOT_ALLOWLISTED: a routed XLSX intent outside the allowlist is refused before dispatch (probe intent; no live one is left)', async () => {
  const tools = require('../../src/server/tools.cjs');
  assert.deepEqual(tools.xlsxRoutedIntents().filter(x => !W.XLSX_WRITE_ALLOWLIST.includes(x)), [], 'the live routed-not-allowlisted set is empty (this test exists because of that)');
  const original = tools.xlsxRoutedIntents;
  tools.xlsxRoutedIntents = () => [...original(), 'zz_routed_probe'];
  try {
    const h = harness();
    const r = await withCountedFetch(h.n, () => h.tool.handler(req([{ intent: 'zz_routed_probe', sheet: 'Budget' }])));
    assert.equal(r.ok, false);
    assert.equal(r.code, 'REFUSED_NOT_ALLOWLISTED', `${r.code} -- ${r.reason}`);
    assert.equal(r.stage, 'PRE_DISPATCH');
    assert.equal(r.dispatched, false);
    assert.deepEqual(ZERO(h.n), NONE);
  } finally { tools.xlsxRoutedIntents = original; }
  const neg = await harness().tool.handler(req([{ intent: 'zz_routed_probe', sheet: 'Budget' }]));
  assert.equal(neg.code, 'REFUSED_UNKNOWN_INTENT', 'without the probe in the routed set the same intent is simply unknown');
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
  const unknown = W.mapExecutionResult(final({ ok: false, editOk: false, mutationCertainty: 'UNCERTAIN', persistenceState: 'UNCERTAIN_MAY_PERSIST',
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
  // alice-1 widened the set by three EXECUTED-stage names (test/acceptance/mona1-failure-reason.test.cjs)
  // GH #76 MODEL 1 (owner decision 2026-10-18) added STORED_CONTENT_MISMATCH (decision 3)
  assert.deepEqual(Object.keys(W.WRITE_ERRORS), ['ACCESS_DENIED', 'TARGET_READ_ONLY', 'STORAGE_WRITE_REJECTED', 'OPERATION_FAILED', 'REFUSED_PRE_MUTATION', 'EXECUTION_FAILED', 'STORED_CONTENT_MISMATCH']);
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

test('M11.4b.1 surface: xlsx_execute_task listed and its tools/call reaches this handler; docx_execute_task (GH #228 proposed state) is listed too, with its OWN handler', async () => {
  const h = harness();
  const { tools } = buildServerTools({ config: CONFIG, configIdentity: 'fixture', writeTool: { execute: async () => { h.n.executor += 1; return { ok: true }; } } });
  assert.deepEqual(exposedTools(), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  // GH #228 (PROPOSED): the DOCX WRITE tool has BOTH release switches on (the capability list AND the tool row)
  assert.deepEqual([...EXPOSED_CAPABILITY_TOOLS], ['xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  assert.equal(TOOL_ROWS.find(r => r.TOOL === 'docx_execute_task').EXPOSE_IN_PRODUCT_MCP, 'YES');
  assert.equal(TOOL_ROWS.find(r => r.TOOL === 'xlsx_execute_task').EXPOSE_IN_PRODUCT_MCP, 'YES');
  assert.deepEqual(Object.keys(tools), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  const s = mcp(tools);
  await s.send('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 't', version: '0' } });
  const list = await s.send('tools/list', {});
  assert.deepEqual(list.result.tools.map(t => t.name), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  const ok = await s.send('tools/call', { name: 'xlsx_execute_task', arguments: req([VALID.write_range]) });
  assert.equal(ok.error, undefined, JSON.stringify(ok));
  // W1d: freeze_panes is allowlisted now (was the refusal example here); GH #138 U2 released clear_print_titles too -- create_pivot
  // (routed, not released) is the refusal example now.
  const refused = await s.send('tools/call', { name: 'xlsx_execute_task', arguments: req([{ intent: 'create_pivot', sheet: 'S' }]) });
  assert.equal(refused.result.structuredContent.stage, 'PRE_DISPATCH');
  // an XLSX step offered to the DOCX tool is refused by the DOCX tool's own pre-dispatch classification, never executed
  const r = await s.send('tools/call', { name: 'docx_execute_task', arguments: req([VALID.write_range]) });
  assert.equal(r.error, undefined);
  assert.equal(r.result.structuredContent.code, 'REFUSED_CROSS_FORMAT');
  assert.equal(r.result.structuredContent.dispatched, false);
  await s.end();
  assert.equal(h.n.executor, 1, 'exactly the valid call reached the execution; the refusal and the DOCX tool did not');
});
