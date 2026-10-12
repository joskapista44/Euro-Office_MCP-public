'use strict';

/**
 * The product MCP tool inventory and its exposure gate (owner decision, owner channel 2026-10-08, points 4-5).
 *
 * The inventory is DERIVED from the canonical adapters (the intents they actually route), never copied from the
 * Talk-bridge tool list. Exposure is the EURO-MCP product exposure gate below (owner, owner channel, M10 closure point 12):
 * a capability tool is listed only if it is in the canonical product scope, accepted, behind requester authorization and
 * part of the final regression -- i.e. named in EXPOSED_CAPABILITY_TOOLS AND every release precondition is PASS.
 * There is deliberately no allow-all and no config switch. Exposed: `product_status` (read-only infrastructure, touches
 * no document), since M11.3b.2 the two READ capability tools `xlsx_read` / `docx_read`, and since M11.4b.1 the XLSX WRITE
 * tool `xlsx_execute_task` -- only while the WRITE release gate below holds. `docx_execute_task` and every routed write
 * capability stay WITHHELD.
 */

const docxIntents = require('../adapters/docx/intents.cjs');
const { xlsxAdapter } = require('../adapters/xlsx/adapter.cjs');

// The capability tools released into the product MCP; adding a name here is a release decision, reviewed in the diff.
// No config switch, no allow-all. M11.3b.2 (owner decision, owner channel 2026-10-10): the two READ tools. M11.4b.1
// (owner decision, owner channel): + the XLSX WRITE tool `xlsx_execute_task`; `docx_execute_task` stays WITHHELD
// (static exact-set guard: test/static/tool-exposure-exact-set-guard.test.cjs).
const EXPOSED_CAPABILITY_TOOLS = Object.freeze(['xlsx_read', 'docx_read', 'xlsx_execute_task']);
// The non-capability release preconditions (owner M10 points 10-11).
//   REQUESTER_AUTHORIZATION: Nextcloud/DS/storage is the access authority, reached with the requester's own credential
//     (owner decision 2026-10-09; M11.1b LIVE: cross-credential IMPOSSIBLE, spoof DENY) -> PASS.
//   CREDENTIAL_PATH: the requester-own broker config route, deployed and LIVE accepted (M11.2) -> PASS.
const RELEASE_PRECONDITIONS = Object.freeze({ REQUESTER_AUTHORIZATION: 'PASS', CREDENTIAL_PATH: 'PASS' });

// The WRITE release gate (owner, owner channel point 3): a WRITE tool is exposed only while the accepted milestones
// are CLOSED, the XLSX WRITE allowlist the handler enforces is EXACTLY the accepted six and DOCX_WRITE_ACCEPTED is empty.
// A widened/narrowed allowlist or a DOCX WRITE intent withholds the WRITE tool again (fail-closed), it never widens release.
const WRITE_RELEASE_PRECONDITIONS = Object.freeze({ 'M11.4a.1': 'CLOSED', 'M11.4a.2': 'CLOSED', 'M11.4a.3': 'CLOSED' });
const ACCEPTED_XLSX_WRITE_ALLOWLIST = Object.freeze(['write_range', 'insert_rows', 'delete_rows', 'rename_sheet', 'unmerge_range', 'set_chart_title']);

/** writeReleaseGate(state?) -> { ok, violations }. `state` overrides are for the static NEG controls only. */
function writeReleaseGate({ preconditions = WRITE_RELEASE_PRECONDITIONS, xlsxAllowlist, docxAllowlist } = {}) {
  if (!xlsxAllowlist || !docxAllowlist) {
    const w = require('./write-tools.cjs'); // lazy: write-tools reaches back into this module lazily as well
    xlsxAllowlist = xlsxAllowlist || w.XLSX_WRITE_ALLOWLIST;
    docxAllowlist = docxAllowlist || w.DOCX_WRITE_ALLOWLIST;
  }
  const violations = [];
  for (const [k, v] of Object.entries(preconditions)) if (v !== 'CLOSED') violations.push(`${k} is ${v}, not CLOSED`);
  for (const k of Object.keys(WRITE_RELEASE_PRECONDITIONS)) if (!(k in preconditions)) violations.push(`${k} missing`);
  const x = [...xlsxAllowlist];
  if (x.length !== ACCEPTED_XLSX_WRITE_ALLOWLIST.length || new Set(x).size !== x.length || !ACCEPTED_XLSX_WRITE_ALLOWLIST.every(i => x.includes(i))) {
    violations.push(`XLSX WRITE allowlist is not exactly the accepted six: ${x.join(',')}`);
  }
  if (docxAllowlist.length !== 0) violations.push(`DOCX_WRITE_ACCEPTED is not empty: ${[...docxAllowlist].join(',')}`);
  return { ok: violations.length === 0, violations };
}

/** The product exposure gate state (read-only, for product_status / --identity). */
function exposureGate() {
  const pre = Object.values(RELEASE_PRECONDITIONS).every(v => v === 'PASS');
  return {
    state: pre && EXPOSED_CAPABILITY_TOOLS.length > 0 ? 'CAPABILITIES_RELEASED' : 'CAPABILITIES_WITHHELD',
    releasedCapabilityTools: [...EXPOSED_CAPABILITY_TOOLS],
    releasePreconditions: RELEASE_PRECONDITIONS,
    writeReleasePreconditions: WRITE_RELEASE_PRECONDITIONS,
    writeReleaseGate: writeReleaseGate().ok ? 'PASS' : 'FAIL',
  };
}

const XLSX_CANDIDATES = ['write_range', 'clear_range', 'format_range', 'layout_range', 'merge_range', 'unmerge_range',
  'freeze_panes', 'unfreeze_panes', 'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns', 'copy_range', 'move_range',
  'sort_range', 'set_validation', 'clear_validation', 'add_conditional_format', 'delete_conditional_format', 'filter_range',
  'clear_filter', 'set_defined_name', 'delete_defined_name', 'set_print_area', 'clear_print_area', 'set_print_titles',
  'clear_print_titles', 'create_sheet', 'rename_sheet', 'delete_sheet', 'move_sheet', 'copy_sheet', 'set_chart', 'rename_chart',
  'set_chart_title', 'set_chart_legend', 'set_chart_axis_title', 'set_page_layout', 'set_print_fit', 'set_header_footer',
  'set_page_break', 'set_first_page_number', 'set_print_centering', 'set_print_scale', 'set_paper_size',
  'set_header_footer_flag', 'create_pivot', 'refresh_pivot'];

/** The xlsx intents the adapter actually routes (planning only -- no editor, no dispatch). */
function xlsxRoutedIntents() {
  return XLSX_CANDIDATES.filter(intent => {
    const r = xlsxAdapter.planOperations({ steps: [{ id: 'x', intent }] });
    return !(r && typeof r.detail === 'string' && /not-migrated/.test(r.detail));
  });
}

const WITHHELD = 'WITHHELD: not released (EXPOSED_CAPABILITY_TOOLS names the READ tools and, since M11.4b.1, xlsx_execute_task; routed capabilities are reached only through those tools, docx_execute_task is withheld)';

/** Capability rows (one per routed intent). */
function capabilityInventory() {
  const rows = [];
  const docx = Array.isArray(docxIntents.INTENTS) ? docxIntents.INTENTS : Object.entries(docxIntents.INTENTS).map(([intent, v]) => ({ intent, ...v }));
  for (const r of docx) {
    rows.push({ TOOL: `docx:${r.intent}`, SOURCE_MODULE: `src/adapters/docx/families (${r.family})`, FORMAT: 'DOCX', READ_WRITE: 'WRITE', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url', CURRENT_ACCEPTANCE: r.status, EXPOSE_IN_PRODUCT_MCP: 'NO', REASON: WITHHELD });
  }
  for (const intent of xlsxRoutedIntents()) {
    rows.push({ TOOL: `xlsx:${intent}`, SOURCE_MODULE: 'src/adapters/xlsx/families', FORMAT: 'XLSX', READ_WRITE: 'WRITE', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url', CURRENT_ACCEPTANCE: 'ROUTED_BY_ADAPTER (M6/M7; released only via the final regression + EXPOSED_CAPABILITY_TOOLS)', EXPOSE_IN_PRODUCT_MCP: 'NO', REASON: WITHHELD });
  }
  return rows;
}

/** The MCP tool surface rows: the transport tools a product MCP would carry. */
const TOOL_ROWS = Object.freeze([
  { TOOL: 'product_status', SOURCE_MODULE: 'src/server/product-status.cjs', FORMAT: 'NONE', READ_WRITE: 'READ', AUTH_REQUIRED: 'NONE', CONFIG_REQUIRED: 'schemaVersion, nextcloud.url', CURRENT_ACCEPTANCE: 'ENTRYPOINT_ACCEPTANCE (test/acceptance/product-mcp-entrypoint.test.cjs)', EXPOSE_IN_PRODUCT_MCP: 'YES', REASON: 'read-only product/runtime identity and health; touches no document; no secret in the answer' },
  { TOOL: 'xlsx_read', SOURCE_MODULE: 'src/server/read-tools.cjs (M11.3a handler + schema) -> driver readLive + src/adapters/xlsx/inventory.cjs', FORMAT: 'XLSX', READ_WRITE: 'READ', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url, playwright.path', CURRENT_ACCEPTANCE: 'M11.3a OFFLINE (test/acceptance/m11-3a-read-tools.test.cjs); accepted ops workbook_overview, read_range, read_validations; M11.3b.2 MCP release (offline acceptance test/acceptance/m11-3b2-read-mcp-release.test.cjs)', EXPOSE_IN_PRODUCT_MCP: 'YES', REASON: 'M11.3b.2 READ release: accepted read operations only, pre-dispatch refusal, requester-own broker route, one session per process' },
  { TOOL: 'docx_read', SOURCE_MODULE: 'src/server/read-tools.cjs (M11.3a handler + schema) -> driver readLive + src/adapters/docx/inventory.cjs', FORMAT: 'DOCX', READ_WRITE: 'READ', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url, playwright.path', CURRENT_ACCEPTANCE: 'M11.3a OFFLINE (test/acceptance/m11-3a-read-tools.test.cjs); accepted ops read_text, read_formatting, read_tables, read_headers_footers, read_bookmarks; M11.3b.2 MCP release (offline acceptance test/acceptance/m11-3b2-read-mcp-release.test.cjs)', EXPOSE_IN_PRODUCT_MCP: 'YES', REASON: 'M11.3b.2 READ release: accepted read operations only, pre-dispatch refusal, requester-own broker route, one session per process' },
  { TOOL: 'xlsx_execute_task', SOURCE_MODULE: 'src/server/write-tools.cjs (M11.4a.1 handler + schema, M11.4a.2 canonical executor wiring) -> src/core/execution + src/adapters/xlsx', FORMAT: 'XLSX', READ_WRITE: 'WRITE', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url, playwright.path', CURRENT_ACCEPTANCE: 'M11.4a.1 OFFLINE handler/schema (test/acceptance/m11-4a1-write-handler.test.cjs): allowlist write_range, insert_rows, delete_rows, rename_sheet, unmerge_range, set_chart_title; M11.4a.2 OFFLINE executor wiring (test/acceptance/m11-4a2-executor-wiring.test.cjs); M11.4a.3 LIVE pre-measurement (docs/M11-4A3-LIVE-PREMEASUREMENT.md); M11.4b.1 OFFLINE MCP release (test/acceptance/m11-4b1-write-mcp-release.test.cjs); installed LIVE acceptance: M11.4b.2', EXPOSE_IN_PRODUCT_MCP: 'YES', REASON: 'M11.4b.1 XLSX WRITE release: exactly the six accepted intents, pre-dispatch refusal, requester-own broker route (WRITE), canonical executor, FORCESAVE_CONFIRMED persistence, one session per process; exposed only while the WRITE release gate holds' },
  { TOOL: 'docx_execute_task', SOURCE_MODULE: 'src/core/execution + src/adapters/docx', FORMAT: 'DOCX', READ_WRITE: 'WRITE', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url, playwright.path', CURRENT_ACCEPTANCE: 'per capability (see capability rows)', EXPOSE_IN_PRODUCT_MCP: 'NO', REASON: WITHHELD },
]);

/**
 * The tools the server lists: the read-only infrastructure tool, plus released capability tools when every precondition
 * passes; a WRITE tool additionally needs the WRITE release gate. `writeGate`: a gate-state override for the static NEG
 * controls only (the server always calls this without arguments).
 */
function exposedTools({ writeGate } = {}) {
  const pre = Object.values(RELEASE_PRECONDITIONS).every(v => v === 'PASS');
  const releasable = r => r.READ_WRITE === 'READ' || (r.FORMAT === 'XLSX' && writeReleaseGate(writeGate).ok); // no DOCX WRITE tool (DOCX_WRITE_ACCEPTED=0)
  return TOOL_ROWS.filter(r => r.EXPOSE_IN_PRODUCT_MCP === 'YES' && (r.FORMAT === 'NONE' && r.READ_WRITE === 'READ' || pre && EXPOSED_CAPABILITY_TOOLS.includes(r.TOOL) && releasable(r))).map(r => r.TOOL);
}

module.exports = { EXPOSED_CAPABILITY_TOOLS, RELEASE_PRECONDITIONS, WRITE_RELEASE_PRECONDITIONS, ACCEPTED_XLSX_WRITE_ALLOWLIST, writeReleaseGate, exposureGate, TOOL_ROWS, capabilityInventory, exposedTools, xlsxRoutedIntents };
