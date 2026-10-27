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
 * tool `xlsx_execute_task` -- only while the WRITE release gate below holds. GH #228 (RELEASED, owner GO 2026-10-19):
 * + `docx_execute_task`, whose release gate is PER FORMAT (docs/GH228-DOCX-WRITE-RELEASE-GATE-PROPOSAL.md) and whose allowlist is
 * exactly ACCEPTED_DOCX_WRITE_ALLOWLIST; every other routed write capability stays WITHHELD.
 */

const docxIntents = require('../adapters/docx/intents.cjs');
const { NEW_INTENTS: docxNewIntents } = require('../adapters/docx/capabilities.cjs');
const { xlsxAdapter } = require('../adapters/xlsx/adapter.cjs');

// The capability tools released into the product MCP; adding a name here is a release decision, reviewed in the diff.
// No config switch, no allow-all. M11.3b.2 (owner decision, owner channel 2026-10-10): the two READ tools. M11.4b.1
// (owner decision, owner channel): + the XLSX WRITE tool `xlsx_execute_task`; `docx_execute_task` stays WITHHELD
// (static exact-set guard: test/static/tool-exposure-exact-set-guard.test.cjs).
// GH #228 (RELEASED, owner GO 2026-10-19): + `docx_execute_task`, served only while the DOCX half of the WRITE release gate holds.
const EXPOSED_CAPABILITY_TOOLS = Object.freeze(['xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
// The non-capability release preconditions (owner M10 points 10-11).
//   REQUESTER_AUTHORIZATION: Nextcloud/DS/storage is the access authority, reached with the requester's own credential
//     (owner decision 2026-10-09; M11.1b LIVE: cross-credential IMPOSSIBLE, spoof DENY) -> PASS.
//   CREDENTIAL_PATH: the requester-own broker config route, deployed and LIVE accepted (M11.2) -> PASS.
const RELEASE_PRECONDITIONS = Object.freeze({ REQUESTER_AUTHORIZATION: 'PASS', CREDENTIAL_PATH: 'PASS' });

// The WRITE release gate (owner, owner channel point 3): a WRITE tool is exposed only while the accepted milestones
// are CLOSED and the allowlist the handler enforces for ITS OWN format is EXACTLY the accepted set of that format
// (ACCEPTED_XLSX_WRITE_ALLOWLIST below; ACCEPTED_DOCX_WRITE_ALLOWLIST, GH #228). The gate is PER FORMAT: a widened/narrowed
// allowlist withholds only the WRITE tool of the format it belongs to (fail-closed), never the other format's tool, and it
// never widens release on its own -- both accepted lists only ever widen by an explicit owner decision.
const WRITE_RELEASE_PRECONDITIONS = Object.freeze({ 'M11.4a.1': 'CLOSED', 'M11.4a.2': 'CLOSED', 'M11.4a.3': 'CLOSED' });
// X-N01 (owner decision 2026-10-12: number format is a CAPABILITY routed through the canonical executor
// onto the product surface): + format_range, released with format.numberFormat ONLY (write-tools.cjs
// FORMAT_RANGE_RELEASED_KEYS); every other format key stays refused pre-dispatch behind B05-WRITE
// W1a (owner decision 2026-10-13): + the 5 intents docs/evidence/m11-4a0/census.json classifies
// ACCEPTED_WRITE (all six evidence axes, FAIL_CLOSED included, cited from the M10 regression) -- merge_range,
// sort_range, add_conditional_format, set_print_area, layout_range.
// W1d (owner decision 2026-10-13 -- same widening rule, W1c LIVE fail-closed evidence for the 11 census-PARTIAL
// intents): + 10 more the regenerated census now classifies ACCEPTED_WRITE -- freeze_panes, unfreeze_panes,
// set_validation, clear_validation, delete_conditional_format, clear_filter, clear_print_area, create_sheet,
// rename_chart, set_chart_axis_title. set_chart_legend is the ONE that stays out (write-tools.cjs's own
// XLSX_WRITE_ALLOWLIST comment names why: PRESERVATION_EVIDENCE still NONE_FOUND).
const ACCEPTED_XLSX_WRITE_ALLOWLIST = Object.freeze(['write_range', 'insert_rows', 'delete_rows', 'rename_sheet', 'unmerge_range', 'set_chart_title', 'format_range',
  'merge_range', 'sort_range', 'add_conditional_format', 'set_print_area', 'layout_range',
  'freeze_panes', 'unfreeze_panes', 'set_validation', 'clear_validation', 'delete_conditional_format',
  'clear_filter', 'clear_print_area', 'create_sheet', 'rename_chart', 'set_chart_axis_title',
  // GH #106 (owner decision 2026-10-14): + clear_range, census ACCEPTED_WRITE with the W1e LIVE fail-closed evidence
  'clear_range',
  // X-O04/11 (owner decision "B", internal tracker): + set_chart_style -- NOT a census-ACCEPTED_WRITE widening
  // (write-tools.cjs's own ACCEPTED_XLSX_WRITE_ALLOWLIST comment explains why it carries no census row); the
  // acceptance evidence is the X-O04/10 DS Bridge /11 READ round-trip instead, restricted at the schema/
  // CONSTRAINTS level to the two targets that makes verifiable (marker, gridlines).
  'set_chart_style',
  // L-04 (owner decision 2026-10-14): + insert_columns / delete_columns, the column twins of the released insert_rows / delete_rows
  'insert_columns', 'delete_columns',
  // L-05 (owner decision 2026-10-15): + delete_sheet, census ACCEPTED_WRITE with the L-05 LIVE measurement as its fail-closed evidence
  'delete_sheet',
  // L-05/2 (owner decision 2026-10-15): + move_sheet, census ACCEPTED_WRITE with the L-05 LIVE measurement's named refusals as its fail-closed evidence
  'move_sheet',
  // L-07 (owner decision L-07): + set_chart, set_chart_legend -- census ACCEPTED_WRITE with the L-07 LIVE measurement (docs/evidence/l07-charts-fail-closed-20261015)
  'set_chart', 'set_chart_legend',
  // L-06 (owner decision L-06): + copy_range -- census ACCEPTED_WRITE; the L-06 LIVE measurement (docs/evidence/l06-range-copy-fail-closed-20261015) covers the new write shapes
  'copy_range',
  // X-O07/4b (internal tracker): + copy_chart -- same class as set_chart_style above (no census
  // row; the acceptance evidence is the X-O07/4b offline+LIVE round, restricted to charts.cjs's
  // own v1 scope: type/series/title/legend/axis-titles/position, never the style targets).
  'copy_chart',
  // L-05/3 (owner decision 2026-10-15): + copy_sheet -- the B03 boundary lifted; census ACCEPTED_WRITE (DS Bridge WRITE worksheet.copy/v1)
  'copy_sheet',
  // L-06/b (owner decision, zero 2026-10-15): + move_range -- census ACCEPTED_WRITE, the SAME
  // write shapes L-06's copy_range measurement already covers; the caller NEVER supplies `expect`
  // (range.cjs's classify() builds it server-side, from its own fresh read, directly before
  // dispatch -- no new executor-level mechanism, no second open/read/close session).
  'move_range',
  // X-O07/3 (internal tracker; X-O07/2 LIVE MEASURED and reviewed): + delete_chart -- same
  // non-census release convention as set_chart_style above; the acceptance evidence is the
  // X-O07/2 DS Bridge /13 WRITE round-trip, LIVE measured end-to-end (session + reopen).
  'delete_chart',
  // L-11/2 (owner decision "A", internal tracker): + filter_range -- SMALL workbooks only; the large-book refusal (owner rule, internal tracker) is unchanged.
  'filter_range',
  // GH #138 U1 (owner GO, the answer to the #141 roadmap question: "Mindegyik mehet, minden fogyjon el"): + set_defined_name / delete_defined_name -- the names
  // family's measured shapes; the READ mirror read_defined_names (GH #134 P2) verifies them.
  'set_defined_name', 'delete_defined_name',
  // L-18/1 (internal tracker): + delete_cells -- a block-shift (up|left) of a sub-range, released for the LIVE-measured
  // scope (docs/L18-DELETE-CELLS-SHIFT-PLAN.md): SCOPED to the block's band, a bare same-sheet cell reference into/behind
  // the block resolved correctly (#REF!/re-pointed), a merge fully inside the block deleted with it. Refused named,
  // PRE_DISPATCH: any merge the block only partially covers, any name/rule/chart/pivot/AutoFilter reaching the sheet, any
  // other formula shape -- none of those has a measurement of its own yet. Retry safety: the SAME anchor mechanism
  // insert/delete_rows/columns already use, scoped to the block's band.
  'delete_cells',
  // GH #138 U2 (same roadmap GO): + set_print_titles / clear_print_titles -- the names family's planPrint(); the READ mirror is
  // read_defined_names's own `titles` projection on the Print_Titles entry.
  'set_print_titles', 'clear_print_titles',
  // GH #138 U3 (same roadmap GO): + set_page_layout -- public setters; verified by read_page_setup's page half.
  'set_page_layout',
  // GH #139 U2 (same #141 roadmap GO): + move_chart -- the charts family's position intent (already
  // adapter-complete and LIVE-measured, X-O06-6); the READ mirror is read_drawings's own pos fields.
  'move_chart',
  // GH #139 U3 (same #141 roadmap GO): + insert_shape / insert_image -- the drawings-insert family's v1
  // scope (already adapter-complete and LIVE-measured, X-O08); the READ mirror is read_drawings.
  'insert_shape', 'insert_image',
  // GH #137 U2 (same #141 roadmap GO; READ mirror read_runs released in U1): + format_run -- the
  // format-run.cjs family's 5-axis per-run font write (already adapter-complete and LIVE-measured, X-F04).
  'format_run',
  // GH #138 U4 (same #141 roadmap GO as GH #138/#139): + the remaining 8 print.cjs intents (fit,
  // header/footer text, manual page breaks, first page number, print centering, scale, paper size,
  // header/footer flags) -- same non-census release convention as set_chart_style/copy_chart/delete_chart
  // above; the acceptance evidence is this release's own offline+LIVE round (already adapter-complete and
  // LIVE-measured, print.cjs's W8b/M7-F2 probes). The READ mirror read_page_setup's `print` field was
  // already released; this entry only wires the WRITE side onto it.
  'set_print_fit', 'set_header_footer', 'set_page_break', 'set_first_page_number', 'set_print_centering',
  'set_print_scale', 'set_paper_size', 'set_header_footer_flag',
  // GH #140 U4 (owner GO, the #141 roadmap): + refresh_pivot -- the pivot family's measured SetSource + RefreshTable path; the READ mirror read_pivots (GH #140 U2) verifies it.
  'refresh_pivot',
  // GH #140 U3 (same #141 roadmap GO; docs/GH140-PIVOT-CONTRACT.md section 4): + create_pivot -- census
  // ACCEPTED_WRITE (docs/evidence/m11-4a0/census.json, DIRECT_INTERNAL_DEPENDENCY none), already adapter-complete
  // and LIVE-measured (M7-H, docs/evidence/M7-H-ACCEPTANCE.md PASS 6/6), only not yet wired on this product
  // surface. The READ mirror is read_pivots (GH #140 U2, already released).
  'create_pivot',
  // GH #301 D4 (standing EURO-MCP GO): + write_array_formula -- SetFormulaArray over an EMPTY area whose ring holds no
  // array member (array-formula.cjs; LIVE design probe docs/evidence/gh301-d4-array-formula-probe-20261025.json)
  'write_array_formula',
  // GH #300 K2 (owner GO 2026-10-26, author rule A): + add_comment -- the author is the authenticated session user
  // (comments.cjs; LIVE design probe docs/evidence/gh300-k2-comment-identity-probe-20261026.json)
  'add_comment',
  // GH #302 G4: + set_shape_text (and insert_shape's optional text), verified by read_drawings' text
  'set_shape_text',
  // GH #302 G5: + set_chart_point_data_label, verified by read_charts' effectiveLabels
  'set_chart_point_data_label',
  // GH #302 G7: + set_shape_rotation, verified by read_drawings' rotation
  'set_shape_rotation',
  // GH #303 R1: + set_sheet_visibility, verified by read_sheet_state's visibility
  'set_sheet_visibility']);

// GH #228 (owner GO 2026-10-19): the DOCX WRITE allowlist the handler enforces is EXACTLY this set. Before the owner's GO this
// is the empty list (state "BEFORE" in docs/GH228-DOCX-WRITE-RELEASE-GATE-PROPOSAL.md); the released state names rewrite_paragraph only.
const ACCEPTED_DOCX_WRITE_ALLOWLIST = Object.freeze(['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'apply_list', 'set_list_level', 'indent_list_item', 'outdent_list_item', 'remove_list', 'replace_table_cell', 'format_text', 'insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table', 'set_page_setup', 'edit_header', 'edit_footer', 'define_style', 'modify_style', 'set_numbering', 'set_list_type', 'add_footnote', 'add_endnote', 'add_link', 'add_bookmark', 'resolve_comment', 'reply_comment', 'add_comment_on_match', 'set_image_size', 'set_image_wrap']); // GH #229/#230/#231: + format_text (docs/GH229..GH231 decision docs); GH #233: + the two insert intents; GH #234: + delete_paragraph; GH #235: + the two move intents; GH #236: + apply_list, set_list_level (set_list_type WITHHELD: on a shared numbering it ends DISPATCHED_OUTCOME_UNKNOWN and the DS persists the unsaved edit -- LIVE g236v1 ROUND6 + guncprobe, 2026-10-19); GH #237: + indent_list_item, outdent_list_item, remove_list; GH #249: + replace_table_cell // RELEASE 3 TRIM (2026-10-19) withheld these eight pending their own LIVE PASS; all eight are since RE-ADDED, each on its own LIVE PASS: define_style (GH #264, g3tri gh264), modify_style (GH #265, g3tri gh265), set_numbering (GH #248, g3tri gh248), set_list_type (GH #236b, g3tri gh236b), add_footnote + add_endnote (GH #267 B, g267v2 incl. R6 through the released handler), add_link (GH #266, release 3, g266v2), add_bookmark (GH #266, release 4, g266v3 ROUND8/8B) -- the array above is current; nothing in it is withheld -- release 4: + resolve_comment, reply_comment, add_comment_on_match (GH #269c/#269d/#269e; handler LIVE PASS g269cv2 / g269dv1 / g269ev1)

/** writeReleaseGate(state?) -> { ok, violations }. `state` overrides are for the static NEG controls only.
 * `format` ('xlsx' | 'docx') narrows the allowlist check to that format's own list; absent = both (the product_status view). */
function writeReleaseGate({ preconditions = WRITE_RELEASE_PRECONDITIONS, xlsxAllowlist, docxAllowlist, format } = {}) {
  if (!xlsxAllowlist || !docxAllowlist) {
    const w = require('./write-tools.cjs'); // lazy: write-tools reaches back into this module lazily as well
    xlsxAllowlist = xlsxAllowlist || w.XLSX_WRITE_ALLOWLIST;
    docxAllowlist = docxAllowlist || w.DOCX_WRITE_ALLOWLIST;
  }
  const violations = [];
  for (const [k, v] of Object.entries(preconditions)) if (v !== 'CLOSED') violations.push(`${k} is ${v}, not CLOSED`);
  for (const k of Object.keys(WRITE_RELEASE_PRECONDITIONS)) if (!(k in preconditions)) violations.push(`${k} missing`);
  const exact = (list, accepted) => list.length === accepted.length && new Set(list).size === list.length && accepted.every(i => list.includes(i));
  const x = [...xlsxAllowlist];
  if ((format === undefined || format === 'xlsx') && !exact(x, ACCEPTED_XLSX_WRITE_ALLOWLIST)) {
    violations.push(`XLSX WRITE allowlist is not exactly the accepted six: ${x.join(',')}`);
  }
  const dx = [...docxAllowlist];
  if ((format === undefined || format === 'docx') && !exact(dx, ACCEPTED_DOCX_WRITE_ALLOWLIST)) {
    violations.push(`DOCX WRITE allowlist is not exactly the accepted ${ACCEPTED_DOCX_WRITE_ALLOWLIST.length}: ${dx.join(',')}`);
  }
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
  'freeze_panes', 'unfreeze_panes', 'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns', 'delete_cells', 'copy_range', 'move_range',
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

const WITHHELD = 'WITHHELD: not released (EXPOSED_CAPABILITY_TOOLS names the READ tools, since M11.4b.1 xlsx_execute_task and, since GH #228, docx_execute_task; routed capabilities are reached only through those tools, and only while they are in that format\'s accepted WRITE allowlist)';

const RELEASED_DOCX = 'RELEASED (GH #228, owner GO 2026-10-19): reached through docx_execute_task, the accepted DOCX allowlist only, behind the per-format WRITE release gate';

/** Capability rows (one per routed intent). GH #262: `docx` now covers BOTH the 28 M4-migrated
 * intents (docxIntents.INTENTS) AND the post-M4 NEW_INTENTS (docx/capabilities.cjs) -- discovered
 * gap, not a design choice: this function iterated INTENTS only, so a NEW_INTENTS capability
 * (set_page_setup, the first of the five ever added to ACCEPTED_DOCX_WRITE_ALLOWLIST) was invisible
 * here even once released -- EXPOSE_IN_PRODUCT_MCP would silently read 'NO' for an actually-exposed
 * tool. docxRoutedIntents() (docx-write-tools.cjs) already unions both lists for routing; this
 * function did not, for the inventory. */
function capabilityInventory() {
  const rows = [];
  const docxMigrated = Array.isArray(docxIntents.INTENTS) ? docxIntents.INTENTS : Object.entries(docxIntents.INTENTS).map(([intent, v]) => ({ intent, ...v }));
  const docx = [...docxMigrated, ...docxNewIntents];
  for (const r of docx) {
    const released = ACCEPTED_DOCX_WRITE_ALLOWLIST.includes(r.intent); // GH #228 (RELEASED): reached only through docx_execute_task
    rows.push({ TOOL: `docx:${r.intent}`, SOURCE_MODULE: `src/adapters/docx/families (${r.family})`, FORMAT: 'DOCX', READ_WRITE: 'WRITE', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url', CURRENT_ACCEPTANCE: r.status, EXPOSE_IN_PRODUCT_MCP: released ? 'YES' : 'NO', REASON: released ? RELEASED_DOCX : WITHHELD });
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
  { TOOL: 'docx_read', SOURCE_MODULE: 'src/server/read-tools.cjs (M11.3a handler + schema) -> driver readLive + src/adapters/docx/inventory.cjs', FORMAT: 'DOCX', READ_WRITE: 'READ', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url, playwright.path', CURRENT_ACCEPTANCE: 'M11.3a OFFLINE (test/acceptance/m11-3a-read-tools.test.cjs); accepted ops read_text, read_formatting, read_tables, read_headers_footers, read_page_setup, read_bookmarks; M11.3b.2 MCP release (offline acceptance test/acceptance/m11-3b2-read-mcp-release.test.cjs)', EXPOSE_IN_PRODUCT_MCP: 'YES', REASON: 'M11.3b.2 READ release: accepted read operations only, pre-dispatch refusal, requester-own broker route, one session per process' },
  { TOOL: 'xlsx_execute_task', SOURCE_MODULE: 'src/server/write-tools.cjs (M11.4a.1 handler + schema, M11.4a.2 canonical executor wiring) -> src/core/execution + src/adapters/xlsx', FORMAT: 'XLSX', READ_WRITE: 'WRITE', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url, playwright.path', CURRENT_ACCEPTANCE: 'M11.4a.1 OFFLINE handler/schema (test/acceptance/m11-4a1-write-handler.test.cjs): allowlist write_range, insert_rows, delete_rows, rename_sheet, unmerge_range, set_chart_title; M11.4a.2 OFFLINE executor wiring (test/acceptance/m11-4a2-executor-wiring.test.cjs); M11.4a.3 LIVE pre-measurement (docs/M11-4A3-LIVE-PREMEASUREMENT.md); M11.4b.1 OFFLINE MCP release (test/acceptance/m11-4b1-write-mcp-release.test.cjs); installed LIVE acceptance: M11.4b.2', EXPOSE_IN_PRODUCT_MCP: 'YES', REASON: 'M11.4b.1 XLSX WRITE release: exactly the six accepted intents, pre-dispatch refusal, requester-own broker route (WRITE), canonical executor, FORCESAVE_CONFIRMED persistence, one session per process; exposed only while the WRITE release gate holds' },
  { TOOL: 'docx_execute_task', SOURCE_MODULE: 'src/core/execution + src/adapters/docx', FORMAT: 'DOCX', READ_WRITE: 'WRITE', AUTH_REQUIRED: 'REQUESTER_OWN_NC_CREDENTIAL (via broker)', CONFIG_REQUIRED: 'nextcloud.url, broker.url, playwright.path', CURRENT_ACCEPTANCE: 'GH #219 OFFLINE handler/contract (test/acceptance/m11-4b-docx-write-handler.test.cjs); GH #221 stored-content gate (test/acceptance/gh221-docx-stored-content-gate.test.cjs); GH #228 OFFLINE release (test/acceptance/gh228-docx-write-release.test.cjs); GH #229 OFFLINE release (test/acceptance/gh229-format-text-write-release.test.cjs); LIVE acceptance: test/live/docx-rewrite-paragraph-product.live.cjs, test/live/gh229-format-text-product.live.cjs', EXPOSE_IN_PRODUCT_MCP: 'YES', REASON: 'GH #228 DOCX WRITE release (owner GO 2026-10-19): exactly the accepted DOCX allowlist (rewrite_paragraph), the per-format WRITE release gate, pre-dispatch refusal of every other intent, requester-own broker route' },
]);

/**
 * The tools the server lists: the read-only infrastructure tool, plus released capability tools when every precondition
 * passes; a WRITE tool additionally needs the WRITE release gate. `writeGate`: a gate-state override for the static NEG
 * controls only (the server always calls this without arguments).
 */
function exposedTools({ writeGate } = {}) {
  const pre = Object.values(RELEASE_PRECONDITIONS).every(v => v === 'PASS');
  // GH #228: the WRITE gate is per format -- an xlsx tool never waits on the docx allowlist, nor the other way round
  const releasable = r => r.READ_WRITE === 'READ' || (r.FORMAT === 'XLSX' && writeReleaseGate({ ...writeGate, format: 'xlsx' }).ok) || (r.FORMAT === 'DOCX' && writeReleaseGate({ ...writeGate, format: 'docx' }).ok);
  return TOOL_ROWS.filter(r => r.EXPOSE_IN_PRODUCT_MCP === 'YES' && (r.FORMAT === 'NONE' && r.READ_WRITE === 'READ' || pre && EXPOSED_CAPABILITY_TOOLS.includes(r.TOOL) && releasable(r))).map(r => r.TOOL);
}

module.exports = { EXPOSED_CAPABILITY_TOOLS, RELEASE_PRECONDITIONS, WRITE_RELEASE_PRECONDITIONS, ACCEPTED_XLSX_WRITE_ALLOWLIST, ACCEPTED_DOCX_WRITE_ALLOWLIST, writeReleaseGate, exposureGate, TOOL_ROWS, capabilityInventory, exposedTools, xlsxRoutedIntents };
