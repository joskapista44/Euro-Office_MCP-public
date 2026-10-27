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
 *        in WRITE mode -> withSession(xlsxAdapter, createExecutor(task, batch)) -> mapExecutionResult)
 *
 * The task is the canonical executor task `{ id, steps: [{ intent, ... }] }`: this module validates and passes it on,
 * it never translates it into a second operation language. Any refused step refuses the WHOLE task before anything
 * starts: no broker call, no browser, no driver, no executor, no write.
 *
 * Exposure: tools.cjs EXPOSED_CAPABILITY_TOOLS does not name this tool -> it is neither listed nor callable
 * (static exact-set guard). There is no DOCX handler: DOCX_WRITE_ACCEPTED is empty and docx_execute_task stays withheld.
 */

const { xlsxAdapter, liveScope } = require('../adapters/xlsx/adapter.cjs');
const { parseRange, colNum, cellsOf, addr, parseCell } = require('../adapters/xlsx/a1.cjs');
const { bareNameProblem, sheetNameProblem } = require('../adapters/xlsx/sheet-name.cjs');
const { validatePlan } = require('../core/execution/plan.cjs');
const docxIntents = require('../adapters/docx/intents.cjs');
const { OPERATIONS: READ_OPERATIONS, CREDENTIAL_FIELDS, REQUESTER_FIELDS, productDriver, validateReadRequest, readOnce, FORMATS } = require('./read-tools.cjs');
const { cellDiffs, productCells } = require('../adapters/xlsx/stored-content.cjs');
const { withSession } = require('../core/session/lifecycle.cjs');
const { sessionKey } = require('./session-registry.cjs');
const { decideReuse } = require('./session-reuse-rule.cjs');
const { createExecutor } = require('../core/execution/executor.cjs');
const { createBrowserRuntime } = require('./browser-runtime.cjs');
const { createSessionGate } = require('./session-gate.cjs');

// The first WRITE release allowlist -- EXACTLY these six (owner 2026-10-10). Widening it is an owner decision.
// W1a (owner decision 2026-10-13 -- widen with LIVE-PASSed intents, per-intent acceptance, deploy separately
// GO'd): + the 5 intents the M11.4a.0 census (docs/evidence/m11-4a0/census.json) classifies
// ACCEPTED_WRITE (all six evidence axes, incl. FAIL_CLOSED, cited from the M10 regression) -- merge_range,
// sort_range, add_conditional_format, set_print_area, layout_range.
// W1d (owner decision 2026-10-13 -- same widening rule, W1c LIVE fail-closed evidence for the 11 census-PARTIAL
// intents, docs/evidence/w1c-fail-closed-20261013/): + 10 more the regenerated census now classifies
// ACCEPTED_WRITE -- freeze_panes, unfreeze_panes, set_validation, clear_validation, delete_conditional_format,
// clear_filter, clear_print_area, create_sheet, rename_chart, set_chart_axis_title. set_chart_legend is the ONLY
// one of the original 11 that STAYS OUT: its FAIL_CLOSED_EVIDENCE is now current, but PRESERVATION_EVIDENCE is
// still NONE_FOUND (the xo04 harness never cites L3/preservationVerification) -- named, not silently done.
// GH #106 (owner decision 2026-10-14): + clear_range (CONTENT only -- number format, fill and bold
// kept, M6 W1b), the census' last missing axis (FAIL_CLOSED_EVIDENCE) supplied by the W1e LIVE run
// (docs/evidence/w1e-clear-range-fail-closed-20261014/). It is the one released way to empty a cell (a merge over a
// non-blank member is refused as data loss, and write_range takes no empty string).
// X-O04/11 (owner decision "B", internal tracker): + set_chart_style -- but NOT the whole intent.
// charts.cjs's set_chart_style carries EIGHT internal targets; only TWO are released here (CONSTRAINTS.
// set_chart_style refuses the other six, named): `marker` (seriesIndex, fill?/outline?, scatter+lineNormal
// series only) and `gridlines` (axis: 'vertical'|'horizontal', outline required, no fill). The other six
// (plotArea/title/legend/chart/series/dataPoint) stay WRITE_ONLY and internal -- this release is scoped to
// what the X-O04/10 DS Bridge /11 READ makes verifiable (classifyStyle reads-and-compares, never
// semanticsNotProven, for these two), not to writing style in general.
// X-O05/3 (internal tracker; reviewed and accepted): + two more set_chart_style targets, same
// release scope (DS Bridge /12 makes them verifiable, same as marker/gridlines on /11) --
// `gapOverlap` (gapWidth 0..500 and/or overlap -100..100, bar charts only) and `seriesOrder`
// (order: the WHOLE series-order permutation in one call, never a single {seriesIndex, order}
// pair -- the product allows only one chart-style write per task, so a per-series shape could
// never implement a two-series swap).
// X-O07/4b (internal tracker): + copy_chart -- charts.cjs's own "recreation" implementation
// (type/series/title/legend/axis-titles read from the source, recreated via the SAME AddChart+
// SetName+SetTitle sequence set_chart's own lower() already uses). The six write-only style
// targets and the four Bridge-verifiable style targets (marker/gridlines/gapOverlap/
// seriesOrder) are NEVER copied in v1 -- always named in the receipt's semanticsNotProven, never
// silently dropped (charts.cjs's own COPY_CHART_ALWAYS_OMITTED). GH #302 G6: the data labels ARE copied (chart level +
// every POINT label, charts.cjs copyLabelPlan); what cannot be is named per copy. Cross-sheet copy is attempted
// but UNMEASURED on this DS engine (charts.cjs's own classifyCopy/lower header comment).
// L-04 (owner decision 2026-10-14, inventory row L-04): + insert_columns / delete_columns -- the structure family (structure.cjs) already
// routes all four row / column intents; insert_rows / delete_rows were released first, the column pair follows the SAME anchor protocol
// (anchor { row, value }: the value of that row's cell in the column that is at `at` NOW for an insert, at `at + count` for a delete).
// Released ONLY for the shapes the LIVE measurement on the prod DS found writable (docs/L04-COLUMNS-RELEASE-20261015.md): a column or a
// run of columns, formulas on the sheet and on other sheets re-pointed, a merge grown / clipped, a styled cell moved with its column.
// The family's own FAIL_CLOSED answers pass through NAMED and are part of the release, not defects: a delete that a formula reaches into
// (xlsx-structure-delete-referenced-by-formulas), a delete of a merge's first column (xlsx-structure-merge-top-left-deleted), a sheet
// carrying a chart (xlsx-observation-truncated) and a cut at the sheet's far edge (xlsx-observation-truncated).
// L-05 (owner decision 2026-10-15, inventory row L-05): + delete_sheet -- census ACCEPTED_WRITE (sheets.cjs). Released for the shapes the LIVE
// measurement on the prod DS found writable (an empty sheet, a sheet with data, a sheet with a chart; the retry is SATISFIED with 0 dispatch; an absent
// sheet is SATISFIED too). The family's own refusal passes through NAMED: a sheet a formula elsewhere names (xlsx-delete-sheet-referenced-by-formulas).
// L-05/2: + move_sheet -- census ACCEPTED_WRITE once the LIVE measurement named its refusals (a sheet relative to itself, an absent reference sheet; docs/evidence/
// l05-sheets-fail-closed-20261015). copy_sheet is NOT part of this release (the owner-accepted B03 boundary).
// L-07 (owner decision L-07): + set_chart (create: bar, lineNormal, area, pie, doughnut, radar, stock; orientation columns / rows) and set_chart_legend (bottom, top, left,
// right, none) -- released for the shapes the LIVE measurement on the prod DS found writable (docs/evidence/l07-charts-fail-closed-20261015); a chart of that name with
// ANOTHER identity (xlsx-chart-exists-with-another-identity), a missing chart (xlsx-chart-not-found), a closed type (scatter) and a one-column range pass through refused, named.
// The style targets (set_chart_style) are NOT part of this release.
// L-05/3 (owner decision 2026-10-15): + copy_sheet -- the B03 TEMPORARY_RC_BOUNDARY is lifted. The write path is the DS Bridge WRITE worksheet.copy/v1
// (the B03 direct internals were removed, M11.5c.4/5), served on the runtimes /8 .. /12 (XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS: the handler is byte-identical on each, pinned by
// b03-write-handler-identity.test.cjs). Released for the shapes the LIVE measurement found writable (a sheet with values + formulas copied to a new name; a retry SATISFIED / 0).
// The family's refusals pass through NAMED: a target name that exists and differs (xlsx-copy-target-exists-and-differs), a source hosting an image / shape, a pivot, a quoted
// name, a runtime the Bridge write is not proven on (xlsx-copy-bridge-write-runtime-not-proven).
// L-06 (owner decision L-06): + copy_range -- released for the shapes the LIVE measurement on the prod DS found writable (docs/evidence/l06-range-copy-fail-closed-20261015): values,
// formulas (translated by the planner, written explicitly after the editor's Copy: the SAVED copy of a formula carried the source's literal text otherwise), formats on the SAME
// sheet; ANOTHER sheet gets the typed values and the translated formulas WITHOUT the editor's Copy (a delayed Copy state overwrote the formulas at the next save), so the cell
// FORMAT does not travel there -- named in the receipt (semanticsNotProven 'cross-sheet-copy-format-not-carried'), never silent. The refusals pass through named: an overlap on
// the same sheet, an occupied target, a merged cell, a formula that leaves the sheet, a source or target block a chart / dependent reaches.
// L-06/b (owner decision, zero 2026-10-15): + move_range -- the caller NEVER supplies `expect`; classify() builds it itself, from
// its OWN fresh read, directly before dispatch (the SAME session the write uses -- range.cjs's own classify()/moveExpectation).
// A blank source (any state that looks already-moved, including a genuine retry) therefore has nothing to build a fresh
// expectation FROM, and is refused named (xlsx-move-already-applied-not-provable), 0/0, BEFORE even checking the target --
// v1 never attempts to verify a retry's target content, and never dispatches a second Cut.
// X-O07/3 (internal tracker; X-O07/2 LIVE MEASURED and reviewed): + delete_chart -- a NEW intent
// (not a set_chart_style target), the same release scope convention: the DS Bridge /13 WRITE
// (chart.drawing.delete.write/v1) makes it verifiable the same way the bridge-backed style
// targets are. Idempotent on retry (an already-absent name is SATISFIED); an ambiguous name
// FAIL_CLOSED, never a guess (X-O07/1 LIVE MEASURED: the handler itself refuses the same way).
// X-O05/5-C (internal tracker; X-O05/5-B round 7b+8d+8k+8m+8q LIVE MEASURED, reviewed and
// accepted): + two more set_chart_style targets -- `markerSymbol` (seriesIndex, symbol? one of
// the 10 ECMA-376 ST_MarkerStyle names this build measured / size? 2..72, scatter+lineNormal
// only -- the SAME scope the already-released `marker` fill/outline target uses) and `lineStyle`
// (seriesIndex, dash? one of the 11 ECMA-376 ST_PresetLineDashVal names / smooth? boolean,
// bar+lineNormal only -- R8q's own LIVE-measured scope). UNLIKE marker(fill/outline)/gridlines/
// gapOverlap/seriesOrder, BOTH are WRITE-ONLY (no DS Bridge READ exists for either field) -- the
// SAME six-original-target discipline: never SATISFIED on retry, always named in the receipt's
// semanticsNotProven (charts.cjs's own styleNotProven), served ONLY on DS /15 (B24, /15-ONLY --
// UNLIKE every other Bxx-WRITE family here, no earlier runtime carries either capability).
const XLSX_WRITE_ALLOWLIST = Object.freeze(['write_range', 'insert_rows', 'delete_rows', 'rename_sheet', 'unmerge_range', 'set_chart_title', 'format_range',
  'merge_range', 'sort_range', 'add_conditional_format', 'set_print_area', 'layout_range',
  'freeze_panes', 'unfreeze_panes', 'set_validation', 'clear_validation', 'delete_conditional_format',
  'clear_filter', 'clear_print_area', 'create_sheet', 'rename_chart', 'set_chart_axis_title', 'clear_range', 'set_chart_style',
  'insert_columns', 'delete_columns',
  'delete_sheet', 'move_sheet',
  'set_chart', 'set_chart_legend',
  'copy_range', 'copy_chart',
  'copy_sheet', 'move_range',
  'delete_chart',
  // L-11/2 (owner decision "A", internal tracker): + filter_range -- for SMALL workbooks only. Released for the shape the
  // family measured (B-D06: ONE field, ONE criteria set: value | two values | two numeric comparisons) with its own X-L05 gate (a manually hidden
  // data row and no active filter -> refused before the write, xlsx-filter-would-unhide-manual-rows). The LARGE-book refusal is UNCHANGED (the owner rule,
  // internal tracker): filter_range is not in the adapter's TIERED_INTENTS and has no bounded variant, so above the read bound it stays refused
  // (IMPLEMENTATION_LIMIT:family-not-migrated-to-the-tiered-observation) -- this entry does not loosen that.
  'filter_range',
  // GH #138 U1 (owner GO, the answer to the #141 roadmap question: "Mindegyik mehet, minden fogyjon el"): + set_defined_name / delete_defined_name --
  // the names family's measured shapes (M6 wave 6, M7 R16): an ABSOLUTE refersTo, workbook level or `scope: '<sheet>'`; a re-point of an existing
  // name, a delete of a name in use, the reserved print names and a name already in another scope stay refused (names.cjs). Both are TIERED_INTENTS,
  // so the large-book path is the tiered observation, not a refusal. The READ mirror is read_defined_names (GH #134 P2).
  'set_defined_name', 'delete_defined_name',
  // L-18/1 (internal tracker): + delete_cells -- a block-shift ('up'|'left') of a sub-range, released for the LIVE-measured
  // scope (the probe's 5 rounds, docs/L18-DELETE-CELLS-SHIFT-PLAN.md): SCOPED to the block's own band, a reference INTO the
  // block correctly #REF!s, a shift-path reference correctly re-points (bare same-sheet cell references only -- any other
  // formula shape, any name/rule/chart/pivot/AutoFilter reaching the sheet, and a merge the block only PARTIALLY covers are
  // refused named, PRE_DISPATCH, blanket -- none of those has a measurement of its own yet, unlike `structure`'s per-cut
  // allowances). Retry safety reuses `structure`'s OWN LIVE-proven anchor mechanism, scoped to the block's band -- never a
  // second shift on an unprovable retry. Large workbooks: not in TIERED_INTENTS, same generic large-book refusal as
  // filter_range (not a guard this family writes itself).
  'delete_cells',
  // GH #138 U2 (same roadmap GO): + set_print_titles / clear_print_titles -- the names family's planPrint() (M7 F Q1/Q2/Q4/Q5, M7 X-P03): a
  // sheet-level Print_Titles name, at least one of rows ('n:m') / columns ('A:B') for the set; the reserved-name and another-scope rules
  // (names.cjs) apply unchanged. UNLIKE set_defined_name/delete_defined_name, these are NOT in TIERED_INTENTS -- the large-book path stays the
  // NAMED refusal (IMPLEMENTATION_LIMIT:family-not-migrated-to-the-tiered-observation), this entry does not loosen that. The READ mirror is
  // read_defined_names's own `titles` projection on the Print_Titles entry (GH #134 P2, defined-names-read.cjs).
  'set_print_titles', 'clear_print_titles',
  // GH #138 U3 (same roadmap GO): + set_page_layout -- the page family's PUBLIC setters (orientation, print gridlines / headings, margins in
  // millimetres, cap 200 mm); the READ mirror is read_page_setup's page half (L-10b). The rest of the print family (fit / scale / paper /
  // header-footer / breaks / first page / centering) stays out: its write path is the direct-internal boundary B10.
  'set_page_layout',
  // GH #139 U2 (same #141 roadmap GO as GH #138; the GH #139 U1 property matrix's own C16 gap, docs/
  // GH139-CHART-DRAWING-SYMMETRY-MATRIX.md): + move_chart -- the charts family's position intent
  // (X-O06-6, already adapter-complete and LIVE-measured: docs/evidence/x-o06-move-chart-live-rebased-
  // PASS-20261012.json), only not yet wired on this product surface. Cell-aligned only (at: { col, row },
  // offset always 0 at this boundary -- charts.cjs's own atProblem()); the READ mirror is read_drawings's
  // own pos fields on the chart kind (GH #139 U1 matrix row C16: chartReads carries NO position, only
  // drawings.observe() does). Like every other chart intent here, NOT in TIERED_INTENTS -- the large-book
  // path stays the NAMED refusal (IMPLEMENTATION_LIMIT:family-not-migrated-to-the-tiered-observation).
  'move_chart',
  // GH #139 U3 (same #141 roadmap GO): + insert_shape / insert_image -- the drawings-insert family's v1 scope
  // (X-O08, drawings-insert.cjs), already adapter-complete and LIVE-measured: docs/evidence/
  // xlsx-x-o08-drawings-insert-live-rebased-PASS-20261012.json. insert_shape: geometry 'rect' only, fill
  // '#RRGGBB' required, an OPTIONAL stroke { width, color } (v2, X-O08-7 LIVE r3 -- WIDTH is the measured
  // READABLE axis, COLOUR is WRITE_ONLY, named in the receipt via the family's own COLOUR_NOT_PROVEN(_WITH_STROKE)
  // -- the SAME notProven vocabulary format.cjs's X-N01b scope already uses). insert_image: PNG data URI only
  // (a http(s):// URL is refused at plan time, never fetched -- outside the measured/approved egress path); the
  // `source` field carries a schema-layer maxLength (1,400,000 chars, a safety margin over the 1,398,126-char
  // exact cap for a 1 MiB PNG with the 'data:image/png;base64,' prefix -- an MCP message is not unbounded, and
  // no LIVE round measured anything near that size). Cell-aligned only (at: { col, row }, offset always 0 at
  // this boundary); size is accepted and written but NEVER an exact-match condition (the family's own ±1 EMU
  // drift measurement, X-O08-3/4). Not in TIERED_INTENTS, not SOLE_STEP -- the drawing claims its name via
  // selector.claims, the SAME composition mechanism charts.cjs already uses, not a cell footprint.
  'insert_shape', 'insert_image',
  // GH #137 U2 (same #141 roadmap GO as #138/#139; "productize BOTH sides together" -- the READ mirror,
  // read_runs, released first in U1): + format_run -- the format-run.cjs family's own 5-axis per-run font
  // write inside one cell (X-F04, already adapter-complete and LIVE-measured: the family's own L1 run-map +
  // L3 control, RUN_MAP_CHAR_CEILING 256 chars TEMPORARY, pending a LIVE timing measurement -- raising it is
  // NOT part of this release). Scope is EXACTLY the 5 measured axes (bold, italic, name, size, color);
  // X-F05's whole-cell effects (underline, strikeout, subscript, superscript) are format_range's own,
  // unrelated capability and stay out. Already in TIERED_INTENTS (adapter.cjs, pre-existing -- this release
  // does not add it there); not SOLE_STEP (a single-cell write, the same class as write_range on one cell).
  'format_run',
  // GH #138 U4 (same #141 roadmap GO as GH #138/#139; completes the print family -- U3's set_page_layout entry
  // explicitly left the rest of print.cjs out): + the remaining 8 print.cjs intents (fit, header/footer text,
  // manual page breaks, first page number, print centering, scale, paper size, header/footer flags). All 8 are
  // already adapter-complete and LIVE-measured (print.cjs's own header comment cites the W8b/M7-F2 probes); the
  // write path is the B10-WRITE DS Bridge sheet.printSettings.write/v1 (/7+, XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS).
  // The READ mirror is read_page_setup's `print` field (fit/scale/paper/header-footer codes/manual breaks/first
  // page/centering, L-10b LIVE) -- already released, this entry only wires the WRITE side onto it. Each schema
  // mirrors the family's own plan() shape; every cross-field correlation this minimal schema cannot state (the
  // axis<->at type for set_page_break, exactly-one-of for set_print_centering, the paper-size table, the
  // slot/flag enums) is re-checked by the family's own plan() and refused named, before dispatch.
  'set_print_fit', 'set_header_footer', 'set_page_break', 'set_first_page_number', 'set_print_centering',
  'set_print_scale', 'set_paper_size', 'set_header_footer_flag',
  // GH #140 U4 (owner GO, the #141 roadmap): + refresh_pivot -- the pivot family's measured SetSource + RefreshTable path (M7 P3): the caller NAMES the new source
  // (pivot / sourceSheet / sourceRange), the old source is never inferred. The family's refusals pass through NAMED (a missing / ambiguous pivot, a header row that is not
  // the pivot's field list, a grown table extent that is not empty). The READ mirror is read_pivots (GH #140 U2); create_pivot is a separate release unit.
  'refresh_pivot',
  // GH #140 U3 (same #141 roadmap GO as GH #137/#138/#139; docs/GH140-PIVOT-CONTRACT.md section 4): + create_pivot --
  // the pivots family's measured shape (M7-H, already adapter-complete and LIVE-measured: docs/evidence/
  // M7-H-ACCEPTANCE.md PASS 6/6, census row create_pivot PRODUCT_RELEASE_CLASS=ACCEPTED_WRITE), only not yet
  // wired on this product surface. FIELDS (pivots.cjs FIELDS, minus intent/id): name/sheet/range/rows/columns/
  // sum/pivotSheet/at. The READ mirror is read_pivots (GH #140 U2, already released).
  'create_pivot',
  // GH #301 D4 (standing EURO-MCP GO 2026-10-15): + write_array_formula -- the public SetFormulaArray over an EMPTY area
  // (2..100 cells), array-formula.cjs. The LIVE design probe (docs/evidence/gh301-d4-array-formula-probe-20261025.json)
  // measured SetFormulaArray silently overwriting values and silently building OVERLAPPING arrays, so the area must be
  // empty and no cell of it or of its one-cell ring may already be an array member (refused named, PRE_DISPATCH); L1
  // proves the extent the same way. The only step of its task; not in TIERED_INTENTS (the generic large-book refusal).
  'write_array_formula',
  // GH #300 K2 (owner GO 2026-10-26, author rule A): + add_comment -- one comment on one cell (comments.cjs). The author
  // is the AUTHENTICATED session user (the signed editor config's user: uid + display name), added by the driver, never a
  // step field; no reliable identity -> not dispatched. A commented cell is refused (a second AddComment is a silent
  // no-op, K1/K2 LIVE probes); L1 is read_comments' own read. The only step of its task.
  'add_comment',
  // GH #302 G4 (owner GO 2026-10-26 block 2): + set_shape_text, and insert_shape's optional `text` -- the public
  // GetContent() paragraph write (shape-text.cjs; LIVE design probe docs/evidence/gh302-g4-shape-text-probe-20261026.json),
  // verified by read_drawings' text (drawing.text.read/v1)
  'set_shape_text',
  // GH #302 G5 (owner GO 2026-10-26 block 2): + set_chart_point_data_label -- the public SetShowPointDataLabel
  // (chart-point-labels.cjs; LIVE design probe docs/evidence/gh302-g5-point-data-label-probe-20261026.json), verified
  // by read_charts' effectiveLabels (chart.dataLabels.effective.read/v1)
  'set_chart_point_data_label',
  // GH #302 G7 (owner GO 2026-10-26 block 2): + set_shape_rotation -- the public SetRotation (shape-rotation.cjs; LIVE
  // design probe docs/evidence/gh302-g7-shape-rotation-probe-20261026.json), verified by read_drawings' rotation within
  // the one measured 1/60000-degree truncation unit
  'set_shape_rotation',
  // GH #303 R1 (owner GO 2026-10-26 block 2): + set_sheet_visibility -- the public SetVisible (sheet-visibility.cjs; LIVE
  // design probe docs/evidence/gh303-r1-sheet-visibility-probe-20261026.json), verified by read_sheet_state's visibility.
  // The engine hides the LAST visible sheet too (measured), so that refusal is ours (read + dispatch); the only step of
  // its task (two hides each judged on the same pre-task read could leave no visible sheet).
  'set_sheet_visibility']);
// X-N01 (owner decision: number format is a CAPABILITY, value and number format are separate semantics):
// format_range is released with its PUBLIC numberFormat key ONLY (SetNumberFormat / GetNumberFormat, W4a); every other
// format key stays behind the temporary direct-internal boundary B05-WRITE and is refused before dispatch, named.
// GitHub #81 (owner request via the owner channel, NAD A80:F87): + border, modes 'cells' / 'outline' / 'inside' /
// 'all' with the four originally measured styles (W4c) -- PUBLIC SetBorders writes, the read is the W4c-gated cell
// reader (format.cjs). 'remove' / the diagonals stay out (not requested); a border step carries no other
// format key (CONSTRAINTS.format_range: exactly one of numberFormat / border).
// GitHub #88 (owner: the cell background color blocks the work, release it first): + fill '#RRGGBB' -- the PUBLIC
// SetFillColor write, read back on the public fill reader (format.cjs, X-F08 measured). fill 'none' (owner request:
// cell colour REMOVAL must also work) -- the ClearFormats + explicit-restore composite (format-fill-clear.cjs,
// X-F08, m7-s15-probe4 LIVE-accepted) -- is released too: same key, same schema pattern, same one-format-key-per-step
// rule. Three refusal classes stay as format-fill-clear.cjs's own module comment already documents: an unobservable
// kept axis, a diagonal border on the target, and a target whose border side faces a readable neighbour (that third
// class is UNMEASURED on this product surface -- whether the restore's SetBorders() disturbs the shared edge is a
// separate, follow-up measurement, not assumed here either way).
// GitHub #89 (owner request via the owner channel, paired with #88): + font, a nested object { color?, bold?,
// italic? }, at least one key -- PUBLIC SetFontColor / SetBold / SetItalic (X-F01/X-F02/X-F03, all CLOSED_PROVEN:
// bold/italic/color are measured on every cell shape the family supports, including a formula cell, no font axis
// is refused any more). name / size: released by L-02/1 (see FORMAT_FONT); font.bold / font.italic accept true or
// false (false targets the SAME already-measured internal cell reader the canonical planner already consumes for
// the true-case drift check, X-F01/S07 -- a new consumer, not a new dependency).
// GitHub #93 (owner request via the owner channel, a NAD F80:F84 centering case): + align, a nested object
// { horizontal?, vertical? } -- PUBLIC SetAlignHorizontal / SetAlignVertical, read through the SAME range.
// appearance.read/v1 bridge capability border already uses (B04 cutover). horizontal: left/center/right/justify;
// vertical: bottom/center/distributed/justify/top (format-maps.cjs H/V, every key the canonical planner's own
// alignExpectation already accepts). rotation / readingOrder / indent / shrink: released by L-03/2 (see FORMAT_ALIGN).
// L-01 (owner GO 2026-10-14, inventory row L-01): + wrap, a boolean -- the PUBLIC SetWrap write, read back on the public GetWrapText reader
// (format.cjs wrap axis, W4a; the family already planned, wrote and verified it, only this product schema held it back).
// true / false are both released; a non-boolean is a schema refusal (REFUSED_SHAPE); a wrap step carries no other format key.
const FORMAT_RANGE_RELEASED_KEYS = Object.freeze(['numberFormat', 'border', 'fill', 'font', 'align', 'wrap']);
// GH #228 (RELEASED, owner GO 2026-10-19 -- docs/GH228-DOCX-WRITE-RELEASE-GATE-PROPOSAL.md): the DOCX WRITE allowlist is exactly
// rewrite_paragraph. It is mirrored by tools.cjs's ACCEPTED_DOCX_WRITE_ALLOWLIST, and the per-format release gate withholds
// docx_execute_task (never xlsx_execute_task) if the two ever differ. BEFORE the owner's GO this list is `[]`.
const DOCX_WRITE_ALLOWLIST = Object.freeze(['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'apply_list', 'set_list_level', 'indent_list_item', 'outdent_list_item', 'remove_list', 'replace_table_cell', 'format_text', 'insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table', 'set_page_setup', 'edit_header', 'edit_footer', 'define_style', 'modify_style', 'set_numbering', 'set_list_type', 'add_footnote', 'add_endnote', 'add_link', 'add_bookmark', 'resolve_comment', 'reply_comment', 'add_comment_on_match', 'set_image_size', 'set_image_wrap']); // GH #229/#230/#231: + format_text; GH #233: + the two insert intents; GH #234: + delete_paragraph; GH #235: + the two move intents; GH #236: + apply_list, set_list_level (set_list_type WITHHELD: on a shared numbering it ends DISPATCHED_OUTCOME_UNKNOWN and the DS persists the unsaved edit -- LIVE g236v1 ROUND6 + guncprobe, 2026-10-19); GH #237: + indent_list_item, outdent_list_item, remove_list; GH #249: + replace_table_cell // RELEASE 3 TRIM (2026-10-19) withheld these eight pending their own LIVE PASS; all eight are since RE-ADDED, each on its own LIVE PASS: define_style (GH #264, g3tri gh264), modify_style (GH #265, g3tri gh265), set_numbering (GH #248, g3tri gh248), set_list_type (GH #236b, g3tri gh236b), add_footnote + add_endnote (GH #267 B, g267v2 incl. R6 through the released handler), add_link (GH #266, release 3, g266v2), add_bookmark (GH #266, release 4, g266v3 ROUND8/8B) -- the array above is current; nothing in it is withheld -- release 4: + resolve_comment, reply_comment, add_comment_on_match (GH #269c/#269d/#269e; handler LIVE PASS g269cv2 / g269dv1 / g269ev1)

// The temporary direct-internal boundaries (M10 legacy set, M11.5 migrates them): refused by name, never routed here.
// GH #138 U4 (2026-10-17): the 8 print.cjs B10 entries REMOVED -- M11.5b.4 already migrated the family's write
// path to the DS Bridge WRITE sheet.printSettings.write/v1 (/7+, print.cjs's own header comment), the SAME
// migration that let set_page_layout (GH #138 U3) drop out of this object before it. The census row these 8
// still carry (docs/evidence/m11-4a0/census.json, TEMPORARY_BOUNDARY/B10) predates that migration and is stale
// for the SAME reason set_page_layout's was -- see test/acceptance/gh138-u4-print-family-release.test.cjs and
// the TEMPORARY_BOUNDARY test's exclusion list in m11-4a1-write-handler.test.cjs.
const TEMPORARY_BOUNDARY = Object.freeze({
  format_range: 'B05-WRITE',
  remove_list: 'B18', outdent_list_item: 'B18', edit_header: 'B21', edit_footer: 'B21',
});

const MAX_STEPS = 50; // the M11.4a.1 release bound (the core plan ceiling is 1000)
const MAX_CELLS = 1000; // = the write_range / unmerge_range planner ceiling (cells.cjs / merge.cjs)
const MAX_TEXT = 32767; // = the write_range planner text ceiling (cells.cjs)

// ---- the step schemas (JSON Schema subset; the SAME objects validate and are published as inputSchema) ----

const ID = Object.freeze({ type: 'string', pattern: '^[A-Za-z0-9_.-]{1,64}$' });
// a sheet named only to ADDRESS cells (write_range / unmerge_range): any valid worksheet name
const SHEET_ADDRESS = Object.freeze({ type: 'string', pattern: '^[^\\u0000-\\u001f\\[\\]:*?/\\\\]{1,31}$' });
// a sheet that a name-only / rename intent names (structure shift, rename, freeze, chart): L-21/2. The schema stays as wide as
// SHEET_ADDRESS on purpose -- the released class (the MEASURED one: letters incl. accents, digits, underscore, single spaces)
// and the NAMED refusal of an apostrophe name are one shared rule, sheetNameProblem (adapters/xlsx/sheet-name.cjs), applied in
// CONSTRAINTS and, again, by the family planners; a narrower pattern here would only hide that reason behind a generic one.
const SHEET_NAME = SHEET_ADDRESS;
// GH #138 U1: a defined-name identifier (names.cjs NAME); cell-reference-shaped names are refused by the family's plan()
// GH #140 U4: a pivot name and the bare (needs-no-quoting) source sheet name, exactly the pivot-refresh family's own NAME / SHEET patterns (families/pivots.cjs)
const PIVOT_NAME = Object.freeze({ type: 'string', pattern: '^[A-Za-z_][A-Za-z0-9_.]{0,99}$' });
const PIVOT_SOURCE_SHEET = Object.freeze({ type: 'string', pattern: '^[A-Za-z_][A-Za-z0-9_]*$' });
const DEFINED_NAME = Object.freeze({ type: 'string', pattern: '^[A-Za-z_\\\\][A-Za-z0-9_.]{0,254}$' });
// GH #140 U3: create_pivot's `name` uses the SAME PIVOT_NAME above (pivots.cjs NAME, the family's own grammar exactly).
const RANGE = Object.freeze({ type: 'string', pattern: '^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$' });
// GH #137 U2: a single A1 cell, no colon alternative -- the SAME bare-cell pattern copy_range/move_range's own
// `target` already uses inline, named here since format_run's `cell` is its own top-level field, not a target.
const CELL = Object.freeze({ type: 'string', pattern: '^[A-Z]{1,3}[1-9][0-9]{0,6}$' });
// X-T01: { text } is the explicit string typing (the native Text cell type): the exact string, whitespace and number- /
// date- / formula-looking content included, never coerced; a plain string stays the untyped shape (coercible text refused)
const CELL_VALUE = Object.freeze({ anyOf: [
  { type: 'string', minLength: 1, maxLength: MAX_TEXT }, // edge whitespace / coercible text: the planner refuses it with the exact reason
  { type: 'number' },
  { type: 'object', properties: { text: { type: 'string', minLength: 1, maxLength: MAX_TEXT } }, required: ['text'], additionalProperties: false },
] });
// W1a: a formula string (cells.cjs formulaProblem: '=' first, <= MAX_FORMULA chars) or null (no formula at this
// position -- the matching `values` cell carries the value instead); external [book] refs / parse failures stay
// with the "alone" planner call in validateWriteRequest, not duplicated here.
const MAX_FORMULA = 8192; // = the write_range planner formula ceiling (cells.cjs)
const COMMENT_MAX_TEXT = require('../adapters/xlsx/families/comments.cjs').MAX_TEXT; // GH #300 K2: = the comments family's own ceiling
// GH #302 G4: a shape's paragraph list (shape-text.cjs MAX_PARAGRAPHS / MAX_CHARS)
const SHAPE_TEXT = Object.freeze({ type: 'array', minItems: 1, maxItems: require('../adapters/xlsx/shape-text.cjs').MAX_PARAGRAPHS, items: { type: 'string', maxLength: require('../adapters/xlsx/shape-text.cjs').MAX_CHARS } });
const FORMULA_CELL = Object.freeze({ anyOf: [{ type: 'string', minLength: 1, maxLength: MAX_FORMULA, pattern: '^=' }, { type: 'null' }] });

const COLUMN_LETTER = Object.freeze({ type: 'string', pattern: '^[A-Z]{1,3}$' });
const COLUMNS_SPAN = Object.freeze({ type: 'string', pattern: '^[A-Z]{1,3}(:[A-Z]{1,3})?$' });
const ROWS_SPAN = Object.freeze({ type: 'string', pattern: '^[1-9][0-9]{0,6}(:[1-9][0-9]{0,6})?$' });
const HEX_COLOR = Object.freeze({ type: 'string', pattern: '^#[0-9A-Fa-f]{6}$' });
// GH #138 U3: a page margin in millimetres (page.cjs MAX_MARGIN_MM)
const PAGE_MARGIN = Object.freeze({ type: 'number', minimum: 0, maximum: 200 });
// GitHub #81: format_range border, edge modes; `sides` only means something with 'outline' (format.cjs refuses
// it with 'inside' / 'all', and a duplicate side, in the planner -- surfaced as REFUSED_SHAPE).
// + 'cells' (every target cell gets the named sides; `sides` REQUIRED by the planner): the one border mode that is
// BOUNDED-eligible on a large workbook (format.cjs boundedEligible, X-L13/8) -- the shared-edge modes above stay
// full-observation only, so on a book over the 45000-cell read bound they refuse xlsx-observation-truncated (LIVE,
// owner's NAD file 2026-10-14); 'cells' with all four sides draws the same inner + outer grid.
// Owner GO (the owner channel): + border mode 'remove' -- X-B04 (format-remove.cjs), the named sides gone, the
// composite running on BOTH cells of every removed shared edge (LIVE-confirmed, P2.0). `style` / `color` are NOT
// required for 'remove' (nothing is drawn; the canonical planner's own `removeExpectation` refuses them as
// extra keys if sent) -- the two kept-required fields for every OTHER mode (style, color) are enforced by the
// canonical planner's `borderExpectation`, not by this static schema, the same way `sides` already is.
// Owner GO (same card, unit 5): + the other 9 measured border styles (X-B06, format-maps.cjs STYLE/API_STYLE,
// CLOSED_PROVEN: m7-b06-probe1 measured all 9 with distinct, collision-free, reopen-stable style codes;
// m7-p47-live2 LIVE-accepted double/hair/dashDot specifically, APPLIED + retry 0/0) -- the original four
// (thin/medium/thick/dashed) stay first in the list, unchanged order.
const FORMAT_BORDER = Object.freeze({ type: 'object', properties: {
  mode: { enum: ['cells', 'outline', 'inside', 'all', 'remove'] },
  style: { enum: ['thin', 'medium', 'thick', 'dashed', 'double', 'hair', 'dashDotDot', 'dashDot', 'dotted', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed'] },
  color: { anyOf: [HEX_COLOR, { const: 'auto' }] },
  sides: { type: 'array', minItems: 1, maxItems: 4, items: { enum: ['top', 'bottom', 'left', 'right'] } },
}, required: ['mode'], additionalProperties: false });
// GitHub #89: format_range font -- { color?, bold?, italic? }, at least one key (CONSTRAINTS.format_range names a
// fully-empty font object). All are boundedEligible (same rule as fill/border 'cells'/numberFormat).
// L-02/1 (owner GO 2026-10-14, inventory row L-02): + name and size -- the PUBLIC SetFontName / SetFontSize writes, read back on the
// public GetCharacters(1, len).GetFont() reader (format.cjs FONT_VALUES, X-F02/X-F03; the family already planned, wrote and verified them).
// MEASURED (m7-s22-probe1, session + two reopens): a fractional size (10.5 / 10.25 / 11.3) and a font name the machine does not have round-trip
// byte-exact (the engine neither rounds nor substitutes). size: a number in [1, 409] (the family's own bound). name: 1..31 characters -- the
// xlsx font-name limit, MEASURED by the L-02/1 LIVE probe: a 100-character name reads 100 in the session and is cut to 31 AFTER the reopen, so a
// longer name is refused here instead of being silently shortened. No control character and no leading / trailing white space (the value reaches
// the editor as a JSON string literal, never as code).
const FONT_NAME = Object.freeze({ type: 'string', minLength: 1, maxLength: 31, pattern: '^[^\\u0000-\\u001f\\u007f\\s]([^\\u0000-\\u001f\\u007f]{0,29}[^\\u0000-\\u001f\\u007f\\s])?$' });
// L-02/2 (owner GO 2026-10-14, inventory row L-02): + underline and strikeout -- the PUBLIC range-level SetUnderline / SetStrikeout writes, read
// back on the public GetCharacters(1, len).GetFont() reader (format-font-effects.cjs, X-F05; the family already planned, wrote and verified them).
// underline: the five MEASURED types (S21 probes + the product path m7-f05-live1); strikeout: a boolean (true reads true, false reads null -- the same
// null/false ambiguity bold / italic have). The write key is `strikeout` (the family's own name); read_range reports it as font.strike.
// subscript / superscript stay UNRELEASED (not requested here; text-cell only and a cross-script rule the family guards).
const FORMAT_FONT = Object.freeze({ type: 'object', properties: {
  color: HEX_COLOR, bold: { type: 'boolean' }, italic: { type: 'boolean' }, name: FONT_NAME, size: { type: 'number', minimum: 1, maximum: 409 },
  underline: { enum: ['single', 'singleAccounting', 'double', 'doubleAccounting', 'none'] }, strikeout: { type: 'boolean' },
}, additionalProperties: false });
// GitHub #93: format_range align -- { horizontal?, vertical? }, at least one key (the canonical planner's own
// alignExpectation names a fully-empty align object, 'align must be a non-empty object'). rotation /
// readingOrder / indent / shrink are NOT released. BoundedEligible (same rule as fill/font/border 'cells'/
// numberFormat).
// L-03/2 (owner GO 2026-10-14, inventory row L-03; LIVE measurement of the four axes on the prod DS /11):
//   rotation      an integer in [-90, 90] or exactly 255 (vertical text) -- the PUBLIC SetOrientation / GetOrientation; any other value silently resets
//                 to 0 in the engine (S24), so it is refused BEFORE the write (format-rotation.cjs). MEASURED WRITABLE (6/6) on /11.
//   readingOrder  'context' | 'ltr' | 'rtl' -- the PUBLIC SetReadingOrder write, read back on the bridge range.alignmentExt.read/v1. MEASURED WRITABLE (4/4).
//   indent        an integer in [0, 15], ONLY together with horizontal 'left' or 'right' IN THE SAME STEP (every other alignment silently zeroes it:
//                 the family refuses it, REFUSED_SHAPE) -- the DS Bridge WRITE (the alignment-ext write capability).
//   shrink        a boolean -- the same Bridge WRITE.
// indent / shrink are written on the runtimes of XLSX_B05_BRIDGE_WRITE_RUNTIME_ENTRIES only (a runtime off that list answers FAIL_CLOSED
// 'xlsx-align-ext-bridge-write-runtime-not-proven', 0 writes). That list was stale at /6 /7 /8 while the handler is byte-identical through /11
// (test/static/b05-write-handler-identity.test.cjs is the warrant for every entry): the prod DS /11 refused indent / shrink 3/3 + 3/3 until it was extended.
const FORMAT_ALIGN = Object.freeze({ type: 'object', properties: {
  horizontal: { enum: ['left', 'center', 'right', 'justify'] },
  vertical: { enum: ['bottom', 'center', 'distributed', 'justify', 'top'] },
  rotation: { anyOf: [{ type: 'integer', minimum: -90, maximum: 90 }, { const: 255 }] },
  readingOrder: { enum: ['context', 'ltr', 'rtl'] },
  indent: { type: 'integer', minimum: 0, maximum: 15 },
  shrink: { type: 'boolean' },
}, additionalProperties: false });

// W1a: add_conditional_format's `rule` -- one object schema per measured rule type (cf.cjs ruleOf / *Rule),
// additionalProperties false on every variant. A value-level detail the planner itself checks (bounds, exact hex
// parsing, the 'xl3Arrows' engine defect) stays with the "alone" planner call, not duplicated here -- the enum
// lists below already exclude the one measured-unsupported icon set by omission.
const CF_OPERAND = Object.freeze({ anyOf: [{ type: 'string', minLength: 1, maxLength: 1024, pattern: '^(?!=).+$' }, { type: 'number' }] });
const cfRule = (type, properties = {}, required = []) => Object.freeze({
  type: 'object', properties: Object.freeze({ type: { const: type }, ...properties }), required: Object.freeze(['type', ...required]), additionalProperties: false,
});
const CF_RULE = Object.freeze({ anyOf: [
  cfRule('cellValue', { operator: { enum: ['greater', 'less', 'equal'] }, formula1: CF_OPERAND, fill: HEX_COLOR }, ['operator', 'formula1']),
  cfRule('cellValue', { operator: { enum: ['between', 'notBetween'] }, formula1: CF_OPERAND, formula2: CF_OPERAND, fill: HEX_COLOR }, ['operator', 'formula1', 'formula2']),
  cfRule('expression', { formula: { type: 'string', minLength: 2, maxLength: 1024, pattern: '^=' }, fill: HEX_COLOR }, ['formula', 'fill']),
  cfRule('aboveAverage', { above: { type: 'boolean' }, numStdDev: { enum: [0, 1] } }, ['above']),
  cfRule('top10', { rank: { type: 'integer', minimum: 1, maximum: 1000 }, percent: { type: 'boolean' }, bottom: { type: 'boolean' } }, ['rank', 'percent', 'bottom']),
  cfRule('unique', { dupe: { const: false } }),
  cfRule('colorScale', { colors: { type: 'array', minItems: 3, maxItems: 3, items: HEX_COLOR } }),
  cfRule('databar', { barColor: HEX_COLOR, barFillType: { enum: ['solid', 'gradient'] }, direction: { enum: ['context', 'rightToLeft'] }, showValue: { type: 'boolean' }, axisPosition: { enum: ['automatic', 'none'] } }),
  cfRule('iconSet', { iconSet: { enum: ['xl3ArrowsGray', 'xl3Flags', 'xl3TrafficLights1', 'xl3TrafficLights2', 'xl3Signs', 'xl3Symbols', 'xl3Symbols2', 'xl4Arrows', 'xl4ArrowsGray', 'xl4RedToBlack', 'xl4CRV', 'xl4TrafficLights', 'xl5Arrows', 'xl5ArrowsGray', 'xl5CRV', 'xl5Quarters', 'xl3Stars', 'xl3Triangles', 'xl5Boxes'] }, reverseOrder: { type: 'boolean' }, showIconOnly: { type: 'boolean' } }),
] });

// W1d: a chart name (set_chart_title already used this pattern inline; rename_chart/set_chart_axis_title reuse it)
const CHART_NAME = Object.freeze({ type: 'string', pattern: '^[A-Za-z_][A-Za-z0-9_. -]{0,99}$' });
const CHART_TITLE = Object.freeze({ type: 'string', minLength: 1, maxLength: 255, pattern: '^[^\\r\\n]+$' });
// GH #140 U3: one of create_pivot's rows/columns/sum fields -- a source header name, one line (pivots.cjs oneLine());
// the "three DIFFERENT names" correlation is the family's own PLAN_CHECKS, not expressible in this schema.
const PIVOT_FIELD_HEADER = Object.freeze({ type: 'string', minLength: 1, maxLength: 255, pattern: '^[^\\r\\n]+$' });
// GH #302 G3: = pivots-v2.cjs FUNCTIONS (the public ApiPivotDataField.SetFunction values the #302 matrix P-03 lists)
const PIVOT_FUNCTIONS = require('../adapters/xlsx/families/pivots-v2.cjs').FUNCTIONS;

// W1d (census ACCEPTED_WRITE): set_validation's fields are FLAT on the step (validation.cjs plan() FIELDS list --
// unlike add_conditional_format's nested `rule`), so the shape here is loose per field; the cross-field
// correlation (which type takes operator+min/max/value vs type 'list' takes items XOR listRange, min<=max,
// items shape) stays with CONSTRAINTS.set_validation + the "alone" planner call, not this schema.
// date was refused pre-dispatch by omission from the type enum (the source reads it back as locale text); L-22d releases it for the dates
// whose read-back text names one serial in every locale (validation.cjs's own comment).
// L-22a (allowed EURO GO; L-22 probe test/live/xlsx-l22-validation-types-probe.live.cjs, MEASURED_OK): + 'time'.
// L-22b (same standing GO; the same probe, custom, MEASURED): + 'custom' -- ABSOLUTE-reference formula only.
// L-22d: + 'date' -- ISO 'YYYY-MM-DD' bounds, only dates that read back unambiguously (the family's dateRule() owns that).
const VALIDATION_TYPE = Object.freeze({ enum: ['wholeNumber', 'decimal', 'textLength', 'list', 'time', 'date', 'custom'] });
const VALIDATION_ALERT = Object.freeze({ enum: ['stop', 'warning', 'information'] });
// GH #301 D1: + equal, notEqual, less, greaterEqual -- the remaining four ONE-axis operators (validation.cjs's own
// ONE table owns the type/value cross-field shape; this enum only gates which strings reach it).
const VALIDATION_OPERATOR = Object.freeze({ enum: ['between', 'notBetween', 'greater', 'lessEqual', 'equal', 'notEqual', 'less', 'greaterEqual'] });
const VALIDATION_ITEM = Object.freeze({ type: 'string', minLength: 1, maxLength: 255 });
// L-22a: min/max/value stay a plain number for every other type; type 'time' takes 24-hour zero-padded
// 'HH:MM' | 'HH:MM:SS' text instead -- the family's own timeRule() owns the authoritative shape and min <= max
// check (this schema only accepts the right JS type per field, never re-derives the ordering).
// L-22d: type 'date' takes ISO 'YYYY-MM-DD' text; calendar validity, the 1900-03-01 floor and the unambiguous-read-back rule are the family's dateRule().
const VALIDATION_BOUND = Object.freeze({ anyOf: [{ type: 'number' }, { type: 'string', pattern: '^([01]\\d|2[0-3]):[0-5]\\d(:[0-5]\\d)?$' }, { type: 'string', pattern: '^\\d{4}-\\d{2}-\\d{2}$' }] });
// L-22b/e: type 'custom' takes `formula` instead of operator/min/max/value -- a '=' prefix only here; the
// bare-same-sheet-reference shape (relative/mixed/absolute all accepted since L-22e) and its reference
// classification are the family's own bareSameSheetRefProblem(), never re-derived at this schema/constraints
// layer (the SAME split as time's day-fraction order, above).
const VALIDATION_FORMULA = Object.freeze({ type: 'string', minLength: 2, maxLength: 1024, pattern: '^=' });

const step = (intent, properties, required) => Object.freeze({
  type: 'object',
  properties: Object.freeze({ intent: { const: intent }, id: ID, ...properties }),
  required: Object.freeze(['intent', ...required]),
  additionalProperties: false,
});

const ROWS = (intent) => step(intent, {
  sheet: SHEET_NAME,
  at: { type: 'integer', minimum: 1, maximum: 1048576 },
  count: { type: 'integer', minimum: 1, maximum: 100 },
  anchor: { type: 'object', properties: { column: { type: 'string', pattern: '^[A-Z]{1,3}$' },
    value: { anyOf: [{ type: 'string', minLength: 1, maxLength: 255 }, { type: 'number' }] } },
  required: ['column', 'value'], additionalProperties: false },
}, ['sheet', 'at', 'count', 'anchor']);

// L-04: the column twin of ROWS -- `at` a column letter, the anchor a ROW number plus the value found there
const COLS = (intent) => step(intent, {
  sheet: SHEET_NAME,
  at: COLUMN_LETTER,
  count: { type: 'integer', minimum: 1, maximum: 100 },
  anchor: { type: 'object', properties: { row: { type: 'integer', minimum: 1, maximum: 1048576 },
    value: { anyOf: [{ type: 'string', minLength: 1, maxLength: 255 }, { type: 'number' }] } },
  required: ['row', 'value'], additionalProperties: false },
}, ['sheet', 'at', 'count', 'anchor']);

const STEP_SCHEMAS = Object.freeze({
  write_range: step('write_range', { sheet: SHEET_ADDRESS, range: RANGE,
    values: { type: 'array', minItems: 1, maxItems: MAX_CELLS, items: { type: 'array', minItems: 1, maxItems: MAX_CELLS, items: CELL_VALUE } },
    // W1a scope +1 (GitHub #74): write_range's own `formulas` field, same shape as values -- at
    // least one of values/formulas is required (CONSTRAINTS.write_range), both may be given together (per-cell,
    // exactly one of a formula or a value; the "alone" planner call checks that correlation)
    formulas: { type: 'array', minItems: 1, maxItems: MAX_CELLS, items: { type: 'array', minItems: 1, maxItems: MAX_CELLS, items: FORMULA_CELL } },
    // GitHub #79: a cell already holding a formula is replaced by a DIFFERENT formula only with this
    // explicit flag (cells.cjs gate); without it today's xlsx-write-over-formula-refused stands. `values`
    // never overwrites a formula, flag or not -- the flag only widens the `formulas` path.
    replaceFormulas: { type: 'boolean' } },
  ['sheet', 'range']),
  insert_rows: ROWS('insert_rows'),
  delete_rows: ROWS('delete_rows'),
  rename_sheet: step('rename_sheet', { sheet: SHEET_NAME, name: SHEET_NAME }, ['sheet', 'name']),
  unmerge_range: step('unmerge_range', { sheet: SHEET_ADDRESS, range: RANGE }, ['sheet', 'range']),
  set_chart_title: step('set_chart_title', { sheet: SHEET_NAME, name: { type: 'string', pattern: '^[A-Za-z_][A-Za-z0-9_. -]{0,99}$' },
    title: { type: 'string', minLength: 1, maxLength: 255, pattern: '^[^\\r\\n]+$' } }, ['sheet', 'name', 'title']),
  format_range: step('format_range', { sheet: SHEET_ADDRESS, range: RANGE,
    // exactly ONE of numberFormat / border / fill / font / align / wrap per step (CONSTRAINTS.format_range names a missing / second key)
    format: { type: 'object', properties: { numberFormat: { type: 'string', minLength: 1, maxLength: 255 }, border: FORMAT_BORDER,
      fill: { type: 'string', pattern: '^(#[0-9A-Fa-f]{6}|none)$' }, font: FORMAT_FONT, align: FORMAT_ALIGN, wrap: { type: 'boolean' } }, additionalProperties: false } },
  ['sheet', 'range', 'format']),
  // W1a (census ACCEPTED_WRITE): merge_range (merge.cjs) -- `across` optional (merge_range only; unmerge_range needs none)
  merge_range: step('merge_range', { sheet: SHEET_ADDRESS, range: RANGE, across: { type: 'boolean' } }, ['sheet', 'range']),
  // W1a (census ACCEPTED_WRITE): sort_range (sort.cjs), one key column; GH #301 D3: OR `keys` (1..3 keys) instead of
  // keyColumn + order -- exactly one of the two shapes (CONSTRAINTS.sort_range)
  sort_range: step('sort_range', { sheet: SHEET_ADDRESS, range: RANGE, keyColumn: COLUMN_LETTER, order: { enum: ['asc', 'desc'] },
    keys: { type: 'array', minItems: 1, maxItems: 3, items: { type: 'object', properties: { column: COLUMN_LETTER, order: { enum: ['asc', 'desc'] } }, required: ['column', 'order'], additionalProperties: false } },
    hasHeaders: { type: 'boolean' } },
  ['sheet', 'range', 'hasHeaders']),
  // W1a (census ACCEPTED_WRITE): add_conditional_format (cf.cjs) -- the measured rule shapes only (CF_RULE)
  add_conditional_format: step('add_conditional_format', { sheet: SHEET_ADDRESS, range: RANGE, rule: CF_RULE }, ['sheet', 'range', 'rule']),
  // W1a (census ACCEPTED_WRITE): set_print_area (names.cjs) -- exactly one of range (one rectangle) / areas (2+,
  // CONSTRAINTS.set_print_area); the family has no cell-count ceiling on a print area, unlike write_range/merge/sort
  // X-P01: `replaces` = the CURRENT area the step overwrites (one range, or its ranges in the stored order) -- a compare-and-swap, names.cjs
  set_print_area: step('set_print_area', { sheet: SHEET_ADDRESS, range: RANGE, areas: { type: 'array', minItems: 2, maxItems: MAX_STEPS, items: RANGE }, replaces: { anyOf: [RANGE, { type: 'array', minItems: 1, maxItems: MAX_STEPS, items: RANGE }] } }, ['sheet']),
  // W1a (census ACCEPTED_WRITE): layout_range (layout.cjs) -- exactly one of columns/rows, exactly one of
  // width/height/hidden/autofit (CONSTRAINTS.layout_range); width/height bounds are exclusive-min, checked there too
  layout_range: step('layout_range', { sheet: SHEET_ADDRESS, columns: COLUMNS_SPAN, rows: ROWS_SPAN, width: { type: 'number' }, height: { type: 'number' }, hidden: { type: 'boolean' }, autofit: { const: true } }, ['sheet']),
  // W1d (census ACCEPTED_WRITE): freeze_panes (freeze.cjs) -- at least one of rows/columns (CONSTRAINTS.freeze_panes)
  freeze_panes: step('freeze_panes', { sheet: SHEET_NAME, rows: { type: 'integer', minimum: 1, maximum: 100000 }, columns: { type: 'integer', minimum: 1, maximum: 1000 } }, ['sheet']), // X-L09: freeze.cjs MAX_ROWS / MAX_COLS
  unfreeze_panes: step('unfreeze_panes', { sheet: SHEET_NAME }, ['sheet']),
  // W1d (census ACCEPTED_WRITE): set_validation (validation.cjs) -- flat fields, loose here; CONSTRAINTS.set_validation
  // enforces which fields belong together (type wholeNumber/decimal/textLength + operator + min/max or value; type
  // list + items XOR listRange)
  // GH #96: `replace` (optional, default false) -- an explicit opt-in to replace an EXISTING, DIFFERENT
  // validation on (or overlapping) the target with the requested one (Delete then Add); without it, today's
  // fail-closed refusal stands unchanged (no silent Modify).
  // L-22c: `items`' real ceiling is the family's own combined list-source LITERAL LENGTH (<=255 chars, the
  // documented Excel "List Source" field limit -- validation.cjs MAX_LIST_LITERAL_CHARS), a cross-field
  // correlation this flat schema cannot express. maxItems here is only the loose, DoS-shaped upper bound
  // (validation.cjs MAX_LIST_ITEMS) -- it must never be SMALLER than the family's own ceiling, or a valid
  // short-item list could be refused here before the family ever sees it.
  set_validation: step('set_validation', { sheet: SHEET_ADDRESS, range: RANGE, type: VALIDATION_TYPE, operator: VALIDATION_OPERATOR, alertStyle: VALIDATION_ALERT,
    min: VALIDATION_BOUND, max: VALIDATION_BOUND, value: VALIDATION_BOUND,
    items: { type: 'array', minItems: 1, maxItems: 128, items: VALIDATION_ITEM }, listRange: RANGE, replace: { type: 'boolean' },
    formula: VALIDATION_FORMULA },
  ['sheet', 'range', 'type', 'alertStyle']),
  clear_validation: step('clear_validation', { sheet: SHEET_ADDRESS, range: RANGE }, ['sheet', 'range']),
  // W1d (census ACCEPTED_WRITE): delete_conditional_format (cf.cjs) -- SAME shape as add_conditional_format
  // (the rule identifies WHICH of possibly several CF rules on the range to delete)
  delete_conditional_format: step('delete_conditional_format', { sheet: SHEET_ADDRESS, range: RANGE, rule: CF_RULE }, ['sheet', 'range', 'rule']),
  // W1d (census ACCEPTED_WRITE): clear_filter (filter.cjs) -- same sheet/range shape as filter_range, no criteria
  clear_filter: step('clear_filter', { sheet: SHEET_ADDRESS, range: RANGE }, ['sheet', 'range']),
  // W1d (census ACCEPTED_WRITE): clear_print_area (names.cjs) -- sheet only, no range/areas (unlike set_print_area)
  clear_print_area: step('clear_print_area', { sheet: SHEET_ADDRESS }, ['sheet']),
  // W1d (census ACCEPTED_WRITE): create_sheet (sheets.cjs) -- the SAME released no-quote name class as rename_sheet
  // (sheets.cjs's own nameProblem() also accepts a wider quoted/unicode class; that wider class stays unreleased)
  create_sheet: step('create_sheet', { name: SHEET_NAME }, ['name']),
  // W1d (census ACCEPTED_WRITE): rename_chart (charts.cjs) -- newName != name enforced in CONSTRAINTS.rename_chart
  rename_chart: step('rename_chart', { sheet: SHEET_NAME, name: CHART_NAME, newName: CHART_NAME }, ['sheet', 'name', 'newName']),
  // W1d (census ACCEPTED_WRITE): set_chart_axis_title (charts.cjs) -- axis 'horizontal' | 'vertical'
  set_chart_axis_title: step('set_chart_axis_title', { sheet: SHEET_NAME, name: CHART_NAME, axis: { enum: ['horizontal', 'vertical'] }, title: CHART_TITLE }, ['sheet', 'name', 'axis', 'title']),
  // GH #106: clear_range -- CONTENT only, one A1 range on one sheet (the clear_validation shape)
  // L-14 / L-17: `scope` = what the clear removes -- 'contents' (default, the value only) | 'formats' (formatting, value kept) | 'all' (both); clear.cjs
  clear_range: step('clear_range', { sheet: SHEET_ADDRESS, range: RANGE, scope: { enum: ['contents', 'formats', 'all'] } }, ['sheet', 'range']),
  // X-O04/11: target restricted to the two released targets at the SCHEMA level already (the other
  // six internal targets are refused by omission, same enum-based refusal CHART_TYPES/legend already
  // use) -- seriesIndex/axis both optional here because WHICH one applies depends on target,
  // enforced by CONSTRAINTS.set_chart_style, not the schema.
  // X-O05/3: + gapOverlap/seriesOrder (same enum-based release pattern). gapWidth/overlap carry
  // their LIVE-measured range here (0..500 / -100..100) -- the schema layer, not CONSTRAINTS,
  // same division HEX_COLOR/outline.width already use above. order's own range (0..length-1) and
  // uniqueness are value-dependent (the length comes from the array itself) and are a
  // CONSTRAINTS-layer check, same as charts.cjs's own plan()/classify() split.
  // X-O05/5-C: + markerSymbol/lineStyle (same enum-based release pattern). symbol/dash carry
  // their own measured enum here -- the schema layer, same division gapWidth/overlap already
  // use above; size's range (2..72) is likewise schema-layer, same as outline.width's own bound.
  set_chart_style: step('set_chart_style', { sheet: SHEET_NAME, name: CHART_NAME, target: { enum: ['marker', 'gridlines', 'gapOverlap', 'seriesOrder', 'markerSymbol', 'lineStyle'] },
    fill: HEX_COLOR, outline: { type: 'object', properties: { width: { type: 'number' }, color: HEX_COLOR }, required: ['width', 'color'], additionalProperties: false },
    seriesIndex: { type: 'integer', minimum: 0 }, axis: { enum: ['vertical', 'horizontal'] },
    gapWidth: { type: 'number', minimum: 0, maximum: 500 }, overlap: { type: 'number', minimum: -100, maximum: 100 },
    order: { type: 'array', items: { type: 'integer', minimum: 0 }, minItems: 1 },
    symbol: { enum: ['circle', 'dash', 'diamond', 'dot', 'none', 'plus', 'square', 'star', 'triangle', 'x'] },
    size: { type: 'integer', minimum: 2, maximum: 72 },
    dash: { enum: ['solid', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot'] },
    smooth: { type: 'boolean' } },
  ['sheet', 'name', 'target']),
  // L-04: insert_columns / delete_columns
  insert_columns: COLS('insert_columns'),
  delete_columns: COLS('delete_columns'),
  // L-05: delete_sheet (sheets.cjs) -- sheet only; the SAME released sheet-name class as rename_sheet / create_sheet
  delete_sheet: step('delete_sheet', { sheet: SHEET_NAME }, ['sheet']),
  // L-05/2: move_sheet (sheets.cjs) -- `sheet` goes before OR after the reference sheet (exactly one of before / after: CONSTRAINTS)
  move_sheet: step('move_sheet', { sheet: SHEET_NAME, before: SHEET_NAME, after: SHEET_NAME }, ['sheet']),
  // L-07: set_chart (charts.cjs) -- the series reference is built unquoted, so the sheet is a BARE name (CONSTRAINTS); the range needs a header row, a category column and data
  set_chart: step('set_chart', { sheet: SHEET_NAME, name: CHART_NAME, range: RANGE,
    type: { type: 'string', enum: ['bar', 'lineNormal', 'area', 'pie', 'doughnut', 'radar', 'stock'] },
    title: { type: 'string', minLength: 1, maxLength: 255, pattern: '^[^\\r\\n]+$' },
    orientation: { type: 'string', enum: ['columns', 'rows'] } }, ['sheet', 'name', 'range', 'type', 'title']),
  // L-07: set_chart_legend (charts.cjs) -- the position the editor accepts; an unsupported one is refused here, before any dispatch
  set_chart_legend: step('set_chart_legend', { sheet: SHEET_NAME, name: CHART_NAME, legend: { type: 'string', enum: ['bottom', 'top', 'left', 'right', 'none'] } }, ['sheet', 'name', 'legend']),
  // L-06: copy_range (range.cjs) -- the source block, the ONE-cell top-left of the target, an optional other sheet; the released sheet-name class on both sheets
  copy_range: step('copy_range', { sheet: SHEET_NAME, range: RANGE, target: { type: 'string', pattern: '^[A-Z]{1,3}[1-9][0-9]{0,6}$' }, targetSheet: SHEET_NAME }, ['sheet', 'range', 'target']),
  // X-O07/4b: copy_chart -- sheet/name are the DESTINATION (charts.cjs's own plan() convention);
  // sourceSheet/sourceName name what to copy FROM; `at` is REQUIRED (never a guessed placement).
  copy_chart: step('copy_chart', { sheet: SHEET_NAME, name: CHART_NAME, sourceSheet: SHEET_NAME, sourceName: CHART_NAME,
    at: { type: 'object', properties: { col: { type: 'integer', minimum: 0 }, row: { type: 'integer', minimum: 0 } }, required: ['col', 'row'], additionalProperties: false } },
  ['sheet', 'name', 'sourceSheet', 'sourceName', 'at']),
  // L-05/3: copy_sheet (sheets.cjs) -- the source sheet and the NEW name; the released sheet-name class on both (CONSTRAINTS: not the same name, case-insensitively)
  copy_sheet: step('copy_sheet', { sheet: SHEET_NAME, name: SHEET_NAME }, ['sheet', 'name']),
  // L-06/b: move_range -- the SAME shape as copy_range, no `expect` field (classify() builds it server-side, see the
  // XLSX_WRITE_ALLOWLIST comment above).
  move_range: step('move_range', { sheet: SHEET_NAME, range: RANGE, target: { type: 'string', pattern: '^[A-Z]{1,3}[1-9][0-9]{0,6}$' }, targetSheet: SHEET_NAME }, ['sheet', 'range', 'target']),
  // X-O07/3 (census not applicable -- same non-census release convention set_chart_style
  // above uses): delete_chart -- just sheet/name, no further fields.
  delete_chart: step('delete_chart', { sheet: SHEET_NAME, name: CHART_NAME }, ['sheet', 'name']),
  // L-11/2: filter_range (filter.cjs) -- sheet + range + the 1-based column inside the range + EXACTLY ONE criteria (CONSTRAINTS.filter_range names a missing /
  // second one). The measured shapes only: one plain value, two different plain values (xlOr), or two numeric comparisons (xlAnd).
  filter_range: step('filter_range', { sheet: SHEET_ADDRESS, range: RANGE, field: { type: 'integer', minimum: 1, maximum: 18278 },
    value: { anyOf: [{ type: 'string', minLength: 1, maxLength: 255 }, { type: 'number' }] },
    // GH #301 D2: values 2..10 (3+ reads back as an xlFilterValues list); top n items; dynamic above / below average
    values: { type: 'array', minItems: 2, maxItems: 10, items: { anyOf: [{ type: 'string', minLength: 1, maxLength: 255 }, { type: 'number' }] } },
    and: { type: 'array', minItems: 2, maxItems: 2, items: { type: 'string', pattern: '^[<>]-?[0-9]+(\\.[0-9]+)?$' } },
    top: { type: 'integer', minimum: 1, maximum: 500 },
    dynamic: { enum: ['aboveAverage', 'belowAverage'] } },
  ['sheet', 'range', 'field']),
  // GH #138 U1: set_defined_name / delete_defined_name (names.cjs) -- name + (set) an ABSOLUTE refersTo ('=Sheet!$A$1' / '=Sheet!$A$1:$B$2', quoted
  // sheet names) + optional scope; the identifier grammar, the reserved names and the reference shape are re-checked by the family's plan().
  set_defined_name: step('set_defined_name', { name: DEFINED_NAME, refersTo: { type: 'string', minLength: 2, maxLength: 300, pattern: '^=' }, scope: SHEET_ADDRESS }, ['name', 'refersTo']),
  delete_defined_name: step('delete_defined_name', { name: DEFINED_NAME, scope: SHEET_ADDRESS }, ['name']),
  // L-18/1: delete_cells (delete-cells.cjs) -- the block to delete, the shift direction, and a retry-safety anchor
  // (the SAME { column|row, value } shape structure.cjs's insert/delete_rows/columns already uses, scoped to the
  // block's own band -- CONSTRAINTS does not check which of column/row is required for which shift; the family's
  // own plan() does, since it also enforces the anchor lying WITHIN the block's band). anchor.value is the value
  // of the cell DIRECTLY BELOW the block (shift 'up') / DIRECTLY TO ITS RIGHT (shift 'left'), in a column/row
  // that falls within the block's own band (the same 'NOW'-position convention as ROWS/COLS above, not the
  // 'at+count' insert-side one) -- it must be UNIQUE within that column/row, or the step refuses, named, rather
  // than guessing which occurrence was meant. If the block is at the very end of the sheet's data (nothing
  // below/right to name as an anchor), there is nothing to anchor on at all: the step refuses, named
  // (`xlsx-delete-cells-anchor-not-found`), by design -- this is not a gap to fill later.
  delete_cells: step('delete_cells', { sheet: SHEET_NAME, range: RANGE, shift: { enum: ['up', 'left'] },
    anchor: { type: 'object', properties: { column: { type: 'string', pattern: '^[A-Z]{1,3}$' }, row: { type: 'integer', minimum: 1, maximum: 1048576 },
      value: { anyOf: [{ type: 'string', minLength: 1, maxLength: 255 }, { type: 'number' }] } },
    required: ['value'], additionalProperties: false } },
  ['sheet', 'range', 'shift', 'anchor']),
  // GH #138 U2: set_print_titles / clear_print_titles (names.cjs planPrint) -- at least one of rows/columns for the set (CONSTRAINTS.set_print_titles);
  // the exact 'n:m' / 'A:B' two-part form, the reserved-name and another-scope rules are re-checked by the family's plan().
  set_print_titles: step('set_print_titles', { sheet: SHEET_ADDRESS, rows: ROWS_SPAN, columns: COLUMNS_SPAN }, ['sheet']),
  clear_print_titles: step('clear_print_titles', { sheet: SHEET_ADDRESS }, ['sheet']),
  // GH #138 U3: set_page_layout (page.cjs) -- sheet + at least one field (CONSTRAINTS.set_page_layout); margins in millimetres, the family's 200 mm cap
  set_page_layout: step('set_page_layout', { sheet: SHEET_ADDRESS, orientation: { enum: ['portrait', 'landscape'] }, printGridlines: { type: 'boolean' }, printHeadings: { type: 'boolean' },
    topMargin: PAGE_MARGIN, bottomMargin: PAGE_MARGIN, leftMargin: PAGE_MARGIN, rightMargin: PAGE_MARGIN }, ['sheet']),
  // GH #139 U2: move_chart (charts.cjs) -- the SAME `at` shape copy_chart already uses (cell-aligned,
  // offset always 0 at this boundary); the family's plan() re-checks the chart exists / is unambiguous.
  move_chart: step('move_chart', { sheet: SHEET_NAME, name: CHART_NAME,
    at: { type: 'object', properties: { col: { type: 'integer', minimum: 0 }, row: { type: 'integer', minimum: 0 } }, required: ['col', 'row'], additionalProperties: false } },
  ['sheet', 'name', 'at']),
  // GH #139 U3: insert_shape / insert_image (drawings-insert.cjs) -- geometry restricted to 'rect' (the family's
  // own GEOMETRIES, v1 scope); fill required, stroke OPTIONAL (v2, width the measured readable axis, colour
  // write-only -- both validated here so a malformed value fails at plan time, not silently written). `at` is the
  // SAME cell-aligned shape move_chart/copy_chart already use. `size` is points, a plain number here (the same
  // division layout_range's width/height and set_chart_style's outline.width already use -- this minimal
  // validator has no exclusiveMinimum; the family's own sizeProblem() is the layer that refuses <= 0, same
  // split everywhere else in this family).
  insert_shape: step('insert_shape', { sheet: SHEET_NAME, name: CHART_NAME, geometry: { enum: ['rect'] },
    at: { type: 'object', properties: { col: { type: 'integer', minimum: 0 }, row: { type: 'integer', minimum: 0 } }, required: ['col', 'row'], additionalProperties: false },
    size: { type: 'object', properties: { width: { type: 'number' }, height: { type: 'number' } }, required: ['width', 'height'], additionalProperties: false },
    fill: HEX_COLOR,
    stroke: { type: 'object', properties: { width: { type: 'number' }, color: HEX_COLOR }, required: ['width', 'color'], additionalProperties: false },
    // GH #302 G4: an optional paragraph list (shape-text.cjs owns the line-break / control / non-empty rules)
    text: SHAPE_TEXT },
  ['sheet', 'name', 'geometry', 'at', 'size', 'fill']),
  // GH #139 U3: insert_image -- PNG data URI ONLY (an http(s):// URL is refused at plan time by the family's own
  // plan(), never fetched; this schema's pattern already excludes it, so the refusal is REFUSED_SHAPE here, before
  // even reaching the family). maxLength: a 1 MiB PNG base64-encodes (with this prefix) to 1,398,126 characters
  // exactly -- 1,400,000 leaves a small, named margin, not an unbounded field.
  insert_image: step('insert_image', { sheet: SHEET_NAME, name: CHART_NAME,
    at: { type: 'object', properties: { col: { type: 'integer', minimum: 0 }, row: { type: 'integer', minimum: 0 } }, required: ['col', 'row'], additionalProperties: false },
    size: { type: 'object', properties: { width: { type: 'number' }, height: { type: 'number' } }, required: ['width', 'height'], additionalProperties: false },
    source: { type: 'string', pattern: '^data:image\\/png;base64,[A-Za-z0-9+/]+=*$', maxLength: 1400000 } },
  ['sheet', 'name', 'at', 'size', 'source']),
  // GH #137 U2: format_run (format-run.cjs) -- `runs` is the SAME { start, length, font } shape the family's
  // own plan() validates (start/length 1-based integers >= 1, maxItems 256 matching RUN_MAP_CHAR_CEILING --
  // a run cannot be shorter than 1 character, so more than 256 runs could not fit anyway). `font` carries
  // ONLY the family's 5 measured axes (bold, italic, name maxLength 100 -- the family's own cap, NOT
  // FORMAT_FONT's 31-char sheet-safe one, since a font name here is never used as a sheet/file name; size
  // [1,409]; color HEX_COLOR). The schema allows an EMPTY font object per run -- this minimal validator has
  // no "at least one of" keyword; the family's own runFontWant() refuses that, named, before dispatch (the
  // SAME split write_range's values/formulas correlation already uses).
  format_run: step('format_run', { sheet: SHEET_NAME, cell: CELL,
    runs: { type: 'array', minItems: 1, maxItems: 256, items: {
      type: 'object',
      properties: { start: { type: 'integer', minimum: 1 }, length: { type: 'integer', minimum: 1 },
        font: { type: 'object', properties: { bold: { type: 'boolean' }, italic: { type: 'boolean' }, name: { type: 'string', minLength: 1, maxLength: 100 }, size: { type: 'number', minimum: 1, maximum: 409 }, color: HEX_COLOR }, additionalProperties: false } },
      required: ['start', 'length', 'font'], additionalProperties: false } } },
  ['sheet', 'cell', 'runs']),
  // GH #138 U4: set_print_fit (print.cjs) -- both numbers required together (the family's own wantBase() --
  // fitToWidth=fitToHeight=0 is fit-OFF, every other pair is fit-ON).
  set_print_fit: step('set_print_fit', { sheet: SHEET_NAME, fitToWidth: { type: 'integer', minimum: 0, maximum: 100 }, fitToHeight: { type: 'integer', minimum: 0, maximum: 100 } },
  ['sheet', 'fitToWidth', 'fitToHeight']),
  // set_header_footer: slot one of the family's 6 measured HF_SLOTS (headerWant()); value a non-empty code
  // (<=255 chars, the family's own cap) or null to clear the slot.
  set_header_footer: step('set_header_footer', { sheet: SHEET_NAME, slot: { enum: ['oddHeader', 'oddFooter', 'firstHeader', 'firstFooter', 'evenHeader', 'evenFooter'] },
    value: { anyOf: [{ type: 'string', minLength: 1, maxLength: 255 }, { type: 'null' }] } },
  ['sheet', 'slot', 'value']),
  // set_page_break: `at` accepts either shape here (a row number or a column letter) -- the family's own
  // breakWant() re-checks the axis<->at correlation and refuses a mismatched one, named ('at must be a row
  // number ... or a column letter'). `remove`, when given, must be exactly `true` (the family's own gate).
  set_page_break: step('set_page_break', { sheet: SHEET_NAME, axis: { enum: ['row', 'column'] },
    at: { anyOf: [{ type: 'integer', minimum: 1, maximum: 1048576 }, { type: 'string', pattern: '^[A-Z]{1,3}$' }] }, remove: { const: true } },
  ['sheet', 'axis', 'at']),
  // set_first_page_number: a positive integer (<=32767, the family's own cap) or null to clear the number.
  set_first_page_number: step('set_first_page_number', { sheet: SHEET_NAME, value: { anyOf: [{ type: 'integer', minimum: 1, maximum: 32767 }, { type: 'null' }] } },
  ['sheet', 'value']),
  // set_print_centering: exactly one of horizontal/vertical -- neither is required at THIS layer (CONSTRAINTS.
  // set_print_centering refuses zero or both, the SAME split set_page_layout's "at least one of" uses).
  set_print_centering: step('set_print_centering', { sheet: SHEET_NAME, horizontal: { type: 'boolean' }, vertical: { type: 'boolean' } },
  ['sheet']),
  // set_print_scale: the family's own pageWant() range (fit-OFF + this percentage, UI asc_SetPrintScale).
  set_print_scale: step('set_print_scale', { sheet: SHEET_NAME, scale: { type: 'integer', minimum: 10, maximum: 400 } },
  ['sheet', 'scale']),
  // set_paper_size: the EXACT editor menu the family's own PAPER table lists (owner OD-C/OD-A) -- any other
  // name is REFUSED_SHAPE here, before the family's own exact-mm lookup even runs.
  set_paper_size: step('set_paper_size', { sheet: SHEET_NAME, paper: { enum: ['US Letter', 'US Legal', 'A4', 'A5', 'B5', 'Envelope #10', 'Envelope DL', 'Tabloid', 'A3', 'Tabloid Oversize', 'ROC 16K', 'Envelope Choukei 3', 'Super B/A3'] } },
  ['sheet', 'paper']),
  // set_header_footer_flag: flag one of the family's 2 measured HF_FLAGS (pageWant()).
  set_header_footer_flag: step('set_header_footer_flag', { sheet: SHEET_NAME, flag: { enum: ['differentFirst', 'differentOddEven'] }, value: { type: 'boolean' } },
  ['sheet', 'flag', 'value']),
  // GH #140 U4: refresh_pivot (families/pivot-refresh.cjs FIELDS: pivot / sourceSheet / sourceRange) -- no layout, no aggregation key exists; the family's plan() re-checks the
  // range (a header row, 1..50 data rows, at least three columns)
  refresh_pivot: step('refresh_pivot', { pivot: PIVOT_NAME, sourceSheet: PIVOT_SOURCE_SHEET, sourceRange: RANGE }, ['pivot', 'sourceSheet', 'sourceRange']),
  // GH #140 U3: create_pivot (pivots.cjs) -- `name`: PIVOT_NAME, the family's own workbook-wide identity grammar.
  // `sheet`/`pivotSheet`: SHEET_NAME -- the family's own narrower "needs no quoting" grammar (pivots.cjs SHEET) is
  // re-checked by plan(), REFUSED_SHAPE, the SAME split every other sheet-carrying intent here already uses.
  // `range`: RANGE; the header-row/data-row-count/column-count shape (1..MAX_DATA_ROWS rows, >=3 columns) is the
  // family's own plan() territory, not duplicated here. rows/columns/sum: PIVOT_FIELD_HEADER. `at`: CELL -- the
  // family's own parseRange(at).count===1 check this already satisfies by construction.
  // GH #302 G3 (create_pivot v2, pivots-v2.cjs): rows / columns may be header LISTS, + pages (a list), data (1..4
  // { field, function }), layout, style; `columns` and `sum` are no longer schema-required -- the family's plan() owns
  // which shape is v1 (rows + columns + sum, one each) and which is v2 (data, no sum), and refuses a mix by name.
  create_pivot: step('create_pivot', { name: PIVOT_NAME, sheet: SHEET_NAME, range: RANGE,
    rows: { anyOf: [PIVOT_FIELD_HEADER, { type: 'array', minItems: 1, maxItems: 4, items: PIVOT_FIELD_HEADER }] },
    columns: { anyOf: [PIVOT_FIELD_HEADER, { type: 'array', minItems: 0, maxItems: 3, items: PIVOT_FIELD_HEADER }] },
    pages: { type: 'array', minItems: 0, maxItems: 3, items: PIVOT_FIELD_HEADER },
    sum: PIVOT_FIELD_HEADER,
    data: { type: 'array', minItems: 1, maxItems: 4, items: { type: 'object', properties: { field: PIVOT_FIELD_HEADER, function: { enum: PIVOT_FUNCTIONS } }, required: ['field', 'function'], additionalProperties: false } },
    layout: { enum: ['Compact', 'Outline', 'Tabular'] },
    style: { type: 'string', pattern: '^PivotStyle(Light|Medium|Dark)([1-9]|1[0-9]|2[0-8])$' },
    pivotSheet: SHEET_NAME, at: CELL },
  ['name', 'sheet', 'range', 'rows', 'pivotSheet', 'at']),
  // GH #301 D4: write_array_formula (array-formula.cjs) -- one formula over the area, the write_range formula rules
  write_array_formula: step('write_array_formula', { sheet: SHEET_ADDRESS, range: RANGE, formula: { type: 'string', minLength: 1, maxLength: MAX_FORMULA, pattern: '^=' } }, ['sheet', 'range', 'formula']),
  // GH #300 K2: add_comment (comments.cjs) -- sheet, ONE cell, the text; NO author field (the session user is the author)
  add_comment: step('add_comment', { sheet: SHEET_ADDRESS, cell: CELL, text: { type: 'string', minLength: 1, maxLength: COMMENT_MAX_TEXT } }, ['sheet', 'cell', 'text']),
  // GH #302 G4: set_shape_text (shape-text family) -- the whole text of one named shape, a paragraph list
  set_shape_text: step('set_shape_text', { sheet: SHEET_NAME, name: CHART_NAME, text: SHAPE_TEXT }, ['sheet', 'name', 'text']),
  // GH #302 G5: set_chart_point_data_label (chart-point-labels family) -- one point's four flags, read_charts' own names
  set_chart_point_data_label: step('set_chart_point_data_label', { sheet: SHEET_NAME, name: CHART_NAME,
    series: { type: 'integer', minimum: 0 }, point: { type: 'integer', minimum: 0 },
    labels: { type: 'object', properties: { showSerName: { type: 'boolean' }, showCatName: { type: 'boolean' }, showVal: { type: 'boolean' }, showPercent: { type: 'boolean' } },
      required: ['showSerName', 'showCatName', 'showVal', 'showPercent'], additionalProperties: false } },
  ['sheet', 'name', 'series', 'point', 'labels']),
  // GH #302 G7: set_shape_rotation (shape-rotation family) -- degrees, 0 <= rotation < 360 (read_drawings' own field name)
  set_shape_rotation: step('set_shape_rotation', { sheet: SHEET_NAME, name: CHART_NAME, rotation: { type: 'number', minimum: 0, exclusiveMaximum: 360 } }, ['sheet', 'name', 'rotation']),
  // GH #303 R1: set_sheet_visibility (sheet-visibility family) -- read_sheet_state's own two values
  set_sheet_visibility: step('set_sheet_visibility', { sheet: SHEET_NAME, visibility: { enum: ['visible', 'hidden'] } }, ['sheet', 'visibility']),
});

// ---- the supported-shape constraints a JSON Schema cannot state (each returns null or why) ----


function rangeProblem(range) {
  const g = parseRange(range);
  if (!g) return 'range must be a plain A1 range, top-left first';
  return g.count > MAX_CELLS ? `range has ${g.count} cells, over the ${MAX_CELLS} ceiling` : null;
}

/** A plain A1 range, no MAX_CELLS ceiling (set_print_area has none in the family, unlike write_range/merge/sort). */
const rangeShapeProblem = range => (parseRange(range) ? null : 'range must be a plain A1 range, top-left first');

const CONSTRAINTS = Object.freeze({
  write_range(s) {
    if (s.values === undefined && s.formulas === undefined) return 'values or formulas is required';
    const why = rangeProblem(s.range);
    if (why) return why;
    const g = parseRange(s.range);
    const fits = m => m === undefined || (m.length === g.rows && m.every(row => row.length === g.cols));
    if (!fits(s.values)) return `values must be a ${g.rows}x${g.cols} matrix matching ${s.range}`;
    if (!fits(s.formulas)) return `formulas must be a ${g.rows}x${g.cols} matrix matching ${s.range}`;
    return null;
  },
  insert_rows: s => sheetNameProblem(s.sheet),
  delete_rows: s => sheetNameProblem(s.sheet),
  rename_sheet: s => sheetNameProblem(s.sheet) || sheetNameProblem(s.name)
    || (s.name.toLowerCase() === s.sheet.toLowerCase() ? 'the new name equals the old one, case-insensitively' : null),
  unmerge_range: s => rangeProblem(s.range),
  set_chart_title: s => sheetNameProblem(s.sheet),
  format_range: s => rangeProblem(s.range)
    || (Object.keys(s.format).length !== 1 ? 'format must carry exactly one of numberFormat / border / fill / font / align / wrap (one key per step)' : null)
    || ('font' in s.format && Object.keys(s.format.font).length === 0 ? 'format.font must carry at least one of color, bold, italic, name, size, underline, strikeout' : null),
  merge_range(s) {
    const why = rangeProblem(s.range);
    if (why) return why;
    const g = parseRange(s.range);
    if (g.count < 2) return 'merge_range needs at least two cells';
    return s.across === true && g.cols < 2 ? 'merge_range across:true needs at least two columns (a single-column row has nothing to merge within itself)' : null;
  },
  sort_range(s) {
    const why = rangeProblem(s.range);
    if (why) return why;
    const g = parseRange(s.range);
    const single = s.keyColumn !== undefined || s.order !== undefined;
    if (single && s.keys !== undefined) return 'give either keys or keyColumn + order, not both';
    if (!single && s.keys === undefined) return 'sort_range needs keyColumn + order, or keys';
    if (single && (s.keyColumn === undefined || s.order === undefined)) return 'keyColumn and order go together';
    const cols = single ? [s.keyColumn] : s.keys.map(k => k.column);
    if (cols.some(c => colNum(c) < g.c1 || colNum(c) > g.c2)) return single ? 'keyColumn must be a column letter inside the range' : 'every key column must be a column letter inside the range';
    if (new Set(cols).size !== cols.length) return 'a key column may appear only once';
    const first = g.r1 + (s.hasHeaders ? 1 : 0);
    return g.r2 - first < 1 ? 'fewer than two data rows' : null;
  },
  // the family has no MAX_CELLS ceiling on the range; the rule's own value-level shape is checked by the schema
  // (CF_RULE) and, for anything the schema cannot express, the "alone" planner call
  add_conditional_format: () => null,
  set_print_area(s) {
    if (('range' in s) === ('areas' in s)) return 'exactly one of range or areas is required';
    const rep = s.replaces === undefined ? null : [].concat(s.replaces).map(rangeShapeProblem).find(Boolean);
    if (rep) return `replaces: ${rep}`;
    return s.range !== undefined ? rangeShapeProblem(s.range) : null;
  },
  layout_range(s) {
    if (('columns' in s) === ('rows' in s)) return 'requires exactly one of columns or rows';
    const props = ['width', 'height', 'hidden', 'autofit'].filter(k => k in s);
    if (props.length !== 1) return 'requires exactly one of width, height, hidden, autofit';
    const kind = 'columns' in s ? 'columns' : 'rows';
    const prop = props[0];
    if (prop === 'width' && kind !== 'columns') return 'width applies to columns';
    if (prop === 'height' && kind !== 'rows') return 'height applies to rows';
    if (prop === 'width' && !(s.width > 0 && s.width <= 255)) return 'width must be a number in (0, 255]';
    if (prop === 'height' && !(s.height > 0 && s.height <= 409)) return 'height must be a number in (0, 409]';
    return null;
  },
  // W1d (census ACCEPTED_WRITE)
  freeze_panes: s => (s.rows === undefined && s.columns === undefined ? 'requires rows and / or columns' : sheetNameProblem(s.sheet)),
  unfreeze_panes: s => sheetNameProblem(s.sheet),
  set_validation(s) {
    const why = rangeProblem(s.range);
    if (why) return why;
    if (s.type === 'list') {
      for (const k of ['operator', 'min', 'max', 'value']) if (k in s) return `type list does not take ${k}`;
      if (('items' in s) === ('listRange' in s)) return 'type list takes exactly one of items | listRange';
      if ('listRange' in s && !parseRange(s.listRange)) return 'listRange must be a plain A1 range';
      return null;
    }
    // L-22b/e: custom -- `formula` only; the bare-same-sheet-reference shape and its classification are the
    // family's own bareSameSheetRefProblem() call, at the "alone" planner step (not re-derived here).
    if (s.type === 'custom') {
      for (const k of ['operator', 'min', 'max', 'value', 'items', 'listRange']) if (k in s) return `type custom does not take ${k}`;
      if (s.formula === undefined) return 'type custom requires formula';
      return null;
    }
    for (const k of ['items', 'listRange', 'formula']) if (k in s) return `${k} belongs to type ${k === 'formula' ? 'custom' : 'list'}`;
    if (s.operator === undefined) return 'operator is required for type wholeNumber / decimal / textLength / time / date';
    if (['between', 'notBetween'].includes(s.operator)) {
      if ('value' in s) return `operator ${s.operator} takes min and max, not value`;
      if (s.min === undefined || s.max === undefined) return `operator ${s.operator} requires min and max`;
      // L-22a: min/max ORDER for type 'time' is the family's own timeRule() call (a day-fraction compare, not
      // this layer's string/number one) -- a bad order still refuses, just at the "alone" planner step below.
      if (!['time', 'date'].includes(s.type) && s.min > s.max) return 'min must be <= max';
    } else {
      if ('min' in s || 'max' in s) return `operator ${s.operator} takes value, not min / max`;
      if (s.value === undefined) return `operator ${s.operator} requires value`;
    }
    if (s.type === 'textLength' && [s.min, s.max, s.value].some(n => n !== undefined && n < 0)) return 'textLength bounds must be non-negative';
    if (!['decimal', 'time', 'date'].includes(s.type) && [s.min, s.max, s.value].some(n => n !== undefined && !Number.isInteger(n))) return `${s.type} bounds must be integers`;
    // L-22a: type 'time' must carry its bounds as the 24-hour TEXT the schema's VALIDATION_BOUND also accepts
    // (a plain number here would mean a different type was meant) -- the family's timeRule() re-checks the shape.
    if (s.type === 'time' && [s.min, s.max, s.value].some(n => n !== undefined && typeof n !== 'string')) return 'time bounds must be HH:MM or HH:MM:SS text';
    // L-22d: type 'date' must carry ISO 'YYYY-MM-DD' text; the family's dateRule() re-checks the calendar and the read-back rule.
    if (s.type === 'date' && [s.min, s.max, s.value].some(n => n !== undefined && typeof n !== 'string')) return 'date bounds must be YYYY-MM-DD text';
    if (!['time', 'date'].includes(s.type) && [s.min, s.max, s.value].some(n => typeof n === 'string')) return `${s.type} bounds must be numbers, not text`;
    return null;
  },
  clear_validation: s => rangeProblem(s.range),
  // the rule's own value-level shape is checked by the schema (CF_RULE), same as add_conditional_format
  delete_conditional_format: () => null,
  clear_filter(s) {
    const why = rangeProblem(s.range);
    if (why) return why;
    const g = parseRange(s.range);
    return g.r2 - g.r1 < 1 ? 'range requires a header row and at least one data row' : null;
  },
  clear_print_area: s => bareNameProblem(s.sheet),
  create_sheet: s => sheetNameProblem(s.name),
  rename_chart: s => sheetNameProblem(s.sheet) || (s.newName === s.name ? 'newName must be a chart name different from name' : null),
  set_chart_axis_title: s => sheetNameProblem(s.sheet),
  clear_range: s => rangeProblem(s.range),
  // X-O04/11: the schema enum already excludes the six internal-only targets; the released
  // targets diverge in shape from here on (gridlines: no fill, outline+axis required, no
  // seriesIndex; marker: seriesIndex required, at least one of fill/outline) -- the SAME per-target
  // divergence charts.cjs's own styleProblem() enforces, mirrored here because this is the schema
  // the product surface sees BEFORE the step ever reaches charts.cjs's plan().
  // X-O05/3: + gapOverlap (bar-only, checked later by classifyStyle against live chart data --
  // this layer only knows the shape) and seriesOrder (order is the whole permutation; its upper
  // bound and uniqueness are checked here since they are a self-contained array property, but its
  // LENGTH against the chart's actual series count is live-dependent, same split as charts.cjs's
  // own plan()/classify()).
  set_chart_style(s) {
    const why = sheetNameProblem(s.sheet);
    if (why) return why;
    if (s.target === 'gridlines') {
      if (s.fill !== undefined) return 'gridlines has no fill, only outline';
      if (s.outline === undefined) return 'outline is required for target gridlines';
      if (s.axis === undefined) return 'axis is required for target gridlines';
      if (s.seriesIndex !== undefined) return 'seriesIndex is not accepted for target gridlines';
      return null;
    }
    if (s.target === 'gapOverlap') {
      if (s.fill !== undefined || s.outline !== undefined) return 'gapOverlap has no fill/outline';
      if (s.axis !== undefined) return 'axis is only accepted for target gridlines';
      if (s.seriesIndex !== undefined) return 'seriesIndex is not accepted for target gapOverlap';
      if (s.order !== undefined) return 'order is not accepted for target gapOverlap';
      if (s.gapWidth === undefined && s.overlap === undefined) return 'at least one of gapWidth/overlap is required';
      return null;
    }
    if (s.target === 'seriesOrder') {
      if (s.fill !== undefined || s.outline !== undefined) return 'seriesOrder has no fill/outline';
      if (s.axis !== undefined) return 'axis is only accepted for target gridlines';
      if (s.seriesIndex !== undefined) return 'seriesIndex is not accepted for target seriesOrder -- order is the whole permutation';
      if (s.gapWidth !== undefined || s.overlap !== undefined) return 'gapWidth/overlap are not accepted for target seriesOrder';
      if (s.order === undefined) return 'order is required for target seriesOrder';
      if (!s.order.every(v => v < s.order.length)) return 'order must contain only integers in 0..length-1';
      if (new Set(s.order).size !== s.order.length) return 'order must not repeat a value (it must be a permutation)';
      return null;
    }
    // X-O05/5-C: markerSymbol/lineStyle are the SAME "own shape, return before the marker
    // fallback" pattern gapOverlap/seriesOrder already established -- both indexed (seriesIndex
    // required), neither has fill/outline/axis/gapWidth/overlap/order.
    if (s.target === 'markerSymbol') {
      if (s.fill !== undefined || s.outline !== undefined) return 'markerSymbol has no fill/outline';
      if (s.axis !== undefined) return 'axis is only accepted for target gridlines';
      if (s.gapWidth !== undefined || s.overlap !== undefined || s.order !== undefined) return 'gapWidth/overlap/order are not accepted for target markerSymbol';
      if (s.dash !== undefined || s.smooth !== undefined) return 'dash/smooth are not accepted for target markerSymbol';
      if (s.seriesIndex === undefined) return 'seriesIndex is required for target markerSymbol';
      if (s.symbol === undefined && s.size === undefined) return 'at least one of symbol/size is required';
      return null;
    }
    if (s.target === 'lineStyle') {
      if (s.fill !== undefined || s.outline !== undefined) return 'lineStyle has no fill/outline';
      if (s.axis !== undefined) return 'axis is only accepted for target gridlines';
      if (s.gapWidth !== undefined || s.overlap !== undefined || s.order !== undefined) return 'gapWidth/overlap/order are not accepted for target lineStyle';
      if (s.symbol !== undefined || s.size !== undefined) return 'symbol/size are not accepted for target lineStyle';
      if (s.seriesIndex === undefined) return 'seriesIndex is required for target lineStyle';
      if (s.dash === undefined && s.smooth === undefined) return 'at least one of dash/smooth is required';
      return null;
    }
    // target === 'marker' (the only remaining enum value after the five returns above)
    if (s.axis !== undefined) return 'axis is only accepted for target gridlines';
    if (s.symbol !== undefined || s.size !== undefined || s.dash !== undefined || s.smooth !== undefined) return 'symbol/size/dash/smooth are only accepted for target markerSymbol/lineStyle';
    if (s.seriesIndex === undefined) return 'seriesIndex is required for target marker';
    if (s.fill === undefined && s.outline === undefined) return 'at least one of fill/outline is required';
    return null;
  },
  // L-04: the column pair (the family's own gate carries every other refusal, named)
  // the same L-21 sheet-name rule as insert_rows / delete_rows (this release merged after L-21 retired SHEET_BARE here)
  insert_columns: s => sheetNameProblem(s.sheet),
  delete_columns: s => sheetNameProblem(s.sheet),
  // L-05: the family's own gate carries every other refusal, named (a sheet a formula names, an unreadable dependent axis)
  delete_sheet: s => sheetNameProblem(s.sheet),
  // L-05/2: the family's own gate carries the rest, named (xlsx-move-reference-not-found, a sheet that is not there)
  move_sheet(s) {
    const why = sheetNameProblem(s.sheet);
    if (why) return why;
    if ((s.before !== undefined) === (s.after !== undefined)) return 'exactly one of before / after must name the reference sheet';
    const ref = s.before !== undefined ? s.before : s.after;
    return sheetNameProblem(ref) || (ref.toLowerCase() === s.sheet.toLowerCase() ? 'a sheet cannot be moved relative to itself' : null);
  },
  // L-07: the family's own gate carries the rest, named (a chart of that name with another identity, a chart that is not there, an unmeasured data shape)
  set_chart(s) {
    const why = bareNameProblem(s.sheet) || rangeProblem(s.range);
    if (why) return why;
    const g = parseRange(s.range);
    return g.r2 - g.r1 < 1 || g.c2 - g.c1 < 1 ? 'range requires a header row, a category column and data' : null;
  },
  set_chart_legend: s => sheetNameProblem(s.sheet),
  // L-06: the family's own gate carries the rest, named (an overlap on the same sheet, an occupied target, a merged cell, a formula that leaves the sheet, a dependent reaching a block)
  copy_range: s => sheetNameProblem(s.sheet) || (s.targetSheet === undefined ? null : sheetNameProblem(s.targetSheet)) || rangeProblem(s.range),
  // X-O07/4b: the family's own plan()/intentProblem() carries every other refusal (source
  // ambiguity, dest==source, at bounds) -- this re-checks both sheet names at the SAME
  // "alone" planner layer insert_columns/delete_columns above already use.
  copy_chart: s => sheetNameProblem(s.sheet) || sheetNameProblem(s.sourceSheet),
  // L-05/3: the family's own gate carries the rest, named (a target name that exists and differs, an image / shape source, a pivot, an unproven runtime)
  copy_sheet: s => sheetNameProblem(s.sheet) || sheetNameProblem(s.name)
    || (s.name.toLowerCase() === s.sheet.toLowerCase() ? 'the copy name equals the source name, case-insensitively' : null),
  // L-06/b: the SAME shape as copy_range's own constraint -- the family's own gate/classify() carry
  // every other refusal (overlap, occupied target, merged cell, a formula that leaves the sheet, a
  // dependent reaching a block, a blank source with nothing to build a fresh expectation from).
  move_range: s => sheetNameProblem(s.sheet) || (s.targetSheet === undefined ? null : sheetNameProblem(s.targetSheet)) || rangeProblem(s.range),
  // X-O07/3: no further cross-field shape to check -- the ambiguous/not-found/runtime-gate
  // checks are all live-dependent (charts.cjs's own classifyDelete), same split plan()/classify()
  // already uses throughout this family.
  delete_chart: s => bareNameProblem(s.sheet),
  // L-11/2: only the range shape here; the family's plan() (which validateWriteRequest runs next) owns the cross-field rules -- a header row and a data row,
  // `field` inside the range columns, exactly one of value / values / and, plain values -- so they are not duplicated.
  filter_range: s => rangeProblem(s.range),
  // GH #138 U1: the scope (a sheet name) only; the family's plan() owns the name grammar, the reserved names and the ABSOLUTE-reference rule
  set_defined_name: s => ('scope' in s ? bareNameProblem(s.scope) : null),
  delete_defined_name: s => ('scope' in s ? bareNameProblem(s.scope) : null),
  // L-18/1: only the sheet/range shape here; the family's own plan() (the "alone" planner call) owns the cross-field
  // rules -- shift enum, anchor.column/row matching shift, anchor within the block's own band, anchor.value shape.
  delete_cells: s => sheetNameProblem(s.sheet) || rangeProblem(s.range),
  // GH #138 U2: at least one of rows/columns for the set (freeze_panes' own CONSTRAINTS pattern); clear takes sheet only
  set_print_titles: s => (s.rows === undefined && s.columns === undefined ? 'requires at least one of rows or columns' : sheetNameProblem(s.sheet)),
  clear_print_titles: s => sheetNameProblem(s.sheet),
  // GH #138 U3: at least one field; the family's plan() re-checks the values
  set_page_layout: s => (['orientation', 'printGridlines', 'printHeadings', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin'].some(k => s[k] !== undefined) ? sheetNameProblem(s.sheet) : 'requires orientation, printGridlines, printHeadings and / or a margin'),
  // GH #139 U2: sheet only -- the schema's { col, row } shape already rules out a non-integer / negative
  // coordinate (copy_chart's own CONSTRAINTS does not re-check `at` either, for the same reason).
  move_chart: s => sheetNameProblem(s.sheet),
  // GH #139 U3: sheet only -- the schema's enum/pattern/{col,row}/{width,height} shapes already rule out an
  // unsupported geometry, a malformed colour or a non-positive size is left to the family's own plan()
  // (sizeProblem), the SAME split move_chart's CONSTRAINTS uses for `at`.
  insert_shape: s => sheetNameProblem(s.sheet),
  insert_image: s => sheetNameProblem(s.sheet),
  // GH #137 U2: sheet only -- the schema's integer/enum/HEX_COLOR shapes already rule out a malformed run;
  // overlap between runs, an out-of-text-bounds run and an empty font object are all the family's own
  // plan()/gate() territory (the "alone" planner call), never re-derived here.
  format_run: s => sheetNameProblem(s.sheet),
  // GH #138 U4: sheet only for 7 of the 8 -- the schema's enum/integer/const shapes already require every
  // other field together where the family needs them (set_print_fit's both numbers, set_header_footer's
  // slot+value, set_page_break's axis+at, set_first_page_number's value, set_print_scale's scale,
  // set_paper_size's paper, set_header_footer_flag's flag+value). The axis<->at type correlation and the
  // exact-mm paper lookup stay the family's own plan()/pageWant() territory, never re-derived here.
  set_print_fit: s => sheetNameProblem(s.sheet),
  set_header_footer: s => sheetNameProblem(s.sheet),
  set_page_break: s => sheetNameProblem(s.sheet),
  set_first_page_number: s => sheetNameProblem(s.sheet),
  // set_print_centering: exactly one of horizontal/vertical -- the ONE correlation this family's shape needs
  // that the other 7 don't (centerWant()'s own "give exactly one of" refusal, mirrored here at REFUSED_SHAPE
  // so a zero- or both-given step fails before any dispatch, not just at the family's own gate).
  set_print_centering: s => sheetNameProblem(s.sheet) || (['horizontal', 'vertical'].filter(k => s[k] !== undefined).length !== 1 ? 'set_print_centering requires exactly one of horizontal / vertical' : null),
  set_print_scale: s => sheetNameProblem(s.sheet),
  set_paper_size: s => sheetNameProblem(s.sheet),
  set_header_footer_flag: s => sheetNameProblem(s.sheet),
  // GH #140 U4: the range shape only; the family's plan() owns the header / data-row / column-count rules, and its classify() the pivot lookup + the header match
  refresh_pivot: s => rangeProblem(s.sourceRange),
  // GH #140 U3: sheet AND pivotSheet (two sheet-carrying fields, unlike every other entry here) -- the schema's
  // string/pattern shapes already rule out a malformed name/range/field/cell; the family's own plan() (PLAN_CHECKS)
  // owns the cross-field correlations (three distinct header names, the "needs no quoting" sheet grammar, the
  // range's row/column-count shape, `at` being exactly one cell).
  create_pivot: s => sheetNameProblem(s.sheet) || sheetNameProblem(s.pivotSheet),
  // GH #301 D4: the sheet/range shape here; the family's own plan() owns the area size and the formula rules
  write_array_formula: s => rangeProblem(s.range),
  // GH #300 K2: the family's own plan() owns the text rules (printable + line breaks, the ceiling)
  add_comment: s => sheetNameProblem(s.sheet),
  set_shape_text: s => sheetNameProblem(s.sheet),
  set_chart_point_data_label: s => sheetNameProblem(s.sheet),
  set_shape_rotation: s => sheetNameProblem(s.sheet),
  set_sheet_visibility: s => sheetNameProblem(s.sheet),
});

// a step that changes the coordinates / sheet graph the other steps would be resolved against: the only step.
// L-19 C (owner-decision C, 2026-10-17): delete_rows is the one exception -- it may repeat within a task, each
// step's anchor read against the state the task's own earlier steps already produced (adapter.cjs planOperations
// carries the same homogeneity rule; the BATCHED_REPEATABLE check right below refuses any OTHER mix).
const SOLE_STEP = Object.freeze(['insert_rows', 'insert_columns', 'delete_columns', 'rename_sheet', 'create_sheet', 'delete_sheet', 'move_sheet', 'copy_range', 'copy_sheet', 'move_range', 'delete_cells', 'write_array_formula', 'add_comment', 'set_sheet_visibility']);

// the only intent allowed to repeat within one task (L-19 C): every OTHER step of a multi-step task must be the
// same intent, or the task is refused -- delete_rows may not share a task with any other intent, including itself
// mixed with a structural sibling (insert_rows etc., already caught by SOLE_STEP above).
const BATCHED_REPEATABLE = Object.freeze(['delete_rows']);

// ---- a minimal validator for the schema subset above ----

const typeOf = v => (Array.isArray(v) ? 'array' : v === null ? 'null' : typeof v);
const isType = (v, t) => (t === 'integer' ? Number.isInteger(v) : t === 'number' ? typeof v === 'number' && Number.isFinite(v) : typeOf(v) === t);

/** -> null | { code, path, reason }. */
function check(schema, v, path) {
  if (schema.anyOf) return schema.anyOf.some(s => check(s, v, path) === null) ? null : { code: 'REFUSED_SHAPE', path, reason: `${path}: unsupported value` };
  if ('const' in schema) return v === schema.const ? null : { code: 'REFUSED_SHAPE', path, reason: `${path} must be ${schema.const}` };
  if (schema.enum) return schema.enum.includes(v) ? null : { code: 'REFUSED_SHAPE', path, reason: `${path} must be one of ${schema.enum.join(', ')}` };
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

const TASK_SCHEMA_KEYS = Object.freeze(['id', 'steps', 'verifyStoredContent']);
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
  // GH #76 MODEL 1 (owner decision 2026-10-18): opt-in, task-level, default false -- see docs/GH76-VERIFY-STORED-CONTENT-OPTION-DESIGN.md
  if (task.verifyStoredContent !== undefined && typeof task.verifyStoredContent !== 'boolean') return refuse('MALFORMED_TASK', 'task.verifyStoredContent must be a boolean');
  if (!Array.isArray(task.steps)) return refuse('MALFORMED_TASK', 'task.steps must be an array');
  if (task.steps.length === 0) return refuse('EMPTY_TASK', 'task.steps is empty');
  if (task.steps.length > MAX_STEPS) return refuse('LIMIT', `task has ${task.steps.length} steps, over the ${MAX_STEPS} release bound`);
  for (const [i, s] of task.steps.entries()) {
    if (!s || typeof s !== 'object' || Array.isArray(s)) return refuse('MALFORMED_TASK', `step ${i} is not an object`, { stepIndex: i });
    const cls = intentRefusal(s.intent, i);
    if (cls) return cls;
    if (s.intent === 'format_range' && s.format && typeof s.format === 'object' && !Array.isArray(s.format)) {
      const held = Object.keys(s.format).filter(k => !FORMAT_RANGE_RELEASED_KEYS.includes(k));
      if (held.length) return refuse('REFUSED_TEMPORARY_BOUNDARY', `step ${i}: format_range -- format.${held.join(', format.')} is not released on the product surface (format_range is released with format.numberFormat, format.border, format.fill, format.font, format.align and format.wrap only; the rest of format_range stays behind the release boundary ${TEMPORARY_BOUNDARY.format_range})`, { stepIndex: i, intent: s.intent });
    }
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
  if (task.steps.length > 1) {
    const batched = task.steps.findIndex(s => BATCHED_REPEATABLE.includes(s.intent));
    if (batched >= 0 && task.steps.some(s => s.intent !== task.steps[batched].intent)) {
      return refuse('REFUSED_COMPOSITION', `step ${batched}: ${task.steps[batched].intent} may only share a task with other ${task.steps[batched].intent} steps`, { stepIndex: batched, intent: task.steps[batched].intent });
    }
  }
  // the canonical planner, pre-dispatch (pure: no session) -- the same call the executor makes in-session
  const planned = xlsxAdapter.planOperations(task);
  if (!planned || !planned.ok) return refuse('REFUSED_COMPOSITION', (planned && planned.detail) || 'the planner refused the task', planned && planned.index !== undefined ? { stepIndex: planned.index } : {});
  const validity = validatePlan(planned.operations);
  if (!validity.ok) return refuse('REFUSED_COMPOSITION', `plan invalid: ${validity.problems.join('; ')}`);
  return Object.freeze({ ok: true, fileId: args.fileId, task: toExecutorTask(task), plannedOperations: planned.operations.length });
}

/** The canonical executor task, verbatim (a deep copy, frozen): no second task language. */
function toExecutorTask(task) {
  const copy = JSON.parse(JSON.stringify({ ...(task.id !== undefined ? { id: task.id } : {}), steps: task.steps, ...(task.verifyStoredContent !== undefined ? { verifyStoredContent: task.verifyStoredContent } : {}) }));
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
        items: { oneOf: XLSX_WRITE_ALLOWLIST.map(i => STEP_SCHEMAS[i]) } },
        // GH #76 MODEL 1 (owner decision 2026-10-18): opt-in post-close stored-content verification, default false
        verifyStoredContent: { type: 'boolean', description: 'after a confirmed save, verify the stored content via a fresh, independent read before returning (default false; adds ~11-13.5s)' } },
        required: ['steps'], additionalProperties: false },
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
  // X-M01: an EXECUTED task that did not succeed always names why (operations[].failure carries the per-step reason)
  OPERATION_FAILED: 'OPERATION_FAILED', // a step stopped the task after the session opened (the applied prefix is in operations[])
  REFUSED_PRE_MUTATION: 'REFUSED_PRE_MUTATION', // a step refused on the planning observation, before ANY dispatch (dispatch 0)
  EXECUTION_FAILED: 'EXECUTION_FAILED', // no single step failed (plan / observation / persistence stage): `reason` = the outcome
  // GH #76 MODEL 1 (owner decision 2026-10-18): verifyStoredContent:true read back the stored file and it did not
  // hold the requested post-state -- a completed proof, never reported as success (decision 3)
  STORED_CONTENT_MISMATCH: 'STORED_CONTENT_MISMATCH',
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

// X-M01 (owner: an empty or uninterpretable failure reason is not acceptable): every operation that did not end APPLIED /
// SATISFIED carries a `failure` with a stable, non-empty `reason` code taken from the canonical register's own evidence
// (the classifier / resolver / authority gate / verifier that stopped it) -- never a fresh guess here.
const SCALAR = v => v === null || ['string', 'number', 'boolean'].includes(typeof v);
const pickScalars = (o, keys) => Object.fromEntries(keys.filter(k => o && SCALAR(o[k]) && o[k] !== null && o[k] !== '').map(k => [k, o[k]]));
const nonEmpty = (...xs) => xs.find(x => typeof x === 'string' && x !== '') || null;
// GH #96: an xlsx-observation-unreadable-getters refusal names WHICH reader failed only in the observation's
// `unsupported` list ([{ field, reason }]) -- not a scalar, so pickScalars dropped it. Carried bounded: at most 10
// entries, scalar field / reason only, each reason cut to 160 characters, plus the total when it was cut.
const MAX_UNSUPPORTED = 10;
function unsupportedOf(o) {
  const u = o && Array.isArray(o.unsupported) ? o.unsupported.filter(x => x && typeof x === 'object') : [];
  if (!u.length) return {};
  const list = u.slice(0, MAX_UNSUPPORTED).map(x => ({ field: String(x.field === undefined ? '' : x.field), reason: String(x.reason === undefined ? '' : x.reason).slice(0, 160) }));
  return { unsupported: list, ...(u.length > MAX_UNSUPPORTED ? { unsupportedTotal: u.length } : {}) };
}

// L-18/1 review: the SAME gap as GH
// #96's `unsupported`, a different shape -- a classification refusal's `unobservable` list (dependencies.cjs's
// own `{ axis, why, host }` entries, already carried by at least two families -- delete-cells.cjs's
// dependentGuard and charts.cjs's drawing/chart-source guards) is an ARRAY, so pickScalars silently dropped
// it too; a caller only ever saw the bare reason code, never WHICH axis was unobservable. Carried bounded,
// the same way: at most 10 entries, scalar fields only, `why` cut to 160 characters.
const MAX_UNOBSERVABLE = 10;
function unobservableOf(o) {
  const u = o && Array.isArray(o.unobservable) ? o.unobservable.filter(x => x && typeof x === 'object') : [];
  if (!u.length) return {};
  const list = u.slice(0, MAX_UNOBSERVABLE).map(x => ({ axis: String(x.axis === undefined ? '' : x.axis), why: String(x.why === undefined ? '' : x.why).slice(0, 160), ...(x.host !== undefined && x.host !== null && x.host !== '' ? { host: String(x.host) } : {}) }));
  return { unobservable: list, ...(u.length > MAX_UNOBSERVABLE ? { unobservableTotal: u.length } : {}) };
}

function failureOf(e, verdicts) {
  const ev = e.evidence || {};
  if (e.status === 'SKIPPED_DUE_TO_PRIOR_FAILURE') return { stage: 'NOT_ATTEMPTED', reason: 'skipped-due-to-prior-failure', ...(ev.stoppedAt ? { stoppedAt: ev.stoppedAt } : {}) };
  if (e.status === 'DISPATCHED_OUTCOME_UNKNOWN') return { stage: 'DISPATCH', reason: nonEmpty(e.uncertainty && e.uncertainty.reason, 'dispatched-outcome-unknown') };
  const inner = ev.evidence && typeof ev.evidence === 'object' ? ev.evidence : {};
  switch (ev.stage) {
    case 'classification':
      return { stage: 'CLASSIFICATION', reason: nonEmpty(inner.reason, inner.threw ? 'classifier-threw' : null, `classification-${String(ev.state || 'fail-closed').toLowerCase().replace(/_/g, '-')}`),
        ...pickScalars(inner, ['cell', 'mergeArea', 'formula', 'detail', 'why']), ...unsupportedOf(inner), ...unobservableOf(inner) };
    case 'resolve': return { stage: 'TARGET_RESOLUTION', reason: nonEmpty(ev.ambiguity, 'target-resolution-failed'), ...pickScalars(ev, ['detail']) };
    case 'targeting-authority': return { stage: 'TARGETING_AUTHORITY', reason: nonEmpty(ev.gate && ev.gate.outcome, 'targeting-authority-failed') };
    case 'targeting-observation': return { stage: 'TARGETING_OBSERVATION', reason: nonEmpty(ev.outcome, 'targeting-observation-failed') };
    case 'ref': return { stage: 'REFERENCE', reason: nonEmpty(ev.reason, 'reference-unresolved') };
    case 'plan-validation': return { stage: 'PLAN', reason: 'plan-invalid' };
    case 'preflight': return { stage: 'PREFLIGHT', reason: nonEmpty(inner.reason, `preflight-${String(ev.state || 'fail-closed').toLowerCase().replace(/_/g, '-')}`),
      ...pickScalars(inner, ['cell', 'mergeArea', 'formula', 'detail', 'why']), ...unsupportedOf(inner), ...(ev.blocking ? { dispatched: false } : {}) };
    default: break;
  }
  if (ev.threw) return { stage: 'DISPATCH', reason: 'dispatch-threw', ...pickScalars(ev, ['detail']) };
  if (verdicts.l1 && verdicts.l1 !== 'PASS') return { stage: 'VERIFICATION', reason: `l1-${String(verdicts.l1).toLowerCase()}` };
  if (verdicts.l3 && verdicts.l3 !== 'PASS') return { stage: 'PRESERVATION', reason: `l3-${String(verdicts.l3).toLowerCase()}` };
  return { stage: 'DISPATCH', reason: nonEmpty(ev.reason, ev.outcome, ev.code, `operation-${String(e.status).toLowerCase()}`), ...pickScalars(ev, ['detail']) };
}

/**
 * GH #76 MODEL 1 (owner decision 2026-10-18): the EXPECTED post-state for the stored-content verify,
 * derived from the task's own write_range steps -- the dominant family in the design's intent->read
 * mapping (docs/GH76-VERIFY-STORED-CONTENT-OPTION-DESIGN.md section 6). Every OTHER intent in the task
 * is named in `notProven`, never guessed at: this is the one family this unit implements a comparator
 * for, not a generic comparator for all 56 write intents. A typed value this unit does not itself
 * decode (anything but a plain string/number/{text}) is likewise named, not silently treated as
 * verified. { bySheet: { sheet: { ref: {value?, formula?, text?} } }, notProven: [...] }
 */
function expectedCellsFromSteps(task) {
  const bySheet = {};
  const notProven = [];
  for (const step of (task && task.steps) || []) {
    if (step.intent !== 'write_range') { notProven.push(`${step.intent}-verify-not-implemented`); continue; }
    const g = parseRange(step.range);
    if (!g) continue; // already refused pre-dispatch (validateWriteRequest) -- defensive only
    const addresses = cellsOf(g);
    const vals = step.values;
    const fs = step.formulas;
    const sheetMap = bySheet[step.sheet] || (bySheet[step.sheet] = {});
    for (let r = 0; r < g.rows; r += 1) {
      for (let c = 0; c < g.cols; c += 1) {
        const a = addresses[r * g.cols + c];
        const f = fs ? fs[r][c] : null;
        const v = vals ? vals[r][c] : null;
        if (f !== null && f !== undefined) { sheetMap[a] = { formula: f }; continue; }
        if (v === null || v === undefined) continue;
        if (typeof v === 'number' || typeof v === 'string') { sheetMap[a] = { value: v }; continue; }
        if (v && typeof v === 'object' && typeof v.text === 'string') { sheetMap[a] = { value: v.text, text: true }; continue; }
        notProven.push(`write_range-typed-value-not-verified:${step.sheet}!${a}`);
      }
    }
  }
  return { bySheet, notProven };
}

/** One A1 range string covering every ref in `refs` on one sheet (a bounding box -- read_range takes one range, not a scattered cell list). */
function rangeCovering(refs) {
  const cells = refs.map(parseCell).filter(Boolean);
  const c1 = Math.min(...cells.map(x => x.c));
  const c2 = Math.max(...cells.map(x => x.c));
  const r1 = Math.min(...cells.map(x => x.r));
  const r2 = Math.max(...cells.map(x => x.r));
  return c1 === c2 && r1 === r2 ? addr(c1, r1) : `${addr(c1, r1)}:${addr(c2, r2)}`;
}

/**
 * GH #76 MODEL 1: the bounded wait for a document-key rotation away from the writer's own key, using
 * the PRODUCT'S OWN config source. The release gate (test/live/stored-content-gate.cjs) uses a separate
 * HTTP transport against the broker for the SAME rule -- two transports, one rule (design section 8,
 * decision 1), not two implementations of the rule itself.
 */
async function awaitKeyRotation(configSource, fileId, preKey, { timeoutMs = 15000, pollMs = 500, now = Date.now, wait = ms => new Promise(res => setTimeout(res, ms)) } = {}) {
  const t0 = now();
  for (;;) {
    let key = null;
    try {
      const cfg = await configSource.getEditorConfig({ fileId, operation: 'WRITE' });
      key = cfg && cfg.document ? cfg.document.key : null;
    } catch { key = null; }
    if (key && preKey && key !== preKey) return { rotated: true, ms: now() - t0 };
    if (now() - t0 >= timeoutMs) return { rotated: false, timedOut: true, ms: now() - t0 };
    await wait(pollMs);
  }
}

/**
 * GH #76 MODEL 1 (owner decision 2026-10-18): the optional post-close stored-content verify. Called
 * AFTER the write's own session has already closed (section 2a: never while held) and with its own,
 * separate lease on the SAME shared gate (section 2b: a second, independent driver, never registered
 * with the #100 registry). Returns the `persistence.storedContent` object to merge into the receipt,
 * or `null` when verification was not requested at all -- then the field stays entirely absent
 * (section 1: "no behavior change at all" for a call that does not opt in).
 */
async function runStoredContentVerify({ v, r, preKey, configSource, createDriver, loadPlaywright, gate, now = Date.now, keyRotation = {} }) {
  if (v.task.verifyStoredContent !== true) return null;
  const preservationScope = (r.preservation && r.preservation.scope) || null;
  const inheritedNotProven = (r.preservation && r.preservation.notProven) || [];
  const unresolved = (status, failureReason, overheadMs, extraNotProven = []) =>
    ({ status, scope: 'requested-post-state', preservationScope, notProven: [...inheritedNotProven, ...extraNotProven], overheadMs, failureReason });
  // requested, but nothing was saved under a confirmed forceSave (a failed write, or a write-free SATISFIED retry):
  // there is no new stored version to verify -- named, never VERIFIED and never a MISMATCH
  if (r.persistence.state !== 'FORCESAVE_CONFIRMED') return unresolved('NOT_APPLICABLE', `no-confirmed-save:${r.persistence.state || 'none'}`, 0);
  const t0 = now();
  const { bySheet, notProven: skipped } = expectedCellsFromSteps(v.task);
  const sheets = Object.keys(bySheet);
  if (sheets.length === 0) return unresolved('UNCERTAIN', 'read-incomplete', now() - t0, skipped);
  const rotation = await awaitKeyRotation(configSource, v.fileId, preKey, keyRotation);
  if (!rotation.rotated) return unresolved('UNCERTAIN', 'key-not-rotated', now() - t0, skipped);
  const lease = gate.tryEnter();
  if (!lease) return unresolved('UNCERTAIN', 'open-failed', now() - t0, skipped);
  let results = []; let incomplete = false;
  try {
    const driver = createDriver('xlsx', { configSource, loadPlaywright });
    lease.track(driver);
    const opened = (await driver.open({ fileId: v.fileId, operation: 'READ' })) || {};
    if (!opened.ok) return unresolved('UNCERTAIN', 'open-failed', now() - t0, skipped);
    try {
      for (const sheet of sheets) {
        const vr = validateReadRequest('xlsx', { fileId: v.fileId, operation: 'read_range', sheet, range: rangeCovering(Object.keys(bySheet[sheet])) });
        if (!vr.ok) { incomplete = true; continue; }
        const out = await readOnce(FORMATS.xlsx, driver, vr, 'xlsx');
        if (!out.ok) { incomplete = true; continue; }
        results.push(out);
      }
    } finally {
      await driver.close().catch(() => {});
    }
  } finally {
    lease.release();
  }
  const overheadMs = now() - t0;
  if (incomplete || results.length === 0) return unresolved('UNCERTAIN', 'read-incomplete', overheadMs, skipped);
  const pc = productCells(results);
  const diffs = [];
  for (const sheet of sheets) for (const [ref, expected] of Object.entries(bySheet[sheet])) diffs.push(...cellDiffs(ref, expected, pc[ref], 'product'));
  if (diffs.length) return { status: 'MISMATCH', scope: 'requested-post-state', preservationScope, notProven: [...inheritedNotProven, ...skipped, ...diffs.map(x => `stored-poststate-mismatch:${x}`)], overheadMs, failureReason: null };
  return { status: 'VERIFIED', scope: 'requested-post-state', preservationScope, notProven: [...inheritedNotProven, ...skipped], overheadMs, failureReason: null };
}

/** Merges a runStoredContentVerify() result into an already-built mapExecutionResult() receipt (additive;
 * MISMATCH is the one case that also flips `ok`/`code` -- decision 3: a completed proof the stored file
 * does not hold the requested post-state is never reported as success). `sv === null` changes nothing. */
function applyStoredContentVerify(receipt, sv) {
  if (!sv) return receipt;
  const persistence = { ...receipt.persistence, storedContent: sv, storedContentVerified: sv.status === 'VERIFIED' };
  if (sv.status === 'MISMATCH') {
    return { ...receipt, ok: false, code: WRITE_ERRORS.STORED_CONTENT_MISMATCH, reason: `stored content does not hold the requested post-state: ${sv.notProven.filter(x => x.startsWith('stored-poststate-mismatch:')).join(', ')}`, persistence };
  }
  return { ...receipt, persistence };
}

/** mapExecutionResult(final, v) -> the MCP result of an executed task (canonical statuses only). */
/** X-N01b: a step's L3 scope as its adapter reported it ({ scope, notProven? }), or null when it reported none. */
function l3Scope(list, id) {
  const v = (list || []).find(x => x.evidence && x.evidence.operationId === id);
  const ev = v && v.evidence && v.evidence.evidence;
  if (!ev || (ev.preservationScope !== 'FULL' && ev.preservationScope !== 'BOUNDED')) return null;
  return ev.preservationScope === 'BOUNDED' ? { scope: 'BOUNDED', proven: ev.proven || [], notProven: ev.notProven || [] } : { scope: 'FULL' };
}

/** X-O08-7 v2 (16fd90b9): a step's L2/whole-task `notProven` ([field, ...]), or null when its
 * check reported none, or no check opted in with an `operationId` (most families do not --
 * `notProven` is an L1/write-semantics axis, DIFFERENT from l3Scope's L3/preservation axis above;
 * the two must never be merged into one list, a FAIL on one says nothing about the other). A step
 * with multiple whole-task checks (should not happen today) unions every matching check's list. */
function semanticsNotProven(checks, id) {
  const matches = (checks || []).filter(c => c.operationId === id && Array.isArray(c.notProven) && c.notProven.length);
  if (!matches.length) return null;
  return [...new Set(matches.flatMap(c => c.notProven))];
}

/**
 * GH #293 L7: the items an operation names -- `affectedItems` (a done list-scope numbering write: every item it changed, or covers on a
 * no-op) and `conflictingItems` (a refused write: the other items it would have changed -- co-members, another list on the definition,
 * a label quoting the level). Strings only, as the family named them; absent when the family named none.
 */
function namedItemsOf(e, done) {
  const ev = e.evidence && typeof e.evidence === 'object' ? e.evidence : {};
  const strs = v => (Array.isArray(v) && v.every(x => typeof x === 'string') ? v : null);
  if (done) {
    const a = strs(ev.affectedItems) || strs(ev.evidence && ev.evidence.affectedItems);
    return a ? { affectedItems: a } : {};
  }
  const inner = ev.evidence && typeof ev.evidence === 'object' ? ev.evidence : {};
  const c = [...new Set([...(strs(inner.coMembers) || []), ...(strs(inner.otherListItems) || []), ...(strs(inner.referencedBy) || [])])];
  return c.length ? { conflictingItems: c } : {};
}

function mapExecutionResult(final, v, toolName = 'xlsx_execute_task') {
  const base = { tool: toolName, fileId: v.fileId, taskId: v.task.id === undefined ? null : v.task.id };
  if (final.openFailure) {
    return { ok: false, stage: 'OPEN', code: mapOpenFailure(final.openFailure), reason: final.openFailure.detail || 'open failed', dispatched: false, ...base };
  }
  const exec = final.execution || {};
  const reason = final.saveEvidence ? final.saveEvidence.reason : null;
  const rejected = storageRejected(reason);
  const layer = (list, id) => { const v = (list || []).find(x => x.evidence && x.evidence.operationId === id); return v ? v.verdict : null; };
  const checks = exec.wholeTaskVerification && exec.wholeTaskVerification.evidence && exec.wholeTaskVerification.evidence.checks;
  const operations = (exec.appliedOperations || []).map(e => {
    const l1 = layer(exec.primitiveVerification, e.operationId);
    const l3 = layer(exec.preservationVerification, e.operationId);
    const scope = l3Scope(exec.preservationVerification, e.operationId);
    // X-O08-7 v2: L1/write-semantics gaps (e.g. a shape's fill/stroke colour, written but never
    // publicly readable back) -- additive, NEVER merged into l3Scope's notProven above (L3/
    // preservation is a different axis: "did everything ELSE stay the same", not "was this field's
    // OWN write verified"). APPLIED/SATISFIED is unaffected -- this names a limit, not a failure.
    const semantics = semanticsNotProven(checks, e.operationId);
    const done = e.status === 'APPLIED' || e.status === 'SATISFIED';
    const named = namedItemsOf(e, done);
    return { id: e.operationId, status: e.status, verification: e.verification || null, l1, l3, ...(scope ? { l3Scope: scope } : {}), ...named,
      ...(semantics ? { semanticsNotProven: semantics } : {}),
      ...(e.uncertainty && e.uncertainty.reason ? { uncertainty: e.uncertainty.reason } : {}),
      ...(done ? {} : { failure: failureOf(e, { l1, l3 }) }) };
  });
  const ok = final.ok === true && !rejected;
  // X-N01b: the task's L3 proof scope -- BOUNDED if ANY step's L3 is bounded (never FULL then), FULL if every step's that
  // reported one is full, null when no step reported a scope (e.g. a write-free retry)
  const scopes = operations.map(o => o.l3Scope).filter(Boolean);
  const preservationScope = scopes.some(x => x.scope === 'BOUNDED') ? 'BOUNDED' : scopes.length ? 'FULL' : null;
  const notProven = [...new Set(scopes.flatMap(x => x.notProven || []))];
  const firstFailed = operations.find(o => o.failure && o.status !== 'SKIPPED_DUE_TO_PRIOR_FAILURE');
  const failCode = rejected ? WRITE_ERRORS.STORAGE_WRITE_REJECTED
    : ok ? null
      : firstFailed ? (firstFailed.failure.stage === 'PREFLIGHT' ? WRITE_ERRORS.REFUSED_PRE_MUTATION : WRITE_ERRORS.OPERATION_FAILED)
        : WRITE_ERRORS.EXECUTION_FAILED;
  const failReason = ok ? null
    : firstFailed ? `operation ${firstFailed.id}: ${firstFailed.failure.reason}` : nonEmpty(final.outcome, final.persistenceOutcome, reason, 'execution-failed');
  return {
    ok,
    stage: 'EXECUTED',
    ...(failCode ? { code: failCode, reason: failReason } : {}),
    ...base,
    outcome: final.outcome,
    editOk: final.editOk === true,
    documentOutcome: final.documentOutcome,
    mutationCertainty: final.mutationCertainty,
    authority: final.authority,
    noOp: final.noOp === true,
    operations,
    wholeTask: exec.wholeTaskVerification ? exec.wholeTaskVerification.verdict : null, // L2
    preservation: { scope: preservationScope, ...(preservationScope === 'BOUNDED' ? { notProven } : {}) }, // L3 scope (X-N01b)
    operationAccounting: exec.operationAccounting || null,
    persistence: {
      outcome: final.persistenceOutcome,
      state: final.persistenceState,
      reason,
      // UNCERTAIN_MAY_PERSIST only: what the state means for the caller (the edit may become durable on its own -- re-read the document)
      ...(final.persistenceState === 'UNCERTAIN_MAY_PERSIST' && final.saveEvidence ? { meaning: final.saveEvidence.meaning } : {}),
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
 *   -> withSession(xlsxAdapter, createExecutor(task, batch)): the canonical lifecycle, executor, verification and
 *      persistence barrier -- nothing here decides APPLIED, saves or retries
 *   -> mapExecutionResult; the lease is released on every exit.
 * `runtime` / `loadPlaywright` / `createDriver` / `session` (the lifecycle's own ready / persistence / now options) are
 * test seams; the product default is the configured browser runtime and productDriver.
 */
/** GH #76: the phase timings of one write (numbers only -- no content, no ids beyond the fileId), for the log. */
function timingOf(final, v) {
  const m = (final && final.metrics) || {};
  const pick = k => (typeof m[k] === 'number' ? m[k] : null);
  return { fileId: v.fileId, steps: (v.task && v.task.steps || []).length, ok: !!(final && final.ok), persistence: (final && final.persistenceState) || null,
    sessionOpenMs: pick('sessionOpenMs'), readyMs: pick('readyMs'), taskElapsedMs: pick('taskElapsedMs'), totalElapsedMs: pick('totalElapsedMs'),
    freshReads: pick('freshReads'), commands: pick('commands'), persistencePolls: pick('persistencePolls') };
}

/**
 * GH #100 U4b: the warm-session options of ONE write call, or null (then the call is exactly the pre-reuse one: no options.reuse, no config fetch).
 * `reuse` = { registry, principal, configSignature } from the product wiring (one process-wide registry). The key is (file, principal, config
 * signature); the fresh document key comes from the SAME config source the open uses, in the same WRITE mode. A registry `take` re-points the gate
 * lease at the session that is now active, so a shutdown during the call closes the right browser.
 */
function warmSessionOptions(reuse, configSource, fileId, lease) {
  if (!reuse || !reuse.registry || reuse.registry.enabled !== true) return null;
  if (!configSource || typeof configSource.getEditorConfig !== 'function') return null;
  const id = String(fileId);
  const key = sessionKey({ fileId: id, principal: reuse.principal, configSignature: reuse.configSignature });
  if (!key) return null;
  const { registry } = reuse;
  return {
    key,
    decide: decideReuse,
    registry: {
      enabled: true,
      hold: (k, driver, opts) => registry.hold(k, driver, opts),
      take: async k => { const taken = await registry.take(k); if (taken) lease.track(taken.driver); return taken; },
    },
    fetchFreshKey: async () => {
      const config = await configSource.getEditorConfig({ fileId: id, operation: 'WRITE' });
      const k = config && config.document ? config.document.key : null;
      return typeof k === 'string' && k !== '' ? k : null;
    },
  };
}

function createWriteExecution({ configSource, browsersPath, gate = createSessionGate(), runtime = null, loadPlaywright = null, createDriver = productDriver, session = {}, onTiming = null, reuse = null } = {}) {
  const rt = runtime || (loadPlaywright ? { check: () => ({ ok: true }), load: loadPlaywright } : createBrowserRuntime({ browsersPath }));
  return async function execute(v) {
    const verifying = v.task.verifyStoredContent === true;
    const lease = gate.tryEnter();
    if (!lease) return Object.freeze({ ok: false, stage: 'ADMISSION', code: 'BUSY', reason: `an editor session is active (MAX_ACTIVE_SESSIONS_PER_PROCESS=${gate.max}); retry after it closes`, dispatched: false });
    let r; let preKey = null;
    try {
      const ready = rt.check();
      if (!ready.ok) return { ok: false, stage: 'RUNTIME', code: ready.code || 'BROWSER_RUNTIME_UNAVAILABLE', reason: ready.detail, dispatched: false };
      // GH #76 MODEL 1: a verifying call never offers or takes a warm session (design section 2a) -- a second,
      // independent driver for the SAME file would race the #100 registry's one-held-session invariant.
      if (verifying && configSource && typeof configSource.getEditorConfig === 'function') {
        try { const cfg = await configSource.getEditorConfig({ fileId: v.fileId, operation: 'WRITE' }); preKey = cfg && cfg.document ? cfg.document.key : null; } catch { preKey = null; }
      }
      const driver = productWriteDriver(createDriver('xlsx', { configSource, loadPlaywright: rt.load }));
      lease.track(driver);
      const warm = verifying ? null : warmSessionOptions(reuse, configSource, v.fileId, lease);
      const options = { driver, fileId: v.fileId, ...(warm ? { reuse: warm } : {}), ...(session.ready ? { ready: session.ready } : {}), ...(session.persistence ? { persistence: session.persistence } : {}), ...(session.now ? { now: session.now } : {}) };
      const final = await withSession(options, xlsxAdapter, createExecutor(v.task, { strategy: 'batch', observation: { live: liveScope(v.task) } }));
      // GH #76: a timing hook must never change the answer -- it is called after the result exists and its failure is swallowed
      if (onTiming) { try { onTiming(timingOf(final, v)); } catch { /* the log is best-effort */ } }
      r = mapExecutionResult(final, v);
    } finally {
      lease.release();
    }
    // GH #76 MODEL 1: the verify opens its OWN, separate lease on the SAME gate, AFTER the write's session has
    // already closed (design section 2) -- never two sessions on the same file at once.
    const sv = await runStoredContentVerify({ v, r, preKey, configSource, createDriver, loadPlaywright: rt.load, gate, ...(session.keyRotation ? { keyRotation: session.keyRotation } : {}) });
    return applyStoredContentVerify(r, sv);
  };
}

/** createWriteTool('xlsx', { execute, ...createWriteExecution options }) -> { description, inputSchema, handler(args) }. */
function createWriteTool(format, { execute = null, ...wiring } = {}) {
  if (format !== 'xlsx') throw new TypeError(`no WRITE handler for ${String(format)} (DOCX_WRITE_ACCEPTED is empty)`);
  const run = execute || createWriteExecution(wiring);
  return {
    description: `XLSX WRITE task on a Nextcloud file as the calling requester (own credential via the broker): the canonical task { id?, steps }, run by the canonical executor and persisted with a confirmed forcesave. Supported intents: ${XLSX_WRITE_ALLOWLIST.join(', ')}; any other intent or shape is refused before dispatch. Most structural intents (insert/delete rows or columns, and the other identity-changing intents) must be the only step of their task; delete_rows is the one exception and may repeat, provided every step of the task is delete_rows -- give each repeat's at/anchor in the coordinates that will be current once the task's OWN earlier steps have already run, not the task's starting coordinates.`,
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
  mapOpenFailure, createWriteTool, createWriteExecution, productWriteDriver, checkNamesDeep, failureOf, FORMAT_RANGE_RELEASED_KEYS, timingOf,
  // GH #219 2/2 (DOCX WRITE contract): format-neutral primitives, reused as-is by src/server/docx-write-tools.cjs --
  // no behaviour change to any existing xlsx caller (both already existed as module-local functions).
  check, refuse, FILE_ID, ID, TASK_SCHEMA_KEYS,
  // GH #220 1/2 (DOCX WRITE lifecycle): warmSessionOptions is the GH #100 U4b wiring itself -- a format-neutral
  // function of (reuse, configSource, fileId, lease) only, already used unchanged above for the xlsx path.
  // Exporting it changes nothing here; it lets docx-write-tools.cjs build the SAME warm-session options object
  // from its OWN `reuse` wiring parameter instead of re-implementing the key/registry/fetchFreshKey closure.
  warmSessionOptions,
  // GH #76 MODEL 1 DOCX port (owner decision "A", the owner channel, 2026-10-19 -- the GH #76 1:1 port to
  // docx_execute_task, NOT the uncertain-branch extension): both
  // `awaitKeyRotation` and `applyStoredContentVerify` are already format-neutral (configSource.getEditorConfig
  // / a plain persistence-object merge, neither touches xlsx vs docx) -- exporting them lets
  // docx-write-tools.cjs's own `runDocxStoredContentVerify` reuse the SAME polling rule and the SAME receipt-
  // merge function, not a second implementation of either.
  awaitKeyRotation, applyStoredContentVerify,
};
