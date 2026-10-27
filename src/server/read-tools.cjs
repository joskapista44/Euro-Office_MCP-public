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
const { projectCfRead } = require('../adapters/xlsx/families/cf-read.cjs');
const { projectLayoutRead, layoutExtent } = require('../adapters/xlsx/families/layout-read.cjs');
const { projectPanesRead } = require('../adapters/xlsx/families/panes-read.cjs');
const { projectSheetStateRead } = require('../adapters/xlsx/families/sheet-state-read.cjs');
const { projectPageSetupRead } = require('../adapters/xlsx/families/page-read.cjs');
const { projectPrintAreaRead } = require('../adapters/xlsx/families/print-area-read.cjs');
const { projectDefinedNamesRead } = require('../adapters/xlsx/families/defined-names-read.cjs');
const { projectFilterRead } = require('../adapters/xlsx/families/filter-read.cjs');
const { projectCommentsRead } = require('../adapters/xlsx/families/comments-read.cjs');
const docxInventory = require('../adapters/docx/inventory.cjs');
const { parseRange, cellsOf } = require('../adapters/xlsx/a1.cjs');
const { reAnchoredFormula, rangeFirstCellAddr } = require('../adapters/xlsx/families/validation.cjs');
const docxIntents = require('../adapters/docx/intents.cjs');
const { createBrowserRuntime } = require('./browser-runtime.cjs');
const { createSessionGate } = require('./session-gate.cjs');
const { sessionKey } = require('./session-registry.cjs');
const { decideReuse } = require('./session-reuse-rule.cjs');

const STATUS = Object.freeze({ ACCEPTED: 'ACCEPTED', PARTIAL: 'PARTIAL', NOT_MEASURED: 'NOT_MEASURED', INTERNAL: 'INTERNAL' });

const SHEET = Object.freeze({ type: 'string', minLength: 1, maxLength: 31, description: 'worksheet name' });
const RANGE = Object.freeze({ type: 'string', pattern: '^\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6}(:\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6})?$', description: 'A1 range on that sheet, no sheet prefix' });
// GH #137 U1: a single A1 cell (no colon alternative -- unlike RANGE, a run-map is keyed by one cell, never a range)
const CELL = Object.freeze({ type: 'string', pattern: '^\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6}$', description: 'a single A1 cell on that sheet, no sheet prefix, not a range' });
const FLAG = Object.freeze({ type: 'boolean' });
// GH #224b (docs/GH224B-READ-STYLES-BY-NAME-DECISION.md): `names` for read_styles -- 1..32 distinct style names (1..128 printable characters each), read through the driver's name-targeted
// style query (doc.GetStyle(name), the SAME read the write family's own L1 uses). It exists because doc.GetAllStyles() is blind to a style created with CreateStyle (CUSTOM_STYLE_LISTED=false).
const STYLE_NAMES = Object.freeze({ type: 'array', items: Object.freeze({ type: 'string', minLength: 1, maxLength: 128 }), minItems: 1, maxItems: 32, uniqueItems: true,
  description: 'style names to read by name (doc.GetStyle); the result lists exactly these names, in this order' });

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
    read_charts: { status: STATUS.ACCEPTED, args: { sheet: SHEET }, optional: ['sheet'],
      evidence: 'name/type/title/series/legend/axis-titles/dataLabels(chart-level): W7, X-O04; effectiveLabels (per-point, point->series->chart precedence) via chart.dataLabels.effective.read/v1: X-O05/4, status flip conditional on this round\'s LIVE acceptance (test/live/xlsx-xo05-4-chart-effective-labels-read.live.cjs) -- revert to PARTIAL if that run does not PASS; series[].markerStyle (scatter+lineNormal only) + chart-level gridlinesColor via chart.series.markerStyle.read/v1 / chart.axis.majorGridlinesColor.read/v1: X-O04/10, served only on bridge patchVersion /11 (a /5-/10 runtime degrades each to its own \'UNOBSERVABLE:CAPABILITY_MISSING\', never FAIL); series[].order via chart.series.order.read/v1: X-O05/1 LIVE WORKS_PERSISTED_READABLE, served only on bridge patchVersion /12 (a /5-/11 runtime degrades to \'UNOBSERVABLE:CAPABILITY_MISSING\'); chart-level gapOverlap (bar-typed charts only, named reason on every other type) via chart.bar.gapOverlap.read/v1: X-O05/2, LIVE re-measured internal-field read-after-reopen, served only on bridge patchVersion /12; series[].markerSymbol+size (scatter+lineNormal only, named reason \'chart-type-not-in-scope:<type>\' on every other type) via chart.series.markerSymbol.read/v1, and series[].lineStyle dash+smooth (bar+lineNormal only, same named-scope convention) via chart.series.lineStyle.read/v1: GH #139 B /16, LIVE write+session+reopen round trip + a named-fail-closed mechanism-probe fix (a series with no marker object at all now fails closed as \'UNOBSERVABLE:no-marker-object\' rather than a null-shaped ok:true), served only on bridge patchVersion /16 (a /5-/15 runtime degrades each to its own \'UNOBSERVABLE:CAPABILITY_MISSING\')' },
    // GH #140 U2 (#141 roadmap GO): PARTIAL -> ACCEPTED -- its own dedicated projection (name-decision
    // docs/GH140-PIVOT-CONTRACT.md, internal tracker decision: a SEPARATE
    // operation, NOT a read_charts.pivots[] widening, so no existing client's string-list shape changes).
    // Projects inventory.cjs's ALREADY-COMPUTED objects axis (pivot-source.cjs withProvenPivots(), the
    // SAME content-proven source create_pivot/refresh_pivot's own gate() has used internally since M7-H
    // -- this unit adds NO new reader, only a product-surface consumer of an existing one, the SAME
    // convention read_runs (X-F04) and read_charts's effectiveLabels already used). name + source
    // (public when readable, else content-proven by pivot-source.cjs's provenSource(): layout + cached
    // body match the named range's own sums, and no OTHER sheet reproduces it -- never a bare
    // sheet-name match); table/layout ride along ONLY when the public source succeeded (an
    // UNOBSERVABLE source carries neither -- NAMED, not guessed, docs/GH140-PIVOT-CONTRACT.md T2/T3).
    // A sheet whose pivots axis itself is unreadable is its OWN named reason, never an empty list
    // standing in for "not read" (the same rule pivot-source.cjs's module comment already states for
    // create_pivot's consumers). An omitted `sheet` reads every sheet's pivots (the SAME '*' convention
    // read_charts/read_drawings already use).
    read_pivots: { status: STATUS.ACCEPTED, args: { sheet: SHEET }, optional: ['sheet'],
      evidence: 'name + content-proven source via pivot-source.cjs provenSource()/withProvenPivots(): M7-H ACCEPTANCE PASS 6/6 (2026-10-04), X-O10 LIVE re-confirmed 2026-10-16 (internal tracker); table/layout ride along only on a publicly-readable source (named absent otherwise, GH #140 U1 finding); a pivots-axis gap is its own named reason, never an empty list; large-workbook limit: pivot-source.cjs MAX_SOURCE_CELLS=10000 on the content-proof path (the SAME ceiling create_pivot/refresh_pivot already enforce); GH #302 G2: a richer pivot is listed with its row / column / page field names, the Values pseudo-field as a position (its GetIndex is -1 -- the only getter that answers on it), every data field with its function and the style name, session + reopen (docs/evidence/gh302-g2-pivot-read-probe-20261026.json); an unproven source is that pivot\'s own named `source`, the write guards\' pivots axis is unchanged' },
    read_drawings: { status: STATUS.ACCEPTED, args: { sheet: SHEET }, optional: ['sheet'],
      evidence: 'parent/kind/name/pos/width/height/rotation/flip: M7 S2b LIVE-measured PUBLIC getters, a stable workbook-wide identity + L3 preservation multiset (docs/evidence/M7-S2B-DRAWING-PROBE.md); a shape\'s stroke width PUBLIC-readable session+reopen (X-O08-7); a shape\'s text via drawing.text.read/v1 (B14/B16); a shape\'s fill/line colour via drawing.fillLineColor.read/v1 (X-O08, bridge /14) -- degrades to its own named \'UNOBSERVABLE:CAPABILITY_MISSING\' on bridge /5-/13, never guessed. PARTIAL -> ACCEPTED: the X-O09 evidence this status previously carried alone is a single, NARROW, non-cumulative 1-EMU image-height normalization on the first product re-save (docs/evidence/xlsx-xo09-product-preservation-20261012.json) -- a preservation nuance scoped to the image height field specifically, unrelated to parent/kind/name/pos/fill/line (same precedent as read_pivots\'s own M7-H note). GH #302/#325 (G1): a chart or shape whose public size getters read an exact 0 x 0 after a reopen (the stored anchor is right, docs/GH302-PIVOT-CHART-DRAWING-LIVE-MATRIX-20261025.md rows C-03/D-02) now degrades width/height to \'UNOBSERVABLE:size-zero-after-reopen\' on this projection -- never the unproven 0.' },
    read_conditional_formats: { status: STATUS.ACCEPTED, args: { sheet: SHEET, range: RANGE },
      evidence: 'range-level CF reader, the MIRROR of the add_conditional_format rule schema: L-08 unit 1 LIVE l08-r1 (driver and public getters read all 56 writable rows, in-session and after reopen); a rule type / value the writer cannot create is NAMED (type unsupported, unrepresentable), never returned as a value; "No Fill" normalized to null in one place; rendered effect X-O02 NOT_MEASURED' },
    read_layout: { status: STATUS.ACCEPTED, args: { sheet: SHEET, range: RANGE },
      evidence: 'column width / row height / hidden, the MIRROR of the layout_range write schema: L-09 unit 1 (the driver reads them through the public GetEntireColumn / GetEntireRow getters, W4b LIVE; a hidden column / row, an unset row height and any out-of-schema value are NAMED, never a number)' },
    read_freeze_panes: { status: STATUS.ACCEPTED, args: { sheet: SHEET },
      evidence: 'frozen rows / columns of the sheet, the MIRROR of freeze_panes: L-10 (the sheet\'s OWN pane through the DS Bridge sheet.freezeState.read/v1, W4b2 + B06); a split / unknown pane or a bridge gap is NAMED, never "not frozen"' },
    // GH #133: every sheet's visibility and the workbook's active sheet. No new driver reader: the sheetVisibility axis (the public ApiWorksheet.GetVisible(), boolean) and the active-sheet name
    // (the public Api.GetActiveSheet().GetName(), carried by the panes block) already exist and ride the tiered path. The getter is boolean: visible | hidden, a very-hidden sheet is not told apart (not measured).
    read_sheet_state: { status: STATUS.ACCEPTED, args: {},
      evidence: 'sheet visibility (visible | hidden, public ApiWorksheet.GetVisible()) + the active sheet (public Api.GetActiveSheet().GetName()), a field-select over the existing X-L10 sheetVisibility axis and the panes block: GH #133; a gap is NAMED per sheet / for the active sheet, never "visible"; no WRITE intent changes either (a hidden sheet refuses a freeze, X-L10)' },
    read_page_setup: { status: STATUS.ACCEPTED, args: { sheet: SHEET },
      evidence: 'page layout (orientation, gridlines, headings, margins mm) + print settings (fit, scale, paper, header / footer codes, manual page breaks, first page number, centering), the MIRROR of set_page_layout and the set_print_* intents: L-10b LIVE on a PRE-SEEDED fixture (the sheet is set up in the OOXML, not by the product writers, which are not released); the page half (public getters) and the print half (sheet.printSettings.read/v1) fail independently; print area / titles are defined names and NOT in this read' },
    read_filter: { status: STATUS.ACCEPTED, args: { sheet: SHEET },
      evidence: 'AutoFilter range + per-field criteria, the MIRROR of filter_range (value | values | and): L-10b LIVE on a PRE-SEEDED fixture (public AutoFilter getters + sheet.autoFilter.fieldIdentity.read/v1); criteria the writer cannot create are NAMED (NO_WRITE_PATH), hidden rows are not read' },
    // GH #300 K1: the cell comments of one sheet, the READ half of the #300 comment family (K2 add_comment is its write).
    read_comments: { status: STATUS.ACCEPTED, args: { sheet: SHEET },
      evidence: 'every comment of the sheet through the PUBLIC ApiComment getters (text / author / userId / timeUtc / resolved + replies), the anchor cell found by a bounded per-cell GetComment scan (ApiComment has no address getter; the used range does not reach a comment-only cell) reconciled against ws.GetComments() -- an unreconciled scan names the sheet, never a partial list; no id (GetId is not stable across a reopen); K1 LIVE probes on DS 9.3.5: editor-made comments session + reopen, and a PRE-SEEDED Excel-shaped workbook (legacy notes, threaded comments with persons / dT / a reply / done) -- docs/evidence/gh300-k1-comments-probe-20261026.json, docs/GH300-K1-XLSX-READ-COMMENTS-DECISION.md' },
    read_print_area: { status: STATUS.ACCEPTED, args: { sheet: SHEET },
      evidence: 'the sheet-level Print_Area, the MIRROR of set_print_area / clear_print_area (names.cjs): GH #134 P1, a field-select over the driver\'s defined-names read (DS Bridge workbook.definedNames.read/v1, the same read set_print_area is verified with); one rectangle -> range, two or more -> areas in the stored order; no print area -> set false; a shape the writer has no word for (a whole column / row, a part on another sheet) or a names gap is NAMED, never a value' },
    read_defined_names: { status: STATUS.ACCEPTED, args: {},
      evidence: 'every defined name the driver reads (DS Bridge workbook.definedNames.read/v1, the read read_print_area and the names family use), each the MIRROR of a set_defined_name step: name, refersTo with its =, scope only for a sheet-level name; Print_Area / Print_Titles carry builtIn, Print_Titles also titles { columns, rows } (the set_print_titles shape) or a NAMED shape gap; hidden and _xlnm-named names are not read (the driver filter); a names gap or more than 500 names is NAMED, never a list. GH #134 P2, LIVE on a PRE-SEEDED fixture' },
    // GH #137 U1 (#141 roadmap GO, "productize BOTH sides together"): the product READ mirror of format_run
    // (src/adapters/xlsx/families/format-run.cjs), released BEFORE format_run's own WRITE (#137's own stated order).
    // The family's own run-map build: GetCharacters(i,1).GetFont() per character, merging consecutive identical
    // font-tuples (bold/italic/name/size/color -- color the public GetRGB() integer, the SAME convention the
    // read_range font.color field already uses). Bounded by the family's RUN_MAP_CHAR_CEILING (256 chars,
    // TEMPORARY) -- a longer text is named 'UNOBSERVABLE:text-too-long', never a guessed/truncated run list; an
    // empty cell is named 'UNOBSERVABLE:empty-cell'; a missing sheet 'UNOBSERVABLE:sheet-missing'. A non-text
    // (e.g. number) cell is NOT a gap -- the driver reads its stringified value like any other text, one run
    // covering the whole string.
    read_runs: { status: STATUS.ACCEPTED, args: { sheet: SHEET, cell: CELL },
      evidence: 'per-run (character start/length) font reader inside one cell, the MIRROR of format_run: X-F04 (format-run.cjs design doc docs/XF04-RUN-FONT-DESIGN-20261007.md, OWNER-APPROVED); the SAME run-map xlsx-driver.cjs readQuery already builds for format_run\'s own L1/L3 (runCells -> runReads, sheet!cell keyed) -- no new driver code, only a new product-surface consumer of an existing reader' },
    read_print_rendering: { status: STATUS.NOT_MEASURED, evidence: 'X-P15, X-P04' },
  }),
  docx: Object.freeze({
    read_text: { status: STATUS.ACCEPTED, args: {}, evidence: 'blocks (heading/paragraph/list/table text) + locator: M4-WAVE1 (M8 124/124), M10 matrix W-08/W-09-3' },
    read_formatting: { status: STATUS.ACCEPTED, args: {}, evidence: 'run/paragraph props + format, hyperlinks, note refs: M4-WAVE1, M4-W11, M4-W3-RICH; defaultText M5 gate 6' },
    read_tables: { status: STATUS.ACCEPTED, args: {}, evidence: 'public table read (cell text/format/paraFormats): M4-W4A-D; GH #222: table-level format (borders/shading/cellMargins/indent/width/style) additive -- the SAME tblPr read W04-4 already measured (table-format.cjs write-side verification), docs/GH222-TABLE-FORMAT-READ-DECISION.md' },
    read_headers_footers: { status: STATUS.ACCEPTED, args: {}, evidence: 'header/footer default/even/title parts: M4-W5 14/14, M5 hf variants 16/16' },
    read_page_setup: { status: STATUS.ACCEPTED, args: {}, evidence: 'GH #223: per-section pageWidth/pageHeight/margins/columnsCount/columnsSpaces/startPageNumber -- the SAME inventory.cjs fields (267-331: measured()/measuredArr(), twips) the W11 page-setup WRITE side already reads back; a DEDICATED operation (not a read_headers_footers extension, docs/GH223-DOCX-PAGE-SETUP-READ-DECISION.md), titlePage/evenAndOdd excluded (B21/#227, a different, runtime-gated axis)' },
    read_bookmarks: { status: STATUS.ACCEPTED, args: {}, evidence: 'M4-W3-RICH' },
    // GH #225 (docs/GH225-NUMBERING-READ-DECISION.md): PARTIAL -> ACCEPTED, the MIRROR of set_numbering/
    // apply_list's own observation (families/numbering.cjs fieldMatches, families/list.cjs listState):
    // present/type/numId/level/levelFormats (Word Bridge docx.numbering.levelFormat.read/v1, the SAME
    // read B19 bridge parity 8/8 already proved) + format.numFmt/lvlText/suff/lvlJc/start (PUBLIC
    // ApiNumbering.ToJSON(), the SAME path W07-1a..1d/W07-3/W07-7 LIVE-measured for the WRITE side's own
    // verification). No new driver code; a gap is per-item and named (unread | inconsistent |
    // UNOBSERVABLE:<field>-not-set), never guessed. DIRECT_INTERNAL_FALLBACK=0 (docx-driver.cjs:232-239,
    // the W-B1920-C1 cutover). SHARED_NUMID_OBSERVABLE=YES: every item's own numId is the grouping key.
    read_numbering: { status: STATUS.ACCEPTED, args: {}, evidence: 'present/type/numId/level/levelFormats: B19 bridge parity 8/8 (docx.numbering.levelFormat.read/v1); format fields (numFmt/lvlText/suff/lvlJc/start): W07-1a..1d/W07-3/W07-7 LIVE (families/numbering.cjs write-side verification, the SAME ToJSON path); DIRECT_INTERNAL_FALLBACK=0 (docx-driver.cjs:232-239); GH #225, docs/GH225-NUMBERING-READ-DECISION.md' },
    // GH #224 (docs/GH224-STYLE-DEFINITIONS-READ-DECISION.md): a DEDICATED operation -- the style DEFINITIONS the
    // canonical inventory already carries (inventory.cjs `document.styles`, keyed by NAME): identity {name, basedOn,
    // styleType} + text {bold, size, color} + paragraph {align, indentLeft, spacingAfter}, each a measured value, the
    // word 'inherited' (the definition itself does not set it -- the inheritance chain is NOT resolved) or a named
    // 'UNOBSERVABLE:<reason>'. SCOPE: the styles doc.GetAllStyles() enumerates -- a custom style it is blind to
    // (W-10 run3) is NOT listed and no completeness is claimed. EFFECTIVE rendered appearance of a paragraph using a
    // style is NOT claimed (the definition only).
    read_styles: { status: STATUS.ACCEPTED, args: { names: STYLE_NAMES }, optional: ['names'], evidence: 'GH #224b (names): a name-targeted read via the driver\'s buildTargetedStyleQuery, a custom style GetAllStyles is blind to included; GH #224: style definitions by name (identity basedOn/styleType + text bold/size/color + paragraph align/indentLeft/spacingAfter) -- the SAME doc.GetStyle(name).ToJSON() read (docx-driver.cjs styleOf, W-10 LIVE run3/run4) the style-definition family\'s own verification uses; scope = the doc.GetAllStyles() enumeration (a custom style it is blind to is not listed, no completeness claimed); the DEFINITION only, NEVER the effective rendered appearance; docs/GH224-STYLE-DEFINITIONS-READ-DECISION.md' },
    read_table_grid: { status: STATUS.PARTIAL, evidence: 'gridCalc B20 bridge parity 5/5, M10 matrix PARTIALLY COVERED' },
    // GH #226 (docs/GH226-NOTES-READ-DECISION.md): the status and the underlying read are UNCHANGED
    // by this unit -- the audit's own finding is that "first paragraph only" had never been
    // MEASURED against a genuinely multi-paragraph note (every prior LIVE probe used a
    // single-paragraph note); the evidence text below replaces an assumption drawn from the
    // getter's own name with the measured fact of WHICH getter is called and WHAT remains unread.
    read_notes: { status: STATUS.ACCEPTED, args: {}, evidence: 'the FULL text of every footnote/endnote: per note the paragraph list through the public GetNext() chain of its first paragraph (docx-driver.cjs); GH #226 LIVE on DS 9.3.5 (gh226chain-live-1792394096767-10bc02): GetNext() === null at every note end (3 footnotes + 3 endnotes of 1-3 paragraphs), every note FULL, fresh and reopen, no write -- docs/GH226-NOTES-READ-DECISION.md. A chain that does not end by null is NAMED incomplete (unobservable), never a guessed text' },
    // GH #268b (docs/GH268B-DOCX-READ-IMAGES-PRODUCT-DECISION.md): RELEASED (PARTIAL -> ACCEPTED) after the g268v1 LIVE probe's SEED_OK was green.
    read_images: { status: STATUS.ACCEPTED, args: {}, evidence: 'doc.GetAllImages() + the page-side projection of the PUBLIC ApiDrawing.ToJSON() (docPr id/name/descr/title, extent EMU, drawingType, wrapType, the paragraph text); the image BYTES never leave the editor; scope: main document content (header / footer / content-control pictures not listed); GH #268 g268v1 SEED_OK + docx-images-read.live.cjs -- docs/GH268B-DOCX-READ-IMAGES-PRODUCT-DECISION.md' },
    // GH #269b (docs/GH269b-DOCX-READ-COMMENTS-DECISION.md): RELEASED (NOT_MEASURED -> ACCEPTED) after the g269bv2 LIVE PASS on DS 9.3.5 (read_comments through the product path: two rows, the author / userId / quote / resolved / reply fields, UTC-ms time, ids stable after a reopen, the anchor a named gap).
    read_comments: { status: STATUS.ACCEPTED, args: {}, evidence: 'GH #269b: the main document\'s comments through the public doc.GetAllComments() (opt-in includeComments; id/text/author/userId/timeUtc/quote/resolved + replies; the range anchor has NO public getter and is a NAMED gap) -- LIVE PASS g269bv2 on DS 9.3.5 (two rows, fields, UTC-ms time, ids stable after a reopen; docs/GH269b-DOCX-READ-COMMENTS-DECISION.md)' },
    // GH #227 (docs/GH227-SECTION-FLAGS-READ-DECISION.md): INTERNAL -> ACCEPTED. The stale evidence this
    // replaces ("internal ApiSection getters") was wrong even at M11.5e.4: since that cutover titlePage/
    // evenAndOdd are read EXCLUSIVELY through the word /3 Bridge READ capability docx.section.variantState.
    // read/v1 (docx-driver.cjs __eoB21Variant/variant()), with NO direct-internal fallback on any failure --
    // the SAME write-side runtime gate (internalGateOf(runtime).sectionVariant, B21_BRIDGE_READ_RUNTIME_IDS)
    // the header/footer family already refuses a variant WRITE on. A per-section flag is one of: a measured
    // boolean, or a NAMED 'UNAVAILABLE:<reason>' string (not-requested | runtime-not-verified | bridge-
    // unavailable), never a guessed default -- the SAME inventory.cjs variantFlag() normalizer the WRITE
    // side's own preservation check (hf.flagChanges) already depends on. DIRECT_INTERNAL_FALLBACK=0.
    read_section_flags: { status: STATUS.ACCEPTED, args: {}, evidence: 'per-section titlePage/evenAndOdd through docx.section.variantState.read/v1 (B21, M11.5e.4 word /3 Bridge READ cutover), runtime-gated (internalGateOf(runtime).sectionVariant); a boolean or a named UNAVAILABLE:<not-requested|runtime-not-verified|bridge-unavailable>, never guessed; DIRECT_INTERNAL_FALLBACK=0; GH #227, docs/GH227-SECTION-FLAGS-READ-DECISION.md' },
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
const arrayOk = (v, spec) => Array.isArray(v) && v.length >= (spec.minItems || 1) && v.length <= (spec.maxItems || 1) && v.every(x => validString(x, spec.items)) && new Set(v).size === v.length;
const typeOk = (v, spec) => (spec.type === 'boolean' ? typeof v === 'boolean' : (spec.type === 'array' ? arrayOk(v, spec) : validString(v, spec)));

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

// GH #137 U1: `cell` must parse to EXACTLY one A1 cell, not a range (the CELL schema pattern already rules out
// the colon form; this catches a malformed single cell, e.g. a column/row out of bounds, the same way
// checkRange does for `range`).
function checkCell(cell) {
  if (cell === undefined) return null;
  const g = parseRange(cell);
  if (!g || g.count !== 1) return refuse('INVALID_ARGS', 'cell must be a single A1 cell, not a range');
  return null;
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
    () => checkFields(op, args.operation, args), () => checkRange(args.range), () => checkCell(args.cell)];
  for (const stage of stages) { const refusal = stage(); if (refusal) return refusal; }
  const out = Object.fromEntries(Object.keys(op.args).filter(k => args[k] !== undefined).map(k => [k, args[k]]));
  if (out.range !== undefined) out.range = out.range.replace(/\$/g, '').toUpperCase();
  if (out.cell !== undefined) out.cell = out.cell.replace(/\$/g, '').toUpperCase();
  return { ok: true, fileId: args.fileId, operation: args.operation, args: Object.freeze(out) };
}

/** The readLive request an XLSX operation needs (the driver's own request fields). read_range is a pure window read, so it
 * is eligible for the tiered observation: above the full-read bound (MAX_READ_CELLS) the driver reads only the requested
 * window instead of a whole-workbook read that truncates before reaching it (GitHub #72: every cell came back
 * `read: false`). `windowStyle` lets a style read of that window through -- a read asserts no preservation, and the window
 * style read is bounded by the window cell count. At or below the bound the full read runs unchanged (tierOf). */
function xlsxRequest(operation, a) {
  // GH #327 (U5): the SAME window also asks the protection axis for its own range -- read_range's own
  // redaction (readRange() below) needs cells.hidden for exactly this sheet+range, not the whole sheet.
  if (operation === 'read_range') return { tiered: true, readOnly: true, windowStyle: a.style === true, windows: [{ sheet: a.sheet, range: a.range, format: a.format === true, style: a.style === true }], protectionRanges: [{ sheet: a.sheet, range: a.range }] };
  if (operation === 'read_validations') return { windows: [], validationRanges: [{ sheet: a.sheet, range: a.range }] };
  if (operation === 'read_conditional_formats') return { windows: [], cfRanges: [{ sheet: a.sheet, range: a.range }] };
  if (operation === 'read_layout') return { windows: [], layout: [{ sheet: a.sheet, ...layoutExtent(a.range) }] };
  if (operation === 'read_freeze_panes') return { windows: [], panes: true };
  // GH #133: every sheet's visibility (the sheetVisibility axis) and the active sheet's name (carried by the panes block); both existing axes.
  if (operation === 'read_sheet_state') return { windows: [], panes: true, sheetVisibility: true };
  if (operation === 'read_print_area') return { windows: [], dependents: true };
  if (operation === 'read_defined_names') return { windows: [], dependents: true };
  if (operation === 'read_page_setup') return { windows: [], pageSheets: [a.sheet], printSheets: [a.sheet] };
  if (operation === 'read_filter') return { windows: [], filterSheets: [a.sheet] };
  if (operation === 'read_comments') return { windows: [], commentSheets: [a.sheet] };
  // X-O05/4: an omitted `sheet` reads every sheet's charts (the driver's own '*' convention, chartAsked.includes('*')) --
  // matching workbook_overview's all-sheets default, not a required argument like read_range/read_validations.
  if (operation === 'read_charts') return { windows: [], chartSheets: [a.sheet || '*'] };
  // GH #140 U2: the SAME all-sheets '*' convention, requesting the driver's objectSheets axis --
  // the one create_pivot/refresh_pivot's own liveExtras() already requests on the WRITE side
  // (adapter.cjs objectSheets), no new driver reader.
  if (operation === 'read_pivots') return { windows: [], objectSheets: [a.sheet || '*'] };
  // X-O08 (/14): an omitted `sheet` reads every sheet's drawings -- the same all-sheets '*' convention
  // read_charts already uses, not a required argument like read_range/read_validations.
  if (operation === 'read_drawings') return { windows: [], drawingSheets: [a.sheet || '*'] };
  // GH #137 U1: the SAME runCells key format-run.cjs's own liveExtras() already builds for the WRITE
  // side (`{ sheet, cell }`), driving the SAME xlsx-driver.cjs readQuery runCells axis -- no new reader.
  if (operation === 'read_runs') return { windows: [], runCells: [{ sheet: a.sheet, cell: a.cell }] };
  return { windows: [] };
}

// `font.color` (GitHub #89, write-tools.cjs FORMAT_RANGE_RELEASED_KEYS): the public
// GetCharacters(1, len).GetFont().GetColor().GetRGB() passthrough -- an integer, or absent on a cell with no
// observable font (format.cjs's own font-unobservable gate). The channel order is RGB (matching
// write-tools.cjs's intOf / colour() convention) -- LIVE-MEASURED on the CAPABILITY path (X-B08,
// test/live/xlsx-xb08-color-roundtrip.live.cjs, LIVE r4, 2026-10-13: 9 colours incl. an asymmetric one,
// #3A7BC8, byte-identical to intOf(hex) after reopen; docs/CAPABILITY-COMPLETION-STATUS-20261012.md X-B08 row).
// The GH #89 product-surface LIVE harness re-confirms the same int through xlsx_execute_task / xlsx_read
// specifically, rather than re-deriving it from nothing.
const CELL_FIELDS = ['v', 'f', 'merge', 'nf', 'fill', 'bold', 'italic', 'wrap', 'rot', 'readingOrder', 'indent', 'shrink', 'font', 'align', 'border'];

// GitHub #72: the cell's value TYPE, so that a text "1." and a number 0.1 are told apart (`v` is always the displayed
// string). formula = a real formula (the inventory already folds a text literal that looks like one into f '');
// empty = no value; number = COUNT 1; text = COUNT 0; unknown = COUNT not read -- never guessed.
const typeOf = c => (c.f ? 'formula' : (c.v === null || c.v === '' ? 'empty' : (c.count === 1 ? 'number' : (c.count === 0 ? 'text' : 'unknown'))));

// GH #327 (U5): a protected sheet's hidden-formula cells -- the formula TEXT is redacted, never the value. A cell
// not of type 'formula' (no `f`, or an unread cell) is untouched either way; only the `f` field itself changes.
const degradeFormulas = (cells, reason) => cells.map(c => (c.type === 'formula' ? { ...c, f: `UNOBSERVABLE:protection-unread:${reason}` } : c));

/** `prot` = inv.protectionRangeReads[`${sheet}!${range}`]: undefined (not read, same as read_validations'
 * own convention) | a degraded 'UNOBSERVABLE:<reason>' string (the WHOLE protection read failed, fail-closed:
 * every formula in the window degrades too, never left readable on a guess) | { protected, forbids, cells }.
 * Unprotected, or protected but the cell's formula is not in `cells.hidden`: unchanged (the owner's
 * condition -- the VALUE is never touched, on any branch). */
function redactHiddenFormulas(cells, prot) {
  if (prot === undefined) return degradeFormulas(cells, 'not-read');
  if (typeof prot === 'string') return degradeFormulas(cells, prot.startsWith('UNOBSERVABLE:') ? prot.slice('UNOBSERVABLE:'.length) : prot);
  if (prot.protected !== true) return cells;
  const hidden = new Set((prot.cells && prot.cells.hidden) || []);
  return cells.map(c => (c.type === 'formula' && hidden.has(c.address) ? { ...c, f: 'REDACTED:hidden-formula-on-protected-sheet' } : c));
}

function readRange(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, range: a.range, error: 'SHEET_NOT_FOUND' };
  // only the requested window's cells, row-major; an unread cell is said so, never defaulted
  const cells = cellsOf(parseRange(a.range)).map(address => {
    const c = inv.cells[xlsxInventory.key(a.sheet, address)];
    return c ? { address, ...Object.fromEntries(CELL_FIELDS.filter(k => k in c).map(k => [k, c[k]])), type: typeOf(c) } : { address, read: false };
  });
  const prot = inv.protectionRangeReads ? inv.protectionRangeReads[`${a.sheet}!${a.range}`] : undefined;
  return { sheet: a.sheet, range: a.range, cells: redactHiddenFormulas(cells, prot) };
}

const XLSX_PROJECTIONS = Object.freeze({
  workbook_overview: inv => ({ sheets: inv.sheets.map(s => ({ name: s.name, index: s.index, used: s.used })) }),
  read_range: readRange,
  read_conditional_formats: projectCfRead,
  read_layout: projectLayoutRead,
  read_freeze_panes: projectPanesRead,
  read_sheet_state: projectSheetStateRead,
  read_page_setup: projectPageSetupRead,
  read_print_area: projectPrintAreaRead,
  read_defined_names: projectDefinedNamesRead,
  read_filter: projectFilterRead,
  read_comments: projectCommentsRead,
  read_validations: (inv, a) => {
    const v = inv.validationReads ? inv.validationReads[`${a.sheet}!${a.range}`] : undefined;
    if (v === undefined) return { sheet: a.sheet, range: a.range, validation: 'UNOBSERVABLE:not-read' };
    if (typeof v !== 'object') return { sheet: a.sheet, range: a.range, validation: v };
    // L-22e: 'activeCell' is an internal field the write-side family's own identity check needs
    // (reAnchoredFormula(), validation.cjs) -- never part of this tool's own public contract. For
    // type 'custom' the raw formula1 the driver read is relative to the RULE'S OWN SHEET's current
    // active cell (measured), not this request's own range -- this tool re-anchors it to the
    // range's own first cell with the SAME helper the write-side identity check uses, BEFORE
    // stripping activeCell, so a caller never sees an active-cell-shifted formula. A fully-absolute
    // formula is unaffected (reAnchoredFormula()'s own fast path). An active cell that cannot be
    // resolved on a relative/mixed formula is a NAMED gap here too -- never the raw, misleading text.
    const { activeCell, ...rest } = v;
    if (rest.type !== 'xlValidateCustom' || typeof rest.formula1 !== 'string') return { sheet: a.sheet, range: a.range, validation: rest };
    const wantAt = rangeFirstCellAddr(a.range);
    const reAnchored = wantAt ? reAnchoredFormula(rest.formula1, activeCell, wantAt) : null;
    return { sheet: a.sheet, range: a.range, validation: { ...rest, formula1: reAnchored === null ? 'UNOBSERVABLE:active-cell-unresolvable' : reAnchored } };
  },
  // GH #137 U1: a plain field-select over inv.runReads (the SAME sheet!cell key format-run.cjs's own runKey()
  // uses) -- the bridge-free, PUBLIC-getter read already happened IN-SESSION (xlsx-driver.cjs readQuery's
  // runCells axis), this projection runs AFTER the session closed and makes no new call. `runMap` is either
  // the driver's own named gap string (e.g. 'UNOBSERVABLE:empty-cell') or `{ text, runs }` -- never a guessed
  // default on a gap, never an empty list standing in for "not read".
  read_runs: (inv, a) => {
    const v = inv.runReads ? inv.runReads[`${a.sheet}!${a.cell}`] : undefined;
    return { sheet: a.sheet, cell: a.cell, runMap: v === undefined ? 'UNOBSERVABLE:not-read' : v };
  },
  // X-O05/4 (design docs/X-O05-3-DATALABEL-READ-CONSUMER-DESIGN-20261013.md, point 1 + point 5 test plan):
  // a plain field-select over inv.chartReads -- the bridge reads already happened IN-SESSION (xlsx-driver.cjs's
  // chartReads), this projection runs AFTER the session closed and makes no new bridge call.
  read_charts: (inv, a) => {
    const reads = inv.chartReads || {};
    const names = a.sheet ? [a.sheet] : Object.keys(reads);
    return { sheets: names.map(name => {
      const r = reads[name];
      // a whole-sheet gap (sheet-missing, or any future whole-axis UNOBSERVABLE) is a string, never guessed as empty
      return typeof r === 'string' ? { sheet: name, charts: null, pivots: null, reason: r } : { sheet: name, charts: r.charts, pivots: r.pivots };
    }) };
  },
  // GH #140 U2 (#141 roadmap GO): a plain field-select over inv.objects -- the SAME content-proven
  // pivot read create_pivot/refresh_pivot's own gate() already consumes internally since M7-H
  // (pivot-source.cjs withProvenPivots()), now projected to the product READ surface for the first
  // time (GH #140 U1's own finding: until this unit, inv.objects had no READ consumer at all). No
  // new driver call: inv.objects is built unconditionally by inventory.cjs whenever objectSheets was
  // requested (xlsxRequest above). A whole-sheet gap (sheet-missing, or the read never asked for this
  // sheet's objects at all) is a string, never guessed as an empty list; a gap on ONE pivot inside an
  // otherwise-good sheet (one bad pivot-ref makes withProvenPivots() name the WHOLE sheet's pivots
  // axis, per its own module comment: "a per-pivot gap inside a list would pass the consumers that
  // check the axis only") surfaces the same way -- `o.pivots` itself a string in that case, not an
  // array -- so this handler treats it identically to a whole-sheet gap, never a partial/guessed list.
  read_pivots: (inv, a) => {
    const objs = inv.objects || {};
    const names = a.sheet ? [a.sheet] : Object.keys(objs);
    return { sheets: names.map(name => {
      const o = objs[name];
      // inv.objects only carries EXISTING sheets, so a name absent from inv.sheets is a missing sheet
      if (o === undefined) return { sheet: name, pivots: null, reason: (inv.sheets || []).some(s => s.name === name) ? 'UNOBSERVABLE:not-read' : 'UNOBSERVABLE:sheet-missing' };
      if (typeof o === 'string') return { sheet: name, pivots: null, reason: o };
      // GH #302 G2: the per-pivot rows when the driver read the axes (pivot-source.cjs pivotDetails) -- a richer pivot is
      // listed with its fields; an unproven source is that row's own named `source`, not the whole sheet's gap
      if (Array.isArray(o.pivotDetails)) return { sheet: name, pivots: o.pivotDetails.map(row => Object.fromEntries(Object.entries(row).filter(([k]) => k !== 'dataRef'))) };
      return typeof o.pivots === 'string' ? { sheet: name, pivots: null, reason: o.pivots } : { sheet: name, pivots: o.pivots };
    }) };
  },
  // X-O08 (/14): the SAME field-select-over-the-driver's-own-read convention read_charts already uses --
  // inv.drawingReads is generated IN-SESSION (xlsx-driver.cjs's drawingReads), this projection runs AFTER the
  // session closed and makes no new bridge call. Every field is already self-describing (a missing/throwing
  // getter or an unserved bridge capability is its own named 'UNOBSERVABLE:<why>' string on that ONE field,
  // never a guessed default) -- the whole per-sheet list is returned as-is, never filtered.
  //
  // GH #302 / #325 (G1 minimum, docs/GH302-PIVOT-CHART-DRAWING-LIVE-MATRIX-20261025.md rows C-03/D-02): the
  // PUBLIC GetWidth()/GetHeight() getters read an EXACT 0 x 0 after a reopen for a chart or a shape, while the
  // stored anchor (xdr:ext) is right -- a wrong value, not a missing one, so dget() upstream cannot name it
  // itself. The richer fix (read the saved package's own xdr:ext instead, same precedent as
  // ooxml/point-data-label-reader.cjs) needs the FILE BYTES; this read path (readOnce -> driver.readLive ->
  // this projection) never has them -- only the in-session public getters -- so that fix would mean plumbing
  // bytes through the whole xlsx_read call chain, out of this unit's scope. Minimum instead: report the
  // honest limit, never the unproven 0, on exactly this one field.
  read_drawings: (inv, a) => {
    const sizeUnobservedAfterReopen = d => (d && typeof d === 'object' && d.width === 0 && d.height === 0
      ? { ...d, width: 'UNOBSERVABLE:size-zero-after-reopen', height: 'UNOBSERVABLE:size-zero-after-reopen' }
      : d);
    const reads = inv.drawingReads || {};
    const names = a.sheet ? [a.sheet] : Object.keys(reads);
    return { sheets: names.map(name => {
      const r = reads[name];
      // a whole-sheet gap (sheet-missing, no-GetAllDrawings, ...) is a string, never guessed as empty
      return typeof r === 'string' ? { sheet: name, drawings: null, reason: r } : { sheet: name, drawings: Array.isArray(r) ? r.map(sizeUnobservedAfterReopen) : r };
    }) };
  },
});

// The DOCX projections. Never returned: gridCalc (PARTIAL), titlePage/evenAndOdd (INTERNAL),
// notes text (PARTIAL), images (NOT_MEASURED), the internal gate.
const orNull = v => (v === undefined ? null : v);
const textOf = b => (b.kind === 'table' ? { rows: b.rows.map(r => r.map(c => c.text)) } : { text: b.text });
const formatOf = n => ({ text: n.text, props: orNull(n.props), format: orNull(n.format), runs: orNull(n.runs), hyperlinks: orNull(n.hyperlinks), noteRefs: orNull(n.noteRefs) });
const unitsOf = blocks => blocks.flatMap(b => (b.kind === 'list' ? b.items.map(it => ({ kind: 'list_item', ...formatOf(it) }))
  : (b.kind === 'paragraph' || b.kind === 'heading' ? [{ kind: b.kind, ...formatOf(b) }] : [])));
// `nestedTables` (the nested tables bound to the holding cell) is projected only when the cell has them (an array of { rows }, recursively, or the driver's named UNOBSERVABLE string).
const cellOf = c => ({ text: c.text, simple: c.simple, format: c.format, paraFormats: c.paraFormats,
  ...(c.nestedTables === undefined ? {} : { nestedTables: Array.isArray(c.nestedTables) ? c.nestedTables.map(nt => ({ rows: nt.rows.map(r => r.map(cellOf)) })) : c.nestedTables }) });

/**
 * GH #225 (DOCX P1 READ, docs/GH225-NUMBERING-READ-DECISION.md): read_numbering, PARTIAL -> ACCEPTED.
 * Every field below is already measured and already consumed by families/numbering.cjs's own L1/L3
 * (the SAME `node.numbering` the canonical inventory (inventory.cjs:161-183) attaches to every
 * paragraph/heading/list-item, unchanged by this unit) -- this projection adds NO new reader, only
 * a new product-surface consumer, the same convention read_runs (X-F04) and read_pivots (GH #140 U2)
 * already used for a PARTIAL -> ACCEPTED flip.
 *
 * `type` ('bullet'|'numbered'|'normal') is NOT part of the Bridge numbering read (`present`/`level`/
 * `numId`/`levelFormats` only, docx-driver.cjs:172-179 validNumbering) -- it is derived EXACTLY the
 * way families/list.cjs's own listState() derives it (list.cjs:188): `block.ordered` (the HTML <ol>
 * vs <ul> the inventory already classified the list block as, inventory.cjs:88) for a list item,
 * `'normal'` for a non-list paragraph/heading whose numbering reads `present:false`. Reusing this
 * exact rule (not reinventing it) is what makes a caller's grouping of list.cjs's own write-side
 * `type` consistent with this read.
 *
 * A gap is per-item and NAMED, never silently dropped or guessed:
 *   numbering: null, numberingGap: 'unread'        -- the Bridge read never attached (ambiguous text
 *                                                      match, b19 parity gap, or this ONE paragraph's
 *                                                      own read failed) -- NOT the same as "not numbered"
 *   numbering: null, numberingGap: 'inconsistent'   -- a list_item block whose OWN numbering read says
 *                                                      present:false (or vice versa) -- a state the
 *                                                      driver should never produce; named, not hidden
 *   numbering: { present: false }                   -- a MEASURED negative: this paragraph carries no
 *                                                      NumPr. Never confused with the two gaps above.
 *   numbering: { present: true, type, numId, level, levelFormats, format }
 *     level            integer (Bridge-read level index) or null (B19 gate note: null IS a measured
 *                       value here, not a gap -- docx-driver.cjs's own validNumbering accepts it)
 *     levelFormats      the 9-slot internal per-level format array (EO_INTERNAL_DEPENDENCY gate,
 *                       inventory.cjs:170), or null when that ONE gate did not verify -- independent
 *                       of `format` below (a DIFFERENT gate; one failing never hides the other)
 *     format            each of numFmt/lvlText/suff/lvlJc/start individually: the measured ToJSON
 *                       value, or 'UNOBSERVABLE:<field>-not-set' when that ONE key was absent
 *                       (inventory.cjs:175-181 already strips anything not independently re-validated
 *                       -- this projection invents nothing beyond what that layer already decided was
 *                       safe to keep). `start` in practice is almost always present: the driver's own
 *                       OOXML-measured default-is-1 rule (docx-driver.cjs:268-275, W07-7 probe) fills
 *                       it in whenever the key was omitted FOR THAT reason -- `format` itself is only
 *                       `'UNOBSERVABLE:numbering-format-not-read'` (the WHOLE object, not per field)
 *                       when the ToJSON read did not run on this paragraph at all (a bridge/runtime
 *                       gap distinct from the per-field case).
 *
 * Shared-numId semantics (SHARED_NUMID_OBSERVABLE): every item's own `numId` IS the observable
 * relationship -- a caller verifying a #248 numbering-definition write groups this operation's own
 * `items` by `numId` and checks every member changed together (the SAME multiset families/
 * numbering.cjs's own verifyPreservation (numbering.cjs:282-305) already builds server-side for its
 * L3 check; this projection does not pre-group, the same "raw per-item, caller aggregates" convention
 * read_runs/read_drawings already use -- grouping 20-200 named items client-side is not a capability
 * gap, so no new grouped shape is added here).
 *
 * DIRECT_INTERNAL_FALLBACK=0: the numId/level/levelFormats/format read is EXCLUSIVELY the Word Bridge
 * capability docx.numbering.levelFormat.read/v1 (docx-driver.cjs:238) plus the PUBLIC
 * ApiNumbering.ToJSON() (docx-driver.cjs:249-278, gated on `numbering.present === true` from that SAME
 * bridge read) -- "this driver no longer touches the internal per-level format reader at all... never
 * a direct-internal fallback" (docx-driver.cjs:232-237, the W-B1920-C1 cutover's own comment). This
 * projection calls neither driver method nor any internal getter itself; it only re-shapes
 * inventory.cjs's already-computed, already-gated `numbering` field.
 */
function numberingTypeOf(block, n) {
  if (n.present === true) return block && block.kind === 'list' ? (block.ordered ? 'numbered' : 'bullet') : 'UNOBSERVABLE:list-block-missing';
  return 'normal';
}
function numberingFormatFieldsOf(fmt) {
  if (!fmt || typeof fmt !== 'object') return 'UNOBSERVABLE:numbering-format-not-read';
  const field = k => (Object.prototype.hasOwnProperty.call(fmt, k) ? fmt[k] : `UNOBSERVABLE:${k}-not-set`);
  return { numFmt: field('numFmt'), lvlText: field('lvlText'), suff: field('suff'), lvlJc: field('lvlJc'), start: field('start') };
}
function numberingOfUnit(kind, block, n) {
  if (!n) return { kind, numbering: null, numberingGap: 'unread' };
  const isListItem = kind === 'list_item';
  if (isListItem ? n.present !== true : n.present !== false) return { kind, numbering: null, numberingGap: 'inconsistent' };
  if (!isListItem) return { kind, numbering: { present: false } };
  return { kind, numbering: {
    present: true, type: numberingTypeOf(block, n), numId: n.numId || null, level: n.level === undefined ? null : n.level,
    levelFormats: n.levelFormats || null, format: numberingFormatFieldsOf(n.format),
    // GH #293 N5: the driver's abstractId is NOT reported -- it is session-local (LIVE: '1_69' before a save + reopen, '700' after,
    // the same definition), so as a read field it would claim a change nobody made. The write guards compare it within one session.
  } };
}
function numberingUnitsOf(blocks) {
  const out = [];
  for (const b of blocks) {
    if (b.kind === 'list') for (const item of b.items || []) out.push({ text: item.text, ...numberingOfUnit('list_item', b, item.numbering) });
    else if (b.kind === 'paragraph' || b.kind === 'heading') out.push({ text: b.text, ...numberingOfUnit(b.kind, b, b.numbering) });
  }
  return out;
}

/**
 * GH #224 (DOCX P2 READ, docs/GH224-STYLE-DEFINITIONS-READ-DECISION.md): read_styles. A plain re-shape of the
 * canonical inventory's `document.styles` (inventory.cjs, keyed by style NAME, never styleId) -- NO new driver
 * read: the driver's styleOf (docx-driver.cjs, W-10) already measured every field below and the style-definition
 * family's own verification consumes the SAME map.
 *
 *   { styles: null }                          -- the style map was never read (no document / stylesUnreadable): a gap, NOT "no styles"
 *   { styles: [{ name, identity, text, paragraph }] }   -- in the inventory's own (GetAllStyles) order
 *     identity   { name, basedOn, next, link, styleType, basedOnName, nextName, linkName } or null when the style's ToJSON was unreadable (then text/paragraph
 *                carry 'UNOBSERVABLE:no-readable-ToJSON' per field -- the driver's own value, passed through)
 *     text       { bold, size (pt), color (hex | 'auto') }
 *     paragraph  { align, indentLeft (pt), spacingAfter (pt) }
 *     each field: a measured value | 'inherited' (NOT set on this definition -- the base-style chain is NOT
 *                 resolved here, so it is NOT a value) | 'UNOBSERVABLE:<reason>'
 *   { name, gap: 'unread' }                    -- a name the enumeration listed but whose definition the driver could
 *                                                 not read back (GetStyle(name) did not return that name): named, not dropped
 *
 * SCOPE and what is NOT claimed: the list is the doc.GetAllStyles() enumeration (the built-in/default styles and any
 * custom style that enumeration happens to list); a session-created custom style GetAllStyles() is blind to (W-10
 * run3) is absent and `scope` says so. The result is the style DEFINITION -- it never claims the EFFECTIVE rendered
 * appearance of a paragraph that uses a style (a modified style's rendering is a different, unproven axis).
 */
function stylesOfDocument(document, args) {
  const names = args && Array.isArray(args.names) ? args.names : null;
  const map = document && document.styles;
  if (names) {
    // GH #224b: name-targeted mode -- exactly the requested names, in the REQUEST order, each read through doc.GetStyle(name) (the targeted query merged them into the style map). A name whose
    // GetStyle(name) did not return a style of that name is { name, gap: 'not-found' }: the style does not exist OR the lookup failed -- the driver cannot tell the two apart (GetStyle falls back to Normal).
    const scope = 'name-targeted doc.GetStyle(name) read of exactly the requested names, in request order (also covers a style GetAllStyles() is blind to); "not-found" = GetStyle(name) returned no style of that name; definition only, not effective rendered appearance';
    if (!map || typeof map !== 'object') return { styles: null, scope };
    return { styles: names.map(name => { const st = Object.prototype.hasOwnProperty.call(map, name) ? map[name] : null; return st ? { name, identity: st.identity, text: st.text, paragraph: st.paragraph } : { name, gap: 'not-found' }; }), scope };
  }
  const scope = 'doc.GetAllStyles() enumeration (a custom style it is blind to is not listed); definition only, not effective rendered appearance';
  if (!map || typeof map !== 'object') return { styles: null, scope };
  return { styles: Object.entries(map).map(([name, st]) => (st ? { name, identity: st.identity, text: st.text, paragraph: st.paragraph } : { name, gap: 'unread' })), scope };
}

/**
 * GH #222 (DOCX P1 READ, docs/GH222-TABLE-FORMAT-READ-DECISION.md): the product-side table-level
 * `format` -- an ADDITIVE field next to `rows` (never replacing it; the cell-level format/
 * paraFormats shape read_tables already returned is untouched). Six axes: borders (top/bottom/
 * left/right/insideH/insideV: value/sz/color), shading, cellMargins (top/right/bottom/left),
 * indent, width, style (identity + name). The source values are inventory.cjs's own
 * tableFormatOf() output (W04-4, src/adapters/docx/inventory.cjs:267-278) -- already measured,
 * per-field UNOBSERVABLE:<reason> on a failed read, `null` on a genuinely absent property; this
 * projection passes borders/shading/cellMargins/indent/width through AS-IS (no re-derivation, no
 * raw Builder object -- inventory.cjs already stripped that) and only adds the style NAME
 * resolution below. Never touches `t.format.style`'s own scalar shape: families/table-format.cjs's
 * write-side verification (fieldMatches 'style' case, line ~182) compares that value directly
 * against a resolved styleId and must keep seeing the SAME scalar it always has.
 */
const STYLE_LIST_UNREAD = 'UNOBSERVABLE:style-list-unread';
const STYLE_ID_NOT_LISTED = 'UNOBSERVABLE:style-id-not-in-list';
/** The REVERSE of table-format.cjs's own resolveStyleId() (name -> id, write side): id -> name,
 * against the SAME per-document tableStyles list (read BEFORE any dispatch). A per-field gap is
 * never guessed: an unread id propagates its own reason (the name could not be resolved either,
 * for the SAME reason); a measured `null` id (no style set) resolves to a measured `null` name --
 * both "no style", neither a read failure; a measured id absent from an unread/missing list is its
 * own distinct, named gap (the id IS known, only the name lookup failed).
 *
 * MEASURED (DS 9.3.5 W04 LIVE evidence, docs/evidence/docx-w04-table-format-live-acceptance-
 * rebased-PASS-20261012.json): the real tblStyle id is itself a STRING ("701", table-format.cjs's
 * own comment at line ~175 says the same -- "String() on both sides"). The gap-check below must
 * therefore key on the UNOBSERVABLE: PREFIX, never on `typeof === 'string'` -- a naive type check
 * would misread every real measured id as a gap. */
function styleNameOf(tableStyles, idVal) {
  if (typeof idVal === 'string' && idVal.startsWith('UNOBSERVABLE:')) return idVal; // propagate the id's own reason
  if (idVal === null) return null; // measured: no style set on this table
  if (!Array.isArray(tableStyles)) return STYLE_LIST_UNREAD; // tableStyles itself is a string (unread) or missing
  const hit = tableStyles.find(s => String(s.styleId) === String(idVal));
  return hit ? hit.name : STYLE_ID_NOT_LISTED;
}
const tableFormatOf = (t, tableStyles) => ({ borders: t.format.borders, shading: t.format.shading, cellMargins: t.format.cellMargins,
  indent: t.format.indent, width: t.format.width, style: { id: t.format.style, name: styleNameOf(tableStyles, t.format.style) } });

/**
 * GH #226 3: the notes of a document, per note: { paragraphs, text (paragraphs joined by a newline), complete, stop }, from the driver's GetNext() chain
 * (inventory.cjs footnoteChains/endnoteChains). A note whose chain is not TRUSTED (complete=false: no GetNext | error | overrun-suspected | bound) is
 * NAMED -- its `unobservable` carries 'UNOBSERVABLE:note-chain-<stop>' and its text is only what was read, never a guessed full text. `complete` is true
 * only when EVERY note of both kinds is complete. An unread kind (null) stays null: the read is not complete.
 */
function notesOfDocument(document) {
  const one = list => (Array.isArray(list)
    ? list.map(c => ({ paragraphs: [...c.paragraphs], text: c.paragraphs.join('\n'), complete: c.complete === true, stop: c.stop,
      ...(c.complete === true ? {} : { unobservable: `UNOBSERVABLE:note-chain-${c.stop}` }) }))
    : null);
  const footnotes = document ? one(document.footnoteChains) : null;
  const endnotes = document ? one(document.endnoteChains) : null;
  const all = [...(footnotes || []), ...(endnotes || [])];
  return { footnotes, endnotes, complete: footnotes !== null && endnotes !== null && all.every(n => n.complete) };
}

/**
 * GH #268b: read_images. A plain field-select over inventory.cjs's own document.images (the driver's page-side projection of ApiDrawing.ToJSON, docx-images-read.cjs): per image { index, alt,
 * title, placement: 'inline' | 'floating', wrap, widthEmu, heightEmu, paragraph }, each a read value, null (genuinely absent) or a named 'UNOBSERVABLE:<reason>'. A wholly unread list is `images: null` with the
 * reason in `gap`. SCOPE: the images doc.GetAllImages() enumerates in the MAIN document content -- header / footer images and content-control pictures are NOT listed and no completeness is claimed for
 * them. The image BYTES are never part of the result (the projection drops them page-side).
 */
function imagesOfDocument(document) {
  const scope = 'doc.GetAllImages() of the main document content (header / footer images and content-control pictures are not listed); alt / title / placement / wrap / size from the public ApiDrawing.ToJSON() projected page-side; the image bytes never leave the editor; the docPr id / name are NOT part of the row (g268v1 LIVE: session-local random values that differ from the saved docPr id / name) -- index is the only key, valid within ONE read';
  const imgs = document ? document.images : undefined;
  if (Array.isArray(imgs)) return { images: imgs.map(r => ({ ...r })), scope };
  return { images: null, gap: typeof imgs === 'string' ? imgs : 'UNOBSERVABLE:unread', scope };
}

/**
 * GH #269b: the comments of the main document, one row per top-level comment: { id, text, author, userId, timeUtc (UTC ms), quote, resolved, replies: [{ text, author, userId }], anchor }.
 * Every field is a measured value or a NAMED 'UNOBSERVABLE:<reason>'. `anchor` is ALWAYS the named gap -- DS 9.3.5 has no public getter for the commented range, `quote` is the
 * text at CREATION time (docs/GH269-DOCX-COMMENTS-READ-WRITE-RESEARCH.md 2.). `comments` null = the read did not happen (never an empty list); `complete` is true only when
 * the list was read and no field of any row is NAMED unobservable.
 */
const COMMENT_ANCHOR_GAP = 'UNOBSERVABLE:no-public-anchor-getter';
function commentsOfDocument(document) {
  const list = document ? document.comments : null;
  if (!Array.isArray(list)) return { comments: null, complete: false };
  const rows = list.map(c => ({ id: c.id, text: c.text, author: c.author, userId: c.userId, timeUtc: c.timeUtc, quote: c.quote, resolved: c.solved,
    replies: Array.isArray(c.replies) ? c.replies.map(r => ({ text: r.text, author: r.author, userId: r.userId })) : c.replies, anchor: COMMENT_ANCHOR_GAP }));
  const named = v => typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0;
  const clean = rows.every(r => !Object.entries(r).some(([k, v]) => k !== 'anchor' && (named(v) || (k === 'replies' && Array.isArray(v) && v.some(x => Object.values(x).some(named))))));
  return { comments: rows, complete: clean };
}

const DOCX_PROJECTIONS = Object.freeze({
  read_text: inv => ({ blocks: inv.blocks.map(b => ({ kind: b.kind, ...(b.kind === 'heading' ? { level: b.level } : {}),
    ...(b.kind === 'list' ? { ordered: b.ordered, items: b.items.map(i => i.text) } : {}), ...textOf(b), fingerprint: b.locator.fingerprint })) }),
  read_formatting: inv => ({ propertyRead: inv.propertyRead, defaultText: inv.document ? inv.document.defaultText : null, units: unitsOf(inv.blocks) }),
  read_tables: inv => ({ tables: inv.document && inv.document.tables
    ? inv.document.tables.map(t => ({ rows: t.rows.map(r => r.map(cellOf)), format: tableFormatOf(t, inv.document.tableStyles) }))
    : null }),
  read_headers_footers: inv => ({ sections: inv.document && inv.document.sections ? inv.document.sections.map(x => ({ index: x.index, header: x.header, footer: x.footer })) : null }),
  // GH #223: a DEDICATED operation (not a read_headers_footers extension -- docs/GH223-DOCX-PAGE-
  // SETUP-READ-DECISION.md), a plain field-select over inventory.cjs's own per-section pageSetup
  // fields (already measured for the W11 WRITE side's own verifyPreservation; no new driver read).
  // titlePage/evenAndOdd are DELIBERATELY excluded -- a different, runtime-gated axis (B21/#227).
  read_page_setup: inv => ({ sections: inv.document && inv.document.sections ? inv.document.sections.map(x => ({
    index: x.index, pageWidth: x.pageWidth, pageHeight: x.pageHeight, margins: x.margins,
    columnsCount: x.columnsCount, columnsSpaces: x.columnsSpaces, startPageNumber: x.startPageNumber,
  })) : null }),
  read_bookmarks: inv => ({ bookmarks: inv.document ? inv.document.bookmarks : null }),
  // GH #225: see the numberingUnitsOf/numberingOfUnit header comment above for the full field contract.
  read_numbering: inv => ({ items: numberingUnitsOf(inv.blocks || []) }),
  // GH #224: see stylesOfDocument's header comment below for the full field contract.
  read_styles: (inv, args) => stylesOfDocument(inv.document, args),
  // GH #268b: see imagesOfDocument's header comment below.
  read_images: inv => imagesOfDocument(inv.document),
  // GH #227: per-section titlePage/evenAndOdd, already normalized by inventory.cjs's variantFlag()
  // (boolean or a named UNAVAILABLE:<reason> string) -- a plain field-select, no new driver code.
  read_section_flags: inv => ({ sections: inv.document && inv.document.sections ? inv.document.sections.map(x => ({
    index: x.index, titlePage: x.titlePage, evenAndOdd: x.evenAndOdd,
  })) : null }),
  // GH #226 3: the notes' FULL text through the driver's GetNext() chain (decision doc docs/GH226-NOTES-READ-DECISION.md). ACTIVE since the read_notes status flip (PARTIAL -> ACCEPTED, the LIVE above) -- notesOfDocument is the plain field-select.
  read_notes: inv => notesOfDocument(inv.document),
  // GH #269b: the main document's comments (docs/GH269b-DOCX-READ-COMMENTS-DECISION.md). ACTIVE since the read_comments status flip (NOT_MEASURED -> ACCEPTED) (the status is the gate: only ACCEPTED operations
  // are offered/dispatched) -- commentsOfDocument is the plain field-select the status flip will serve, after the g269v1 LIVE probe is green.
  read_comments: inv => commentsOfDocument(inv.document),
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
  // GH #227: includeSectionVariantState is requested ONLY for read_section_flags -- the B21 Bridge READ
  // (docx.section.variantState.read/v1) and its runtime gate stay unexercised for every other docx operation,
  // the SAME opt-in pattern GH #225 established for numberingFormat below.
  // GH #225: numberingFormat (the extra per-paragraph ToJSON() call, docx-driver.cjs:160/240-279) is requested
  // ONLY for read_numbering -- every other docx operation keeps paying zero for it, unchanged by this unit.
  docx: { inventory: docxInventory, read: (d, v) => d.readLive(null, { includeSectionVariantState: v.operation === 'read_section_flags', numberingFormat: v.operation === 'read_numbering', includeComments: v.operation === 'read_comments', ...(v.operation === 'read_styles' && v.args && v.args.names ? { styleNames: v.args.names } : {}) }), project: (p, v) => DOCX_PROJECTIONS[v.operation](p, v.args), complete: p => p.propertyRead === true && typeof p.richUnobservable !== 'string' },
});

/** GH #110: why a read is not complete -- truncation and the inventory's unsupported notes, as 'field: reason'. */
const incompleteReasons = primary => [...(primary.truncated === true ? ['truncated'] : []), ...(typeof primary.richUnobservable === 'string' ? [`rich: ${primary.richUnobservable}`] : []), ...(Array.isArray(primary.unsupported) ? primary.unsupported.map(u => `${u.field}: ${u.reason}`) : [])];

/** One live read on an OPEN driver: readLive -> normalizeInventory -> projection. Never throws. */
async function readOnce(f, driver, v, format) {
  try {
    const r = await f.read(driver, v);
    if (!r || !r.ok) return readFailed((r && r.detail) || 'read failed');
    const primary = f.inventory.normalizeInventory(r.payload);
    if (!primary) return readFailed('the live read produced no inventory');
    const complete = f.complete(primary);
    // GH #110: an incomplete read says why (the observation's own notes), never a bare complete:false
    const incomplete = complete ? null : incompleteReasons(primary);
    return { ok: true, tool: `${format}_read`, fileId: v.fileId, operation: v.operation, complete, ...(incomplete && incomplete.length ? { incomplete } : {}), result: f.project(primary, v) };
  } catch (e) {
    return readFailed(String((e && e.message) || e));
  }
}

const closeDetail = closed => (closed && closed.detail ? String(closed.detail) : null);

/**
 * GH #100 U4c: the warm-session options of ONE read call, or null (then the call is exactly the pre-reuse one:
 * no registry.take/hold, no extra config fetch). `reuse` = { registry, principal, configSignature } -- the SAME
 * shape and the SAME key (fileId, principal, configSignature) write-tools.cjs's warmSessionOptions builds, so a
 * read and a write on the same file+requester contend for the ONE held slot rather than evicting each other
 * (docs/GH100-U4C-READ-REUSE-DESIGN.md sec. 3). Unlike the write side, a read never calls `withSession`; this
 * mirrors (does not reuse) lifecycle.cjs's private admitWarmSession/endCallKeepingSession admission/release shape,
 * because a read has no handle/opened descriptor of its own to carry through an executor (sec. 1, decision B).
 */
function readWarmSession(reuse, configSource, fileId, lease) {
  if (!reuse || !reuse.registry || reuse.registry.enabled !== true) return null;
  if (!configSource || typeof configSource.getEditorConfig !== 'function') return null;
  const id = String(fileId);
  const key = sessionKey({ fileId: id, principal: reuse.principal, configSignature: reuse.configSignature });
  if (!key) return null;
  const { registry } = reuse;
  return {
    key,
    async take() {
      let taken;
      try { taken = await registry.take(key); } catch { return null; }
      if (!taken) return null;
      lease.track(taken.driver);
      let decision;
      try {
        const driver = taken.driver;
        const health = await driver.health();
        const openedKey = driver.documentKey();
        let freshKey = null;
        try {
          const config = await configSource.getEditorConfig({ fileId: id, operation: 'READ' });
          const k = config && config.document ? config.document.key : null;
          freshKey = typeof k === 'string' && k !== '' ? k : null;
        } catch { freshKey = null; } // a failed fetch is a "no", never "unchanged"
        decision = taken.opened
          ? decideReuse({ health, openedKey, freshKey, lastRelease: taken.lastRelease, requestedOperation: 'READ', openedOperation: taken.opened.operation })
          : { reuse: false, reason: 'no-recorded-open-descriptor' };
        if (decision.reuse) {
          const reset = driver.resetForNextCall();
          if (!reset || reset.ok !== true) decision = { reuse: false, reason: 'reset-failed' };
        }
      } catch {
        decision = { reuse: false, reason: 'reuse-check-threw' };
      }
      // carry the ORIGINAL openedAt/calls/opened descriptor through -- a kept
      // chain must still age toward maxAge/maxCalls (R6), and a session another operation opened
      // (e.g. WRITE) must come back with that SAME descriptor so a later take() of either operation
      // still sees who really opened it, not what THIS call happened to be.
      if (decision.reuse) return { driver: taken.driver, reused: true, reason: decision.reason, openedAt: taken.openedAt, calls: taken.calls, opened: taken.opened };
      try { await taken.driver.close(); } catch { /* best-effort: a dead driver is simply not served */ }
      return null;
    },
    /**
     * hold(driver, session) -> true if kept (the caller must NOT close it), false if NOT kept (the
     * caller must close it -- hold() never closes it itself, so the driver is closed EXACTLY once).
     * ONLY a call whose read actually succeeded may be offered, and ONLY on
     * a REAL driver.release() (never a fabricated confirmation) -- a write-free call accepts a
     * null/confirmed release, exactly the same acceptance rule write-tools.cjs's endCallKeepingSession
     * uses (wroteVerified=false here, a read never dispatches). `session` = { openedAt, calls, opened }
     * from the ORIGINAL open (fresh) or the ORIGINAL taken descriptor (reused) -- never reset to
     * "now"/1 (point 3, R6).
     */
    async hold(driver, session) {
      let released;
      try { released = await driver.release(); } catch (err) { released = { ok: false, detail: String((err && err.message) || err) }; }
      const confirmation = (released && released.confirmation) || null;
      const clean = released && released.ok === true && (!confirmation || confirmation.state === 'confirmed');
      if (!clean) return false;
      let held;
      try {
        held = await registry.hold(key, driver, { openedAt: session.openedAt, calls: session.calls, lastRelease: { ok: true, confirmation }, opened: session.opened });
      } catch { held = { held: false, reason: 'registry-hold-threw' }; }
      return !!(held && held.held === true);
    },
  };
}

/**
 * createReadTool(format, { configSource, browsersPath, gate, runtime, loadPlaywright, createDriver, reuse }) -> { description, inputSchema, handler(args) }
 *
 * M11.3b.2 request path (each step refuses BEFORE the next one starts):
 *   validate (pre-dispatch: schema, credential/requester fields, operation class)
 *   -> admission: the process-wide session gate (MAX_ACTIVE_SESSIONS_PER_PROCESS=1) -> BUSY, no queue
 *   -> browser runtime check (pinned playwright-core / revision under playwright.path) -> BROWSER_RUNTIME_UNAVAILABLE
 *   -> GH #100 U4c: if `reuse` is wired and enabled, try the held warm session first (readWarmSession.take()) ->
 *      driver open (the broker call happens here) only on a miss -> readLive -> projection ->
 *      a clean call holds the session for the next call (readWarmSession.hold()) instead of closing it ->
 *      any other exit closes it, exactly as before reuse existed -> lease released.
 * `gate`: the product passes ONE gate shared by every tool of the process. `runtime` / `loadPlaywright` / `createDriver`
 * are test seams; the product default is the configured browser runtime and productDriver (broker configSource only).
 */
function createReadTool(format, { configSource, browsersPath, gate = createSessionGate(), runtime = null, loadPlaywright = null, createDriver = productDriver, reuse = null } = {}) {
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
        const warm = readWarmSession(reuse, configSource, v.fileId, lease);
        const taken = warm ? await warm.take() : null;
        let driver; let session;
        if (taken) {
          driver = taken.driver;
          session = { openedAt: taken.openedAt, calls: taken.calls + 1, opened: taken.opened };
        } else {
          driver = createDriver(format, { configSource, loadPlaywright: rt.load });
          lease.track(driver);
          const opened = (await driver.open({ fileId: v.fileId, operation: 'READ' })) || {};
          if (!opened.ok) return { ok: false, stage: 'OPEN', code: opened.code || 'OPEN_FAILED', reason: opened.detail || 'open failed', dispatched: false };
          session = { openedAt: Date.now(), calls: 1, opened: { operation: 'READ', id: opened.id, frame: opened.frame || null, apiWhere: opened.apiWhere || null } };
        }
        const out = await readOnce(f, driver, v, format);
        // only a call whose read actually SUCCEEDED may be offered back --
        // a failed read closes, exactly as a failed write never reaches endCallKeepingSession.
        if (warm && out.ok === true && (await warm.hold(driver, session))) {
          out.close = null; // held, not closed -- the next call's take()/open() carries the real close outcome, if any
        } else {
          out.close = closeDetail(await driver.close().catch(e => ({ detail: String((e && e.message) || e) })));
        }
        return out;
      } finally {
        lease.release();
      }
    },
  };
}

// readOnce/FORMATS: GH #76, the optional product-side verifyStoredContent (write-tools.cjs) opens its own
// targeted read on an independently-opened driver -- exported so it calls the SAME read path this tool
// uses, not a second one.
module.exports = { OPERATIONS, STATUS, validateReadRequest, inputSchema, createReadTool, productDriver, xlsxRequest, writeIntents, CREDENTIAL_FIELDS, REQUESTER_FIELDS, readOnce, FORMATS, readRange };
