'use strict';

const { BRIDGE_CAPABILITIES, BRIDGE_CLIENT_SRC, BRIDGE_HYDRATIONS, BRIDGE_HYDRATE_CLIENT_SRC, BRIDGE_WRITES, BRIDGE_WRITE_CLIENT_SRC, BRIDGE_SERIES_STYLE_WRITE_CLIENT_SRC, BRIDGE_COPY_WRITE_CLIENT_SRC, BRIDGE_CHART_DELETE_WRITE_CLIENT_SRC } = require('./ds-bridge.cjs');
const { ACCEPTED_RUNTIMES, XLSX_RUNTIME_IDS, XLSX_B05_BRIDGE_WRITE_RUNTIME_ENTRIES, XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B10_BRIDGE_WRITE_RUNTIME_ENTRIES, XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B03_BRIDGE_WRITE_RUNTIME_ENTRIES, XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B22_BRIDGE_WRITE_RUNTIME_ENTRIES, XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B23_BRIDGE_WRITE_RUNTIME_ENTRIES, XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B24_BRIDGE_WRITE_RUNTIME_ENTRIES, XLSX_B24_BRIDGE_WRITE_RUNTIME_IDS, cacheKeyOf, HANDSHAKE_BODY, classifyHandshake } = require('../../core/runtime-identity.cjs');

const { startConfigHost, readDsOpenState, openWaitFailureDetail } = require('./config-host.cjs');
const { observePage, settleSaveAck, dsIdentity, confirmForceSave, closeResult } = require('./save-ack.cjs');
const { defaultWait } = require('../../core/session/open.cjs');
const { resolveConfigSource } = require('./config-source.cjs');
const { wireBody, decodeWire } = require('./wire-codec.cjs');

/**
 * Euro-Office XLSX editor transport -- the M1 driver contract on the spreadsheet editor (M6).
 *
 *   open / close / forceSave / canSave / readLive / dispatch / readPackage
 *
 * TRANSPORT only: M1 decides when each is called, the XLSX adapter decides what a command means.
 * The open/close/forceSave/canSave/dispatch path is the DOCX driver's (docx-driver.cjs), re-expressed
 * for the spreadsheet editor frame; it is duplicated rather than shared on purpose, so the LIVE-proven
 * DOCX transport is not touched by M6. The route is the one the owner fixed: the NC eurooffice config
 * over HTTPS and the Document Server it names -- no SSH, no .10.
 *
 * STATUS: written from the source's measured getters (range-reader.cjs, xlsx-merge-observer.cjs,
 * workbook-inspector.cjs); the M6 wave-0 LIVE probe (test/live/xlsx-wave0-probe.live.cjs) must
 * confirm them on the live server before this driver carries a LIVE acceptance.
 *
 * READS ARE READ-ONLY BY CONSTRUCTION: readQuery() builds a body that calls only getters, ForEach and
 * WorksheetFunction.COUNT (static gate: test/static/xlsx-driver-read-query.test.cjs).
 */

const MAX_READ_CELLS = 45000;
// GH #300 K1: the read_comments anchor scan (per-cell GetComment, no public address getter) and the comment count it covers
const MAX_COMMENT_SCAN_CELLS = 20000;
const MAX_COMMENTS = 1000;

/**
 * M7 S32 slice 1 (owner GO 2026-10-03, owner channel; contract docs/evidence/M7-S32-TIERED-CONTRACT.md): above
 * MAX_READ_CELLS and up to MAX_TIERED_CELLS used cells, a write_range / clear_range task is observed in TIERS -- every
 * cell of the task's windows (T1, the reader below) plus a per-sheet digest of the content outside them (T2). An
 * IMPLEMENTATION limit of this product, not an editor limit.
 */
const MAX_TIERED_CELLS = 50000;

/**
 * M7 S32 slice 2 (owner GO 2026-10-03, owner channel; plan docs/M7-S32B-FORMULA-INDEX-PLAN.md): T3, the FORMULA INDEX --
 * every formula cell outside the task's windows (address + formula text, no value), read in the same scan as the T2
 * digest, for a tiered request that needs the dependency analysis. The ceiling is a starting IMPLEMENTATION limit:
 * over it the observation is incomplete and the task is refused before any write.
 */
const MAX_FORMULA_INDEX = 5000;
/** The extras a tiered request may carry (slice 2): the dependency readers and the AutoFilter read of named sheets. */
// X-T01b: formatSheets too -- on a tiered read the digest (T2) carries each cell's number format on those sheets
// X-L13/6 (fdb66c26's code-level survey): runCells (keyed by the exact requested sheet+cell pair), panes +
// sheetVisibility (per-SHEET, not per-cell), cfRanges and validationRanges (each keyed by the exact requested
// sheet+range pair) -- none of the four scans the used range; their cost is bounded by the TASK's own request
// shape, not by the workbook's used-cell count.
// GH #327 (U5): protectionRanges joins the same (keyed by the exact requested sheet+range pair) group -- the
// bridge handler's own range walk is capped at 10000 non-empty cells, so its cost is bounded the same way.
// GH #326 (U4b): protectionSheets joins too -- it never walks a range at all (sheet.protection.read/v1
// without `range`), so its cost is bounded by the task's own sheet count, not even the 10000-cell cap.
// X-L13/8 (the (b)-group BOUNDED classification, owner decision): `layout` joins the list too, on a
// DIFFERENT justification -- its per-sheet cost is capped by MAX_LAYOUT_ITEMS (2000, maxCol+maxRow), and a
// sheet over that cap degrades to a NAMED 'UNOBSERVABLE:layout-truncated' entry for that sheet, fail-closed/
// fail-named rather than silently truncating the whole read. Below the 45000-cell bound the SAME reader already
// runs unconditionally for every sheet (readQuery's `layout` axis is not itself gated by tier) -- tiering a
// request that ASKS for `layout` therefore does not change what the reader does, only whether the REQUEST as a
// whole is still eligible above the bound. *** THIS DOES NOT RESOLVE filter_range's STOP (X-L13/6): that
// family's OWN gate() needs a precise per-row hidden-state read that layout's maxCol/maxRow summary cannot
// serve -- a separate, still-open question, unaffected by this entry. *** `layout_range` needs `layout` in its
// OWN bounded request too (not just its full one): without it, not even its own write can be proven under
// BOUNDED (layout.cjs's gate()/verifyPrimitive() read the op's own target through `primary.layout[sheet]`).
const TIERED_EXTRAS = Object.freeze(['dependents', 'filterSheets', 'formatSheets', 'runCells', 'panes', 'sheetVisibility', 'cfRanges', 'validationRanges', 'layout', 'protectionRanges', 'protectionSheets', 'commentSheets', 'sessionUser']);

/**
 * The live read: the sheet list with used ranges, every cell of every used range, plus the requested
 * windows ([{ sheet, range }]), stopping at MAX_READ_CELLS (then `truncated`). Per cell: GetValue,
 * GetFormula, COUNT of the single-cell range (number 1 / text 0), the MergeArea address. A getter that
 * is missing or throws is recorded in `unsupported` -- the adapter then treats the read as incomplete.
 */
/**
 * Cell-level BOLD / ITALIC (B01 / B02) -- MIGRATED to the DS Bridge (owner GO, owner channel, internal tracker point 8;
 * LIVE parity: B01 pilot #ref, B02 Wave A #ref). The 9.3.4 cell SDK has no public cell-level font flag getter, so
 * the axis is read through the advertised `textStyle.bold.read/v1` / `textStyle.italic.read/v1` capabilities -- the
 * direct `range.getFont(false).getBold/getItalic` call and its build allowlist are GONE. The bridge's own gates
 * (bridge present, major 1, manifest build == the editor build, capability advertised, result shape) replace the
 * allowlist; any gap is 'UNOBSERVABLE:<reason>' for that axis only, NEVER false, and there is NO fallback to a direct
 * internal read.
 */

/**
 * B03-WRITE through the DS Bridge (M11.5c.4, owner: M11.5 continuous mandate, internal tracker): the driver's second
 * line behind the adapter's sheets.B03_BRIDGE_WRITE_RUNTIMES (the same /8 .. /12 list, L-05) -- a copy_sheet command (always a
 * worksheet.copy/v1 command, meta.bridgeCopyWrite) is refused WITHOUT touching the editor on any other runtime. The former
 * EO_INTERNAL_COPY_RUNTIMES direct-internal path (the frame-side asc_copyWorksheet unit) is gone; there is no fallback.
 */
const B03_BRIDGE_WRITE_RUNTIMES = XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS;

/**
 * Cell alignment (horizontal / vertical) and borders (B04) -- MIGRATED to the DS Bridge `range.appearance.read/v1`
 * (internal tracker point 8; LIVE parity #ref incl. the row-2 shared edge). The direct getAlign/getBorderFull read for
 * these two axes and EO_INTERNAL_APPEARANCE_BUILDS are gone; the align-EXT axes (B05 READ) read through their own
 * capability, `range.alignmentExt.read/v1` (Wave B READ, below).
 */

/**
 * align.readingOrder / align.indent / align.shrink READ (B05 READ, X-F07b/c) -- MIGRATED to the DS Bridge
 * `range.alignmentExt.read/v1` (Wave B READ, patch family /4, docs/DS-BRIDGE-WAVE-B-READ-DESIGN-20261008.md 3.1). The
 * direct getAlign / getReadingOrder / getIndent / getShrinkToFit read and EO_INTERNAL_ALIGN_EXT_RUNTIMES are gone. The
 * WRITE half (setIndent / setShrinkToFit, families/format.cjs) is still internal (a later wave).
 */

/**
 * A sheet's OWN freeze panes WITHOUT activating it (B06) -- MIGRATED to the DS Bridge `sheet.freezeState.read/v1`
 * (internal tracker point 8; LIVE parity #ref on an active AND a non-active sheet). The public GetFreezePanes() reads the
 * ACTIVE sheet, so it never replaced this read; the direct worksheet-model read and EO_INTERNAL_PANE_BUILDS are gone.
 */

/**
 * An AutoFilter filter's FIELD identity (B07) -- MIGRATED to the DS Bridge `sheet.autoFilter.fieldIdentity.read/v1`
 * (internal tracker point 8; LIVE parity #ref). The public filter wrapper has no field getter; the private
 * `filter.ColId` read and EO_INTERNAL_FILTER_BUILDS are gone. A missing field is never "no filter".
 */

/**
 * The PUBLIC page-layout getters (orientation, gridlines, headings, the four margins) -- a public API, and whether
 * they read the TARGET sheet or the ACTIVE one had to be measured (W4b2 found GetFreezePanes reading the active
 * sheet). PROVEN on 9.3.4-9475e9f1 by the W8b probe (2026-09-29, case L4, docs/evidence/m6-w8b-probe-20260929.json):
 * with another sheet active, each sheet read its OWN layout, in session and after reopen. Any other build: refused.
 */
const EO_PAGE_LAYOUT_PER_SHEET_RUNTIMES = XLSX_RUNTIME_IDS;

/**
 * The print settings READ (B09: fit width / height and the fitToPage flag, the header/footer texts and flags, the row
 * breaks, the first page number, centering, scale, paper size) -- MIGRATED to the DS Bridge `sheet.printSettings.read/v1`
 * (Wave B READ, patch family /4, design 3.2). The value keeps the W8b-measured normalisations; the direct PagePrintOptions
 * / headerFooter / rowBreaks read and EO_INTERNAL_PRINT_RUNTIMES are gone.
 */
/**
 * The canonical persisted width of a REQUESTED column width (B17, the engine's pure charCountToModelColWidth ->
 * modelColWidthToColWidth -> colWidthToCharCount chain) -- MIGRATED to the DS Bridge `sheet.columnWidth.canonicalize/v1`
 * (internal tracker point 8; LIVE parity #ref: 8.43 / 15 / 20.5 / 33.3 -> 8.43 / 15 / 20.43 / 33.29). Read only, before
 * any write; the direct worksheet-function read and EO_INTERNAL_COLWIDTH_CANON_BUILDS are gone.
 */

/**
 * B10-WRITE through the DS Bridge (M11.5b.4, owner: M11.5 continuous mandate, internal tracker): the driver's second
 * line behind the adapter's print.B10_BRIDGE_WRITE_RUNTIMES (the same /7-only list) -- a print command (always a
 * sheet.printSettings.write/v1 command, meta.bridgePrintWrite) is refused WITHOUT touching the editor on any other runtime.
 * The former EO_INTERNAL_PRINT_WRITE_RUNTIMES direct-internal path is gone; there is no fallback to it.
 */
const B10_BRIDGE_WRITE_RUNTIMES = XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS;
/**
 * B05-WRITE through the DS Bridge (M11.5a.4, owner GO owner channel, internal tracker): the driver's second line
 * behind the adapter's format.B05_BRIDGE_WRITE_RUNTIMES (the same /6-only list) -- a format command that carries a B05
 * Bridge WRITE (meta.bridgeAlignExtWrite) is refused WITHOUT touching the editor on any other runtime.
 */
const B05_BRIDGE_WRITE_RUNTIMES = XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS;
/**
 * gapOverlap/seriesOrder-WRITE through the DS Bridge (X-O05/3, internal tracker; reviewed and
 * accepted): the driver's second line behind the adapter's charts.B22_BRIDGE_WRITE_RUNTIMES
 * (the same /12 .. /14 list, the handler byte-identical on each) -- a set_chart_style command that
 * carries one of these two Bridge WRITEs (meta.bridgeChartStyleWrite) is refused WITHOUT touching
 * the editor on any other runtime.
 */
const B22_BRIDGE_WRITE_RUNTIMES = XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS;
/**
 * chart-delete-WRITE through the DS Bridge (X-O07/3, internal tracker; X-O07/2 LIVE MEASURED and
 * reviewed): the driver's second line behind the adapter's charts.B23_BRIDGE_WRITE_RUNTIMES (the
 * same /13 .. /14 list, the handler byte-identical on each) -- a delete_chart command
 * (meta.bridgeChartDeleteWrite) is refused WITHOUT touching the editor on any other runtime.
 */
const B23_BRIDGE_WRITE_RUNTIMES = XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS;
/**
 * marker-symbol/line-style-WRITE through the DS Bridge (X-O05/5-C, internal tracker; X-O05/5-B
 * round 7b+8d+8k+8m+8q LIVE MEASURED and reviewed): the driver's second line behind the
 * adapter's charts.B24_BRIDGE_WRITE_RUNTIMES (/15-ONLY -- UNLIKE B22/B23, neither
 * chart.series.marker.write/v1 nor chart.series.lineStyle.write/v1 exists on any earlier
 * runtime) -- a set_chart_style command carrying `markerSymbol`/`lineStyle`
 * (meta.bridgeSeriesStyleWrite) is refused WITHOUT touching the editor on any other runtime.
 */
const B24_BRIDGE_WRITE_RUNTIMES = XLSX_B24_BRIDGE_WRITE_RUNTIME_IDS;

/**
 * The in-editor B05 client, prepended ONLY to a command that writes B05 (every other command stays byte-identical):
 *   __eoB05Gate(sheet)              the Bridge WRITE gate, ONCE, before any mutation of the command
 *   __eoB05Write(cell, args, req)   ONE write() for one cell, only the requested keys
 *   __eoB05Fail(cell, w)            the failure report: the bridge reason, partial / applied, the cells already written,
 *                                   and the OBSERVED indent / shrink of those cells and the failed one (a fresh Bridge
 *                                   READ, diagnostics only -- the executor's own fresh observation stays the authority)
 */
function b05BridgeWritePrelude() {
  const W = BRIDGE_WRITES.ALIGN_EXT_WRITE; const R = BRIDGE_CAPABILITIES.ALIGN_EXT_READ;
  return `${BRIDGE_CLIENT_SRC}${BRIDGE_WRITE_CLIENT_SRC}
  let __b05 = null; const __b05done = [];
  const __eoB05Gate = sheet => { const g = __eoBridgeWriteGate(${JSON.stringify(W.id)}, ${W.version}, ${JSON.stringify(XLSX_B05_BRIDGE_WRITE_RUNTIME_ENTRIES)}); if (!g.ok) return { ok: false, stage: 'bridge-write-gate', reason: g.reason, mutation: 'NONE' }; __b05 = { bridge: g.bridge, sheet }; return { ok: true }; };
  const __eoB05Write = (cell, args, requested) => { const w = __eoBridgeWrite(__b05.bridge, ${JSON.stringify(W.id)}, Object.assign({ sheet: __b05.sheet, cell }, args), requested); if (w.ok) __b05done.push(cell); return w; };
  const __eoB05Fail = (cell, w) => { const observed = {}; for (const a of __b05done.concat([cell])) { const r = __eoBridgeRead(${JSON.stringify(R.id)}, ${R.version}, ${JSON.stringify(ACCEPTED_RUNTIMES.xlsx)}, { sheet: __b05.sheet, cell: a }, ${JSON.stringify(R.valueType)}); observed[a] = r.ok ? { indent: r.value.indent, shrink: r.value.shrink } : 'UNOBSERVABLE:' + r.reason; } return { ok: false, stage: 'bridge-write', cell, reason: w.reason, uncertain: w.uncertain === true, partial: w.partial === true, applied: w.applied || [], cellsDone: __b05done.slice(), observed }; };
`;
}

/**
 * The in-editor B10 client, prepended ONLY to a print command (every other command stays byte-identical):
 *   __eoB10Gate(sheet)              the Bridge WRITE gate, ONCE, before any mutation of the command
 *   __eoB10Write(args, requested)   the ONE write() of the command
 *   __eoB10Fail(w)                  the failure report: the bridge reason, partial / applied, and the OBSERVED print
 *                                   settings of the sheet (a fresh Bridge READ, diagnostics only -- the executor's own
 *                                   fresh observation stays the authority)
 */
function b10BridgeWritePrelude() {
  const W = BRIDGE_WRITES.PRINT_SETTINGS_WRITE; const R = BRIDGE_CAPABILITIES.PRINT_SETTINGS_READ;
  return `${BRIDGE_CLIENT_SRC}${BRIDGE_WRITE_CLIENT_SRC}
  let __b10 = null;
  const __eoB10Gate = sheet => { const g = __eoBridgeWriteGate(${JSON.stringify(W.id)}, ${W.version}, ${JSON.stringify(XLSX_B10_BRIDGE_WRITE_RUNTIME_ENTRIES)}); if (!g.ok) return { ok: false, stage: 'bridge-write-gate', reason: g.reason, mutation: 'NONE' }; __b10 = { bridge: g.bridge, sheet }; return { ok: true }; };
  const __eoB10Write = (args, requested) => __eoBridgeWrite(__b10.bridge, ${JSON.stringify(W.id)}, Object.assign({ sheet: __b10.sheet }, args), requested);
  const __eoB10Fail = w => { const r = __eoBridgeRead(${JSON.stringify(R.id)}, ${R.version}, ${JSON.stringify(ACCEPTED_RUNTIMES.xlsx)}, { sheet: __b10.sheet }, ${JSON.stringify(R.valueType)}); return { ok: false, stage: 'bridge-write', reason: w.reason, uncertain: w.uncertain === true, partial: w.partial === true, applied: w.applied || [], observed: r.ok ? r.value : 'UNOBSERVABLE:' + r.reason }; };
`;
}

/**
 * The dispatch result of a B05 command, from the editor result: ok ONLY when the command completed AND its own body
 * returned { ok: true } (every Bridge WRITE accepted). A partial receipt -> partiallyApplied (the core records it
 * uncertain); any other failure -> not ok. Never APPLIED on the receipt alone: L1 still runs on the fresh READ.
 */
function b05DispatchResult(r, tag = 'b05') {
  const res = r && r.result;
  if (r && r.ok === true && res && res.ok === true) return { ok: true, dispatched: true, outcome: r.outcome, error: null };
  if (r && r.ok === true && res && res.stage === 'bridge-write' && res.partial === true) return { ok: true, dispatched: true, partiallyApplied: true, outcome: `${tag}-bridge-write-partial`, evidence: res, error: null };
  return { ok: false, dispatched: true, outcome: r && r.ok === true ? `${tag}-bridge-write-failed` : (r && r.outcome) || `${tag}-command-failed`, evidence: res || null, error: (r && r.error) || null };
}
/** The same rule for a B10 (print) command. */
const b10DispatchResult = r => b05DispatchResult(r, 'b10');
/** The same rule for a chart-style gapOverlap/seriesOrder WRITE command. */
const chartStyleDispatchResult = r => b05DispatchResult(r, 'chartStyle');

/**
 * X-O05/3 (internal tracker; reviewed and accepted): the in-editor chart-style client, prepended
 * ONLY to a set_chart_style command that writes gapOverlap/seriesOrder (every other command stays
 * byte-identical). Parametrized by the specific WRITE capability, unlike B05/B10's one-prelude-
 * per-capability shape -- X-O05/3 has TWO (chart.bar.gapOverlap.write/v1,
 * chart.series.order.write/v1) sharing the identical mechanism, same convention bridgeAxis already
 * uses for multiple READ capabilities through one helper:
 *   __eoChartStyleGate(sheet, chart)       the Bridge WRITE gate, ONCE, before any mutation
 *   __eoChartStyleWrite(args, requested)   the ONE write() of the command
 *   __eoChartStyleFail(w)                  the failure report: bridge reason, partial/applied --
 *                                          NO diagnostic re-read here (unlike B05/B10): the
 *                                          adapter's own classify()/chartReads loop already
 *                                          re-observes fresh on the next round, the SAME path
 *                                          'marker'/'gridlines' already rely on for their own
 *                                          bridge-backed WRITE targets.
 */
function chartStyleBridgeWritePrelude(cap) {
  return `${BRIDGE_CLIENT_SRC}${BRIDGE_WRITE_CLIENT_SRC}
  let __csw = null;
  const __eoChartStyleGate = (sheet, chart) => { const g = __eoBridgeWriteGate(${JSON.stringify(cap.id)}, ${cap.version}, ${JSON.stringify(XLSX_B22_BRIDGE_WRITE_RUNTIME_ENTRIES)}); if (!g.ok) return { ok: false, stage: 'bridge-write-gate', reason: g.reason, mutation: 'NONE' }; __csw = { bridge: g.bridge, sheet, chart }; return { ok: true }; };
  const __eoChartStyleWrite = (args, requested) => __eoBridgeWrite(__csw.bridge, ${JSON.stringify(cap.id)}, Object.assign({ sheet: __csw.sheet, chart: __csw.chart }, args), requested);
  const __eoChartStyleFail = w => ({ ok: false, stage: 'bridge-write', reason: w.reason, uncertain: w.uncertain === true, partial: w.partial === true, applied: w.applied || [] });
`;
}

/**
 * X-O05/5-C (internal tracker; X-O05/5-B round 7b+8d+8k+8m+8q LIVE MEASURED, reviewed and
 * accepted): the in-editor marker-symbol/line-style client, prepended ONLY to a set_chart_style
 * command carrying `markerSymbol` or `lineStyle` (every other command stays byte-identical). TWO
 * capabilities (chart.series.marker.write/v1, chart.series.lineStyle.write/v1) share this one
 * gate, the SAME convention `chartStyleBridgeWritePrelude` above already uses for gapOverlap/
 * seriesOrder -- but on the B24 runtime list (/15-ONLY: neither capability exists on /12../14,
 * unlike B22), not B22's:
 *   __eoSeriesStyleGate(sheet, chart)       the Bridge WRITE gate, ONCE, before any mutation
 *   __eoSeriesStyleWrite(args, requested)   the ONE write() of the command
 *   __eoSeriesStyleFail(w)                  the failure report: bridge reason, partial/applied
 * The actual bridge write call and its receipt-shape validation live in ds-bridge.cjs's own
 * `BRIDGE_SERIES_STYLE_WRITE_CLIENT_SRC` (`__eoBridgeSeriesStyleWrite`) -- the SAME architectural
 * boundary every other Bridge write respects (the m11-5a2 dual-runtime-compat static test's own
 * pin: that one call site exists ONLY in ds-bridge.cjs, nowhere else under `src/`). This prelude
 * only wires the gate and forwards to it, exactly like `chartDeleteBridgeWritePrelude` above
 * forwards to `__eoBridgeChartDeleteWrite`.
 */
function seriesStyleBridgeWritePrelude(cap) {
  return `${BRIDGE_CLIENT_SRC}${BRIDGE_WRITE_CLIENT_SRC}${BRIDGE_SERIES_STYLE_WRITE_CLIENT_SRC}
  let __ssw = null;
  const __eoSeriesStyleGate = (sheet, chart) => { const g = __eoBridgeWriteGate(${JSON.stringify(cap.id)}, ${cap.version}, ${JSON.stringify(XLSX_B24_BRIDGE_WRITE_RUNTIME_ENTRIES)}); if (!g.ok) return { ok: false, stage: 'bridge-write-gate', reason: g.reason, mutation: 'NONE' }; __ssw = { bridge: g.bridge, sheet, chart }; return { ok: true }; };
  const __eoSeriesStyleWrite = (args, requested) => __eoBridgeSeriesStyleWrite(__ssw.bridge, ${JSON.stringify(cap.id)}, Object.assign({ sheet: __ssw.sheet, chart: __ssw.chart }, args), requested);
  const __eoSeriesStyleFail = w => ({ ok: false, stage: 'bridge-write', reason: w.reason, uncertain: w.uncertain === true, partial: w.partial === true, applied: w.applied || [] });
`;
}

/**
 * X-O07/3 (internal tracker; X-O07/2 LIVE MEASURED and reviewed): the in-editor chart-delete
 * client, prepended ONLY to a delete_chart command (every other command stays byte-identical):
 *   __eoChartDeleteGate(sheet, chart)   the Bridge WRITE gate, ONCE, before any mutation
 *   __eoChartDeleteWrite()              the ONE write() of the command
 *   __eoChartDeleteFail(w)              the failure report: bridge reason, mutation NONE/UNKNOWN
 *                                       -- NO diagnostic re-read here, same reasoning
 *                                       chartStyleBridgeWritePrelude's own header comment
 *                                       records: the adapter's own classify()/chartReads loop
 *                                       re-observes fresh on the next round.
 */
function chartDeleteBridgeWritePrelude() {
  const W = BRIDGE_WRITES.CHART_DELETE_WRITE;
  return `${BRIDGE_CLIENT_SRC}${BRIDGE_WRITE_CLIENT_SRC}${BRIDGE_CHART_DELETE_WRITE_CLIENT_SRC}
  let __cdw = null;
  const __eoChartDeleteGate = (sheet, chart) => { const g = __eoBridgeWriteGate(${JSON.stringify(W.id)}, ${W.version}, ${JSON.stringify(XLSX_B23_BRIDGE_WRITE_RUNTIME_ENTRIES)}); if (!g.ok) return { ok: false, stage: 'bridge-write-gate', reason: g.reason, mutation: 'NONE' }; __cdw = { bridge: g.bridge, sheet, chart }; return { ok: true }; };
  const __eoChartDeleteWrite = () => __eoBridgeChartDeleteWrite(__cdw.bridge, ${JSON.stringify(W.id)}, { sheet: __cdw.sheet, chart: __cdw.chart });
  const __eoChartDeleteFail = w => ({ ok: false, stage: 'bridge-write', reason: w.reason, uncertain: w.uncertain === true, mutation: w.mutation });
`;
}
/** The same dispatch rule for a chart-delete WRITE command. */
const chartDeleteDispatchResult = r => b05DispatchResult(r, 'chartDelete');
/** The same rule for a marker-symbol/line-style WRITE command (X-O05/5-C). */
const seriesStyleDispatchResult = r => b05DispatchResult(r, 'seriesStyle');

/**
 * The in-editor B03 client, prepended ONLY to a copy_sheet command (every other command stays byte-identical):
 *   __eoB03Copy(sheet, name)   the Bridge WRITE gate, then ONE worksheet.copy/v1 write() through __eoBridgeCopyWrite
 *                              (ds-bridge.cjs: the receipt rule; a refusal after the copy call is uncertain, mutation
 *                              UNKNOWN). Never retried here; the core records the outcome and re-reads.
 */
function b03BridgeWritePrelude() {
  const W = BRIDGE_WRITES.WORKSHEET_COPY_WRITE;
  return `${BRIDGE_CLIENT_SRC}${BRIDGE_WRITE_CLIENT_SRC}${BRIDGE_COPY_WRITE_CLIENT_SRC}
  const __eoB03Copy = (sheet, name) => {
    const g = __eoBridgeWriteGate(${JSON.stringify(W.id)}, ${W.version}, ${JSON.stringify(XLSX_B03_BRIDGE_WRITE_RUNTIME_ENTRIES)});
    if (!g.ok) return { ok: false, stage: 'bridge-write-gate', reason: g.reason, mutation: 'NONE' };
    const w = __eoBridgeCopyWrite(g.bridge, ${JSON.stringify(W.id)}, { sheet, name });
    return w.ok ? { ok: true, index: w.index } : Object.assign({ stage: 'bridge-write' }, w);
  };
`;
}
/** The same dispatch rule for a B03 (copy_sheet) command: ok ONLY when the body returned { ok: true }. */
const b03DispatchResult = r => b05DispatchResult(r, 'b03');

/**
 * GH #326 (U4b), the race condition: prepended ONLY to a command whose `meta.protectionRaceCheck` is set (every
 * other command stays byte-identical) -- a command dispatches on a document the PLANNING observation already
 * found either unprotected, or protected and specifically allowed (an unlocked target, or a named `forbids` flag
 * false). Right before the command's OWN first setter runs, `sheet.protection.read/v1` is re-read, in THIS SAME
 * editor session, for every `{sheet, range?}` the gate checked -- the SAME args shape it was read with at planning
 * (a range-level check stays range-level; a sheet-level one stays sheet-level). The comparison is not a re-run of
 * the gate's own predicate: it is the RAW snapshot, byte-for-byte (JSON.stringify) -- stricter, and it needs no
 * per-intent logic duplicated in-editor. ANY difference (protected flipped either way, a forbids flag, an unlocked
 * cell, or the read becoming unobservable) returns BEFORE the `try` block below even starts, so no setter in the
 * rest of `command.body` is ever reached -- 'protection-changed-before-mutation', no DispatchResult wrapper (the
 * generic `dispatch()` branch, and the b05-style ones that happen to wrap this command too, both already treat a
 * non-{ok:true} body result as a clean failure; evidence.outcome still names the real reason either way).
 */
function protectionRacePrelude(checks) {
  const R = BRIDGE_CAPABILITIES.PROTECTION_READ;
  // its own block: a bridge-write prelude in front of it declares the SAME top-level `const __eoBridgeRead`
  // (BRIDGE_CLIENT_SRC), and two of those in one callCommand body is a SyntaxError -- the block shadows instead
  return `{${BRIDGE_CLIENT_SRC}
  const __eoProtChecks = ${JSON.stringify(checks)};
  let __eoProtFail = null;
  for (const __c of __eoProtChecks) {
    const __args = __c.range !== undefined ? { sheet: __c.sheet, range: __c.range } : { sheet: __c.sheet };
    const __r = __eoBridgeRead(${JSON.stringify(R.id)}, ${R.version}, ${JSON.stringify(ACCEPTED_RUNTIMES.xlsx)}, __args, ${JSON.stringify(R.valueType)});
    // the SAME shape the baseline was stored in: protectionReads (sheet-only) strips \`cells\` entirely
    // (U4a's own scope); protectionRangeReads (a range given) keeps it whole -- a bare re-read's extra
    // \`cells: null\` on the sheet-only axis would otherwise mismatch a snapshot that never carried it.
    const __got = !__r.ok ? 'UNOBSERVABLE:' + __r.reason : (__c.range !== undefined ? __r.value : { protected: __r.value.protected, forbids: __r.value.forbids });
    if (JSON.stringify(__got) !== JSON.stringify(__c.expected)) { __eoProtFail = { sheet: __c.sheet, range: __c.range || null }; break; }
  }
  if (__eoProtFail) return { ok: false, stage: 'protection-race', outcome: 'protection-changed-before-mutation', sheet: __eoProtFail.sheet, range: __eoProtFail.range };
}
`;
}

/**
 * A sheet's COLUMN page breaks (B12) -- MIGRATED to the DS Bridge `sheet.columnBreaks.read/v1` (internal tracker point 8;
 * LIVE parity #ref). Same contract as before (a present null -> []; each element { id, man, min, max } validated; any
 * gap 'UNOBSERVABLE:<why>'); the direct `worksheet.colBreaks.getBreaks()` read and EO_INTERNAL_COLBREAK_BUILDS are gone.
 */

/**
 * A pivot's data reference (B13) -- MIGRATED to the DS Bridge `pivot.state.read/v1` (Wave B READ, patch family /4, design
 * 3.3), addressed by sheet + the pivot's public name (a non-unique name is refused, never one picked). Used, as before, as
 * the input of the content proof (src/adapters/xlsx/pivot-source.cjs) and as every pivot's raw identity (pivotRefs, M7 P3
 * refresh_pivot). MEASURED limit, unchanged by the Bridge: after a source-sheet rename it still names the OLD sheet --
 * hence never trusted by name alone. The direct `pivot.asc_getDataRef()` read and EO_INTERNAL_PIVOT_RUNTIMES are gone.
 */
const MAX_PIVOT_SOURCE_CELLS = 10000;
/**
 * A drawing's anchor mode and computed bottom-right anchor (B14, `drawingBase.Type` / `.boundsFromTo.to`) -- MIGRATED to
 * the DS Bridge `drawing.anchor.read/v1` (internal tracker point 8; LIVE parity #ref). The bridge refuses an
 * un-hydrated geometry itself (`geometry-not-hydrated`), so it is read after the B15 hydration below, as before; the
 * direct drawingBase read and EO_INTERNAL_DRAWING_ANCHOR_BUILDS are gone.
 */
/**
 * A NON-ACTIVE sheet's view (B15, KIND=RUNTIME_HYDRATION) -- MIGRATED to the DS Bridge `sheet.model.ensureReady/v1`
 * (patch family /5; owner GO the owner, owner channel 2026-10-09 msg 33257, internal tracker). A drawing's geometry exists only
 * through its sheet's view, which the editor creates lazily (at load only for the active sheet); the bridge runs the
 * editor's own view creator behind its hydrate() entry -- not a READ, no document data returned, no document/structure/
 * history/dirty/file change (LIVE: docs/evidence/b15-hydration-live-20261009/). Called ONLY before the drawing read of
 * the anchor-asking (structure) request. Any bridge refusal (runtime, capability, the bridge's own reason) makes that
 * sheet's drawings 'UNOBSERVABLE:hydration-<why>', and a drawing whose geometry is still unset reads
 * 'UNOBSERVABLE:BRIDGE:geometry-not-hydrated' -- both refuse the structural step before the write. There is NO
 * fallback to a direct call: the direct WorkbookView.getWorksheet call and EO_INTERNAL_SHEET_HYDRATION_RUNTIMES are gone.
 */
/**
 * A shape's text WITHOUT creating one (B16) -- MIGRATED to the DS Bridge `drawing.text.read/v1` (internal tracker point 8;
 * LIVE parity #ref: a text shape and a text-less shape, history 0). The public GetContent() still writes a text body
 * on a text-less shape, so it stays off the read path; the direct getDocContent() read and
 * EO_INTERNAL_SHAPE_TEXT_BUILDS are gone.
 */
/** The in-command pivot reader; the data reference through pivot.state.read/v1 (B13). */
const PIVOT_READ = `
  // M7 P9: each pivot in its own try -- the public source (it throws after any reopen, W3g / P-A) and the RAW proof the
  // adapter decides on (pivot-source.cjs): the Bridge data reference (pivot.state.read/v1), the field names, the cached body and table
  // (public), the named range's values and every OTHER sheet's values at the same address. Read only; a part that
  // throws is its own 'UNOBSERVABLE:<why>'.
  const grid = r => { const v = r.GetValue(); const s = y => (y === undefined || y === null ? '' : String(y)); return Array.isArray(v) ? v.map(x => (Array.isArray(x) ? x.map(s) : [s(x)])) : [[s(v)]]; };
  const fieldNames = f => { try { return (f() || []).map(x => x.GetName()); } catch (err) { return 'UNOBSERVABLE:' + String(err).slice(0, 80); } };
  // GH #302 G3 (LIVE probe run 2): a data field's SOURCE field through GetSourceName ('Amount' for 'Sum of Amount2' holding a
  // Count) -- session + reopen -- so nothing parses the display name.
  // X-O10 r6 (zero LIVE, 2026-10-15): GetFunction (on an ACTUAL data field) and
  // GetLayoutForm/GetLayoutCompactRow/GetLayoutSubtotals (on a pivot field) are now read-measured
  // -- same style as fieldNames (the whole list is its own UNOBSERVABLE if the getter itself
  // throws or is not a list), PLUS a per-field try so one field's getter throwing never hides the
  // others (each sub-value its own 'UNOBSERVABLE:<why>').
  const dataFieldFunctions = f => { try { return (f() || []).map(x => { let name; try { name = x.GetName(); } catch (e) { name = 'UNOBSERVABLE:' + String(e).slice(0, 80); } let fn; try { fn = x.GetFunction(); } catch (e) { fn = 'UNOBSERVABLE:' + String(e).slice(0, 80); } let field; try { field = x.GetSourceName(); } catch (e) { field = 'UNOBSERVABLE:' + String(e).slice(0, 80); } return { name, function: fn, field }; }); } catch (err) { return 'UNOBSERVABLE:' + String(err).slice(0, 80); } };
  const pivotFieldLayout = f => { try { return (f() || []).map(x => { let name; try { name = x.GetName(); } catch (e) { name = 'UNOBSERVABLE:' + String(e).slice(0, 80); } let form; try { form = x.GetLayoutForm(); } catch (e) { form = 'UNOBSERVABLE:' + String(e).slice(0, 80); } let compactRow; try { compactRow = x.GetLayoutCompactRow(); } catch (e) { compactRow = 'UNOBSERVABLE:' + String(e).slice(0, 80); } let subtotals; try { subtotals = x.GetLayoutSubtotals(); } catch (e) { subtotals = 'UNOBSERVABLE:' + String(e).slice(0, 80); } return { name, form, compactRow, subtotals }; }); } catch (err) { return 'UNOBSERVABLE:' + String(err).slice(0, 80); } };
  // GH #302 G2 (LIVE probe docs/evidence/gh302-g2-pivot-read-probe-20261026.json): the row / column / page field names of a
  // richer pivot. With two or more data fields the row or column list carries the "Values" pseudo-field: its GetName (and
  // every other getter but GetIndex) throws, its GetIndex() is -1 (the stored <field x="-2"/>), session and reopen alike.
  // So a name that throws on an item whose GetIndex() is -1 is the Values position, never a guessed name; any other
  // throwing item is that item's own 'UNOBSERVABLE:<why>'. A list getter that throws is the list's own reason string.
  const axisList = f => { let xs; try { xs = f() || []; } catch (err) { return 'UNOBSERVABLE:' + String(err).slice(0, 80); } if (!Array.isArray(xs)) return 'UNOBSERVABLE:not-a-list'; const names = []; let values = null; xs.forEach((x, i) => { let nm; try { nm = x.GetName(); } catch (e) { nm = null; } if (typeof nm === 'string') { names.push(nm); return; } let ix; try { ix = x.GetIndex(); } catch (e) { ix = null; } if (ix === -1 && values === null) values = i + 1; else names.push('UNOBSERVABLE:field-name-unread'); }); return { names, values }; };
  // GetPageFields() itself THROWS on a pivot with no page field (G3 LIVE probe: "reading 'map'" of null) -- the commonest
  // case. Then the page fields are the pivot fields whose GetOrientation() is 'Filters' (the G2 probe: a page field reads
  // 'Filters'), in GetPosition() order; if that read fails too, the list is named.
  const pagesByOrientation = p => { try { const out = []; for (const f of (p.GetPivotFields() || [])) { if (f.GetOrientation() === 'Filters') out.push([f.GetPosition(), f.GetName()]); } return { names: out.sort((a, b) => a[0] - b[0]).map(x => x[1]), values: null }; } catch (err) { return null; } };
  const pivotAxes = p => {
    const rows = axisList(() => p.GetRowFields()); const cols = axisList(() => p.GetColumnFields());
    let pages = axisList(() => p.GetPageFields());
    if (typeof pages === 'string') pages = pagesByOrientation(p) || pages;
    const at = (l, axis) => (typeof l === 'object' && l.values ? { axis, position: l.values } : null);
    return { rows: typeof rows === 'string' ? rows : rows.names, cols: typeof cols === 'string' ? cols : cols.names, pages: typeof pages === 'string' ? pages : pages.names, values: at(rows, 'rows') || at(cols, 'columns') };
  };
  const cellsIn = range => { const part = a => { let c = 0; let i = 0; while (i < a.length && a[i] >= 'A' && a[i] <= 'Z') { c = c * 26 + a.charCodeAt(i) - 64; i += 1; } return [c, Number(a.slice(i))]; }; const [x, y] = range.split(':').map(part); return (y[0] - x[0] + 1) * (y[1] - x[1] + 1); };
  const pivotState = (p, n) => bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.PIVOT_STATE_READ)}, { sheet: n, pivot: p.GetName() });
  const pivotProof = (p, st) => {
    if (typeof st === 'string') return st;
    const dataRef = st.dataRef;
    const i = dataRef.lastIndexOf('!');
    if (i < 1) return 'UNOBSERVABLE:pivot-dataref-shape';
    let sheet = dataRef.slice(0, i);
    if (sheet.length > 1 && sheet[0] === "'" && sheet[sheet.length - 1] === "'") sheet = sheet.slice(1, -1).split("''").join("'");
    const range = dataRef.slice(i + 1).split('$').join('');
    if (range.indexOf(':') < 0 || !(cellsIn(range) > 0) || cellsIn(range) > ${MAX_PIVOT_SOURCE_CELLS}) return 'UNOBSERVABLE:pivot-source-size';
    const src = Api.GetSheet(sheet);
    const values = src && src.GetName() === sheet ? axis(() => grid(src.GetRange(range))) : 'UNOBSERVABLE:pivot-source-sheet-missing';
    const others = sheets.filter(x => x.row.name !== sheet).map(x => ({ sheet: x.row.name, values: axis(() => grid(x.ws.GetRange(range))) }));
    return { dataRef, fields: fieldNames(() => p.GetPivotFields()), rows: fieldNames(() => p.GetRowFields()), cols: fieldNames(() => p.GetColumnFields()), data: fieldNames(() => p.GetDataFields()),
      body: { range: addrOf(p.GetDataBodyRange()), values: grid(p.GetDataBodyRange()) }, table: { range: addrOf(p.GetTableRange1()), values: grid(p.GetTableRange1()) },
      source: { sheet, range, values }, others };
  };
  const pivotRead = (p, n) => {
    const out = { name: p.GetName() };
    try { out.source = addr(p.GetSource()); } catch (err) { out.source = 'UNOBSERVABLE:' + String(err).slice(0, 80); }
    // M7 P1: the table's extent and the layout (public getters, both survive a reopen -- P-A2 / P-A3)
    try { out.table = addrOf(p.GetTableRange1()); } catch (err) { out.table = 'UNOBSERVABLE:' + String(err).slice(0, 80); }
    out.layout = { rows: fieldNames(() => p.GetRowFields()), cols: fieldNames(() => p.GetColumnFields()), data: fieldNames(() => p.GetDataFields()) };
    // X-O10 r6 (zero LIVE): the data field's aggregation FUNCTION + each pivot field's
    // layout-form state, both read-measured (session + reopen); read_pivots stays PARTIAL for an
    // unrelated reason (M7-H, the reopened source identity), not for these two fields.
    out.functions = dataFieldFunctions(() => p.GetDataFields());
    out.fieldLayout = pivotFieldLayout(() => p.GetPivotFields());
    // GH #302 G2: the axes (row / column / page names + where the Values pseudo-field sits) and the style name
    out.axes = pivotAxes(p);
    try { out.style = p.GetStyleName(); } catch (err) { out.style = 'UNOBSERVABLE:' + String(err).slice(0, 80); }
    // M7 P3 refresh_pivot (owner GO 2026-10-04, owner channel: "pontos dataRef-egyezes"): the approved P9 READ and
    // the field list for EVERY pivot, in session and reopened -- the explicit rebind is verified on the exact reference
    const st = axis(() => pivotState(p, n));
    out.dataRef = typeof st === 'string' ? st : st.dataRef;
    out.fields = fieldNames(() => p.GetPivotFields());
    if (out.source.indexOf('UNOBSERVABLE') === 0) out.proof = axis(() => pivotProof(p, st));
    return out;
  };
`;

/**
 * B11 (MIGRATED to the DS Bridge, owner GO owner channel; LIVE-accepted on the served /3 4b6fa653, evidence
 * docs/evidence/b11-v3-*-live-acceptance-20261007.json): the dependency inventory delete/copy/insert need -- every
 * defined name and every data validation of every sheet -- read through `workbook.definedNames.read/v1` and
 * `range.validations.read/v1`, in ONE callCommand. The returned shape is the one the adapter has always consumed:
 *   names:       [{ name, ref, scopeSheet }]  -- scopeSheet null = workbook scope
 *   validations: { <sheet>: [{ type, ranges: [{ r1, c1, r2, c2 }], formula1, formula2 }] }
 * The names keep the production filter the old internal reader applied (hidden, '_xlnm'-NAMED and ref-less names
 * dropped -- by NAME, not by the bridge's isXLNM-based builtIn flag; table names and the prefix-stripped Print_Area /
 * Print_Titles KEPT) and its name order; the validations are each sheet's full set (a whole-sheet query).
 * Any bridge gap (capability missing, wrong build/version, malformed value) makes THAT axis
 * 'UNOBSERVABLE:<reason>' -- never an empty list, and there is NO direct-internal fallback.
 */
const WHOLE_SHEET = 'A1:XFD1048576';
function dependentsBridgeBody() {
  return `${BRIDGE_CLIENT_SRC}
const rd = (cap, args) => __eoBridgeRead(cap.id, cap.version, ${JSON.stringify(ACCEPTED_RUNTIMES.xlsx)}, args, cap.valueType);
const out = { names: null, validations: null };
try {
  const dn = rd(${JSON.stringify(BRIDGE_CAPABILITIES.DEFINED_NAMES_READ)}, {});
  if (!dn.ok) out.names = 'UNOBSERVABLE:' + dn.reason;
  else if (!Array.isArray(dn.value)) out.names = 'UNOBSERVABLE:bridge-names-not-a-list';
  else {
    // the OLD reader's exact filter: ref present, not hidden, name WITHOUT '_xlnm'. NOT the bridge's builtIn flag -- the engine
    // stores Print_Area / Print_Titles with isXLNM=true and the '_xlnm.' prefix STRIPPED, and the names family reads them here
    // (LIVE-measured, m7-f regression 2026-10-07 18:51: filtering on builtIn made set_print_area unverifiable).
    const keep = dn.value.filter(e => e.refersTo !== null && e.hidden !== true && e.name.indexOf('_xlnm') === -1);
    const scopeKey = e => (e.scope.kind === 'workbook' ? -1 : e.scope.index);
    keep.sort((a, b) => (a.name > b.name ? 1 : a.name < b.name ? -1 : scopeKey(a) - scopeKey(b)));
    out.names = keep.map(e => ({ name: e.name, ref: e.refersTo, scopeSheet: e.scope.kind === 'workbook' ? null : e.scope.sheet }));
  }
} catch (err) { out.names = 'UNOBSERVABLE:bridge-names-threw'; }
try {
  const v = {};
  for (const ws of (Api.GetSheets() || [])) {
    const sheet = ws.GetName();
    const r = rd(${JSON.stringify(BRIDGE_CAPABILITIES.VALIDATIONS_READ)}, { sheet, range: ${JSON.stringify(WHOLE_SHEET)} });
    if (!r.ok) { out.validations = 'UNOBSERVABLE:' + r.reason; break; }
    if (!Array.isArray(r.value)) { out.validations = 'UNOBSERVABLE:bridge-validations-not-a-list'; break; }
    v[sheet] = r.value.map(x => ({ type: x.type, ranges: x.ranges, formula1: x.formula1, formula2: x.formula2 }));
  }
  if (out.validations === null) out.validations = v;
} catch (err) { out.validations = 'UNOBSERVABLE:bridge-validations-threw'; }
return out;`;
}

/**
 * Column / row layout (wave 4, W4b-measured PUBLIC getters): per column GetEntireColumn().GetColumnWidth()
 * and GetHidden(), per row GetEntireRow().GetRowHeight() and GetHidden(). Read only when a layout step
 * asks (`layout: [{ sheet, maxCol, maxRow }]`), for EVERY sheet over the larger of its used range and the
 * requested extent (L3 compares every other column / row). A getter that throws, or more than
 * MAX_LAYOUT_ITEMS columns + rows, makes THAT sheet's axis a reason string -- never a partial list.
 */
const MAX_LAYOUT_ITEMS = 2000;

/**
 * X-O05/4 (design docs/X-O05-3-DATALABEL-READ-CONSUMER-DESIGN-20261013.md, point 2): the per-chart
 * budget for `chart.dataLabels.effective.read/v1` bridge calls (one call per (series, point) --
 * see `chartReads` below). Above this, same discipline as MAX_LAYOUT_ITEMS: the WHOLE chart's
 * `effectiveLabels` becomes a named reason string, never a partial per-point list.
 * LIVE-MEASURED (test/live/xlsx-xo05-4-chart-effective-labels-read.live.cjs, two independent runs):
 * 150 bridge calls (a 50-series x 3-point chart) cost a CHART_READ_DELTA_MS of -16 and -25 against
 * a ~80ms baseline read -- i.e. the 150-call cost is BELOW the run-to-run noise floor of the
 * measurement itself, not separable from it. This value (300, double the measured-safe 150) is a
 * margin on a floor we could only bound, not pin down -- the EXACT per-point cost stays NEM-MERT;
 * a tighter number would need a repeated, order-alternated run with a median, which has not been
 * done.
 */
const MAX_CHART_LABEL_POINTS = 300;

/**
 * M7 tier F2 (owner GO 2026-10-03, owner channel, point 3) + X-P07 (OD-A, owner channel): the print READ
 * widened for the Q9 / Q11 / Q12 / Q14 / Q15 WRITE shapes -- the odd FOOTER, VERTICAL centering, the first / even texts
 * and their two enabling flags, and (always, with them) the column breaks. Projected from the same
 * sheet.printSettings.read/v1 value (B09); generated only when asked, so no other query shape changes.
 */
const PRINT_WIDE_FIELDS = ['oddFooter', 'vCenter', 'firstHeader', 'firstFooter', 'evenHeader', 'evenFooter', 'diffFirst', 'diffOddEven'].map(k => `, ${k}: v.${k}`).join('');

/** The D2 column-break field of the print READ: generated only when asked. B12 cutover: the column breaks through sheet.columnBreaks.read/v1 (same { id, man, min, max } list, null -> []). */
const colBreaksField = asked => (asked ? `colBreaks: axis(() => bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.COL_BREAKS_READ)}, { sheet: n })), ` : '');

/** A driver-side boolean, as the literal text to interpolate into a generated script (one shared
 * ternary, not one per flag -- qlty, readQuery complexity). */
const boolText = v => (v ? 'true' : 'false');

/** readQuery's optional, build-gated editor-side fragments (kept out of readQuery itself): '' unless switched on. */
const pick = (on, js, otherwise) => (on === true ? js : otherwise);
const when = (on, js) => pick(on, js, '');
const HYDRATE_JS = `${BRIDGE_HYDRATE_CLIENT_SRC}  // the sheet's view through the DS Bridge (B15 /5), BEFORE any getter (a size read before it computes 0 and may stick)
  const hydrate = ws => {
    const r = __eoBridgeHydrate(${JSON.stringify(BRIDGE_HYDRATIONS.SHEET_MODEL_READY.id)}, ${BRIDGE_HYDRATIONS.SHEET_MODEL_READY.version}, ${JSON.stringify(ACCEPTED_RUNTIMES.xlsx)}, { sheet: ws.GetName() });
    return r.ok ? null : 'UNOBSERVABLE:hydration-' + r.reason;
  };
`;
const HYDRATE_CALL_JS = "\n      const unhydrated = hydrate(ws);\n      if (unhydrated) return unhydrated;";

function readQuery(windows = [], { panes = false, formatSheets = [], styleSheets = [], objectSheets = [], conditional = false, layout = null, validationRanges = [], cfRanges = [], runCells = [], arrayCells = [], sheetVisibility = false, filterSheets = [], chartSheets = [], pageSheets = [], pagePerSheet = false, printSheets = [], colBreaks = false, drawingSheets = [], usedRanges = true, printWide = false, drawingAnchor = false, hydrateSheets, protectionSheets = [], protectionRanges = [], commentSheets = []} = {}) {
  return `
try {
  const LIMIT = ${MAX_READ_CELLS};
  const windows = ${JSON.stringify((windows || []).map(w => ({ sheet: String(w.sheet), range: String(w.range), format: w.format === true || w.style === true, style: w.style === true, ...(w.nf === true ? { nf: true } : {}) })))};
  const formatSheets = ${JSON.stringify((formatSheets || []).map(String))};
  const styleSheets = ${JSON.stringify((styleSheets || []).map(String))};
  const objectSheetsAsked = ${JSON.stringify((objectSheets || []).map(String))};
  const unsupported = [];
  const cells = [];
  const seen = {};
  let truncated = false;
  const wf = Api.WorksheetFunction;
  if (!wf || typeof wf.COUNT !== 'function') unsupported.push({ field: 'WorksheetFunction.COUNT', reason: 'unavailable' });
  const addrOf = r => String(r.GetAddress(false, false, 'xlA1', false));
  // The DS Bridge client (fail-closed, no fallback) and the cell-level font flags through it (B01 / B02).
${BRIDGE_CLIENT_SRC}  const bridgeFontFlag = (out, cap, key) => {
    const r = __eoBridgeRead(cap.id, cap.version, ${JSON.stringify(ACCEPTED_RUNTIMES.xlsx)}, { sheet: out.sheet, cell: out.a }, 'boolean');
    out[key] = r.ok ? r.value : 'UNOBSERVABLE:' + r.reason;
  };
  // Any other migrated axis: the capability's value, or 'UNOBSERVABLE:<reason>' -- never a guessed default, no fallback.
  const bridgeAxis = (cap, args) => {
    const r = __eoBridgeRead(cap.id, cap.version, ${JSON.stringify(ACCEPTED_RUNTIMES.xlsx)}, args, cap.valueType);
    return r.ok ? r.value : 'UNOBSERVABLE:' + r.reason;
  };
  // Number format and fill: only for window cells that ask for them (a clear must prove it kept them).
  const addFormat = (out, cell) => {
    try { out.nf = cell.GetNumberFormat(); } catch (err) { unsupported.push({ field: out.a + '.GetNumberFormat', reason: String(err) }); }
    try { const c = cell.GetFillColor(); out.fill = c && typeof c.GetRGB === 'function' ? c.GetRGB() : (c === null || c === undefined ? null : 'S:' + String(c)); }
    catch (err) { unsupported.push({ field: out.a + '.GetFillColor', reason: String(err) }); }
    // Cell-level bold + italic through the DS Bridge (B01 / B02): EACH gap is its OWN reason string, never false, and
    // never an 'unsupported' entry (that would void the whole observation instead of just that one axis).
    bridgeFontFlag(out, ${JSON.stringify(BRIDGE_CAPABILITIES.BOLD_READ)}, 'bold');
    bridgeFontFlag(out, ${JSON.stringify(BRIDGE_CAPABILITIES.ITALIC_READ)}, 'italic');
  };
  // Style (wave 4, format_range): wrap (PUBLIC GetWrapText, W4a); the font through the characters (PUBLIC
  // GetCharacters(1, len).GetFont(), W4a2 -- a cell whose first / last character disagrees with the whole is
  // mixed-run: UNOBSERVABLE); X-F02 (m7-s31b-probe1/2 MEASURED): an EMPTY cell's font is read
  // the SAME way, at (1, 1) -- no separate reader, no internal dependency; a throw there stays UNOBSERVABLE,
  // never a guessed default; alignment and borders through the build-gated EO_INTERNAL_DEPENDENCY READ (W4c).
  // A gap is a reason string on that axis, never a default.
  const addStyle = (out, cell) => {
    try { const w = cell.GetWrapText(); out.wrap = typeof w === 'boolean' ? w : 'UNOBSERVABLE:wrap-not-boolean'; } catch (err) { out.wrap = 'UNOBSERVABLE:threw'; }
    // X-F07a (LIVE m7-s24-probe1/2): rotation, PUBLIC GetOrientation() -- unlike
    // align.h/v this is NOT read through the DS Bridge (no internal dependency); a NEVER-
    // formatted empty cell's reader throws (MEASURED), tagged UNOBSERVABLE like every other gap,
    // never defaulted.
    try { const o = cell.GetOrientation(); out.rot = typeof o === 'number' ? o : 'UNOBSERVABLE:rot-not-number'; } catch (err) { out.rot = 'UNOBSERVABLE:threw'; }
    out.font = (() => {
      let text;
      try { text = String(cell.GetValue()); } catch (err) { return 'UNOBSERVABLE:value-threw'; }
      const fontAt = (start, len) => { const f = cell.GetCharacters(start, len).GetFont(); const c = f.GetColor(); return { bold: f.GetBold(), italic: f.GetItalic(), name: f.GetName(), size: f.GetSize(), color: c && typeof c.GetRGB === 'function' ? c.GetRGB() : null, underline: f.GetUnderline(), strike: f.GetStrikethrough(), sub: f.GetSubscript(), sup: f.GetSuperscript() }; };
      // X-F02 (m7-s31b-probe1/2 + p24/m7-s19-probe1 MEASURED): an empty cell's font
      // is observable through the SAME public reader, at (1, 1) -- no new internal call. A throw
      // (or anything other than the expected shape, which fontAt's own getters would throw on
      // too) stays UNOBSERVABLE, never a guessed default.
      if (text === '') {
        try { return fontAt(1, 1); } catch (err) { return 'UNOBSERVABLE:threw'; }
      }
      try {
        const all = fontAt(1, text.length);
        const same = x => JSON.stringify(x) === JSON.stringify(all);
        if (text.length > 1 && !(same(fontAt(1, 1)) && same(fontAt(text.length, 1)))) return 'UNOBSERVABLE:mixed-runs';
        return all;
      } catch (err) { return 'UNOBSERVABLE:threw'; }
    })();
    // B04 cutover: horizontal / vertical alignment and the border -- with the row-2 / column-B shared edge the engine's own
    // getBorderFull skips (M7 S4) -- through range.appearance.read/v1. ONE capability, so a gap in either part makes BOTH
    // axes unobservable (stricter than the two former independent reads, never a default).
    const ap = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.APPEARANCE_READ)}, { sheet: out.sheet, cell: out.a });
    if (typeof ap === 'string') { out.align = ap; out.border = ap; } else { out.align = ap.align; out.border = ap.border; }
    // X-F07b/c (B05 READ cutover): readingOrder / indent / shrink through range.alignmentExt.read/v1. A never-set
    // readingOrder is the semantic default 0 (S31c MEASURED, mapped by the capability). ONE capability, so a gap makes ALL
    // THREE axes 'UNOBSERVABLE:<reason>' (stricter than the former per-key reads, never a default).
    const ae = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.ALIGN_EXT_READ)}, { sheet: out.sheet, cell: out.a });
    if (typeof ae === 'string') { out.readingOrder = ae; out.indent = ae; out.shrink = ae; } else { out.readingOrder = ae.readingOrder; out.indent = ae.indent; out.shrink = ae.shrink; }
  };
  // X-T01: the number format ALONE (PUBLIC GetNumberFormat) -- a write_range target's Text '@' type, without the fill and
  // the DS Bridge font-flag reads a full format read adds (the plain write path gains no extra read)
  const addNf = (out, cell) => { try { out.nf = cell.GetNumberFormat(); } catch (err) { unsupported.push({ field: out.a + '.GetNumberFormat', reason: String(err) }); } };
  const readRange = (ws, range, fmt, sty, nfOnly) => {
    range.ForEach(cell => {
      if (truncated) return;
      const a = addrOf(cell);
      const k = ws.GetName() + '\\u0000' + a;
      if (seen[k] !== undefined) { if (fmt) addFormat(cells[seen[k]], cell); else if (nfOnly && !('nf' in cells[seen[k]])) addNf(cells[seen[k]], cell); if (sty) addStyle(cells[seen[k]], cell); return; }
      if (cells.length >= LIMIT) { truncated = true; return; }
      seen[k] = cells.length;
      const out = { sheet: ws.GetName(), a };
      try { out.v = cell.GetValue(); } catch (err) { unsupported.push({ field: a + '.GetValue', reason: String(err) }); }
      try { out.f = cell.GetFormula(); } catch (err) { unsupported.push({ field: a + '.GetFormula', reason: String(err) }); }
      try { out.count = wf && typeof wf.COUNT === 'function' ? wf.COUNT(cell) : null; } catch (err) { unsupported.push({ field: a + '.COUNT', reason: String(err) }); }
      try { const m = cell.MergeArea; out.merge = m && typeof m.GetAddress === 'function' ? addrOf(m) : null; if (out.merge === null) unsupported.push({ field: a + '.MergeArea', reason: 'unavailable' }); }
      catch (err) { unsupported.push({ field: a + '.MergeArea', reason: String(err) }); }
      if (fmt) addFormat(out, cell);
      else if (nfOnly) addNf(out, cell);
      if (sty) addStyle(out, cell);
      cells.push(out);
    });
  };
  const sheets = Api.GetSheets().map(ws => {
    const ur = ws.GetUsedRange();
    const used = ur ? addrOf(ur) : null;
    return { ws, row: { name: ws.GetName(), index: ws.GetIndex(), used } };
  });
${usedRanges ? `  for (const s of sheets) if (s.row.used) readRange(s.ws, s.ws.GetRange(s.row.used), formatSheets.includes(s.row.name) || styleSheets.includes(s.row.name), styleSheets.includes(s.row.name));
` : '  // tiered observation (T1): the windows only -- the content outside them is the digest\'s (T2)\n'}  for (const w of windows) {
    const ws = Api.GetSheet(w.sheet);
    if (!ws || ws.GetName() !== w.sheet) continue;
    readRange(ws, ws.GetRange(w.range), w.format, w.style, w.nf);
  }
  // Objects of the named sheets (W3c-measured public readers): chart series references, pivot sources,
  // sheet-level defined names. A reader that is missing or throws makes THAT axis a reason string --
  // never an empty list (owner rule: a failing reader is not an empty dependency set).
  const objects = {};
  const addr = r => String(r.GetAddress(true, true, 'xlA1', true)).replace(/^\\[[^\\]]*\\]/, '');
  const axis = f => { try { return f(); } catch (err) { return 'UNOBSERVABLE:' + String(err).slice(0, 80); } };
  // '*' = every sheet (a delete must see the dependents of the whole workbook)
  const objectSheets = objectSheetsAsked.includes('*') ? sheets.map(s => s.row.name) : objectSheetsAsked;
${PIVOT_READ}  for (const n of objectSheets) {
    const ws = Api.GetSheet(n);
    if (!ws || ws.GetName() !== n) continue;
    objects[n] = {
      charts: axis(() => (ws.GetAllCharts() || []).map(c => ({ type: c.GetChartType(), series: (c.GetAllSeries() || []).map(x => ({ name: x.GetName(), values: x.GetValues(), cat: x.GetCatFormula(), x: x.GetXValues() })) }))),
      pivots: axis(() => (ws.GetAllPivotTables() || []).map(p => pivotRead(p, n))),
      names: axis(() => (ws.GetDefNames() || []).filter(x => x && typeof x.GetName === 'function' && x.GetName()).map(x => ({ name: x.GetName(), refersTo: x.GetRefersTo() }))),
    };
  }
  // Conditional formats of EVERY sheet (PUBLIC, W3e: whole-sheet GetFormatConditions -> GetItem(i)).
  // Only the measured types' getters are called (xlExpression W3e; xlCellValue + GetOperator W6 C1/C2, M7 S1);
  // GH #141 (docs/GH141-CF-DEPENDENCY-OBSERVABILITY-20261018.md): aboveAverage/top10/unique have no
  // formula concept at all (their Formula1/Formula2 getters throw), so only GetAppliesTo is called for
  // them -- already LIVE-measured safe for every CF type via cf.cjs's deleteBody re-identification.
  // colorScale/databar/iconSet (their CFVO stops can be formula-typed, never read here) stay bare.
  const conditional = ${conditional ? '{}' : 'undefined'};
  if (conditional) for (const s of sheets) {
    conditional[s.row.name] = axis(() => {
      const fc = s.ws.GetRange('A1:XFD1048576').GetFormatConditions();
      const n = fc.GetCount();
      if (typeof n !== 'number') throw new Error('GetCount-not-a-number');
      const out = [];
      for (let i = 1; i <= n; i += 1) {
        const x = fc.GetItem(i);
        if (!x) throw new Error('GetItem-null');
        const type = x.GetType();
        if (type === 'xlExpression') out.push({ type, appliesTo: addr(x.GetAppliesTo()), formula1: x.GetFormula1(), formula2: x.GetFormula2() });
        else if (type === 'xlCellValue') out.push({ type, operator: x.GetOperator(), appliesTo: addr(x.GetAppliesTo()), formula1: x.GetFormula1(), formula2: x.GetFormula2() });
        else if (type === 'xlAboveAverageCondition' || type === 'xlTop10' || type === 'xlUniqueValues') out.push({ type, appliesTo: addr(x.GetAppliesTo()), formula1: null, formula2: null });
        else out.push({ type });
      }
      return out;
    });
  }
  const layoutAsked = ${JSON.stringify(layout === null ? null : (layout || []).map(x => ({ sheet: String(x.sheet), maxCol: Number(x.maxCol) || 0, maxRow: Number(x.maxRow) || 0, canon: (x.canon || []).map(Number).filter(Number.isFinite) })))};
  let layout;
  if (layoutAsked) {
    layout = {};
    const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    const colNum = t => t.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    for (const s of sheets) {
      const name = s.row.name;
      const m = s.row.used ? /^([A-Z]+)(\\d+)(?::([A-Z]+)(\\d+))?$/.exec(s.row.used.replace(/\\$/g, '')) : null;
      let maxCol = m ? colNum(m[3] || m[1]) : 0;
      let maxRow = m ? Number(m[4] || m[2]) : 0;
      const canonAsked = [];
      for (const x of layoutAsked) if (x.sheet === name) { maxCol = Math.max(maxCol, x.maxCol); maxRow = Math.max(maxRow, x.maxRow); canonAsked.push(...x.canon); }
      if (maxCol + maxRow > ${MAX_LAYOUT_ITEMS}) { layout[name] = 'UNOBSERVABLE:layout-truncated'; continue; }
      layout[name] = axis(() => {
        const cols = {};
        const rows = {};
        for (let c = 1; c <= maxCol; c += 1) { const e = s.ws.GetRange(colName(c) + '1').GetEntireColumn(); cols[colName(c)] = { width: e.GetColumnWidth(), hidden: e.GetHidden() }; }
        for (let r = 1; r <= maxRow; r += 1) { const e = s.ws.GetRange('A' + r).GetEntireRow(); rows[String(r)] = { height: e.GetRowHeight(), hidden: e.GetHidden() }; }
        if (!canonAsked.length) return { cols, rows };
        // B17 cutover: the requested widths' canonical persisted values through sheet.columnWidth.canonicalize/v1 (pure
        // arithmetic, no write); keyed by the requested value as before -- any gap or a length mismatch is unobservable
        const canonR = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.COL_WIDTH_CANON_READ)}, { sheet: name, values: canonAsked });
        return { cols, rows, canon: typeof canonR === 'string' ? canonR : (Array.isArray(canonR) && canonR.length === canonAsked.length ? Object.fromEntries(canonAsked.map((x, i) => [String(x), canonR[i]])) : 'UNOBSERVABLE:canon-unexpected-shape') };
      });
    }
  }
  let panes;
  let active;
${panes ? `  // Freeze panes (wave 4): the ACTIVE sheet name (public) and every sheet's OWN pane through the DS Bridge (B06),
  // never by activating a sheet. frozen = { state: 'frozen', row0, col0 } (the first UNfrozen cell, 0-based);
  // none = { state: 'none' }; any gap 'UNOBSERVABLE:<why>' -- never 'none'.
  {
    active = axis(() => Api.GetActiveSheet().GetName());
    panes = {};
    for (const s of sheets) {
      // B06 cutover: the sheet's OWN pane through sheet.freezeState.read/v1 (the same tri-state, never by activating)
      panes[s.row.name] = axis(() => bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.FREEZE_READ)}, { sheet: s.row.name }));
    }
  }
` : ''}${sheetVisibility ? `  // X-L10 (freeze on a HIDDEN sheet): every sheet's visibility through the PUBLIC ApiWorksheet.GetVisible(), read only
  // when a freeze step asks (the saved file's activeTab follows the freeze TARGET, so a hidden target opens as the
  // visible active sheet -- the freeze family refuses a hidden target). true | false | 'UNOBSERVABLE:<why>'.
  const visibility = {};
  for (const s of sheets) {
    visibility[s.row.name] = axis(() => { const ws = Api.GetSheet(s.row.name); const v = ws && typeof ws.GetVisible === 'function' ? ws.GetVisible() : undefined; return typeof v === 'boolean' ? v : 'UNOBSERVABLE:visibility-not-boolean'; });
  }
` : ''}  // Validation (wave 6, set / clear_validation): the PUBLIC range-level reader (the source's measured getters, B-D16).
  // { type, alertStyle, operator, formula1, formula2, activeCell } | { absent: true } (no validation object, or type null /
  // 'xlValidateInputOnly' -- the source's absence) | 'UNOBSERVABLE:<why>'
  // L-22e (r1-r3 LIVE measurement, xlsx-l22e-custom-relative-ref-probe.live.cjs): type
  // 'custom's Formula1 reports the formula AS IF WRITTEN AT the RULE'S OWN SHEET's current active cell, not the
  // range's own anchor -- 'activeCell' is read in the SAME call as GetFormula1 (never inferred after the fact,
  // never from a DIFFERENT sheet's or the globally-active sheet's own notion of "active"), so the family's own
  // identity check (reAnchoredFormula()) can re-anchor the read text deterministically. 'UNOBSERVABLE:<why>' if
  // GetActiveCell() itself throws or returns nothing -- unmeasured whether this sheet being the globally-active
  // one matters; the family fails closed on it either way, never guessing a cell.
  const validationReads = {};
  const validationAsked = ${JSON.stringify((validationRanges || []).map(x => ({ sheet: String(x.sheet), range: String(x.range) })))};
  for (const w of validationAsked) {
    validationReads[w.sheet + '!' + w.range] = axis(() => {
      const ws = Api.GetSheet(w.sheet);
      if (!ws || ws.GetName() !== w.sheet) return 'UNOBSERVABLE:sheet-missing';
      const v = ws.GetRange(w.range).GetValidation();
      if (!v) return { absent: true };
      const type = v.GetType();
      if (type === null || type === undefined || type === 'xlValidateInputOnly') return { absent: true };
      let activeCell;
      try { const ac = ws.GetActiveCell(); activeCell = ac ? ac.GetAddress(false, false, 'xlA1', false) : 'UNOBSERVABLE:active-cell-absent'; }
      catch (err) { activeCell = 'UNOBSERVABLE:active-cell-threw'; }
      return { type, alertStyle: v.GetAlertStyle(), operator: v.GetOperator(), formula1: v.GetFormula1(), formula2: v.GetFormula2(), activeCell };
    });
  }
  // Conditional formats on a range (wave 6, add / delete_conditional_format; M7 X-O01 widened to
  // type-branch, so a type OTHER than xlCellValue/xlExpression no longer calls getters that throw
  // on it -- M7-S29 LIVE probe, GetOperator/GetFormula1/GetFormula2 throw on every other measured
  // type). xlCellValue/xlExpression keep the ORIGINAL shape (B-D34) unchanged. The other MEASURED
  // types (xlAboveAverageCondition/xlTop10/xlUniqueValues/xlDatabar/xlIconSets/xlColorScale) read
  // type + appliesTo + their own MEASURED scalar getters only (CF_SCALAR_FIELDS). An EMPTY type
  // ('' -- xlUniqueValues after SetDupeUnique('xlDuplicate'), M7-S29b LIVE probe) makes the WHOLE
  // RANGE unobservable (fail-closed): a rule this reader cannot identify must never be silently
  // dropped from the list a write-time identity check relies on.
  const CF_SCALAR_FIELDS = {
    xlAboveAverageCondition: [['aboveBelow', 'GetAboveBelow'], ['numStdDev', 'GetNumStdDev']],
    xlTop10: [['rank', 'GetRank'], ['percent', 'GetPercent'], ['topBottom', 'GetTopBottom']],
    xlUniqueValues: [['dupeUnique', 'GetDupeUnique']],
    // X-O01 (LIVE phase 1/2 probes): the appearance getters, all PUBLIC and read back after a reopen
    xlDatabar: [['barColor', 'GetBarColor'], ['barFillType', 'GetBarFillType'], ['direction', 'GetDirection'], ['showValue', 'GetShowValue'], ['axisPosition', 'GetAxisPosition']],
    xlIconSets: [['iconSet', 'GetIconSet'], ['showIconOnly', 'GetShowIconOnly'], ['reverseOrder', 'GetReverseOrder']],
    xlColorScale: [],
  };
  // X-O01: a colour scale's stop colours (GetColorScaleCriteria()[i].GetColor(), PUBLIC), as one comparable JSON string
  const cfStopColors = x => JSON.stringify((x.GetColorScaleCriteria() || []).map(c => { const v = c.GetColor(); return v && typeof v.GetRGB === 'function' ? v.GetRGB() : null; }));
  const cfReads = {};
  const cfAsked = ${JSON.stringify((cfRanges || []).map(x => ({ sheet: String(x.sheet), range: String(x.range) })))};
  for (const w of cfAsked) {
    cfReads[w.sheet + '!' + w.range] = axis(() => {
      const ws = Api.GetSheet(w.sheet);
      if (!ws || ws.GetName() !== w.sheet) return 'UNOBSERVABLE:sheet-missing';
      const fc = ws.GetRange(w.range).GetFormatConditions();
      const k = fc.GetCount();
      if (typeof k !== 'number') throw new Error('GetCount-not-a-number');
      const list = [];
      for (let i = 1; i <= k; i += 1) {
        const x = fc.GetItem(i);
        const type = x.GetType();
        if (type === null || type === undefined || type === '') return 'UNOBSERVABLE:cf-type-empty';
        if (type === 'xlCellValue' || type === 'xlExpression') {
          const fill = x.GetFillColor();
          list.push({ type, operator: x.GetOperator(), formula1: x.GetFormula1(), formula2: x.GetFormula2(), fill: fill && typeof fill.GetRGB === 'function' ? fill.GetRGB() : null, appliesTo: addrOf(x.GetAppliesTo()) });
          continue;
        }
        const entry = { type, appliesTo: addrOf(x.GetAppliesTo()) };
        for (const [key, getter] of (CF_SCALAR_FIELDS[type] || [])) {
          const v = x[getter]();
          entry[key] = v && typeof v.GetRGB === 'function' ? v.GetRGB() : v;
        }
        if (type === 'xlColorScale') entry.stopColors = cfStopColors(x);
        list.push(entry);
      }
      return list;
    });
  }
${(runCells || []).length ? `  // X-F04 (M7 backlog, format_run): a cell's run-map, PUBLIC and
  // character-by-character -- GetCharacters(i, 1).GetFont(), the SAME fontAt(1,1)/(len,1)
  // boundary-check this file's own addStyle() font reader already uses, just called on EVERY
  // character and merging consecutive identical font-tuples. 'UNOBSERVABLE:text-too-long' is
  // its own named ceiling (256 accepted as a TEMPORARY value, pending a LIVE
  // timing measurement) -- NOT a guess at a longer text's runs, and NOT a silent truncation.
  const RUN_MAP_CHAR_CEILING = 256;
  const runReads = {};
  const runAsked = ${JSON.stringify((runCells || []).map(x => ({ sheet: String(x.sheet), cell: String(x.cell) })))};
  for (const w of runAsked) {
    runReads[w.sheet + '!' + w.cell] = axis(() => {
      const ws = Api.GetSheet(w.sheet);
      if (!ws || ws.GetName() !== w.sheet) return 'UNOBSERVABLE:sheet-missing';
      const cell = ws.GetRange(w.cell);
      let text;
      try { text = String(cell.GetValue()); } catch (err) { return 'UNOBSERVABLE:value-threw'; }
      if (text === '') return 'UNOBSERVABLE:empty-cell';
      if (text.length > RUN_MAP_CHAR_CEILING) return 'UNOBSERVABLE:text-too-long';
      const charFont = i => {
        const f = cell.GetCharacters(i, 1).GetFont();
        const c = f.GetColor();
        return { bold: f.GetBold(), italic: f.GetItalic(), name: f.GetName(), size: f.GetSize(), color: c && typeof c.GetRGB === 'function' ? c.GetRGB() : null };
      };
      const runs = [];
      let cur = null;
      for (let i = 1; i <= text.length; i += 1) {
        const ft = charFont(i);
        const k = JSON.stringify(ft);
        if (cur && cur.key === k) { cur.length += 1; } else { cur = { start: i, length: 1, font: ft, key: k }; runs.push(cur); }
      }
      return { text, runs: runs.map(({ start, length, font }) => ({ start, length, font })) };
    });
  }
` : ''}${(arrayCells || []).length ? `  // GH #301 D4 (write_array_formula): each asked cell's PUBLIC GetFormulaArray(), raw -- the LIVE design probe
  // (docs/evidence/gh301-d4-array-formula-probe-20261025.json) measured it is NOT an extent reader (a value cell
  // returns its value, a plain formula null, an empty cell '', every cell of an array the array's formula), so the
  // family decides membership from it together with the cell's own GetFormula; nothing is interpreted here.
  const arrayReads = {};
  const arrayAsked = ${JSON.stringify((arrayCells || []).map(x => ({ sheet: String(x.sheet), cell: String(x.cell) })))};
  for (const w of arrayAsked) {
    arrayReads[w.sheet + '!' + w.cell] = axis(() => {
      const ws = Api.GetSheet(w.sheet);
      if (!ws || ws.GetName() !== w.sheet) return 'UNOBSERVABLE:sheet-missing';
      const fa = ws.GetRange(w.cell).GetFormulaArray();
      return fa === undefined ? null : fa;
    });
  }
` : ''}  // AutoFilter (wave 6, filter_range / clear_filter): the PUBLIC readers (GetAutoFilter -> GetRange, GetFilters -> each
  // GetOn / GetOperator / GetCriteria1 / GetCriteria2) and the field identity through the DS Bridge (B07).
  // { range: null } = no AutoFilter; any gap -> 'UNOBSERVABLE:<why>' for that sheet.
  const filterReads = {};
  const filterAsked = ${JSON.stringify((filterSheets || []).map(String))};
  // '*' = every sheet (M7 S2a: a structural step's L3 claims every sheet's AutoFilter)
  for (const n of (filterAsked.includes('*') ? sheets.map(s => s.row.name) : filterAsked)) {
    filterReads[n] = axis(() => {
      const ws = Api.GetSheet(n);
      if (!ws || ws.GetName() !== n) return 'UNOBSERVABLE:sheet-missing';
      const af = ws.GetAutoFilter();
      if (!af) return { range: null, filters: [] };
      const fs = af.GetFilters() || [];
      // a never-filtered sheet: an AutoFilter object with NO range and no filters (W5-W9 LIVE run 1, 2026-09-29) --
      // that is "no AutoFilter"; filters without a range are not a shape we measured
      const afRange = af.GetRange();
      if (!afRange) return fs.length === 0 ? { range: null, filters: [] } : 'UNOBSERVABLE:autofilter-filters-without-range';
      const filters = [];
      // B07 cutover: each filter's FIELD (1-based column within the range) through sheet.autoFilter.fieldIdentity.read/v1,
      // in GetFilters() order; the bridge must see the SAME range and the same number of filters, else the sheet is
      // unobservable -- never a guessed field
      const ids = fs.length ? bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.FILTER_FIELDS_READ)}, { sheet: n }) : { range: addrOf(afRange), fields: [] };
      if (typeof ids === 'string') return ids;
      if (!ids || ids.range !== addrOf(afRange) || !Array.isArray(ids.fields) || ids.fields.length !== fs.length) return 'UNOBSERVABLE:field-identity-mismatch';
      fs.forEach((f, i) => filters.push({ field: ids.fields[i], on: f.GetOn(), operator: f.GetOperator(), criteria1: f.GetCriteria1(), criteria2: f.GetCriteria2() }));
      return { range: addrOf(afRange), filters };
    });
  }
${(commentSheets || []).length ? `  // GH #300 K1 (read_comments): every comment of the asked sheets through the PUBLIC getters (GetText / GetAuthorName /
  // GetUserId / GetTimeUTC / IsSolved, the replies through GetRepliesCount / GetReply). ApiComment has NO address getter
  // (GetQuoteText is the cell VALUE), so the anchor is found by a per-cell GetRange(a).GetComment() scan -- and the used
  // range does NOT reach a comment-only cell (K1 LIVE probe 3: used A1:C5, notes on E6 and J30). The scan is a bounded
  // rectangle from A1 (at least A1:Z200, the used range + 10 columns / 100 rows), row-major, stopping when it has found as
  // many anchors as ws.GetComments() counts; fewer -> the whole sheet is NAMED (never a partial list). GetId() is not
  // read: it is not stable across a reopen (K1 LIVE probe 1), the cell is the key. GetTime() is not read (0 on every
  // measured comment); GetTimeUTC() is 0 on a comment that carries no time (an API-made one) -- passed as read.
  const commentReads = {};
  const commentAsked = ${JSON.stringify((commentSheets || []).map(String))};
  for (const n of commentAsked) {
    commentReads[n] = axis(() => {
      const ws = Api.GetSheet(n);
      if (!ws || ws.GetName() !== n) return 'UNOBSERVABLE:sheet-missing';
      const expected = (ws.GetComments() || []).length;
      if (expected === 0) return { comments: [] };
      if (expected > ${MAX_COMMENTS}) return 'IMPLEMENTATION_LIMIT:comments-over-${MAX_COMMENTS}';
      const ur = ws.GetUsedRange();
      const end = ur ? addrOf(ur).split(':').pop() : 'A1';
      const m = /^([A-Z]+)(\\d+)$/.exec(end);
      const usedCol = m ? m[1].split('').reduce((x, ch) => x * 26 + ch.charCodeAt(0) - 64, 0) : 1;
      const usedRow = m ? Number(m[2]) : 1;
      const maxCol = Math.min(16384, Math.max(26, usedCol + 10));
      const maxRow = Math.min(1048576, Math.max(200, usedRow + 100), Math.floor(${MAX_COMMENT_SCAN_CELLS} / maxCol));
      if (maxRow < usedRow) return 'IMPLEMENTATION_LIMIT:comment-scan-over-${MAX_COMMENT_SCAN_CELLS}-cells';
      const name = c => { let x = ''; for (let k = c; k > 0; k = Math.floor((k - 1) / 26)) x = String.fromCharCode(65 + ((k - 1) % 26)) + x; return x; };
      const unob = e => 'UNOBSERVABLE:' + String((e && e.message) || e).slice(0, 80);
      const pick = (o, k, f) => { try { o[k] = f(); } catch (e) { o[k] = unob(e); } };
      const rows = [];
      for (let r = 1; r <= maxRow && rows.length < expected; r += 1) {
        for (let c = 1; c <= maxCol && rows.length < expected; c += 1) {
          const cell = name(c) + r;
          const cm = ws.GetRange(cell).GetComment();
          if (!cm) continue;
          const row = { cell };
          pick(row, 'text', () => cm.GetText()); pick(row, 'author', () => cm.GetAuthorName()); pick(row, 'userId', () => cm.GetUserId());
          pick(row, 'timeUtc', () => cm.GetTimeUTC()); pick(row, 'solved', () => cm.IsSolved());
          row.replies = [];
          try {
            const k = cm.GetRepliesCount();
            for (let i = 0; i < k; i += 1) { const rp = cm.GetReply(i); const rr = {}; pick(rr, 'text', () => rp.GetText()); pick(rr, 'author', () => rp.GetAuthorName()); pick(rr, 'userId', () => rp.GetUserId()); pick(rr, 'timeUtc', () => rp.GetTimeUTC()); row.replies.push(rr); }
          } catch (e) { row.replies = unob(e); }
          rows.push(row);
        }
      }
      if (rows.length < expected) return 'UNOBSERVABLE:comment-anchor-outside-scan:' + rows.length + '/' + expected + ':A1:' + name(maxCol) + maxRow;
      return { comments: rows };
    });
  }
` : ''}  // Charts (wave 7, set_chart / rename_chart): the PUBLIC chart getters (GetName, GetChartType, GetTitle, every
  // series' name / values / categories) and the pivot NAMES (GetSource is not read: it throws after a reopen, audit
  // PV4). '*' = every sheet; a reader that is missing or throws -> 'UNOBSERVABLE:<why>' for that sheet.
  // X-O04: the chart read also gets legend: c.GetLegendPos(), right after title.
  // X-O05 (LIVE MEASURED): GetHorAxisTitle/GetVerAxisTitle/GetDataLabels are called directly, no
  // per-field try/catch -- the control-chart finding showed they do NOT throw on an untouched chart (null / null /
  // an all-false object), so a throw here is a genuine "whole sheet unobservable" signal, same as every other getter.
  // X-O05/4 (design docs/X-O05-3-DATALABEL-READ-CONSUMER-DESIGN-20261013.md): effectiveLabels, the PER-POINT
  // effective data-label state (point -> series -> chart precedence) through chart.dataLabels.effective.read/v1,
  // via the SAME bridgeAxis helper every other migrated axis in this file uses -- never a new call mechanism.
  // Over MAX_CHART_LABEL_POINTS: the WHOLE chart's effectiveLabels becomes one named reason string, never a
  // partial per-point list (the MAX_LAYOUT_ITEMS precedent). A per-point bridge failure (BRIDGE_UNAVAILABLE,
  // RUNTIME_/PATCH_VERSION_MISMATCH, CAPABILITY_MISSING, ...) surfaces verbatim as THAT point's own
  // 'UNOBSERVABLE:<reason>' value (bridgeAxis's own convention) -- it does not affect any other point, series,
  // chart or the chart's own name/type/title/legend/series/axis-title/dataLabels fields.
  // X-O05/4 round 2 (LIVE-measured): GetValues() on a REAL chart is a FORMULA STRING
  // ('$B$2:$C$2', or a bare single-cell ref with no colon), NEVER an array -- round 1's Array.isArray guard
  // silently read 0 points for every series. The point count is derived from the formula's own range span
  // instead: a single-row range (same row, columns differ) -> the column span; a single-column range (same
  // column, rows differ) -> the row span; a bare single-cell ref (no colon) or a degenerate same-cell range
  // -> 1. Anything else (a 2D range, or a shape colNumOf/the regex does not recognize) makes the point count
  // UNKNOWN for that series -- and an unknown count on ANY series makes the WHOLE chart's effectiveLabels
  // 'UNOBSERVABLE:point-count-unknown', never a guessed {} (the same "never a partial/guessed list" discipline
  // as the over-cap branch, just for the opposite failure direction).
  const colNumOf = t => t.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const pointCountOf = values => {
    if (Array.isArray(values)) return values.length;
    if (typeof values !== 'string') return null;
    const body = values.replace(/^.*!/, ''); // drop an optional ('Sheet'!/Sheet!) prefix -- the range itself decides
    if (/^\\$?[A-Za-z]+\\$?\\d+$/.test(body)) return 1; // a bare single-cell reference, no colon
    const m = /^\\$?([A-Za-z]+)\\$?(\\d+):\\$?([A-Za-z]+)\\$?(\\d+)$/.exec(body);
    if (!m) return null;
    const [, c1, r1, c2, r2] = m;
    if (c1 === c2 && r1 !== r2) return Math.abs(Number(r2) - Number(r1)) + 1;
    if (r1 === r2 && c1 !== c2) return Math.abs(colNumOf(c2) - colNumOf(c1)) + 1;
    if (c1 === c2 && r1 === r2) return 1;
    return null; // a 2D range -- not a single-axis point count
  };
  const chartReads = {};
  const chartAsked = ${JSON.stringify((chartSheets || []).map(String))};
  for (const n of (chartAsked.includes('*') ? sheets.map(s => s.row.name) : chartAsked)) {
    chartReads[n] = axis(() => {
      const ws = Api.GetSheet(n);
      if (!ws || ws.GetName() !== n) return 'UNOBSERVABLE:sheet-missing';
      const charts = (ws.GetAllCharts() || []).map(c => {
        const name = c.GetName();
        // LIVE (X-O03/7 r2): the PUBLIC GetValues() THROWS on a scatter series -- read bare, it took the whole sheet's
        // chart read down. Each per-series public getter is read on its own (axis): a throw is THAT field's reason.
        const series = (c.GetAllSeries() || []).map(x => ({ name: axis(() => x.GetName()), values: axis(() => x.GetValues()), cat: axis(() => x.GetCatFormula()) }));
        const pointCounts = series.map(s => pointCountOf(s.values));
        let effectiveLabels;
        if (pointCounts.some(p => p === null)) {
          effectiveLabels = 'UNOBSERVABLE:point-count-unknown';
        } else if (pointCounts.reduce((sum, p) => sum + p, 0) > ${MAX_CHART_LABEL_POINTS}) {
          effectiveLabels = 'UNOBSERVABLE:chart-label-points-truncated';
        } else {
          effectiveLabels = {};
          series.forEach((s, si) => {
            for (let pi = 0; pi < pointCounts[si]; pi += 1) {
              effectiveLabels[si + ':' + pi] = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.CHART_LABELS_EFFECTIVE_READ)}, { sheet: n, chart: name, series: si, point: pi });
            }
          });
        }
        const type = c.GetChartType();
        // X-O03/6b: scatter's own PUBLIC GetValues() (above, 'values') is a formula string or
        // THROWS (X-O03/1, internal tracker) -- the actual Y-value array is read through the Bridge instead,
        // on scatter-type charts only, one call per series (never per-point: a flat array, no
        // point-level precedence to resolve, unlike effectiveLabels above).
        if (type === 'scatter') {
          series.forEach((s, si) => { s.scatterY = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.SCATTER_Y_READ)}, { sheet: n, chart: name, series: si }); });
        }
        // X-O04/10 (design docs/X-O04-9-CHART-STYLE-READ-BRIDGE-DESIGN-20261014.md, reviewed and
        // accepted): a series' own marker fill+outline colour, read through the internal
        // model (no PUBLIC getter exists). Scoped to the two chart types the LIVE chain was
        // actually confirmed on (X-O04/6+/8 fixtures: scatter + lineNormal) -- never a guess on an
        // untested chart type.
        if (type === 'scatter' || type === 'lineNormal') {
          series.forEach((s, si) => { s.markerStyle = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.MARKER_STYLE_READ)}, { sheet: n, chart: name, series: si }); });
          // GH #139 B /16 (design docs/GH139-B16-CHART-STYLE-READ-DESIGN.md, LIVE MEASURED):
          // symbol/size off the SAME series.marker object markerStyle reads
          // fill/outline off -- same scope (scatter + lineNormal), never a guess on an untested type.
          series.forEach((s, si) => { s.markerSymbol = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.MARKER_SYMBOL_READ)}, { sheet: n, chart: name, series: si }); });
        } else {
          // the owner's "unsupported chart/type fail-closed, named" requirement (pre-deploy mechanism probe): an out-of-scope type must NOT simply omit the field -- that is
          // indistinguishable from "never ran". Named explicitly, per series, per chart type.
          series.forEach(s => { s.markerSymbol = 'UNOBSERVABLE:chart-type-not-in-scope:' + type; });
        }
        // GH #139 B /16: dash/smooth, read through the series' own spPr.ln.prstDash / smooth
        // fields (LIVE MEASURED). Scoped to the WRITE side's own measured chart
        // types for dash (bar + lineNormal, charts.cjs LINE_STYLE_CHART_TYPES) -- smooth's
        // narrower scope (lineNormal only) is named by the handler itself (a bar series' own
        // smooth field is UNDEFINED, the measurement's own negative control), never guessed here.
        if (type === 'bar' || type === 'lineNormal') {
          series.forEach((s, si) => { s.lineStyle = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.LINE_STYLE_READ)}, { sheet: n, chart: name, series: si }); });
        } else {
          // same named-fail-closed requirement as markerSymbol above, for the dash/smooth scope.
          series.forEach(s => { s.lineStyle = 'UNOBSERVABLE:chart-type-not-in-scope:' + type; });
        }
        // X-O05/2 (X-O05/1 LIVE r4: the order getter proven session+reopen-readable): a series'
        // own c:order index, read through its internal model's own order getter. Unscoped by
        // chart type, unlike markerStyle above -- X-O05/1 proved the setter+getter pair on a BAR
        // chart specifically, but the internal series' own set/get order pair (via
        // private_GetSeries) is the SAME convention X-O04/6+/10 already used across chart types;
        // a chart type whose internal series lacks the method degrades to its own named reason
        // (getorder-missing), never a guess.
        series.forEach((s, si) => { s.order = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.SERIES_ORDER_READ)}, { sheet: n, chart: name, series: si }); });
        // Gridline colour is unscoped by chart type (unlike markerStyle above): a chart type with
        // no axes (e.g. pie) degrades to its own named reason (axis-getters-missing) through the
        // SAME bridge call, never a guessed null.
        const gridlinesColor = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.GRIDLINES_COLOR_READ)}, { sheet: n, chart: name });
        // X-O05/2 (the other X-O05/1 axis; LIVE re-measured): gapWidth/overlap, read through the
        // chart-type's own internal data field (no method, no public getter -- the LIVE
        // re-measurement confirmed it survives a reopen). Called unconditionally, same as
        // gridlinesColor above -- the Bridge handler itself checks the chart is bar-typed (the
        // LIVE-confirmed scope) and names its own reason on every other type, never a guess.
        const gapOverlap = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.GAP_OVERLAP_READ)}, { sheet: n, chart: name });
        return { name, type, title: c.GetTitle(), legend: c.GetLegendPos(), series,
          horAxisTitle: c.GetHorAxisTitle(), verAxisTitle: c.GetVerAxisTitle(), dataLabels: c.GetDataLabels(), effectiveLabels, gridlinesColor, gapOverlap };
      });
      const pivots = (ws.GetAllPivotTables() || []).map(p => p.GetName());
      return { charts, pivots };
    });
  }
  // Drawings (M7 S2b, owner GO 2026-10-01, docs/evidence/M7-S2B-DRAWING-PROBE.md): the PUBLIC drawing getters of every
  // drawing (image / shape / chart) -- each required field read on its own, a missing / throwing getter
  // 'UNOBSERVABLE:<why>' for THAT field, never null / [] / a default. Generated only when asked. '*' = every sheet.
  const drawingReads = {};
${(drawingSheets || []).length ? `  const drawingAsked = ${JSON.stringify((drawingSheets || []).map(String))};
  const dget = (o, k) => {
    if (!o || typeof o[k] !== 'function') return 'UNOBSERVABLE:no-' + k;
    try { const v = o[k](); return v === undefined ? 'UNOBSERVABLE:' + k + '-undefined' : v; } catch (e) { return 'UNOBSERVABLE:' + k + '-threw'; }
  };
  // B14 / B16 cutover: a drawing's anchor and a shape's text through drawing.anchor.read/v1 / drawing.text.read/v1. The
  // bridge finds a drawing BY NAME (the first match), so a name that is not unique on its sheet (or not a string) is
  // UNOBSERVABLE here and the bridge is not asked -- never a read of a possibly different drawing.
  const dbyName = (d, n, names, cap) => {
    const nm = dget(d, 'GetName');
    if (typeof nm !== 'string' || names.filter(x => x === nm).length !== 1) return 'UNOBSERVABLE:drawing-name-ambiguous';
    return bridgeAxis(cap, { sheet: n, drawing: nm });
  };
  const dtext = (d, n, names) => dbyName(d, n, names, ${JSON.stringify(BRIDGE_CAPABILITIES.DRAWING_TEXT_READ)});
${drawingAnchor === true ? `  const danchor = (d, n, names) => dbyName(d, n, names, ${JSON.stringify(BRIDGE_CAPABILITIES.DRAWING_ANCHOR_READ)});
` : ''}${when(hydrateSheets, HYDRATE_JS)}  for (const n of (drawingAsked.includes('*') ? sheets.map(s => s.row.name) : drawingAsked)) {
    drawingReads[n] = axis(() => {
      const ws = Api.GetSheet(n);
      if (!ws || ws.GetName() !== n) return 'UNOBSERVABLE:sheet-missing';${when(hydrateSheets, HYDRATE_CALL_JS)}
      if (typeof ws.GetAllDrawings !== 'function') return 'UNOBSERVABLE:no-GetAllDrawings';
      const list = ws.GetAllDrawings();
      if (!Array.isArray(list)) return 'UNOBSERVABLE:drawings-not-a-list';
      const names = list.map(d => dget(d, 'GetName'));
      return list.map(d => {
        const pos = dget(d, 'GetPosition');
        const ps = dget(d, 'GetParentSheet');
        const kind = dget(d, 'GetClassType');
        // X-O08-7 (16fd90b9, LIVE r3): a shape's GetLine().GetWidth() is PUBLIC-readable, same session
        // AND after reopen -- GetLine().GetFill() colour is NOT (opaque UniFill, never go past GetType).
        const line = kind === 'shape' ? dget(d, 'GetLine') : 'UNOBSERVABLE:no-public-content-getter';
        return { kind, name: dget(d, 'GetName'), parent: typeof ps === 'string' ? ps : dget(ps, 'GetName'),
          pos: typeof pos === 'string' ? pos : (pos && typeof pos === 'object' ? { fromCol: pos.fromCol, colOffset: pos.colOffset, fromRow: pos.fromRow, rowOffset: pos.rowOffset } : 'UNOBSERVABLE:position-unexpected-shape'),
          width: dget(d, 'GetWidth'), height: dget(d, 'GetHeight'), rotation: dget(d, 'GetRotation'), flipH: dget(d, 'GetFlipH'), flipV: dget(d, 'GetFlipV'),
          strokeWidth: typeof line === 'object' ? dget(line, 'GetWidth') : line,
          // X-O08 (zero LIVE): fill/line RGB -- the public GetFill()/GetLine().GetFill() give
          // the TYPE only (strokeWidth above), never the byte value; this capability reads the
          // internal model for the value, scoped to shapes only (the measured subject).
          fillLineColor: kind === 'shape' ? bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.FILL_LINE_COLOR_READ)}, { sheet: n, drawing: dget(d, 'GetName') }) : 'UNOBSERVABLE:no-public-content-getter',
          text: kind === 'shape' ? dtext(d, n, names) : 'UNOBSERVABLE:no-public-content-getter'${drawingAnchor === true ? ', anchor: danchor(d, n, names)' : ''} };
      });
    });
  }
` : ''}  // Page layout (wave 8, set_page_layout): the PUBLIC getters, served only on a build where the per-sheet read is
  // proven; otherwise 'UNOBSERVABLE:per-sheet-read-not-verified'. '*' = every sheet.
  const pageReads = {};
  const pageAsked = ${JSON.stringify((pageSheets || []).map(String))};
  for (const n of (pageAsked.includes('*') ? sheets.map(s => s.row.name) : pageAsked)) {
    pageReads[n] = axis(() => {
      if (!${pagePerSheet ? 'true' : 'false'}) return 'UNOBSERVABLE:per-sheet-read-not-verified';
      const ws = Api.GetSheet(n);
      if (!ws || ws.GetName() !== n) return 'UNOBSERVABLE:sheet-missing';
      return { orientation: ws.GetPageOrientation(), printGridlines: ws.GetPrintGridlines(), printHeadings: ws.GetPrintHeadings(),
        topMargin: ws.GetTopMargin(), bottomMargin: ws.GetBottomMargin(), leftMargin: ws.GetLeftMargin(), rightMargin: ws.GetRightMargin() };
    });
  }
  // Print settings (wave 8, the print family): sheet.printSettings.read/v1 (B09); generated only when asked, and only the
  // ASKED fields are projected from the capability's value. '*' = every sheet.
  const printReads = {};
${(printSheets || []).length ? `  const printAsked = ${JSON.stringify((printSheets || []).map(String))};
  for (const n of (printAsked.includes('*') ? sheets.map(s => s.row.name) : printAsked)) {
    printReads[n] = axis(() => {
      const ws = Api.GetSheet(n);
      if (!ws || ws.GetName() !== n) return 'UNOBSERVABLE:sheet-missing';
      const v = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.PRINT_SETTINGS_READ)}, { sheet: n });
      if (typeof v === 'string') return v;
      return { fitW: v.fitW, fitH: v.fitH, fitFlag: v.fitFlag, oddHeader: v.oddHeader,
        rowBreaks: v.rowBreaks, ${colBreaksField(colBreaks || printWide)}useFirstPage: v.useFirstPage, firstPage: v.firstPage, hCenter: v.hCenter, scale: v.scale, width: v.width, height: v.height${printWide ? PRINT_WIDE_FIELDS : ''} };
    });
  }
` : ''}  // Sheet protection (GH #326 / #327, U4a client-wiring only): sheet.protection.read/v1 (B/17); generated only
  // when asked, '*' = every sheet. { protected, forbids } -- a range/cells is out of THIS unit's scope (the #326
  // guard / #327 redaction units add it later). On a /16 runtime bridgeAxis degrades the whole axis to its own
  // named 'UNOBSERVABLE:CAPABILITY_MISSING' string, never a guessed value.
  const protectionReads = {};
${(protectionSheets || []).length ? `  const protectionAsked = ${JSON.stringify((protectionSheets || []).map(String))};
  for (const n of (protectionAsked.includes('*') ? sheets.map(s => s.row.name) : protectionAsked)) {
    protectionReads[n] = axis(() => {
      const ws = Api.GetSheet(n);
      if (!ws || ws.GetName() !== n) return 'UNOBSERVABLE:sheet-missing';
      const v = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.PROTECTION_READ)}, { sheet: n });
      return typeof v === 'string' ? v : { protected: v.protected, forbids: v.forbids };
    });
  }
` : ''}  // Sheet protection, RANGE-scoped (GH #327, U5): the SAME capability with a range -- { protected, forbids,
  // cells: { range, nonEmpty, unlocked, hidden } }, keyed '<sheet>!<range>' (the SAME per-request-range convention
  // cfReads/validationReads already use, so a tiered request stays eligible via TIERED_EXTRAS). On a /16 runtime
  // (or any other read gap) the WHOLE entry degrades to bridgeAxis's own 'UNOBSERVABLE:<reason>' string --
  // read_range's own redaction (read-tools.cjs) fails the window's formula field closed on it, never guessing.
  const protectionRangeReads = {};
${(protectionRanges || []).length ? `  const protectionRangeAsked = ${JSON.stringify((protectionRanges || []).map(x => ({ sheet: String(x.sheet), range: String(x.range) })))};
  for (const w of protectionRangeAsked) {
    protectionRangeReads[w.sheet + '!' + w.range] = axis(() => {
      const ws = Api.GetSheet(w.sheet);
      if (!ws || ws.GetName() !== w.sheet) return 'UNOBSERVABLE:sheet-missing';
      const v = bridgeAxis(${JSON.stringify(BRIDGE_CAPABILITIES.PROTECTION_READ)}, { sheet: w.sheet, range: w.range });
      return typeof v === 'string' ? v : { protected: v.protected, forbids: v.forbids, cells: v.cells };
    });
  }
` : ''}  return { ok: true, sheets: sheets.map(s => s.row), cells, truncated, unsupported, objects, conditional, layout, panes, active, validationReads, cfReads, ${(runCells || []).length ? 'runReads, ' : ''}${(arrayCells || []).length ? 'arrayReads, ' : ''}${sheetVisibility ? 'visibility, ' : ''}filterReads, ${(commentSheets || []).length ? 'commentReads, ' : ''}chartReads, pageReads, printReads, drawingReads, protectionReads, protectionRangeReads };
} catch (err) { return { ok: false, error: String(err) }; }
`;
}

/** browser.close() with its failure REPORTED, never swallowed: { ok, detail }. No browser -> nothing to close. */
async function closeBrowser(browser) {
  if (!browser) return { ok: true, detail: 'no-browser' };
  return browser.close().then(() => ({ ok: true, detail: 'torn-down' }),
    err => ({ ok: false, detail: `browser-close-failed: ${String(err && err.message || err).slice(0, 160)}` }));
}

/** Read-only: the sheet list with the used ranges, nothing else (the tier decision's input). */
function usedProbeQuery() {
  return `
try {
  return { ok: true, sheets: Api.GetSheets().map(ws => { const ur = ws.GetUsedRange(); return { name: ws.GetName(), index: ws.GetIndex(), used: ur ? String(ur.GetAddress(false, false, 'xlA1', false)) : null }; }) };
} catch (err) { return { ok: false, error: String(err) }; }
`;
}

/** A plain A1 cell or range ('B2', 'A1:J500', '$' allowed) -> { r1, c1, r2, c2, count, text }; null otherwise. Local on
 * purpose: the transport does not import the adapter's A1 module. */
function parseRange(text) {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(String(text || '').toUpperCase());
  if (!m) return null;
  const col = x => [...x].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const a = { c: col(m[1]), r: Number(m[2]) };
  const b = m[3] ? { c: col(m[3]), r: Number(m[4]) } : a;
  const g = { c1: Math.min(a.c, b.c), r1: Math.min(a.r, b.r), c2: Math.max(a.c, b.c), r2: Math.max(a.r, b.r) };
  if (g.r1 < 1) return null;
  return { ...g, count: (g.c2 - g.c1 + 1) * (g.r2 - g.r1 + 1) };
}
const rangeText = g => { const n = c => { let s = ''; for (let x = c; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s; return s; };
  return g.c1 === g.c2 && g.r1 === g.r2 ? `${n(g.c1)}${g.r1}` : `${n(g.c1)}${g.r1}:${n(g.c2)}${g.r2}`; };

const cellCount = range => { const g = range ? parseRange(range) : null; return g ? g.count : 0; };
const bboxOf = (a, b) => ({ c1: Math.min(a.c1, b.c1), r1: Math.min(a.r1, b.r1), c2: Math.max(a.c2, b.c2), r2: Math.max(a.r2, b.r2) });

/**
 * The tier of an observation request (pure). `bound`: per sheet, the bounding box of the used range together with the
 * request's windows on that sheet, summed -- a write lands only inside its windows and a clear only shrinks, so every
 * observation of one task stays under this bound and sees the same tier.
 *   full     bound <= MAX_READ_CELLS                        -> today's read, unchanged
 *   tiered   bound <= MAX_TIERED_CELLS and an ELIGIBLE request: the adapter marked it `tiered` (a task of write_range /
 *            clear_range steps only), windows with format flags only, nothing else asked, at most MAX_READ_CELLS window
 *            cells
 *   bounded  X-N01b (owner decision A): the request itself is NOT eligible, but it names a `boundedVariant` that is --
 *            the variant is read instead (tier 'tiered', bounded: true). Only above MAX_READ_CELLS: below it the full
 *            request always runs, so a bounded proof never replaces a full one that fits the bound.
 *   refused  anything else above MAX_READ_CELLS -> today's read (it truncates) + an explicit IMPLEMENTATION_LIMIT note
 */
function tierOf(sheets, request = {}) {
  const windows = request.windows || [];
  const bound = boundOf(sheets, windows);
  if (bound <= MAX_READ_CELLS) return { tier: 'full', bound };
  if (bound > MAX_TIERED_CELLS) {
    // GH #110: the ceiling guards the T2 digest scan; a read-only window request needs no digest -- its window read alone
    if (request.readOnly === true && !refusalOf(request, windows)) return { tier: 'window', bound, windowCells: windows.reduce((n, w) => n + cellCount(w.range), 0) };
    return { tier: 'refused', bound, reason: `workbook-over-${MAX_TIERED_CELLS}-cells` };
  }
  const reason = refusalOf(request, windows);
  if (!reason) return { tier: 'tiered', bound, windowCells: windows.reduce((n, w) => n + cellCount(w.range), 0) };
  const v = request.boundedVariant;
  const vReason = v ? refusalOf(v, v.windows || []) : null;
  if (v && !vReason) return { tier: 'tiered', bound, bounded: true, fullRefusedBecause: reason, windowCells: (v.windows || []).reduce((n, w) => n + cellCount(w.range), 0) };
  return { tier: 'refused', bound, reason: v ? `${reason}; bounded variant: ${vReason}` : reason };
}

/** Per sheet, the bounding box of the used range together with the windows on that sheet; summed over the sheets. */
function boundOf(sheets, windows) {
  let bound = 0;
  for (const s of sheets || []) {
    const boxes = [s.used, ...windows.filter(w => w.sheet === s.name).map(w => w.range)].map(parseRange).filter(Boolean);
    if (!boxes.length) continue;
    const box = boxes.reduce(bboxOf);
    bound += (box.c2 - box.c1 + 1) * (box.r2 - box.r1 + 1);
  }
  return bound;
}

/** GH #107 (a): the adapter's note on a multi-step task with neither path ('' when there is none). */
function mixedTaskNote(m) {
  if (!m) return '';
  const part = (name, xs) => (xs && xs.length ? `; ${name}: ${xs.join(', ')}` : '');
  return `; mixed task --${part('not tiered', m.notTiered).slice(1)}${part('no bounded variant', m.noBoundedVariant)}${m.splittable ? '; each step alone has a tiered or a bounded path: send one task per family' : ''}`;
}

/** Why a request above MAX_READ_CELLS may NOT be tiered (null = eligible). */
function refusalOf(request, windows) {
  if (request.tiered !== true) return `family-not-migrated-to-the-tiered-observation${mixedTaskNote(request.mixedTask)}`;
  // noFormulaIndex (GH #96 f/u) asks for LESS, not for a read; mixedTask (GH #107) is a refusal note, not a read
  const META = ['windows', 'tiered', 'windowStyle', 'boundedVariant', 'noFormulaIndex', 'mixedTask', 'readOnly'];
  const isAsked = ([k, v]) => !META.includes(k) && v !== null && v !== undefined && v !== false && !(Array.isArray(v) && v.length === 0);
  const asked = Object.entries(request).filter(isAsked).map(([k]) => k).filter(k => !TIERED_EXTRAS.includes(k)).sort();
  if (asked.length) return `tiered-observation-reads-windows-only: ${asked.join(',')}`;
  // a window's own style read is bounded by the window cell count; only a bounded variant asks for it (X-N01b)
  if (request.windowStyle !== true && windows.some(w => w.style === true)) return 'tiered-observation-reads-windows-only: window-style';
  return windows.reduce((n, w) => n + cellCount(w.range), 0) > MAX_READ_CELLS ? `task-windows-over-${MAX_READ_CELLS}-cells` : null;
}

/** SHA-256 over a JS string (UTF-8), pure JS, as editor source -- no editor or library dependency. */
const SHA256_SRC = `
function __utf8(s) { const out = []; for (let i = 0; i < s.length; i += 1) { let c = s.charCodeAt(i);
  if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length) { const d = s.charCodeAt(i + 1); if (d >= 0xdc00 && d <= 0xdfff) { c = 0x10000 + ((c - 0xd800) << 10) + (d - 0xdc00); i += 1; } }
  if (c < 0x80) out.push(c); else if (c < 0x800) out.push(0xc0 | (c >> 6), 0x80 | (c & 63));
  else if (c < 0x10000) out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  else out.push(0xf0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63)); } return out; }
function __sha256(str) {
  const K = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
  const H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
  const b = __utf8(str); const bitLen = b.length * 8; b.push(0x80); while (b.length % 64 !== 56) b.push(0);
  const hi = Math.floor(bitLen / 0x100000000), lo = bitLen >>> 0;
  b.push((hi >>> 24) & 255, (hi >>> 16) & 255, (hi >>> 8) & 255, hi & 255, (lo >>> 24) & 255, (lo >>> 16) & 255, (lo >>> 8) & 255, lo & 255);
  const W = new Array(64); const rotr = (x, n) => (x >>> n) | (x << (32 - n));
  for (let o = 0; o < b.length; o += 64) {
    for (let t = 0; t < 16; t += 1) W[t] = (b[o + 4 * t] << 24) | (b[o + 4 * t + 1] << 16) | (b[o + 4 * t + 2] << 8) | b[o + 4 * t + 3];
    for (let t = 16; t < 64; t += 1) { const s0 = rotr(W[t - 15], 7) ^ rotr(W[t - 15], 18) ^ (W[t - 15] >>> 3); const s1 = rotr(W[t - 2], 17) ^ rotr(W[t - 2], 19) ^ (W[t - 2] >>> 10); W[t] = (W[t - 16] + s0 + W[t - 7] + s1) | 0; }
    let [a, bb, c, d, e, f, g, h] = H;
    for (let t = 0; t < 64; t += 1) { const S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25); const ch = (e & f) ^ (~e & g); const t1 = (h + S1 + ch + K[t] + W[t]) | 0;
      const S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22); const mj = (a & bb) ^ (a & c) ^ (bb & c); const t2 = (S0 + mj) | 0;
      h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = bb; bb = a; a = (t1 + t2) | 0; }
    H[0] = (H[0] + a) | 0; H[1] = (H[1] + bb) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0; H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
  }
  return H.map(x => (x >>> 0).toString(16).padStart(8, '0')).join('');
}`;

/**
 * T2 (read-only): per sheet, the digest of the content OUTSIDE the windows. Canonical entry, NON-BLANK cells only, in
 * the used range's ForEach order: [A1, type, value, formula]. type: 'formula' (GetFormula starts with '=') | 'number'
 * (COUNT exactly 1) | 'text' (COUNT exactly 0); any other COUNT reply is UNOBSERVABLE -> an `unsupported` entry (the
 * observation is incomplete: no write before, no APPLIED after); a cell with an empty value and no formula is blank and is not an entry (a write that grows, or a
 * clear that shrinks, the used range only moves blank coordinates). value: GetValue as a string, EMPTY for a formula
 * cell (its value may follow its inputs, as in the cell-level L3). `canonical: true` returns the canonical text too
 * (verification only).
 */
function digestQuery(windows = [], { canonical = false, formulaIndex = false, nfSheets = [] } = {}) {
  // X-T01b: on the sheets in nfSheets each entry also carries the cell's number format (PUBLIC GetNumberFormat), and a
  // blank cell with a non-General format is an entry -- a format change outside the windows changes the hash. Without
  // nfSheets the generated body is byte-identical to the pre-X-T01b one.
  const nfs = (nfSheets || []).map(String);
  const rects = (windows || []).map(w => { const g = parseRange(w.range); return g ? { sheet: String(w.sheet), r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, range: rangeText(g) } : null; }).filter(Boolean);
  return `
${SHA256_SRC}
try {
  const LIMIT = ${MAX_TIERED_CELLS};
  const rects = ${JSON.stringify(rects)};
  const wf = Api.WorksheetFunction;
  if (!wf || typeof wf.COUNT !== 'function') return { ok: false, error: 'WorksheetFunction.COUNT unavailable' };
  const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
  let scanned = 0; let truncated = false; const unsupported = []; const out = [];${formulaIndex ? '\n  const formulas = []; let formulaIndexOver = false;' : ''}${nfs.length ? `\n  const nfSheets = ${JSON.stringify(nfs)};` : ''}
  for (const ws of Api.GetSheets()) {
    const name = ws.GetName();${nfs.length ? '\n    const withNf = nfSheets.includes(name);' : ''}
    const ur = ws.GetUsedRange();
    const used = ur ? String(ur.GetAddress(false, false, 'xlA1', false)) : null;
    const mine = rects.filter(r => r.sheet === name);
    const entries = [];
    if (used) ws.GetRange(used).ForEach(cell => {
      if (truncated) return;
      scanned += 1;
      if (scanned > LIMIT) { truncated = true; return; }
      const a = String(cell.GetAddress(false, false, 'xlA1', false));
      const m = /^([A-Z]+)(\\d+)$/.exec(a);
      if (!m) { unsupported.push({ field: name + '!' + a + '.address', reason: 'unparseable' }); return; }
      const col = colNum(m[1]); const row = Number(m[2]);
      if (mine.some(r => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2)) return;
      let v; let f; let n;
      try { v = cell.GetValue(); f = cell.GetFormula(); n = wf.COUNT(cell); } catch (err) { unsupported.push({ field: name + '!' + a, reason: String(err) }); return; }
      const isFormula = typeof f === 'string' && f.startsWith('=');${nfs.length ? `
      let nf = null;
      if (withNf) { try { nf = cell.GetNumberFormat(); } catch (err) { unsupported.push({ field: name + '!' + a + '.GetNumberFormat', reason: String(err) }); return; } }` : ''}${formulaIndex ? `
      // T3 (slice 2): every formula cell outside the windows, up to the ceiling; over it the index is incomplete
      if (isFormula) { if (formulas.length < ${MAX_FORMULA_INDEX}) formulas.push([name, a, f]); else formulaIndexOver = true; }` : ''}
      if (!isFormula && (v === '' || v === null || v === undefined)${nfs.length ? " && !(withNf && nf !== 'General')" : ''}) return;
      // the type of a non-formula cell rests on COUNT: only the calibrated 0 (text) / 1 (number) is an answer; any
      // other reply is UNOBSERVABLE and makes the observation incomplete -- never normalized to text
      if (!isFormula && n !== 0 && n !== 1) { unsupported.push({ field: name + '!' + a + '.COUNT', reason: 'UNOBSERVABLE:type (COUNT ' + (n === undefined ? 'undefined' : JSON.stringify(n)) + ')' }); return; }
      entries.push([a, isFormula ? 'formula' : (n === 1 ? 'number' : 'text'), isFormula ? '' : String(v), isFormula ? f : ''${nfs.length ? '].concat(withNf ? [String(nf)] : [])' : ']'});
    });
    const text = JSON.stringify(entries);
    const row = { sheet: name, used, excluded: mine.map(r => r.range), count: entries.length, hash: __sha256(text) };${nfs.length ? '\n    if (withNf) row.nf = true;' : ''}
    if (${canonical ? 'true' : 'false'}) row.canonical = text;
    out.push(row);
  }
  return { ok: true, digests: out, scanned: Math.min(scanned, LIMIT), truncated, unsupported${formulaIndex ? ', formulas, formulaIndexOver' : ''} };
} catch (err) { return { ok: false, error: String(err) }; }
`;
}

/**
 * The tier decision for one live read: a read-only probe of the used ranges, then tierOf(). A failed probe keeps the
 * full read (today's behaviour). `run` executes one editor body.
 */
async function tierFor(run, request) {
  const probe = await run(usedProbeQuery());
  return probe.ok && probe.result && probe.result.ok ? tierOf(probe.result.sheets, request || {}) : { tier: 'full', bound: null, probe: 'failed' };
}

/** T1 (the windows, the product's own cell reader) + T2 (the digest of the rest), in the same session. */
/** GH #96 follow-up: the T3 formula index is read with the dependency readers, unless the request opts out (noFormulaIndex). */
const formulaIndexWanted = request => !!(request && request.dependents) && !(request && request.noFormulaIndex);

async function tieredRead(run, windows, { formulaIndex = false, ...opts }, tier, runtime) {
  const r1 = await run(readQuery(windows, { ...opts, usedRanges: false }));
  const p = r1.ok ? r1.result : null;
  if (!p || p.ok !== true) return { ok: false, detail: (p && p.error) || r1.outcome || 'tiered window read failed' };
  const r2 = await run(digestQuery(windows, { formulaIndex, nfSheets: opts.formatSheets || [] }));
  const g = r2.ok ? r2.result : null;
  if (!g || g.ok !== true) return { ok: false, detail: (g && g.error) || r2.outcome || 'tiered digest read failed' };
  return { ok: true, authority: 'LIVE_READ', payload: tieredPayload(p, g, tier, runtime) };
}

/** A refused request above MAX_READ_CELLS keeps today's (truncated) read, with the limit named explicitly. */
const withLimitNote = (unsupported, tier) => unsupported.concat([{ field: 'observation', reason: `IMPLEMENTATION_LIMIT:${tier.reason} (bound ${tier.bound})` }]);

/** The tiered payload from the window read (T1) and the digest read (T2) of one observation (pure). */
function tieredPayload(p, g, tier, runtime) {
  const unsupported = p.unsupported.concat(g.unsupported);
  if (g.truncated) unsupported.push({ field: 'observation', reason: `IMPLEMENTATION_LIMIT:digest-scan-over-${MAX_TIERED_CELLS}-cells` });
  if (g.formulaIndexOver) unsupported.push({ field: 'formula-index', reason: `IMPLEMENTATION_LIMIT:formula-index-over-${MAX_FORMULA_INDEX}-formulas` });
  const out = { sheets: p.sheets, cells: p.cells, truncated: p.truncated || g.truncated, unsupported, objects: p.objects,
    tier: 'tiered', bound: tier.bound, digests: g.digests, scanned: g.scanned, runtime, ...(tier.bounded ? { bounded: true } : {}) };
  // slice 2: the formula index (T3) and the reads a dependency-guarded family needs, when they were asked
  if (Array.isArray(g.formulas)) out.formulaIndex = g.formulas;
  if (p.conditional !== undefined) out.conditional = p.conditional;
  if (p.filterReads && Object.keys(p.filterReads).length) out.filterReads = p.filterReads;
  if (p.commentReads && Object.keys(p.commentReads).length) out.commentReads = p.commentReads;
  // X-L13/6 fix: the same per-family extras readLive's non-tiered branch already forwards (lines ~1280-1284) --
  // T1's readQuery() returns these on `p` whenever TIERED_EXTRAS asked for them, but the tiered branch dropped
  // them on the floor, so eligibility passed while the family's own gate/classify still saw "not-read"
  if (p.panes !== undefined) { out.panes = p.panes; out.active = p.active; }
  if (p.validationReads && Object.keys(p.validationReads).length) out.validationReads = p.validationReads;
  if (p.cfReads && Object.keys(p.cfReads).length) out.cfReads = p.cfReads;
  if (p.protectionRangeReads && Object.keys(p.protectionRangeReads).length) out.protectionRangeReads = p.protectionRangeReads;
  if (p.protectionReads && Object.keys(p.protectionReads).length) out.protectionReads = p.protectionReads;
  if (p.runReads && Object.keys(p.runReads).length) out.runReads = p.runReads;
  if (p.visibility && typeof p.visibility === 'object') out.visibility = p.visibility;
  // X-L13/8 (zero LIVE r1 finding): the SAME class of bug the X-L13/6 fix above already names --
  // `layout` joined TIERED_EXTRAS but this branch never forwarded it, so layout_range's own L1
  // (gate()/verifyPrimitive(), which reads primary.layout[sheet]) saw "not-read" under BOUNDED.
  if (p.layout !== undefined) out.layout = p.layout;
  return out;
}

/**
 * GH #300 K2: the session's comment author -- the editor config's editorConfig.user, which Nextcloud signs for the
 * requester's OWN credential (the broker route) and the editor session itself runs as (K2 LIVE probe: DocInfo / User
 * equal the config's id / name). `id` is the stable, instance-prefixed Nextcloud uid; `name` the display name derived
 * from it. Never a tool argument. null when either is missing or empty: no reliable identity -> no comment.
 */
function sessionUserOf(config) {
  const u = config && config.editorConfig && config.editorConfig.user;
  const id = u && typeof u.id === 'string' ? u.id.trim() : '';
  const name = u && typeof u.name === 'string' ? u.name.trim() : '';
  return id && name ? Object.freeze({ id, name }) : null;
}

function createEuroOfficeXlsxDriver({ url, user, pass, loadPlaywright, timeoutMs = 90000, pollMs = 100, commandTimeoutMs = 15000,
  // the fixed bound for the DS save acknowledgement at close (P1, owner GO owner channel); never extended at run time
  closeAckTimeoutMs = 30000,
  // forceSave (owner GO, owner channel): the fixed bounds for the precondition (the last save acked) and the confirmation
  forceSaveAckTimeoutMs = 30000, forceSaveConfirmTimeoutMs = 30000,
  // TEST SEAMS for the LIVE negative controls only (a reader on an unproven build / a broken reader);
  // production callers never pass them
  // (B11, Bridge): `dependencyBody(build)` returns the callCommand body of the dependency read
  dependencyBody = dependentsBridgeBody,
  // TEST SEAM for the M9-2 O1 LIVE dual-path parity harness only: 'wire' (default, production) or 'object' (the
  // pre-O1 transfer), or a function returning one of them per read; production callers never pass it
  readTransport = 'wire',
  // M11.2: the editor-config source (config-source.cjs). The PRODUCT passes the broker source and NO user/pass;
  // absent = the legacy Basic-auth source built from url/user/pass (test/harness only)
  configSource = null }) {
  const source = resolveConfigSource({ configSource, url, user, pass });
  const state = { browser: null, page: null, frame: null, host: null, config: null, runtimeId: null, cacheKey: null, runtimeRefusal: null, fileId: null, saveAck: null, mutated: false, forceSaveOp: null };

  // M11.2: the editor config comes from the config source seam (config-source.cjs) -- the driver never holds a
  // credential of its own; without a configSource the legacy Basic-auth source (test/harness only) is used
  async function fetchConfig(fileId, operation) {
    return source.getEditorConfig({ fileId: String(fileId), operation: operation || null });
  }

  // M11.3b.1: the session's loopback config host is token-protected (config-host.cjs)
  const startHost = startConfigHost;

  const inEditor = fn => state.frame.evaluate(fn);

  /**
   * The session's DS runtime identity -- ONE canonical bridge handshake (core/runtime-identity.cjs), the only
   * compatibility authority. The editor URL's cache key is recorded for diagnostics only. A failed handshake leaves
   * runtimeId null: every runtime-gated path refuses (fail-closed); the bridge reads gate themselves per call.
   */
  async function establishRuntime() {
    state.cacheKey = cacheKeyOf(state.frame.url());
    const r = await runInEditor(HANDSHAKE_BODY).catch(err => ({ ok: false, outcome: `threw:${String(err && err.message || err).slice(0, 80)}` }));
    const c = r && r.ok ? classifyHandshake(r.result, ACCEPTED_RUNTIMES.xlsx) : { ok: false, reason: `HANDSHAKE_FAILED:${(r && r.outcome) || 'unknown'}` };
    state.runtimeId = c.ok ? c.runtimeId : null;
    state.runtimeRefusal = c.ok ? null : c.reason;
  }

  /** The B11 dependency inventory through the DS Bridge, in the SAME session right after the public read, before any dispatch. */
  async function addDependents(payload) {
    const r = await runInEditor(dependencyBody()).catch(err => ({ ok: false, outcome: `threw:${String(err && err.message || err).slice(0, 80)}` }));
    const d = r && r.ok ? r.result : null;
    payload.dependents = d && typeof d === 'object' && 'names' in d && 'validations' in d
      ? { names: d.names, validations: d.validations }
      : `UNOBSERVABLE:dependency-read-failed:${(r && r.outcome) || 'unexpected-shape'}`;
    // true only when neither axis was refused by the bridge's own identity / capability gate
    const gated = x => typeof x === 'string' && /^UNOBSERVABLE:(BRIDGE_UNAVAILABLE|MANIFEST_MALFORMED|BRIDGE_VERSION_MISMATCH|RUNTIME_VERSION_MISMATCH|PATCH_VERSION_MISMATCH|CAPABILITY_MISSING|CAPABILITY_VERSION_MISMATCH)$/.test(x);
    payload.runtime.dependencyRuntime = typeof payload.dependents === 'object' && !gated(payload.dependents.names) && !gated(payload.dependents.validations);
  }

  /**
   * M9-2 O1: a READ body (readQuery / digestQuery) through the lossless wire codec (wire-codec.cjs) -- the body is
   * unchanged, its result crosses the editor boundary as one string and is decoded to the same value. Any decode
   * failure is a FAILED command (the live read fails), never a partial payload.
   */
  async function runRead(body) {
    const mode = typeof readTransport === 'function' ? readTransport() : readTransport;
    if (mode === 'object') return runInEditor(body);
    if (mode !== 'wire') return { ok: false, outcome: `read-transport-unknown:${String(mode)}` };
    const r = await runInEditor(wireBody(body));
    if (!r.ok) return r;
    const d = decodeWire(r.result);
    return d.ok ? { ...r, result: d.value } : { ok: false, outcome: d.reason };
  }

  async function runInEditor(body) {
    return state.frame.evaluate(({ body: src, limit }) => new Promise(resolve => {
      const e = window.editor || (window.Asc || {}).editor;
      if (!e || typeof e.callCommand !== 'function') { resolve({ ok: false, outcome: 'callcommand-unavailable' }); return; }
      let done = false;
      const finish = v => { if (!done) { done = true; clearTimeout(timer); resolve(v); } };
      const timer = setTimeout(() => finish({ ok: false, outcome: 'command-callback-timeout' }), limit);
      try { e.callCommand(new Function(src), true, r => finish({ ok: true, outcome: 'command-completed', result: r === undefined ? null : r })); }
      catch (err) { finish({ ok: false, outcome: 'call-threw', error: String(err && (err.message || err)).slice(0, 500) }); }
    }), { body, limit: commandTimeoutMs });
  }

  /**
   * Closes the browser and the loopback host. A browser.close() failure is REPORTED, not swallowed (owner review of
   * 582a6e8): the close result carries it and a close with a failed teardown is never ok. The host is closed either way.
   */
  async function teardown() {
    const torn = await closeBrowser(state.browser);
    if (state.host) await state.host.close();
    state.browser = state.page = state.frame = state.host = state.saveAck = state.forceSaveOp = null;
    return torn;
  }

  const aliveNow = () => (state.page && state.frame ? inEditor(() => true).then(() => true, () => false) : Promise.resolve(false));

  const readCanSave = () => canSaveNow().then(r => (r && r.ok ? r.canSave : null), () => null);

  async function canSaveNow() {
    return inEditor(() => {
      const e = window.editor || (window.Asc || {}).editor;
      try {
        const v = typeof e.asc_isDocumentCanSave === 'function' ? e.asc_isDocumentCanSave() : e.canSave;
        return typeof v === 'boolean' ? { ok: true, canSave: v } : { ok: false, canSave: null, detail: 'canSave not boolean' };
      } catch (err) { return { ok: false, canSave: null, detail: String(err) }; }
    });
  }

  return {
    async open({ fileId, operation = null }) {
      const loaded = loadPlaywright();
      if (!loaded.ok) return { ok: false, detail: loaded.detail || 'playwright unavailable' };
      try {
        state.fileId = String(fileId);
        state.config = await fetchConfig(fileId, operation);
        state.host = await startHost(state.config);
        state.browser = await loaded.pw.chromium.launch();
        state.page = await state.browser.newPage({ viewport: { width: 1400, height: 900 } });
        // the editor <-> DS socket of THIS session: the config's DS host and document key (the socket's shardkey)
        state.mutated = false;
        state.forceSaveOp = null;
        state.saveAck = observePage(state.page, dsIdentity(state.config));
        await state.page.goto(state.host.url, { waitUntil: 'domcontentloaded', timeout: timeoutMs });
        try {
          await state.page.waitForFunction(() => window.__ready === true || window.__error, null, { timeout: timeoutMs, polling: pollMs });
        } catch {
          // GitHub #77: this is reachable only when NEITHER __ready NOR __error ever became truthy in the
          // whole window (onError sets __error, so it resolves the wait above, not this catch) -- read back
          // onWarning/onInfo, or confirm literally nothing arrived
          const e = new Error(openWaitFailureDetail(await readDsOpenState(state.page)));
          e.name = 'EditorWaitTimeout';
          throw e;
        }
        // Condition-based: the editor is open once its live spreadsheet model names its active sheet.
        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
          state.frame = state.page.frames().find(f => /spreadsheeteditor/.test(f.url())) || null;
          if (state.frame) {
            const r = await runInEditor('return Api.GetActiveSheet().GetName();').catch(() => null);
            if (r && r.ok && typeof r.result === 'string' && r.result) { await establishRuntime(); return { ok: true, id: `eo-${fileId}-${Date.now()}` }; }
          }
          await state.page.waitForTimeout(pollMs);
        }
        // GitHub #77: the ready/error wait above resolves as soon as DS reports EITHER __ready or __error --
        // an onError does NOT fail-fast (whether onDocumentReady still arrives afterward is NEM-MERT, so a
        // successful open must not be cut short), which means a real onError lands HERE, not in the catch
        // below. Read back whatever DS reported before the generic message swallows it.
        // the DS description is page-reported text -- redacted like every other detail that leaves open()
        const dsDetail = (state.host ? state.host.redact : (x => x))(openWaitFailureDetail(await readDsOpenState(state.page)));
        await teardown();
        return { ok: false, detail: `live spreadsheet model not ready (${dsDetail})` };
      } catch (err) {
        // M11.3b.1: a navigation error carries the host URL -- the session token is redacted before anything leaves
        const redact = state.host ? state.host.redact : (x => x);
        await teardown();
        // a controlled config-source refusal (M11.2) keeps its product error class; the editor was never started
        if (err && err.name === 'ConfigSourceError') return { ok: false, code: err.code, detail: err.detail || err.code };
        // GitHub #77: the ready/error wait timed out -- err.message already names the DS's own code/description,
        // what onWarning/onInfo recorded, or 'no-ds-event' (config-host.cjs openWaitFailureDetail)
        if (err && err.name === 'EditorWaitTimeout') return { ok: false, detail: redact(err.message) };
        return { ok: false, detail: redact(String(err && err.message || err)) };
      }
    },

    /**
     * close -- P1 (owner GO, owner channel): BEFORE requestClose and the browser teardown, wait (bounded by
     * closeAckTimeoutMs, condition-based) until the DS has acknowledged the editor's last saveChanges and the
     * editor has nothing left to send. requestClose comes after the wait on purpose: the editor's own
     * onRequestClose calls asc_stopSaving() while the document still counts as modified. A missing event, an
     * unpairable send/ack or the timeout gives ok:false (a not-verified close); the browser is closed anyway.
     * No rewrite, no unlock. A verified ack is NOT storage proof: the stored file is a separate check.
     */
    async close() {
      // the observer and the forceSave operation window are kept until the END of the close (owner review of 2ab3aca)
      const obs = state.page ? state.saveAck : null;
      const op = state.forceSaveOp;
      const pre = await settleSaveAck(obs, { timeoutMs: closeAckTimeoutMs, pollMs, wait: defaultWait, mutated: state.mutated, readCanSave });
      const mark = obs ? obs.mark() : null;
      let lifecycle = { ok: false, detail: 'no editor' };
      if (state.page) {
        lifecycle = await state.page.evaluate(async () => {
          const d = window.__docEditor;
          if (!d || typeof d.requestClose !== 'function') return { ok: false, detail: 'requestClose-unavailable' };
          try { const r = d.requestClose(); if (r && typeof r.then === 'function') await r; return { ok: true, detail: 'requestClose-called' }; }
          catch (err) { return { ok: false, detail: String(err) }; }
        }).catch(err => ({ ok: false, detail: String(err) }));
      }
      const torn = await teardown();
      await defaultWait(0); // deliver the websocket events queued during the teardown before the final verdict
      // the FINAL close-ack and confirmation: every event observed during requestClose and the teardown counts; a
      // failed requestClose or a failed teardown is a failed close, and no save confirmation tied to it stands
      return closeResult(obs, { pre, mark, op, lifecycle, torn });
    },

    /**
     * release -- GH #100 U1 (UNUSED until the session registry unit): the CALL END of a session that STAYS open. It is close() minus the
     * two destructive steps: the same bounded wait for the DS acknowledgement of the editor's last saveChanges and the same re-evaluation of
     * the forceSave confirmation (finishClose / closeResult), but NO requestClose and NO teardown -- the browser, the loopback host and the
     * observer stay. `ok` therefore means "the call ended with every change sent and acked" and, with `confirmation`, carries the same save
     * claim a close would; a release that is not ok is a failed call end, and the caller must close() the session instead of reusing it.
     * close() itself is untouched.
     */
    async release() {
      const obs = state.page ? state.saveAck : null;
      const op = state.forceSaveOp;
      const pre = await settleSaveAck(obs, { timeoutMs: closeAckTimeoutMs, pollMs, wait: defaultWait, mutated: state.mutated, readCanSave });
      const mark = obs ? obs.mark() : null;
      const live = await aliveNow();
      return closeResult(obs, { pre, mark, op, lifecycle: { ok: live, detail: live ? 'session-kept' : 'no live editor' }, torn: { ok: true, detail: 'kept' } });
    },

    /**
     * resetForNextCall -- GH #100 U1: the per-call state a kept session must not carry into the next call. `mutated` is "a write was attempted
     * in this call" (a close needs an acked save from there) and `forceSaveOp` is the last forceSave operation window; both belong to ONE call.
     * The save observer is NOT reset: its pairing is per DS socket and per send, measured to confirm every forceSave of a session (probe, 3 of 3).
     */
    resetForNextCall() {
      state.mutated = false;
      state.forceSaveOp = null;
      return { ok: !!(state.page && state.frame) };
    },

    /**
     * health -- GH #100 U3: what the call-start check of a KEPT session needs, in one read-only call that never throws. `canSave` is null when
     * it could not be read (and null is never read as "nothing unsaved"); `ds.open` is an authenticated editor <-> DS socket that is still open
     * (the co-editing sync of a foreign editor's write travels on it, measured LIVE); `anomalies` counts the observer's anomalies since open.
     */
    async health() {
      const live = await aliveNow();
      const canSave = live ? await readCanSave() : null;
      const ev = state.saveAck ? state.saveAck.evidence() : null;
      const sockets = ev ? ev.dsSockets : [];
      return { alive: live, canSave, ds: { open: sockets.some(x => x.open && x.authed), sockets: sockets.length, busy: sockets.some(x => x.busy) }, anomalies: ev ? ev.anomalies.length : null };
    },

    /** documentKey -- GH #100 U3: the DS document key of the config this session was OPENED with (null when none). Compared in memory, never logged. */
    documentKey() {
      const k = state.config && state.config.document ? state.config.document.key : null;
      return typeof k === 'string' && k !== '' ? k : null;
    },

    /** alive -- GH #100 U1: is the kept session's editor frame still answering (a cheap evaluate)? false on any failure, never throws. */
    async alive() {
      return aliveNow();
    },

    /**
     * forceSave -- FORCESAVE_CONFIRMED (owner GO, owner channel), on the P1 observer, in this order:
     *   1. wait (bounded) until every change is sent and the LAST send is an acked final chunk on this session's
     *      DS socket, and the editor has nothing left to send -- no request is made before that;
     *   2. only then the editor's forceSave() (one forceSaveStart);
     *   3. wait (bounded) for THIS request's own code 0 and success:true, both carrying that ack's save time.
     * A failed precondition: ok:false, no request. A session with nothing saved (no mutation, no save traffic): no
     * request is forced just to get a confirmation. Any code other than 0, success:false, a missing or unpairable
     * answer, a newer send, or the timeout: confirmed:false. Never a rewrite. confirmed:true means the DS reported
     * that save stored -- FORCESAVE_CONFIRMED, not PERSISTED on its own.
     */
    async forceSave() {
      const r = await confirmForceSave(state.page && state.saveAck, { readCanSave, mutated: state.mutated, pollMs, wait: defaultWait,
        ackTimeoutMs: forceSaveAckTimeoutMs, confirmTimeoutMs: forceSaveConfirmTimeoutMs,
        requestForceSave: () => inEditor(() => {
          const e = window.editor || (window.Asc || {}).editor;
          if (!e || typeof e.forceSave !== 'function') return { ok: false, detail: 'forceSave-unavailable' };
          return { ok: true, sent: e.forceSave() === true };
        }) });
      if (r.forceSave && r.forceSave.op) state.forceSaveOp = r.forceSave.op; // re-evaluated at close
      return r;
    },

    async canSave() {
      return canSaveNow();
    },

    async readLive(_handle, request = {}) {
      const pagePerSheet = EO_PAGE_LAYOUT_PER_SHEET_RUNTIMES.includes(state.runtimeId);
      const drawingAnchor = !!(request && request.drawingAnchor);
      const hydrateSheets = drawingAnchor; // the bridge client gates the runtime itself (no driver-side runtime list)
      // the tier decision (M7 S32 slice 1): one read-only probe of the used ranges first; X-N01b: a bounded tier reads the
      // request's boundedVariant instead of the request (only above the read bound -- tierOf)
      const tier = await tierFor(runInEditor, request);
      if (tier.bounded) request = request.boundedVariant;
      const wantDeps = !!(request && request.dependents);
      const opts = { panes: !!(request && request.panes), validationRanges: (request && request.validationRanges) || [], cfRanges: (request && request.cfRanges) || [], runCells: (request && request.runCells) || [], arrayCells: (request && request.arrayCells) || [], sheetVisibility: !!(request && request.sheetVisibility), filterSheets: (request && request.filterSheets) || [], chartSheets: (request && request.chartSheets) || [], pageSheets: (request && request.pageSheets) || [], pagePerSheet, printSheets: (request && request.printSheets) || [], printWide: true, colBreaks: !!(request && request.colBreaks), drawingSheets: (request && request.drawingSheets) || [], drawingAnchor, hydrateSheets, protectionSheets: (request && request.protectionSheets) || [], protectionRanges: (request && request.protectionRanges) || [], commentSheets: (request && request.commentSheets) || [], styleSheets: (request && request.styleSheets) || [], formatSheets: (request && request.formatSheets) || [], objectSheets: wantDeps ? ['*'] : ((request && request.objectSheets) || []), conditional: wantDeps, layout: (request && request.layout) || null };
      if (tier.tier === 'window') {
        const w = await runRead(readQuery((request && request.windows) || [], { ...opts, usedRanges: false }));
        const p = w.ok ? w.result : null;
        if (!p || p.ok !== true) return { ok: false, detail: (p && p.error) || w.outcome || 'window read failed' };
        return { ok: true, authority: 'LIVE_READ', payload: { sheets: p.sheets, cells: p.cells, truncated: p.truncated, unsupported: p.unsupported, objects: p.objects, tier: 'window', bound: tier.bound, runtime: { identity: state.runtimeId, cacheKey: state.cacheKey, fileId: state.fileId || null } } };
      }
      if (tier.tier === 'tiered') {
        const t = await tieredRead(runRead, (request && request.windows) || [], { ...opts, formulaIndex: formulaIndexWanted(request) }, tier, { identity: state.runtimeId, cacheKey: state.cacheKey, fileId: state.fileId || null });
        if (t.ok && wantDeps) await addDependents(t.payload);
        if (t.ok && request && request.sessionUser) t.payload.sessionUser = sessionUserOf(state.config) || 'UNOBSERVABLE:no-session-user';
        return t;
      }
      const r = await runRead(readQuery((request && request.windows) || [], opts));
      const p = r.ok ? r.result : null;
      if (!p || p.ok !== true) return { ok: false, detail: (p && p.error) || r.outcome || 'live read failed' };
      const payload = { sheets: p.sheets, cells: p.cells, truncated: p.truncated, unsupported: p.unsupported, objects: p.objects, runtime: { identity: state.runtimeId, cacheKey: state.cacheKey, fileId: state.fileId || null } };
      if (tier.tier === 'refused' && p.truncated) payload.unsupported = withLimitNote(p.unsupported, tier);
      if (p.layout !== undefined) payload.layout = p.layout;
      if (p.panes !== undefined) { payload.panes = p.panes; payload.active = p.active; }
      if (p.validationReads && Object.keys(p.validationReads).length) payload.validationReads = p.validationReads;
      if (p.cfReads && Object.keys(p.cfReads).length) payload.cfReads = p.cfReads;
      if (p.protectionRangeReads && Object.keys(p.protectionRangeReads).length) payload.protectionRangeReads = p.protectionRangeReads;
      if (p.protectionReads && Object.keys(p.protectionReads).length) payload.protectionReads = p.protectionReads;
      if (p.runReads && Object.keys(p.runReads).length) payload.runReads = p.runReads;
      if (p.arrayReads && Object.keys(p.arrayReads).length) payload.arrayReads = p.arrayReads;
      if (p.visibility && typeof p.visibility === 'object') payload.visibility = p.visibility;
      if (p.filterReads && Object.keys(p.filterReads).length) payload.filterReads = p.filterReads;
      if (p.commentReads && Object.keys(p.commentReads).length) payload.commentReads = p.commentReads;
      if (request && request.sessionUser) payload.sessionUser = sessionUserOf(state.config) || 'UNOBSERVABLE:no-session-user';
      if (p.chartReads && Object.keys(p.chartReads).length) payload.chartReads = p.chartReads;
      if (p.drawingReads && Object.keys(p.drawingReads).length) payload.drawingReads = p.drawingReads;
      if (p.pageReads && Object.keys(p.pageReads).length) payload.pageReads = p.pageReads;
      if (p.printReads && Object.keys(p.printReads).length) payload.printReads = p.printReads;
      if (wantDeps) {
        payload.conditional = p.conditional;
        await addDependents(payload);
      }
      return { ok: true, authority: 'LIVE_READ', payload };
    },

    async dispatch(_handle, command) {
      state.mutated = true; // any write attempt: from here a close needs an observed, acked save (owner review point 3)
      const b03 = !!(command && command.meta && command.meta.family === 'sheets' && command.meta.action === 'copy_sheet');
      if (b03 && !(command.meta.bridgeCopyWrite === true && B03_BRIDGE_WRITE_RUNTIMES.includes(state.runtimeId))) {
        return { ok: false, dispatched: false, outcome: 'b03-bridge-write-runtime-not-proven', runtime: state.runtimeId, refusal: state.runtimeRefusal };
      }
      const b10 = !!(command && command.meta && command.meta.family === 'print');
      if (b10 && !(command.meta.bridgePrintWrite === true && B10_BRIDGE_WRITE_RUNTIMES.includes(state.runtimeId))) {
        return { ok: false, dispatched: false, outcome: 'b10-bridge-write-runtime-not-proven', runtime: state.runtimeId, refusal: state.runtimeRefusal };
      }
      const b05 = !!(command && command.meta && command.meta.family === 'format' && command.meta.bridgeAlignExtWrite === true);
      if (b05 && !B05_BRIDGE_WRITE_RUNTIMES.includes(state.runtimeId)) {
        return { ok: false, dispatched: false, outcome: 'b05-bridge-write-runtime-not-proven', runtime: state.runtimeId, refusal: state.runtimeRefusal };
      }
      // X-O05/3: the capability ID travels in meta (not a bare boolean, unlike b03/b05/b10) --
      // set_chart_style has TWO distinct chart-style Bridge WRITEs sharing this one gate.
      const chartStyleCapId = command && command.meta && command.meta.family === 'charts' ? command.meta.bridgeChartStyleWrite : null;
      const chartStyleCap = chartStyleCapId ? Object.values(BRIDGE_WRITES).find(c => c.id === chartStyleCapId) : null;
      if (chartStyleCapId && (!chartStyleCap || !B22_BRIDGE_WRITE_RUNTIMES.includes(state.runtimeId))) {
        return { ok: false, dispatched: false, outcome: 'chart-style-bridge-write-runtime-not-proven', runtime: state.runtimeId, refusal: state.runtimeRefusal };
      }
      // X-O07/3: a single boolean gate, like b03/b05/b10 above (unlike chart-style's two-capability
      // lookup) -- delete_chart has exactly one Bridge WRITE behind it.
      const b23 = !!(command && command.meta && command.meta.family === 'charts' && command.meta.bridgeChartDeleteWrite === true);
      if (b23 && !B23_BRIDGE_WRITE_RUNTIMES.includes(state.runtimeId)) {
        return { ok: false, dispatched: false, outcome: 'b23-bridge-write-runtime-not-proven', runtime: state.runtimeId, refusal: state.runtimeRefusal };
      }
      // X-O05/5-C: the SAME meta-carries-the-capability-ID shape as chart-style above (two
      // capabilities, one gate) -- but on B24_BRIDGE_WRITE_RUNTIMES (/15-ONLY), a DIFFERENT meta
      // key (bridgeSeriesStyleWrite) so this never collides with the B22 chart-style gate above.
      const seriesStyleCapId = command && command.meta && command.meta.family === 'charts' ? command.meta.bridgeSeriesStyleWrite : null;
      const seriesStyleCap = seriesStyleCapId ? Object.values(BRIDGE_WRITES).find(c => c.id === seriesStyleCapId) : null;
      if (seriesStyleCapId && (!seriesStyleCap || !B24_BRIDGE_WRITE_RUNTIMES.includes(state.runtimeId))) {
        return { ok: false, dispatched: false, outcome: 'series-style-bridge-write-runtime-not-proven', runtime: state.runtimeId, refusal: state.runtimeRefusal };
      }
      if (!command || command.kind !== 'callCommand' || typeof command.body !== 'string') {
        return { ok: false, dispatched: false, detail: 'unsupported command kind' };
      }
      // GH #326 (U4b): the adapter's own race-check, in FRONT of every other prelude (they only declare
      // helpers, never execute on their own, so running first/last among themselves does not matter --
      // but THIS one returns early on a mismatch, so it must run before anything with a side effect does).
      const raceChecks = command.meta && command.meta.protectionRaceCheck;
      // GH #300 K2: a comment write gets its author from the SESSION (the signed editor config), never from the command;
      // no reliable identity -> the write is not dispatched
      let authorPrelude = '';
      if (command.meta && command.meta.sessionAuthor === true) {
        const author = sessionUserOf(state.config);
        if (!author) return { ok: false, dispatched: false, outcome: 'comment-author-identity-unavailable' };
        authorPrelude = `const __sessionAuthor = Object.freeze(${JSON.stringify(author)});\n`;
      }
      const body = (raceChecks ? protectionRacePrelude(raceChecks) : '') + authorPrelude + command.body;
      if (b05) return b05DispatchResult(await runInEditor(b05BridgeWritePrelude() + body));
      if (b10) return b10DispatchResult(await runInEditor(b10BridgeWritePrelude() + body));
      if (b03) return b03DispatchResult(await runInEditor(b03BridgeWritePrelude() + body));
      if (chartStyleCap) return chartStyleDispatchResult(await runInEditor(chartStyleBridgeWritePrelude(chartStyleCap) + body));
      if (b23) return chartDeleteDispatchResult(await runInEditor(chartDeleteBridgeWritePrelude() + body));
      if (seriesStyleCap) return seriesStyleDispatchResult(await runInEditor(seriesStyleBridgeWritePrelude(seriesStyleCap) + body));
      const r = await runInEditor(body);
      return { ok: r.ok === true, dispatched: true, outcome: r.outcome, error: r.error || null };
    },

    async readPackage() {
      const docUrl = state.config && state.config.document && state.config.document.url;
      if (!docUrl) return { ok: false, detail: 'no document url in the editor config' };
      const r = await fetch(docUrl);
      if (!r.ok) return { ok: false, detail: `package HTTP ${r.status}` };
      return { ok: true, bytes: Buffer.from(await r.arrayBuffer()) };
    },
  };
}

module.exports = { createEuroOfficeXlsxDriver, sessionUserOf, TIERED_EXTRAS, readQuery, usedProbeQuery, digestQuery, tierOf, tieredPayload, withLimitNote, formulaIndexWanted, SHA256_SRC, MAX_TIERED_CELLS, dependentsBridgeBody, EO_PAGE_LAYOUT_PER_SHEET_RUNTIMES, B10_BRIDGE_WRITE_RUNTIMES, B05_BRIDGE_WRITE_RUNTIMES, B03_BRIDGE_WRITE_RUNTIMES, b05BridgeWritePrelude, b05DispatchResult, b10BridgeWritePrelude, b10DispatchResult, b03BridgeWritePrelude, b03DispatchResult, MAX_READ_CELLS, MAX_COMMENT_SCAN_CELLS, MAX_COMMENTS, MAX_LAYOUT_ITEMS, MAX_CHART_LABEL_POINTS, B22_BRIDGE_WRITE_RUNTIMES, chartStyleBridgeWritePrelude, chartStyleDispatchResult, B23_BRIDGE_WRITE_RUNTIMES, chartDeleteBridgeWritePrelude, chartDeleteDispatchResult, B24_BRIDGE_WRITE_RUNTIMES, seriesStyleBridgeWritePrelude, seriesStyleDispatchResult, protectionRacePrelude };
