'use strict';

const { parseRange, rangeText, addr, colName } = require('../a1.cjs');
const { sheetNameProblem } = require('../sheet-name.cjs');
const { cellAt } = require('../inventory.cjs');
const cells = require('./cells.cjs');
const drawings = require('../drawings.cjs');
const { XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B24_BRIDGE_WRITE_RUNTIME_IDS } = require('../runtime.cjs');

/**
 * X-O05/3 (internal tracker; reviewed and accepted): gapOverlap/seriesOrder go through the DS
 * Bridge WRITE capabilities (src/drivers/eurooffice/ds-bridge.cjs's BRIDGE_WRITES.GAP_OVERLAP_WRITE
 * / SERIES_ORDER_WRITE) -- neither has a PUBLIC setter. gapOverlap's internal method
 * (setGapWidth()/setOverlap()) IS the persisting path (PRE-DEPLOY LIVE, internal tracker: a direct
 * data-field assignment -- the read side's own path -- APPEARS to apply in-session but does not
 * survive a reopen; the setter does, the SAME setter X-O05/1's own measurement used). order's
 * internal method (setOrder()) is likewise not public. The capability IDs are literal here, not
 * imported: the adapter tree imports nothing outside itself (runtime.cjs's own header comment),
 * the SAME boundary XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS already respects for the runtime-id list.
 */
const GAP_OVERLAP_WRITE_ID = 'chart.bar.gapOverlap.write/v1';
const SERIES_ORDER_WRITE_ID = 'chart.series.order.write/v1';
const B22_BRIDGE_WRITE_RUNTIMES = XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS;
/**
 * X-O07/3 (internal tracker; X-O07/2 LIVE MEASURED and reviewed): delete_chart goes through the
 * DS Bridge WRITE capability chart.drawing.delete.write/v1 -- the public chart API has no delete
 * method of its own (the X-O07 reflection probe found only series-level RemoveSeria). Unlike
 * gapOverlap/seriesOrder above, only ONE capability sits behind this gate, so the driver dispatch
 * uses a bare boolean (meta.bridgeChartDeleteWrite, like bridgeCopyWrite/bridgeAlignExtWrite/
 * bridgePrintWrite) rather than a capability-ID lookup -- the id itself never needs to travel
 * through this file.
 */
const B23_BRIDGE_WRITE_RUNTIMES = XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS;
/**
 * X-O05/5-C (internal tracker; X-O05/5-B round 7b+8d+8k+8m+8q LIVE MEASURED, reviewed and
 * accepted): a chart series' own marker symbol/size (chart.series.marker.write/v1) and line
 * dash/smooth (chart.series.lineStyle.write/v1) -- TWO MORE DS Bridge WRITE capabilities,
 * served ONLY on /15 (UNLIKE B22/B23: neither exists on /12../14 at all, runtime.cjs's own
 * header comment). WRITE-ONLY, like the six ORIGINAL style targets (X-O04/1+2) -- neither
 * field is readable back through any capability this product has (markerStyle.read/v1 covers
 * fill/outline, a DIFFERENT pair of fields on the same `series.marker` object; no read exists
 * for dash/smooth at all) -- `styleNotProven` below names both, every time, never
 * semanticsNotProven-free the way marker(fill/outline)/gridlines/gapOverlap/seriesOrder are.
 */
const MARKER_SYMBOL_WRITE_ID = 'chart.series.marker.write/v1';
const LINE_STYLE_WRITE_ID = 'chart.series.lineStyle.write/v1';
const B24_BRIDGE_WRITE_RUNTIMES = XLSX_B24_BRIDGE_WRITE_RUNTIME_IDS;
/** The two WRITE-ONLY, B24-gated style targets (`markerSymbol`/`lineStyle`) -- distinct from
 * `CHART_STYLE_BRIDGE_WRITE_TARGETS` above (gapOverlap/seriesOrder), which ARE read-and-compare
 * verifiable. Never added to that set: these two are NEVER SATISFIED, always APPLY, the SAME
 * write-only discipline the six ORIGINAL targets (STYLE_TARGETS) already follow. */
const CHART_SERIES_STYLE_WRITE_ONLY_TARGETS = Object.freeze(['markerSymbol', 'lineStyle']);
/** ECMA-376 ST_MarkerStyle, minus auto/picture (round8d/build-rev15.py's own measured set) --
 * 'none' goes through the handler's own setNoneMarker() path, server-side; this client-side enum
 * only validates the REQUEST shape, mirroring the ALREADY-RELEASED `symbol` enum convention
 * nowhere else in this file (new to this target). */
const MARKER_SYMBOL_ENUM = Object.freeze(['circle', 'dash', 'diamond', 'dot', 'none', 'plus', 'square', 'star', 'triangle', 'x']);
/** ECMA-376 ST_PresetLineDashVal, the full 11-value enum (build-rev15.py's own DASH_ENUM, same order). */
const LINE_DASH_ENUM = Object.freeze(['solid', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot']);
/** lineStyle's own measured chart-type scope: R8q's LIVE probe proved DASH on a LINE series WITH
 * an existing line (L1), a LINE series with NO existing line (L2, a theme default), AND a BAR
 * series' own line (B1) -- bar and lineNormal are therefore `dash`'s measured scope, the SAME
 * "never wider than what was LIVE-measured" discipline MARKER_CHART_TYPES below already follows
 * for the fill/outline marker target. markerSymbol reuses MARKER_CHART_TYPES itself
 * (round8/round8d's own symbol-write-and-readback measurement ran on scatter + lineNormal
 * charts -- the SAME `series.marker` sub-object the fill/outline target already scopes there).
 *
 * `smooth` is NARROWER: R8q's B1 (bar) case requested dash ONLY, never smooth -- so `smooth` on
 * a bar series is UNMEASURED, and X-O05/5-C's own first LIVE round (internal tracker) found it
 * fails at the real engine (CSer.setSmooth does not exist on a bar series; the DS handler's own
 * `typeof s.setSmooth !== 'function'` guard returns `setsmooth-missing`, which the product's
 * generic bridge-write client folds into an UNEXPECTED-shaped, uncertain dispatch refusal --
 * DISPATCHED_OUTCOME_UNKNOWN, never a clean FAILED). See LINE_STYLE_SMOOTH_CHART_TYPES below:
 * `smooth` is refused BEFORE dispatch, named, on anything outside its OWN (narrower) measured
 * scope -- the SAME discipline this comment already describes for `dash`, just a second,
 * independent gate (a chart type can be in one list and not the other: bar carries `dash` but
 * not `smooth`).
 */
const LINE_STYLE_CHART_TYPES = Object.freeze(['bar', 'lineNormal']);
/** `smooth`'s OWN, narrower measured scope -- see the comment above `LINE_STYLE_CHART_TYPES`.
 * Only `lineNormal` is LIVE-proven (R8q's L1/L2); `bar` carries `dash` but not `smooth`. */
const LINE_STYLE_SMOOTH_CHART_TYPES = Object.freeze(['lineNormal']);

/**
 * XLSX `charts` family: set_chart (create) / rename_chart (M6 wave 7).
 *
 * The reader is PUBLIC (the driver's chartReads: every chart's GetName / GetChartType / GetTitle / series name, values
 * and category formulas, plus the pivot NAMES, of every sheet). Audit group C: CH1 / CH18 MIGRATE; CH2 SOURCE_DEFECT --
 * the source never read the data range, so a chart with the same name / type / title / size over ANOTHER range
 * "satisfied" the request; CH4 the size reads 0x0 after a reopen (so the size is not part of the contract); CH20 there
 * is no public chart Delete on this build (W7 probe K5: no Delete / Remove on the chart or the drawing -- delete_chart
 * is not an intent).
 *
 * Contract (the W7 probe, test/live/xlsx-w7-objects-probe.live.cjs, measured the series form; LIVE-accepted):
 *   set_chart { sheet, name, range, type: 'bar' | 'lineNormal' | 'area' | 'pie' | 'doughnut', title } -- the
 *   data shape W3h measured and nothing wider: series in
 *   COLUMNS, the first row text headers (the series names), the first column text categories, a numeric body, a
 *   sheet name that needs no quoting. Identity = type + title + every series' name / values / categories, derived
 *   from the range. SATISFIED: exactly one chart of that name on the sheet with that identity. Refused: a chart of
 *   that name with another identity (no silent modify), more than one, an unmeasured data shape.
 *   rename_chart { sheet, name, newName }: APPLY one `name` and no `newName`; SATISFIED no `name` and one `newName`;
 *   anything else refused. The renamed chart keeps its identity (L1).
 * L1 the chart; L3 every cell, every OTHER chart of every sheet, every pivot name. Retry SATISFIED 0/0.
 *
 * M7 S5 C1 (owner M7 GO 2026-10-03; audit P C1): `orientation: 'rows'` -- the series in ROWS over the same data shape,
 * AddChart(ref, true, ...). MEASURED in session and after a reopen (W7 probe K2, m6-w7-probe-20260929.json): per data
 * row r, name = the row's first cell, values = the row's body, cat = the header row's body. The orientation is part of
 * the identity (it changes every series), so a chart of the same name in the other orientation is refused as another
 * identity -- never silently modified. 'columns' (the default) is unchanged.
 *
 * X-O03 (implementation phase): `type` beyond 'bar' -- LIVE-measured (the
 * X-O03 series-semantics probe) on the SAME multi-column data shape: bar/lineNormal/area/
 * pie/doughnut/radar/stock all hold exactly the bar contract's series description
 * (name/values/cat per column) and persist across a reopen -- these seven are open here.
 *
 * radar's read-side TYPE is a MEASURED quirk, not a drift: `GetChartType()` (the driver's
 * chartReads source for this field) returns the literal string "unknown" for a radar chart,
 * never "radar" (LIVE, X-O03/2 + X-O03/3); stock
 * reads back correctly as "stock". `READBACK_TYPE`/`readbackType()` below hold this ONE exception
 * fixed -- the identity check compares against the measured read value, never against the
 * requested type blindly. (An EARLIER version of this paragraph claimed stock "does not persist
 * with this data shape across a reopen" -- that was WRONG: X-O03/2's LIVE measurement
 * found stock's series (GetName/GetValues/GetCatFormula) exactly consistent between
 * creation and a fresh reopen on the M1:O4 block, same as radar and every other open type here.)
 *
 * `scatter` stays explicitly CLOSED: its Y-value series data has no PUBLIC getter at all on this
 * build (X-O03/2 -- `GetValues()` THREWs, and an exhaustive enumeration of every
 * other `Get*` method on the series' full prototype chain found no working replacement; only the
 * X values read back, via `GetXValues()`). `surface` stays CLOSED: `AddChart` itself THREWs for it
 * (X-O03/1) -- both are measured failures, not an oversight.
 * The remaining named candidates (stacked/3D/percent variants) are simply NOT YET measured on
 * this data shape; they stay closed until their own LIVE acceptance, per the same rule.
 *
 * X-O04 (implementation phase): `set_chart_legend { sheet, name, legend }` -- LIVE-
 * measured: `SetLegendPos` accepts exactly 'bottom' | 'top' | 'left' | 'right' | 'none', persists
 * across a reopen; an unsupported value (measured: 'topRight', and any invalid string) is
 * REJECTED WITHOUT THROWING -- the setter returns `false` and the position stays unchanged. The
 * write therefore checks the return value explicitly (never assumed from the absence of a throw,
 * the same lesson as W-09's `SetShd`). No creation-time legend field: `set_chart` never measured
 * one, so legend is post-creation-only here, exactly like `set_chart_title` is for the title.
 *
 * X-O05 (LIVE MEASURED; the untouched control chart, same report): `set_chart_axis_title` --
 * SetHorAxisTitle(text) / SetVerAxisTitle(text) (ONE argument, plain text; GetHorAxisTitle/GetVerAxisTitle read it
 * back with a trailing CRLF the same way GetTitle does, trimmed by the SAME `trimTitle`). An untouched chart reads
 * BOTH axis titles as `null` (not a throw), so `trimTitle(null)` (`''`) is the measured "no title" baseline.
 * DATA LABELS ARE NOT A WRITE INTENT (OWNER-GATE / LIMIT / #ref): SetShowDataLabels writes the
 * CHART-level c:dLbls, but every series keeps its own explicit all-0 c:dLbls (written at AddChart), which OOXML lets
 * override the chart level -- the effective (rendered) state is not what GetDataLabels reads, and no public reader
 * exists for the series or point level (ApiSeries has no label method; only ApiChart.SetShowPointDataLabel writes per
 * point). The chart-level labels are still READ (GetDataLabels) so every chart write proves it left them unchanged.
 * Identity: exactly one chart of the name; APPLY when the requested field differs from the fresh read;
 * SATISFIED when it already matches. L1 also proves every OTHER chart field this family tracks (type, title, every
 * series, the other axis title, the data-label flags) stayed put -- the same "seriesKept" pattern set_chart_title
 * already uses, widened to the two new fields.
 *
 * X-O06-6 (position measured by the p115-o06 chart size/position probe v2):
 * `move_chart { sheet, name, at: { col, row } }` -- `chart.SetPosition(col, 0, row, 0)` is the
 * MEASURED call shape (the probe's own ACT script, cited verbatim, not guessed). Cell-aligned
 * ONLY, offset always 0 at the Builder boundary -- the SAME X-O08-3-measured scope `drawings-
 * insert.cjs` already uses for insert_shape/insert_image (design instruction: size is NOT
 * part of this intent; the ±1 EMU drift is the SAME open X-O06 owner-tolerance question, not
 * reopened here).
 *
 * UNLIKE every other post-create classifier here (title/legend/axis-title/rename), `p.chartReads`
 * has NO position field at all (confirmed by reading the driver's own chartReads fragment --
 * {name, type, title, legend, series, horAxisTitle, verAxisTitle, dataLabels}, nothing else) --
 * position comes ONLY from the generic drawing reader (`drawings.observe()`, kind: 'chart'), the
 * SAME source `verifyPreservation` below already reads for every OTHER chart's L3. `classifyMove`/
 * the `move_chart` branch of `verifyPrimitive` therefore take the full observation `p` (the
 * `postCreate` dispatch's 4th argument, added for this one intent; every existing postCreate
 * classifier ignores it, unchanged behaviour) and call `drawings.observe()` directly, rather than
 * `gate()`'s `p.chartReads`.
 *
 * `drawings.cjs`'s `drawingDiffs()` gained one new, OPTIONAL hook (`sizeBlind`, default `() =>
 * false` -- every existing caller unaffected): nulls `width`/`height` on BOTH sides of the diff
 * for a record `sizeBlind()` matches, so the moved chart's own size (which may drift, same as
 * X-O08's measured sizes) is never exact-compared, while every OTHER drawing's full identity
 * (size included) stays exactly compared, unchanged.
 *
 * X-O04/1+2 chart STYLE (1bafc2ca, a LIVE measurement round): `set_chart_style { sheet, name,
 * target, fill, outline, seriesIndex, dataPointIndex }` -- WRITE-ONLY, by design: `ApiChart` has
 * NO public getter for any of these fields (confirmed, X-O04/1's exhaustive Get-/Set-prefixed
 * enumeration of the full prototype chain). Six targets, each with a measured-working fill+outline setter
 * pair (X-O04/1: plotArea/title/legend measured directly with a single-argument Fill/Stroke call;
 * X-O04/2: series/dataPoint measured with their documented multi-argument index form):
 *   plotArea   SetPlotAreaFill(oFill)              / SetPlotAreaOutLine(oStroke)
 *   title      SetTitleFill(oFill)                 / SetTitleOutLine(oStroke)
 *   legend     SetLegendFill(oFill)                / SetLegendOutLine(oStroke)
 *   chart      Fill(oFill)                         / SetOutLine(oStroke)
 *   series     SetSeriesFill(oFill, i, false)       / SetSeriesOutLine(oStroke, i, false)
 *   dataPoint  SetDataPointFill(oFill, i, j, false) / SetDataPointOutLine(oStroke, i, j, false)
 * `chart`'s pair (the bare `Fill`/`SetOutLine` names) is DERIVED, not independently LIVE-named:
 * X-O04/1's own full-prototype enumeration (the SAME reflection `m7-s2b-probe-run4` already
 * recorded for `ApiChart`) is a CLOSED, exhaustive set of every Fill/Outline-shaped name on the
 * object -- six already-claimed pairs (plotArea/title/legend/series/dataPoint/marker) plus exactly
 * one unclaimed Fill-shaped name (`Fill`) and one unclaimed Outline-shaped name (`SetOutLine`);
 * by elimination on a CLOSED set, not a guess among open possibilities. The runtime still never
 * trusts this: `lower()` checks the setter's OWN return value (`=== true`), the SAME fail-closed
 * check `set_chart_legend` already uses for `SetLegendPos` -- a wrong name here fails VISIBLY
 * (FAIL_CLOSED, named 'xlsx-chart-style-setter-rejected'), never silently.
 *
 * X-O05/3 (internal tracker; X-O05/1+/2 LIVE VERIFIED, reviewed and accepted): `gapOverlap`
 * { gapWidth?, overlap? } (BAR charts only -- the LIVE-measured scope, same discipline as
 * `marker`'s MARKER_CHART_TYPES) and `seriesOrder` { order: [...] } are a THIRD shape, different
 * again from the six ORIGINAL targets and from `marker`/`gridlines`: there is no PUBLIC setter and
 * (for gapOverlap) no internal method either -- only the internal DATA field X-O05/2's READ
 * capability already reads. Both go through a dedicated DS Bridge WRITE
 * (chart.bar.gapOverlap.write/v1 / chart.series.order.write/v1, GAP_OVERLAP_WRITE/SERIES_ORDER_WRITE
 * above), served on /12 .. /14 (B22_BRIDGE_WRITE_RUNTIMES, the handler byte-identical on each) -- a
 * runtime off that list refuses BEFORE touching the editor (classifyStyle's own runtime gate, the
 * SAME "never touch the editor on an unsupported runtime" discipline B05-WRITE established). Both ARE verifiable (read-and-
 * compare through the SAME `chartReads` fields X-O05/2 already attaches: `chart.gapOverlap` /
 * `series[].order`), so neither is ever `semanticsNotProven`, same as `marker`/`gridlines`.
 *
 * `seriesOrder` is a WHOLE-CHART write, NOT per-series (X-O05/3 design review, reviewed and
 * accepted): `order` is the FULL permutation, index = the PUBLIC GetAllSeries() index, value = its
 * new c:order -- e.g. `{ target: 'seriesOrder', order: [1, 0, 2] }` swaps series 0 and 1 and keeps
 * series 2. A single-series `{ seriesIndex, order }` shape was rejected: with one chart-style write
 * per task (X-O04/11 r2, REFUSED_COMPOSITION), a one-series step can NEVER swap two series -- the
 * first write always collides with the OTHER series' still-unmoved order. The whole-array shape
 * fixes this by construction: the caller names EVERY series explicitly in one call, so there is no
 * implicit side effect on an unnamed series and no collision is even possible (`order` must be a
 * permutation of 0..n-1, checked before the write, named FAIL_CLOSED otherwise --
 * 'xlsx-chart-style-series-order-not-a-permutation' / '-length-mismatch'). Any COLLIDING
 * intermediate state the internal setOrder calls pass through stays INSIDE the one Bridge write
 * (the final per-series order is what gets verified, through the SAME /12 READ, one `order` key
 * per series). The internal model's own behaviour on a repeated/colliding intermediate value is
 * NOT LIVE-measured -- the LIVE harness measures a 2-series swap and a 3-series cycle.
 *
 * REFUSED, measured non-functional AS OF X-O04/1+2 (round 1, `bAllMarkers`/`bAll*` = `false`):
 * `SetMarkerFill`/`SetMarkerOutLine` returned `false` on BOTH 'bar' and 'lineNormal'; the four
 * gridline setters returned `true` for ANY argument (even a deliberately wrong string) with no
 * distinguishable OOXML effect. *** SUPERSEDED, X-O04/11 (internal tracker; X-O04/4 round 2 +
 * X-O04/8 LIVE): with `bAllMarkers`/`bAll*` = `true`, the written colour WORKS_PERSISTED on both
 * 'scatter' and 'lineNormal', proven end-to-end through the X-O04/10 DS Bridge /11 READ round-trip
 * (`chart.series.markerStyle.read/v1` / `chart.axis.majorGridlinesColor.read/v1`) -- both are now
 * released product WRITE targets, see the VERIFICATION paragraph below. ***
 *
 * VERIFICATION is deliberately NARROWER than every other post-create intent here for the SIX
 * ORIGINAL targets (`plotArea`/`title`/`legend`/`chart`/`series`/`dataPoint`): identity (exactly
 * one chart of the name) is all `verifyPrimitive` can ever confirm -- there is no getter to prove
 * the STYLE value itself landed. `classify()` therefore NEVER returns SATISFIED for these six
 * (every other field is read-and-compared before deciding APPLY vs SATISFIED; these cannot be read
 * at all) -- a retry re-dispatches the SAME idempotent setter call every time, and the RECEIPT must
 * not claim the field verified: `wholeTaskChecks` names every written field in `notProven` (e.g.
 * `chart-style-plotarea-fill`), threaded into `operations[i].semanticsNotProven` by
 * `write-tools.cjs`'s `mapExecutionResult` (16fd90b9's own mechanism, reused unchanged here -- NOT
 * re-derived, NOT a second notProven axis).
 *
 * X-O04/11: `marker` { seriesIndex, fill?, outline? } (scatter/lineNormal series only, matching the
 * LIVE-proven fixtures -- the SAME scope the X-O04/10 READ capability already uses) and `gridlines`
 * { axis: 'vertical' | 'horizontal', outline } (no `fill` -- a gridline has no fill, only its own
 * stroke colour) are DIFFERENT from the six above: they ARE verifiable, through the SAME
 * `chartReads` fields X-O04/10 already attaches to every observation (`series[].markerStyle` /
 * `gridlinesColor`, the SAME `bridgeAxis` path, no new live call here) -- `classifyStyle` and the
 * `set_chart_style` PRIMITIVE read-and-compare the bridge value exactly like every OTHER
 * postCreate classifier, so these two CAN return SATISFIED, and their write is NEVER
 * `semanticsNotProven` (`styleNotProven` excludes them). A pre-/11 runtime (or any chart type
 * `marker` is not scoped to) makes the bridge field `UNOBSERVABLE:<reason>` or absent -- FAIL_CLOSED,
 * named, NEVER a silent write-only fallback to the six-target behaviour above. The raw setter
 * return value is NOT checked for these two (the gate() of the dispatch-time boolean the six
 * targets rely on was exactly the signal X-O04/1+2 round 1 proved unreliable for them -- the bridge
 * read is the only trusted signal, by design, not an oversight).
 */

const INTENTS = Object.freeze(['set_chart', 'rename_chart', 'set_chart_title', 'set_chart_legend', 'set_chart_axis_title', 'move_chart', 'set_chart_style', 'delete_chart', 'copy_chart']);
const CHART_TYPES = Object.freeze(['bar', 'lineNormal', 'area', 'pie', 'doughnut', 'radar', 'stock']);
// X-O04/1+2: the six measured-working style targets, each { fill, outline } setter name pair.
// X-O04/11: `marker` joins this map (same { fill, outline, indexed } shape, indexed like
// series/dataPoint) -- now verifiable, see the header comment. `gridlines` does NOT fit this shape
// (no fill, no index, the setter NAME itself depends on the requested `axis`) and is handled as its
// own case throughout (GRIDLINE_SETTERS below), the SAME "not in STYLE_TARGETS" pattern
// CHART_TYPES already uses for scatter/surface at the enum level.
const STYLE_TARGETS = Object.freeze({
  plotArea: { fill: 'SetPlotAreaFill', outline: 'SetPlotAreaOutLine', indexed: false },
  title: { fill: 'SetTitleFill', outline: 'SetTitleOutLine', indexed: false },
  legend: { fill: 'SetLegendFill', outline: 'SetLegendOutLine', indexed: false },
  chart: { fill: 'Fill', outline: 'SetOutLine', indexed: false },
  series: { fill: 'SetSeriesFill', outline: 'SetSeriesOutLine', indexed: true },
  dataPoint: { fill: 'SetDataPointFill', outline: 'SetDataPointOutLine', indexed: true },
  marker: { fill: 'SetMarkerFill', outline: 'SetMarkerOutLine', indexed: true },
});
// X-O04/11: LIVE-measured setter names, keyed by the PUBLIC `axis` the step requests -- the
// internal cross-binding the X-O04/8 READ capability found (a vertical setter's colour lands on
// the horizontal internal getter) is entirely a quirk of the GETTER side; the WRITE setter names
// are exactly what they say (no cross-binding on this side).
const GRIDLINE_SETTERS = Object.freeze({ vertical: 'SetMajorVerticalGridlines', horizontal: 'SetMajorHorizontalGridlines' });
// X-O04/11: marker is scoped to the two chart types the LIVE fixtures proved (the SAME scope the
// X-O04/10 READ capability uses) -- an untested type (e.g. 'bar') refuses named, never attempted.
const MARKER_CHART_TYPES = Object.freeze(['scatter', 'lineNormal']);
// X-O07/4b: copy_chart v1 NEVER attempts these -- the six WRITE-ONLY style targets (no getter
// exists at all, X-O04/1), dataLabels (not a write intent on this product at all, X-O05), and the
// four BRIDGE-verifiable style targets X-O04/11/X-O05/3 released (deferred to a follow-up card,
// so a bridge-write failure never entangles with copy_chart's own APPLY/FAIL_CLOSED decision).
// X-O05/5-C: + markerSymbol/lineStyle -- copy_chart's scope was fixed BEFORE these two existed,
// and nothing here revisits that scope; they join the six WRITE-ONLY targets above (no getter),
// not the four BRIDGE-verifiable ones.
// Named here, unconditionally, in wholeTaskChecks' `notProven` -- never a silent omission.
// GH #302 G6: 'chart-data-labels' LEFT this list -- the copy now carries the data labels (copyLabelPlan below); what it
// cannot carry is named per copy instead ('chart-series-level-data-labels' / 'chart-point-data-labels-unobservable').
const COPY_CHART_ALWAYS_OMITTED = Object.freeze([
  'chart-style-plotarea', 'chart-style-title', 'chart-style-legend', 'chart-style-chart', 'chart-style-series', 'chart-style-datapoint',
  'chart-style-marker', 'chart-style-gridlines', 'chart-style-gapoverlap', 'chart-style-seriesorder',
  'chart-style-markersymbol', 'chart-style-linestyle',
]);
const isUnobservableStyle = v => typeof v === 'string' && v.startsWith('UNOBSERVABLE:');
const STYLE_FILL_RE = /^#?[0-9A-Fa-f]{6}$/;
const STYLE_EMU_PER_PT = 12700; // same conversion drawings-insert.cjs's size/stroke already use
const styleEmu = pt => Math.round(pt * STYLE_EMU_PER_PT);
// X-O03/3 (LIVE MEASURED): `GetChartType()` reads a radar chart back as
// the literal string "unknown", never "radar" -- every other open type (stock included) reads
// back as its own requested name. Narrow and type-keyed on purpose: this maps the ONE measured
// exception, not a general "unknown means mine" rule (an UNRELATED chart genuinely of an
// unmeasured type would also read "unknown", and must NOT be mistaken for a radar chart).
const READBACK_TYPE = Object.freeze({ radar: 'unknown' });
const readbackType = type => READBACK_TYPE[type] || type;
const LEGEND_POS = Object.freeze(['bottom', 'top', 'left', 'right', 'none']);
const AXES = Object.freeze({ horizontal: 'SetHorAxisTitle', vertical: 'SetVerAxisTitle' });
const LABEL_KEYS = Object.freeze(['showSerName', 'showCatName', 'showVal', 'showPercent']);
const SHEET = /^[A-Za-z_][A-Za-z0-9_]*$/;
const NAME = /^[A-Za-z_][A-Za-z0-9_. -]{0,99}$/;

const oneLine = x => typeof x === 'string' && x !== '' && !/[\r\n]/.test(x);

// the accepted fields per intent (set_chart is the only one that creates and reads cells)
const FIELDS = Object.freeze({
  set_chart: ['intent', 'id', 'sheet', 'name', 'range', 'type', 'title', 'orientation'],
  set_chart_title: ['intent', 'id', 'sheet', 'name', 'title'],
  set_chart_legend: ['intent', 'id', 'sheet', 'name', 'legend'],
  set_chart_axis_title: ['intent', 'id', 'sheet', 'name', 'axis', 'title'],
  rename_chart: ['intent', 'id', 'sheet', 'name', 'newName'],
  move_chart: ['intent', 'id', 'sheet', 'name', 'at'],
  // X-O07/4b: `sheet`/`name` are the DESTINATION (the generic sheet/name checks below already
  // validate them, same as every other intent); `sourceSheet`/`sourceName` name what to copy FROM.
  // `at` is REQUIRED (never guesses a destination position, the SAME discipline move_chart's own
  // required `at` already established).
  copy_chart: ['intent', 'id', 'sheet', 'name', 'sourceSheet', 'sourceName', 'at'],
  // X-O05/3: gapWidth/overlap (target gapOverlap) and order (target seriesOrder) join the fields
  // fill/outline/seriesIndex/dataPointIndex/axis already cover.
  // X-O05/5-C: symbol/size (target markerSymbol) and dash/smooth (target lineStyle) join them.
  set_chart_style: ['intent', 'id', 'sheet', 'name', 'target', 'fill', 'outline', 'seriesIndex', 'dataPointIndex', 'axis', 'gapWidth', 'overlap', 'order', 'symbol', 'size', 'dash', 'smooth'],
  // X-O07/3: delete_chart -- just sheet/name, the same minimal shape rename_chart's "from" side uses.
  delete_chart: ['intent', 'id', 'sheet', 'name'],
});

const atProblem = at => {
  if (!at || typeof at !== 'object') return 'at must be an object { col, row }';
  if (!Number.isInteger(at.col) || at.col < 0) return 'at.col must be a non-negative integer';
  if (!Number.isInteger(at.row) || at.row < 0) return 'at.row must be a non-negative integer';
  return null;
};

const nonNegIdx = v => Number.isInteger(v) && v >= 0;

/** `set_chart_style`'s own checks: a measured target, at least one of fill/outline, and the
 * index fields `series`/`dataPoint` actually require. X-O04/11: `gridlines` has its own shape
 * (no fill, outline required, `axis` instead of an index) -- handled as its own branch, returning
 * before any of the six-target (`STYLE_TARGETS`) checks below run. */
function styleProblem(step) {
  if (step.target === 'gridlines') {
    if (step.fill !== undefined) return 'gridlines has no fill, only outline';
    if (step.outline === undefined) return 'outline is required for target gridlines';
    if (!step.outline || typeof step.outline !== 'object') return 'outline must be an object { width, color }';
    if (!Number.isFinite(step.outline.width) || step.outline.width <= 0) return 'outline.width must be a positive number of points';
    if (typeof step.outline.color !== 'string' || !STYLE_FILL_RE.test(step.outline.color)) return 'outline.color must be a 6-hex colour, e.g. \'#RRGGBB\'';
    if (!Object.hasOwn(GRIDLINE_SETTERS, step.axis)) return `axis must be one of ${Object.keys(GRIDLINE_SETTERS).join(', ')}`;
    if (step.seriesIndex !== undefined || step.dataPointIndex !== undefined) return 'seriesIndex/dataPointIndex are not accepted for target gridlines';
    if (step.symbol !== undefined || step.size !== undefined || step.dash !== undefined || step.smooth !== undefined) return 'symbol/size/dash/smooth are not accepted for target gridlines';
    return null;
  }
  // X-O05/3: gapOverlap/seriesOrder are the SAME "own shape, return before STYLE_TARGETS" pattern
  // gridlines already established -- neither has fill/outline/axis; gapOverlap takes gapWidth/
  // overlap (at least one, each range-checked -- the LIVE-measured bounds).
  if (step.target === 'gapOverlap') {
    if (step.fill !== undefined || step.outline !== undefined) return 'gapOverlap has no fill/outline';
    if (step.axis !== undefined) return 'axis is only accepted for target gridlines';
    if (step.seriesIndex !== undefined || step.dataPointIndex !== undefined) return 'seriesIndex/dataPointIndex are not accepted for target gapOverlap';
    if (step.gapWidth === undefined && step.overlap === undefined) return 'at least one of gapWidth/overlap is required';
    if (step.gapWidth !== undefined && !(Number.isFinite(step.gapWidth) && step.gapWidth >= 0 && step.gapWidth <= 500)) return 'gapWidth must be a number between 0 and 500';
    if (step.overlap !== undefined && !(Number.isFinite(step.overlap) && step.overlap >= -100 && step.overlap <= 100)) return 'overlap must be a number between -100 and 100';
    if (step.symbol !== undefined || step.size !== undefined || step.dash !== undefined || step.smooth !== undefined) return 'symbol/size/dash/smooth are not accepted for target gapOverlap';
    return null;
  }
  // X-O05/3 (design review, reviewed and accepted): `order` is the WHOLE permutation (index =
  // GetAllSeries() index, value = the new c:order), never a single { seriesIndex, order } pair --
  // one chart-style write per task makes a single-series step unable to ever swap two series (the
  // first write collides with the still-unmoved other one). Validated as a shape HERE (an array,
  // every element a distinct integer in 0..length-1 -- a self-contained check, no live chart data
  // needed yet); the array's LENGTH against the chart's ACTUAL series count is live-dependent and
  // checked in classifyStyle instead, the same split plan()/classify() already uses for
  // seriesIndex bounds elsewhere in this file.
  if (step.target === 'seriesOrder') {
    if (step.fill !== undefined || step.outline !== undefined) return 'seriesOrder has no fill/outline';
    if (step.axis !== undefined) return 'axis is only accepted for target gridlines';
    if (step.seriesIndex !== undefined || step.dataPointIndex !== undefined) return 'seriesIndex/dataPointIndex are not accepted for target seriesOrder -- order is the whole permutation';
    if (!Array.isArray(step.order) || step.order.length === 0) return 'order must be a non-empty array (a permutation of 0..n-1)';
    if (!step.order.every(v => Number.isInteger(v) && v >= 0 && v < step.order.length)) return 'order must contain only integers in 0..length-1';
    if (new Set(step.order).size !== step.order.length) return 'order must not repeat a value (it must be a permutation)';
    if (step.symbol !== undefined || step.size !== undefined || step.dash !== undefined || step.smooth !== undefined) return 'symbol/size/dash/smooth are not accepted for target seriesOrder';
    return null;
  }
  // X-O05/5-C (reviewed and accepted): `markerSymbol` { seriesIndex, symbol?, size? } and
  // `lineStyle` { seriesIndex, dash?, smooth? } are the SAME "own shape, return before
  // STYLE_TARGETS" pattern gridlines/gapOverlap/seriesOrder already established -- neither has
  // fill/outline/axis, both ARE indexed (unlike gapOverlap/seriesOrder, which carry no index at
  // all: the whole chart, not one series).
  if (step.target === 'markerSymbol') {
    if (step.fill !== undefined || step.outline !== undefined) return 'markerSymbol has no fill/outline';
    if (step.axis !== undefined) return 'axis is only accepted for target gridlines';
    if (step.dataPointIndex !== undefined) return 'dataPointIndex is not accepted for target markerSymbol';
    if (step.gapWidth !== undefined || step.overlap !== undefined || step.order !== undefined) return 'gapWidth/overlap/order are not accepted for target markerSymbol';
    if (step.dash !== undefined || step.smooth !== undefined) return 'dash/smooth are not accepted for target markerSymbol';
    if (!nonNegIdx(step.seriesIndex)) return 'seriesIndex must be a non-negative integer';
    if (step.symbol === undefined && step.size === undefined) return 'at least one of symbol/size is required';
    if (step.symbol !== undefined && !MARKER_SYMBOL_ENUM.includes(step.symbol)) return `symbol must be one of ${MARKER_SYMBOL_ENUM.join(', ')}`;
    if (step.size !== undefined && !(Number.isInteger(step.size) && step.size >= 2 && step.size <= 72)) return 'size must be an integer between 2 and 72';
    return null;
  }
  if (step.target === 'lineStyle') {
    if (step.fill !== undefined || step.outline !== undefined) return 'lineStyle has no fill/outline';
    if (step.axis !== undefined) return 'axis is only accepted for target gridlines';
    if (step.dataPointIndex !== undefined) return 'dataPointIndex is not accepted for target lineStyle';
    if (step.gapWidth !== undefined || step.overlap !== undefined || step.order !== undefined) return 'gapWidth/overlap/order are not accepted for target lineStyle';
    if (step.symbol !== undefined || step.size !== undefined) return 'symbol/size are not accepted for target lineStyle';
    if (!nonNegIdx(step.seriesIndex)) return 'seriesIndex must be a non-negative integer';
    if (step.dash === undefined && step.smooth === undefined) return 'at least one of dash/smooth is required';
    if (step.dash !== undefined && !LINE_DASH_ENUM.includes(step.dash)) return `dash must be one of ${LINE_DASH_ENUM.join(', ')}`;
    if (step.smooth !== undefined && typeof step.smooth !== 'boolean') return 'smooth must be a boolean';
    return null;
  }
  if (!Object.hasOwn(STYLE_TARGETS, step.target)) return `target must be one of ${Object.keys(STYLE_TARGETS).join(', ')}, gridlines, gapOverlap, seriesOrder, markerSymbol, lineStyle`;
  const t = STYLE_TARGETS[step.target];
  if (step.axis !== undefined) return 'axis is only accepted for target gridlines';
  if (step.gapWidth !== undefined || step.overlap !== undefined || step.order !== undefined) return 'gapWidth/overlap/order are only accepted for target gapOverlap/seriesOrder';
  if (step.symbol !== undefined || step.size !== undefined || step.dash !== undefined || step.smooth !== undefined) return 'symbol/size/dash/smooth are only accepted for target markerSymbol/lineStyle';
  if (step.fill === undefined && step.outline === undefined) return 'at least one of fill/outline is required';
  if (step.fill !== undefined && (typeof step.fill !== 'string' || !STYLE_FILL_RE.test(step.fill))) return 'fill must be a 6-hex colour, e.g. \'#RRGGBB\'';
  if (step.outline !== undefined) {
    if (!step.outline || typeof step.outline !== 'object') return 'outline must be an object { width, color }';
    if (!Number.isFinite(step.outline.width) || step.outline.width <= 0) return 'outline.width must be a positive number of points';
    if (typeof step.outline.color !== 'string' || !STYLE_FILL_RE.test(step.outline.color)) return 'outline.color must be a 6-hex colour, e.g. \'#RRGGBB\'';
  }
  if (t.indexed) { if (!nonNegIdx(step.seriesIndex)) return 'seriesIndex must be a non-negative integer'; }
  else if (step.seriesIndex !== undefined) return 'seriesIndex is only accepted for target series/dataPoint/marker';
  if (step.target === 'dataPoint') { if (!nonNegIdx(step.dataPointIndex)) return 'dataPointIndex must be a non-negative integer'; }
  else if (step.dataPointIndex !== undefined) return 'dataPointIndex is only accepted for target dataPoint';
  return null;
}

/** The intent-specific checks of plan(): { detail } on a problem, else { g } (set_chart's parsed range, or null). */
function intentProblem(step, at) {
  if (['set_chart', 'set_chart_title', 'set_chart_axis_title'].includes(step.intent) && !oneLine(step.title)) return { detail: `${at}: title must be one non-empty line` };
  if (step.intent === 'set_chart_axis_title' && !Object.hasOwn(AXES, step.axis)) return { detail: `${at}: axis must be 'horizontal' or 'vertical'` };
  if (step.intent === 'set_chart_legend' && !LEGEND_POS.includes(step.legend)) return { detail: `${at}: legend must be one of ${LEGEND_POS.join(', ')}` };
  if (step.intent === 'rename_chart' && (typeof step.newName !== 'string' || !NAME.test(step.newName) || step.newName === step.name)) return { detail: `${at}: newName must be a chart name different from name` };
  if (step.intent === 'move_chart') { const e = atProblem(step.at); if (e) return { detail: `${at}: ${e}` }; }
  if (step.intent === 'set_chart_style') { const e = styleProblem(step); if (e) return { detail: `${at}: ${e}` }; }
  if (step.intent === 'copy_chart') {
    const e = atProblem(step.at); if (e) return { detail: `${at}: ${e}` };
    if (typeof step.sourceSheet !== 'string') return { detail: `${at}: sourceSheet must be a sheet name` };
    const why = sheetNameProblem(step.sourceSheet); // cellFree: the same bare-sheet-name class as the destination
    if (why) return { detail: `${at}: sourceSheet ${JSON.stringify(step.sourceSheet)}: ${why}` };
    if (typeof step.sourceName !== 'string' || !NAME.test(step.sourceName)) return { detail: `${at}: sourceName must be a chart name` };
    if (step.sourceSheet === step.sheet && step.sourceName === step.name) return { detail: `${at}: destination (sheet+name) must differ from the source` };
    return { g: null };
  }
  if (step.intent !== 'set_chart') return { g: null };
  const g = parseRange(step.range);
  if (!g || g.r2 - g.r1 < 1 || g.c2 - g.c1 < 1) return { detail: `${at} requires a plain A1 range with a header row, a category column and data` };
  if (!CHART_TYPES.includes(step.type)) return { detail: `${at}: type must be one of ${CHART_TYPES.join(', ')}` };
  if ('orientation' in step && !['columns', 'rows'].includes(step.orientation)) return { detail: `${at}: orientation must be 'columns' (default) or 'rows'` };
  return { g };
}

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const set = step.intent === 'set_chart';
  const retitle = step.intent === 'set_chart_title';
  const relegend = step.intent === 'set_chart_legend';
  const axisTitle = step.intent === 'set_chart_axis_title';
  const cellFree = !set; // every non-creation chart intent touches no cells
  const fields = FIELDS[step.intent] || [];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  // L-21/2: a non-creation chart intent only reaches Api.GetSheet(name) (cellFree: no cell, no reference text) -> the shared released
  // sheet-name class (sheet-name.cjs); set_chart BUILDS the series reference `${sheet}!$A$1...` unquoted, so it keeps the bare class
  if (typeof step.sheet !== 'string') return { ok: false, detail: `${at}: sheet must be a sheet name` };
  if (cellFree) {
    const why = sheetNameProblem(step.sheet);
    if (why) return { ok: false, detail: `${at}: sheet ${JSON.stringify(step.sheet)}: ${why}` };
  } else if (!SHEET.test(step.sheet)) return { ok: false, detail: `${at}: sheet must be a name that needs no quoting in a reference (a quoted sheet reference is not measured yet)` };
  if (typeof step.name !== 'string' || !NAME.test(step.name)) return { ok: false, detail: `${at}: name must be a chart name (letters, digits, _ . - space; starting with a letter or _)` };
  const checked = intentProblem(step, at);
  if (checked.detail) return { ok: false, detail: checked.detail };
  const { g } = checked;
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'charts',
      // a chart claims its NAME(s) on the sheet (two steps on one chart name clash); set_chart also reads its data
      // range (a write into it clashes), a rename touches no cells
      selector: Object.freeze({ sheet: step.sheet, range: set ? rangeText(g) : 'A1', chart: step.name, ...(cellFree ? { cellFree: true } : {}),
        claims: Object.freeze([`chart:${step.sheet}:${step.name}`, ...(step.intent === 'rename_chart' ? [`chart:${step.sheet}:${step.newName}`] : [])]) }),
      name: step.name,
      newName: step.intent === 'rename_chart' ? step.newName : null,
      // X-O07/4b: `title`/`legend`/`type`/`grid`/`inRows` for copy_chart stay null at PLAN time (we
      // do not know the source's values yet) -- classifyCopy MUTATES this SAME operation object
      // with the source's freshly-read values once it has them, verified against how the executor
      // actually threads `op` (queue.push({ op, ... }) in executor.cjs's per-group loop uses the
      // SAME reference classifyOperation received, all the way to lower()): no change to the
      // plan/classify/lower contract, no new field added to the executor itself.
      title: set || retitle || axisTitle ? step.title : null,
      legend: relegend ? step.legend : null,
      // X-O04/11: set_chart_style's `gridlines` target also carries an `axis` ('vertical' |
      // 'horizontal'), the SAME field name set_chart_axis_title already uses for a different
      // purpose -- disambiguated by op.intent at every read site, same convention as `title` above.
      axis: axisTitle ? step.axis : (step.intent === 'set_chart_style' && step.axis !== undefined ? step.axis : null),
      type: set ? step.type : null,
      at: (step.intent === 'move_chart' || step.intent === 'copy_chart') ? Object.freeze({ col: step.at.col, row: step.at.row }) : null,
      inRows: set && step.orientation === 'rows',
      grid: g ? Object.freeze({ ...g }) : null,
      // X-O07/4b: whether the source chart has NO separate category column/row at all -- stays
      // null at plan time (see the comment above), populated by classifyCopy from reconstructGrid.
      noCat: null,
      // X-O07/4b: the copy source -- a SEPARATE sheet+name pair from the destination (selector.sheet/`name` above).
      sourceSheet: step.intent === 'copy_chart' ? step.sourceSheet : null,
      sourceName: step.intent === 'copy_chart' ? step.sourceName : null,
      styleTarget: step.intent === 'set_chart_style' ? step.target : null,
      styleFill: step.intent === 'set_chart_style' && step.fill !== undefined ? step.fill.replace(/^#/, '').toUpperCase() : null,
      styleOutline: step.intent === 'set_chart_style' && step.outline !== undefined ? Object.freeze({ width: step.outline.width, color: step.outline.color.replace(/^#/, '').toUpperCase() }) : null,
      seriesIndex: step.intent === 'set_chart_style' && step.seriesIndex !== undefined ? step.seriesIndex : null,
      dataPointIndex: step.intent === 'set_chart_style' && step.dataPointIndex !== undefined ? step.dataPointIndex : null,
      // X-O05/3: gapOverlap's two fields (either or both) and seriesOrder's single field.
      styleGapWidth: step.intent === 'set_chart_style' && step.gapWidth !== undefined ? step.gapWidth : null,
      styleOverlap: step.intent === 'set_chart_style' && step.overlap !== undefined ? step.overlap : null,
      styleOrder: step.intent === 'set_chart_style' && step.order !== undefined ? Object.freeze(step.order.slice()) : null,
      // X-O05/5-C: markerSymbol's two fields (either or both) and lineStyle's two fields.
      styleSymbol: step.intent === 'set_chart_style' && step.symbol !== undefined ? step.symbol : null,
      styleSize: step.intent === 'set_chart_style' && step.size !== undefined ? step.size : null,
      styleDash: step.intent === 'set_chart_style' && step.dash !== undefined ? step.dash : null,
      styleSmooth: step.intent === 'set_chart_style' && step.smooth !== undefined ? step.smooth : null,
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const bare = x => String(x === null || x === undefined ? '' : x).replace(/^=/, '');
const trimTitle = x => String(x === null || x === undefined ? '' : x).replace(/[\r\n]+$/, '');
const abs = (sheet, c, r) => `${sheet}!$${colName(c)}$${r}`;

/**
 * The series the bar chart over the range has. Columns (W3h): per data column, name = the header, values = the column's
 * body, cat = the first column's body. Rows (W7 K2): per data row, name = the row's first cell, values = the row's body,
 * cat = the header row's body.
 */
function expectedSeries(op) {
  const { sheet } = op.selector;
  const g = op.grid;
  const out = [];
  if (op.inRows) {
    for (let r = g.r1 + 1; r <= g.r2; r += 1) out.push({ name: abs(sheet, g.c1, r), values: `${abs(sheet, g.c1 + 1, r)}:$${colName(g.c2)}$${r}`, cat: `${abs(sheet, g.c1 + 1, g.r1)}:$${colName(g.c2)}$${g.r1}` });
    return out;
  }
  for (let c = g.c1 + 1; c <= g.c2; c += 1) out.push({ name: abs(sheet, c, g.r1), values: `${abs(sheet, c, g.r1 + 1)}:$${colName(c)}$${g.r2}`, cat: `${abs(sheet, g.c1, g.r1 + 1)}:$${colName(g.c1)}$${g.r2}` });
  return out;
}

/** A single cell is read back by the DS engine WITHOUT the ':' range syntax (LIVE measured,
 * X-O07/4b r1: a one-data-row chart's values read as "$F$2", never "$F$2:$F$2") -- a range
 * string and a bare-cell string must not be compared as if interchangeable. */
const rangeOrCell = (sheet, c1, r1, c2, r2) => (c1 === c2 && r1 === r2 ? abs(sheet, c1, r1) : `${abs(sheet, c1, r1)}:$${colName(c2)}$${r2}`);

/** X-O07/4b (copy_chart): the inverse of reconstructGrid()'s shape, mirroring expectedSeries()
 * but (a) honoring op.noCat -- a chart with NO separate category column/row at all, where every
 * cell in the grid is itself a series' own name/values and cat reads '' -- and (b) using
 * rangeOrCell so a single-cell values/cat span is never emitted as a 1-cell range. set_chart's own
 * expectedSeries() is untouched -- this is copy_chart-specific (reconstructGrid has no set_chart
 * counterpart, so there is no shared contract to preserve). LIVE r1 (2026-10-15, zero): a
 * category-less 3-series chart (series.cat === '', values single-cell) failed BOTH ways before
 * this existed -- reconstructGrid rejected it outright, and even with a correct grid,
 * expectedSeries's unconditional colon-range would have mismatched the DS's bare-cell read. */
function expectedCopySeries(op) {
  const { sheet } = op.selector;
  const g = op.grid;
  const out = [];
  if (op.inRows) {
    const r1 = op.noCat ? g.r1 : g.r1 + 1;
    for (let r = r1; r <= g.r2; r += 1) {
      out.push({
        name: abs(sheet, g.c1, r),
        values: rangeOrCell(sheet, g.c1 + 1, r, g.c2, r),
        cat: op.noCat ? '' : rangeOrCell(sheet, g.c1 + 1, g.r1, g.c2, g.r1),
      });
    }
    return out;
  }
  const c1 = op.noCat ? g.c1 : g.c1 + 1;
  for (let c = c1; c <= g.c2; c += 1) {
    out.push({
      name: abs(sheet, c, g.r1),
      values: rangeOrCell(sheet, c, g.r1 + 1, c, g.r2),
      cat: op.noCat ? '' : rangeOrCell(sheet, g.c1, g.r1 + 1, g.c1, g.r2),
    });
  }
  return out;
}

/**
 * X-O07/4b (copy_chart): the EXACT inverse of expectedSeries() -- given a chart's OWN read series
 * (the fingerprint()'d, sheet-prefixed absolute formula strings), reconstruct the bounding grid
 * {c1,r1,c2,r2} and orientation. This only works for a chart whose series follow expectedSeries()'s
 * own contiguous shape (one header row + one fixed category column for 'columns', or one header
 * column + one fixed category row for 'rows') -- ANY other shape (a hand-built chart over
 * discontiguous ranges, say) is NOT reconstructible and refuses named, never guessed.
 * Orientation is detected from TWO signals, in order: (1) with 2+ series, whether every series'
 * `name` cell shares a row (columns) or a column (rows) -- robust even for a single-data-column/
 * single-data-row chart; (2) with exactly one series, the shape of `cat` (if present) or else
 * `values` itself (a multi-row, single-column range means columns; a single-row, multi-column
 * range means rows) -- the ONLY case where the degenerate 1-data-column-vs-1-data-row ambiguity
 * could otherwise arise.
 * X-O07/4b LIVE r1 (2026-10-15, zero): a VALID, category-less chart (every series' own `cat`
 * reads '') is a real shape, not a parse failure -- every series in the chart must agree on
 * having (or lacking) a cat range; a MIX is refused named (xlsx-chart-copy-series-cat-mismatch),
 * never guessed. When none have one, there is no separate category column/row at all: the grid's
 * c1 (columns) or r1 (rows) bound is the FIRST series' own name cell, not a column/row to its
 * side -- the resulting op.noCat flag tells expectedCopySeries() the same thing on the way back.
 */
function reconstructGrid(sheet, series) {
  if (!Array.isArray(series) || series.length === 0) return { ok: false, reason: 'xlsx-chart-copy-no-series' };
  const prefix = `${sheet}!`;
  const stripSheet = ref => (typeof ref === 'string' && ref.startsWith(prefix) ? ref.slice(prefix.length) : null);
  const parse = ref => { const stripped = stripSheet(ref); return stripped === null ? null : parseRange(stripped); };
  const hasCat = s => typeof s.cat === 'string' && s.cat.length > 0;
  const catFlags = series.map(hasCat);
  const allHaveCat = catFlags.every(Boolean);
  const noCat = catFlags.every(x => !x);
  if (!allHaveCat && !noCat) return { ok: false, reason: 'xlsx-chart-copy-series-cat-mismatch' };
  const nameCells = series.map(s => parse(bare(s.name)));
  const valuesRanges = series.map(s => parse(bare(s.values)));
  if (nameCells.some(x => !x) || valuesRanges.some(x => !x)) return { ok: false, reason: 'xlsx-chart-copy-series-ref-not-parseable' };
  const catRanges = allHaveCat ? series.map(s => parse(bare(s.cat))) : null;
  if (allHaveCat && catRanges.some(x => !x)) return { ok: false, reason: 'xlsx-chart-copy-series-ref-not-parseable' };
  let inRows;
  if (series.length >= 2) {
    const sameRow = nameCells[0].r1 === nameCells[1].r1;
    const sameCol = nameCells[0].c1 === nameCells[1].c1;
    if (sameRow && !sameCol) inRows = false;
    else if (sameCol && !sameRow) inRows = true;
    else return { ok: false, reason: 'xlsx-chart-copy-orientation-not-measured' };
  } else {
    const shape = allHaveCat ? catRanges[0] : valuesRanges[0];
    if (shape.rows === 1 && shape.cols > 1) inRows = true;
    else if (shape.cols === 1 && shape.rows > 1) inRows = false;
    else return { ok: false, reason: 'xlsx-chart-copy-orientation-not-measured' };
  }
  const last = series.length - 1;
  if (inRows) {
    if (nameCells[0].c1 !== nameCells[last].c1) return { ok: false, reason: 'xlsx-chart-copy-series-shape-mismatch' };
    const r1 = noCat ? nameCells[0].r1 : catRanges[0].r1;
    return { ok: true, inRows: true, noCat, g: { c1: nameCells[0].c1, r1, c2: valuesRanges[0].c2, r2: nameCells[last].r1 } };
  }
  if (nameCells[0].r1 !== nameCells[last].r1) return { ok: false, reason: 'xlsx-chart-copy-series-shape-mismatch' };
  const c1 = noCat ? nameCells[0].c1 : catRanges[0].c1;
  return { ok: true, inRows: false, noCat, g: { c1, r1: nameCells[0].r1, c2: nameCells[last].c1, r2: valuesRanges[0].r2 } };
}

// X-O05: an untouched chart reads BOTH axis titles null and every data-label flag false (LIVE
// MEASURED, the control chart) -- THAT is the baseline: null -> '' (no title), four booleans as read.
// a maintainer review: anything else (undefined, a missing getter, a non-boolean flag) is UNOBSERVABLE, never defaulted --
// each such read gets its OWN token, so it never compares equal to anything (every identity / kept check fails closed).
let unobservedSeq = 0;
const UNOBSERVED = what => `UNOBSERVABLE:${what}:${unobservedSeq += 1}`;
const isUnobserved = v => typeof v === 'string' && v.startsWith('UNOBSERVABLE:');
const axisTitleOf = t => (t === null ? '' : typeof t === 'string' ? trimTitle(t) : UNOBSERVED('axis-title'));
const labelsOf = l => (l !== null && typeof l === 'object' && LABEL_KEYS.every(k => typeof l[k] === 'boolean')
  ? Object.fromEntries(LABEL_KEYS.map(k => [k, l[k]])) : UNOBSERVED('data-labels'));
const NO_LABELS = Object.freeze(Object.fromEntries(LABEL_KEYS.map(k => [k, false])));
const pointFlagsOf = e => (e !== null && typeof e === 'object' && LABEL_KEYS.every(k => typeof e[k] === 'boolean') ? Object.fromEntries(LABEL_KEYS.map(k => [k, e[k]])) : null);

/**
 * GH #302 G6: the source's data labels a copy can carry -- the G6 LIVE design probe
 * (docs/evidence/gh302-g6-copy-chart-labels-probe-20261026.json): a fresh chart given the source's chart-level flags
 * (SetShowDataLabels -> GetDataLabels) and every POINT-level effective label (SetShowPointDataLabel) reads back EXACTLY
 * the source's dataLabels AND effectiveLabels (session, reopen, the saved <c:dLbl>), in either call order. A point at
 * any other level is carried only all off (a fresh chart's own default); one that SHOWS a flag there has no public
 * setter (no series-level writer exists) and is named per copy, never claimed. effectiveLabels unobservable (the bridge
 * capability missing, over the point cap): no point is attempted, named the same way.
 */
function copyLabelPlan(src) {
  const eff = src.effectiveLabels;
  const unobservable = Object.freeze({ points: null, keys: null, skip: null, notProven: Object.freeze(['chart-point-data-labels-unobservable']) });
  if (eff === null || typeof eff !== 'object') return unobservable;
  const keys = Object.keys(eff).sort();
  const points = [];
  const skip = [];
  for (const k of keys) {
    const f = pointFlagsOf(eff[k]);
    if (!f) return unobservable;
    if (eff[k].level === 'POINT') points.push(Object.freeze({ key: k, series: Number(k.split(':')[0]), point: Number(k.split(':')[1]), flags: Object.freeze(f) }));
    else if (LABEL_KEYS.some(x => f[x])) skip.push(k);
  }
  return Object.freeze({ points: Object.freeze(points), keys: Object.freeze(keys), skip: Object.freeze(skip), notProven: Object.freeze(skip.length ? ['chart-series-level-data-labels'] : []) });
}

/** the copy's effectiveLabels against the plan: every POINT label exact, every other point all off (except `skip`) */
function copiedPointsMatch(dest, lp) {
  if (!lp || !lp.points) return true;
  const eff = dest.effectiveLabels;
  if (eff === null || typeof eff !== 'object' || !same(Object.keys(eff).sort(), lp.keys)) return false;
  return lp.keys.every(k => {
    const f = pointFlagsOf(eff[k]);
    if (!f) return false;
    const want = lp.points.find(x => x.key === k);
    if (want) return eff[k].level === 'POINT' && same(f, want.flags);
    return lp.skip.includes(k) || LABEL_KEYS.every(x => !f[x]);
  });
}
const fingerprint = c => ({
  type: c.type, title: trimTitle(c.title), series: (c.series || []).map(x => ({ name: bare(x.name), values: bare(x.values), cat: bare(x.cat) })),
  horAxisTitle: axisTitleOf(c.horAxisTitle), verAxisTitle: axisTitleOf(c.verAxisTitle), labels: labelsOf(c.dataLabels),
});
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
// a freshly-created chart's measured baseline: no axis titles, every data-label flag false (same as above)
const wanted = op => ({ type: readbackType(op.type), title: op.title, series: expectedSeries(op), horAxisTitle: '', verAxisTitle: '', labels: NO_LABELS });

/** The data shape the contract measured, or why not. */
function shapeProblem(p, op) {
  const { sheet } = op.selector;
  const g = op.grid;
  for (let r = g.r1; r <= g.r2; r += 1) {
    for (let c = g.c1; c <= g.c2; c += 1) {
      const cell = cellAt(p, sheet, addr(c, r));
      if (!cell) return { reason: 'xlsx-chart-cell-unread', cell: addr(c, r) };
      if (r === g.r1 && c === g.c1) continue;
      if (cell.f !== '' || cell.count === null) return { reason: 'xlsx-chart-data-shape-not-measured', cell: addr(c, r), why: 'formula-or-unreadable' };
      const text = cell.count === 0 && cell.v !== null && cell.v !== '';
      if ((r === g.r1 || c === g.c1) ? !text : cell.count !== 1) return { reason: 'xlsx-chart-data-shape-not-measured', cell: addr(c, r), why: r === g.r1 || c === g.c1 ? 'header-or-category-not-text' : 'value-not-numeric' };
    }
  }
  return null;
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const reads = p.chartReads;
  const bad = !reads ? ['not-read'] : p.sheets.map(x => x.name).filter(n => !reads[n] || typeof reads[n] !== 'object' || !Array.isArray(reads[n].charts) || !Array.isArray(reads[n].pivots));
  if (bad.length) return { ok: false, reason: 'xlsx-chart-unobservable', sheets: bad.slice(0, 20), why: reads ? bad.map(n => reads[n]).slice(0, 5) : 'not-read' };
  return { ok: true, charts: reads[op.selector.sheet].charts };
}

const named = (charts, n) => charts.filter(c => c.name === n);

/** M7 S2b: the L3 claims every sheet's drawings -- an unreadable one refuses the write (only APPLY is gated). */
function drawingBlocked(p) {
  const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  return o.complete ? null : { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-charts-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length } };
}

function classify(p, op) {
  const c = classifyChart(p, op);
  return c.state === 'APPLY' ? (drawingBlocked(p) || c) : c;
}

function classifyRename(g, op, where) {
  const from = named(g.charts, op.name).length;
  const to = named(g.charts, op.newName).length;
  if (from === 1 && to === 0) return { state: 'APPLY', evidence: { ...where, newName: op.newName } };
  if (from === 0 && to === 1) return { state: 'SATISFIED', evidence: { ...where, newName: op.newName, proof: 'fresh-live-chart-under-the-new-name' } };
  return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-rename-state-conflict', ...where, newName: op.newName, withName: from, withNewName: to } };
}

/**
 * X-O06-6: `p.chartReads` has NO position field -- the identity (exactly one chart of the name)
 * is still confirmed through it (`g.charts`, consistent with every other post-create classifier),
 * but the POSITION itself comes from `drawings.observe()` (the generic drawing reader, kind
 * 'chart'), the SAME source `verifyPreservation` already reads for every chart's L3.
 */
function classifyMove(g, op, where, p) {
  const hits = named(g.charts, op.name);
  if (hits.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: hits.length ? 'xlsx-chart-name-ambiguous' : 'xlsx-chart-not-found', ...where, count: hits.length } };
  const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  if (!o.complete) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-charts-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length } };
  const recs = (o.bySheet[op.selector.sheet] || []).filter(r => r.kind === 'chart' && r.name === op.name);
  if (recs.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: recs.length ? 'xlsx-chart-drawing-ambiguous' : 'xlsx-chart-drawing-not-found', ...where, count: recs.length } };
  const r = recs[0];
  const atTarget = r.fromCol === op.at.col && r.fromRow === op.at.row && r.colOffset === 0 && r.rowOffset === 0;
  return atTarget
    ? { state: 'SATISFIED', evidence: { ...where, at: op.at, proof: 'fresh-live-drawing-position' } }
    : { state: 'APPLY', evidence: { ...where, at: op.at, now: { fromCol: r.fromCol, fromRow: r.fromRow } } };
}

/**
 * X-O07/3: delete_chart's own classify -- idempotent on retry (SATISFIED if the name is already
 * gone: "a chart named X should not exist" is already true, the SAME end-state-matches-goal
 * reasoning classifyRename's own "to" side already uses), FAIL_CLOSED on an ambiguous name
 * (never guess which one to delete -- X-O07/2 LIVE MEASURED: the handler itself refuses the same
 * way), and the B23 Bridge WRITE runtime gate BEFORE touching the editor on an unsupported
 * runtime (the SAME discipline classifyStyle's own B22 gate established).
 */
function classifyDelete(g, op, where, p) {
  const hits = named(g.charts, op.name);
  if (hits.length === 0) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-chart-absent' } };
  if (hits.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-name-ambiguous', ...where, count: hits.length } };
  const runtime = p && p.runtime ? p.runtime.identity : null;
  if (!B23_BRIDGE_WRITE_RUNTIMES.includes(runtime)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-delete-bridge-write-runtime-not-proven', ...where, runtime, proven: B23_BRIDGE_WRITE_RUNTIMES } };
  }
  return { state: 'APPLY', evidence: { ...where } };
}

/**
 * X-O07/4b: diagram copy as RECREATION -- the internal clone path (CChartSpace copy()/
 * getCopyWithSourceFormatting() + addToDrawingObjects) is a NAMED, measured boundary (X-O07/4
 * r1-r4: four isolated LIVE rounds, the last throwing "Cannot read properties of undefined
 * (reading 'length')" even after the clone's own drawing-context was populated from the source --
 * it does not close). copy_chart instead reads the source's type/series/title/legend/axis-titles
 * (ALREADY in `g.charts`/the source-sheet's own `p.chartReads`, no new live capability) and
 * recreates it via the SAME AddChart+SetName+SetTitle sequence set_chart's own lower() already
 * uses, plus SetLegendPos/SetHorAxisTitle/SetVerAxisTitle for the fields that have one -- entirely
 * INTERNAL calls, independent of whether set_chart/set_chart_legend are themselves released on the
 * product surface (set_chart's OWN lower() already calls SetTitle the same way, unreleased or not).
 *
 * v1 SCOPE, NAMED (never silently claimed wider than this):
 *   COPIED: type, title, every series (the SOURCE's own absolute cell references -- the copy keeps
 *     pointing at the SOURCE's data, the same way a human "duplicate chart" does NOT duplicate the
 *     underlying cells), legend, both axis titles, position (`at`, REQUIRED -- never a guessed
 *     default placement), and (GH #302 G6) the data labels: the chart-level flags and every POINT-level effective
 *     label (copyLabelPlan; a series-level label that shows a flag has no setter and is named, not claimed).
 *   NEVER COPIED, by construction (callers must re-apply these separately, same as creating a
 *     fresh chart): the six WRITE-ONLY style targets (plotArea/title/legend/chart/series/dataPoint
 *     fill+outline -- NO getter exists at all, confirmed exhaustively, X-O04/1); marker/gridlines/
 *     gapOverlap/seriesOrder (the four BRIDGE-verifiable style targets X-O04/11/X-O05/3 released --
 *     deferred to a follow-up card, not attempted here, so a bridge-write failure never entangles
 *     with this intent's own APPLY/FAIL_CLOSED decision). `wholeTaskChecks` below NAMES every one
 *     of these in `notProven` so the receipt never implies more than v1 actually does.
 *   UNMEASURED: whether THIS DS engine accepts AddChart's `ref` argument naming a DIFFERENT sheet
 *     than the one `AddChart` is called on (sourceSheet !== destSheet) -- the LIVE harness is what
 *     decides this, same "ship the attempt, report honestly" discipline X-O03's scatter/stock
 *     closed-vs-open types already used. The SAME-SHEET path (sourceSheet === destSheet) is the
 *     SAME mechanism set_chart's own lower() already uses, with no new risk.
 *
 * `reconstructGrid()` runs HERE (classify-time, against the FRESHEST read) -- a refusal here means
 * the editor is never touched (the same "never dispatch on an unreconstructable source" discipline
 * B05-WRITE's runtime gate already established for a DIFFERENT kind of pre-dispatch refusal).
 * classifyCopy then MUTATES `op` (grid/inRows/type/title/legend/copyHorAxisTitle/
 * copyVerAxisTitle) with these freshly-derived values -- `lower()` reads them directly, no second
 * reconstruction inside the command body. This relies on `op` being the SAME object reference
 * classify/resolve/lower all receive (verified by reading executor.cjs: `queue.push({ op, ... })`
 * carries the identical reference classifyOperation was called with, all the way to
 * `adapter.executePrimitive(handle, q.op, ...)` -> `familyFor(op).lower(op, target)`) -- no change
 * to the plan/classify/lower contract, no new field threaded through the executor itself.
 */
function classifyCopy(g, op, where, p) {
  const srcG = gate(p, { selector: { sheet: op.sourceSheet } });
  if (!srcG.ok) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-copy-source-sheet-unobservable', source: { sheet: op.sourceSheet, chart: op.sourceName }, detail: srcG } };
  const srcHits = named(srcG.charts, op.sourceName);
  if (srcHits.length !== 1) {
    return { state: 'FAIL_CLOSED', evidence: { reason: srcHits.length ? 'xlsx-chart-copy-source-ambiguous' : 'xlsx-chart-copy-source-not-found', source: { sheet: op.sourceSheet, chart: op.sourceName }, count: srcHits.length } };
  }
  const src = srcHits[0];
  const rg = reconstructGrid(op.sourceSheet, src.series);
  if (!rg.ok) return { state: 'FAIL_CLOSED', evidence: { reason: rg.reason, source: { sheet: op.sourceSheet, chart: op.sourceName } } };
  // mutate op with the source's freshly-read values -- see the header comment above.
  op.grid = Object.freeze({ ...rg.g });
  op.inRows = rg.inRows;
  op.noCat = rg.noCat;
  op.type = src.type;
  op.title = trimTitle(src.title);
  op.legend = src.legend;
  op.copyHorAxisTitle = axisTitleOf(src.horAxisTitle);
  op.copyVerAxisTitle = axisTitleOf(src.verAxisTitle);
  // GH #302 G6: the data labels the copy carries (copyLabelPlan); an unreadable chart level refuses -- it would be lost
  op.copyLabels = labelsOf(src.dataLabels);
  if (isUnobserved(op.copyLabels)) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-copy-source-labels-unobservable', source: { sheet: op.sourceSheet, chart: op.sourceName } } };
  op.copyLabelPlan = copyLabelPlan(src);

  const destHits = named(g.charts, op.name);
  if (destHits.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-name-ambiguous', ...where, count: destHits.length } };
  const srcFp = fingerprint(src);
  if (destHits.length === 1) {
    const destFp = fingerprint(destHits[0]);
    // identity AND legend must match the SOURCE exactly (no silent partial-modify) -- the same
    // "no silent modify" discipline classifyChart's own base case already uses for set_chart.
    if (!same(destFp, srcFp) || destHits[0].legend !== src.legend || !copiedPointsMatch(destHits[0], op.copyLabelPlan)) {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-copy-dest-exists-with-another-identity', ...where, now: { ...destFp, legend: destHits[0].legend }, source: { ...srcFp, legend: src.legend } } };
    }
    const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
    if (!o.complete) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-charts-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length } };
    const recs = (o.bySheet[op.selector.sheet] || []).filter(r => r.kind === 'chart' && r.name === op.name);
    if (recs.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: recs.length ? 'xlsx-chart-drawing-ambiguous' : 'xlsx-chart-drawing-not-found', ...where, count: recs.length } };
    const atOk = recs[0].fromCol === op.at.col && recs[0].fromRow === op.at.row && recs[0].colOffset === 0 && recs[0].rowOffset === 0;
    return atOk
      ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-chart-copy-matches-source', source: { sheet: op.sourceSheet, chart: op.sourceName } } }
      : { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-copy-dest-exists-with-another-identity', ...where, now: { at: { fromCol: recs[0].fromCol, fromRow: recs[0].fromRow } }, wantedAt: op.at } };
  }
  return { state: 'APPLY', evidence: { ...where, source: { sheet: op.sourceSheet, chart: op.sourceName }, at: op.at, grid: op.grid, inRows: op.inRows } };
}

function classifyChart(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, chart: op.name };
  const postCreate = { set_chart_title: classifyTitle, set_chart_legend: classifyLegend, set_chart_axis_title: classifyAxisTitle, rename_chart: classifyRename, move_chart: classifyMove, set_chart_style: classifyStyle, delete_chart: classifyDelete, copy_chart: classifyCopy }[op.intent];
  if (postCreate) return postCreate(g, op, where, p);
  const shape = shapeProblem(p, op);
  if (shape) return { state: 'FAIL_CLOSED', evidence: shape };
  const hits = named(g.charts, op.name);
  if (hits.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-name-ambiguous', ...where, count: hits.length } };
  if (hits.length === 1) {
    return same(fingerprint(hits[0]), wanted(op))
      ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-chart-with-the-identity' } }
      : { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-exists-with-another-identity', ...where, now: fingerprint(hits[0]) } };
  }
  return { state: 'APPLY', evidence: { ...where, range: op.selector.range } };
}

/**
 * M7 C3 set_chart_title (owner M7 GO 2026-10-03; tier D probe: SetTitle on an existing chart, the series unchanged, the
 * title read back with a trailing CR LF -- session = reopen): exactly one chart of the name; SATISFIED when its title
 * already reads the requested one.
 */
function classifyTitle(g, op, where) {
  const hits = named(g.charts, op.name);
  if (hits.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: hits.length ? 'xlsx-chart-name-ambiguous' : 'xlsx-chart-not-found', ...where, count: hits.length } };
  if (trimTitle(hits[0].title) === op.title) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-chart-title' } };
  return { state: 'APPLY', evidence: { ...where, title: op.title, now: trimTitle(hits[0].title) } };
}

/**
 * X-O04 set_chart_legend (LIVE-measured): exactly one chart of the name; SATISFIED when its legend
 * position already reads the requested one.
 */
function classifyLegend(g, op, where) {
  const hits = named(g.charts, op.name);
  if (hits.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: hits.length ? 'xlsx-chart-name-ambiguous' : 'xlsx-chart-not-found', ...where, count: hits.length } };
  if (hits[0].legend === op.legend) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-chart-legend' } };
  return { state: 'APPLY', evidence: { ...where, legend: op.legend, now: hits[0].legend } };
}

/** X-O05 (LIVE MEASURED): same shape as classifyTitle, keyed by the requested axis. */
function classifyAxisTitle(g, op, where) {
  const hits = named(g.charts, op.name);
  if (hits.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: hits.length ? 'xlsx-chart-name-ambiguous' : 'xlsx-chart-not-found', ...where, count: hits.length } };
  const now = axisTitleOf(op.axis === 'horizontal' ? hits[0].horAxisTitle : hits[0].verAxisTitle);
  if (isUnobserved(now)) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-axis-title-unobservable', ...where, axis: op.axis } };
  if (now === op.title) return { state: 'SATISFIED', evidence: { ...where, axis: op.axis, proof: 'fresh-live-chart-axis-title' } };
  return { state: 'APPLY', evidence: { ...where, axis: op.axis, title: op.title, now } };
}

/**
 * X-O04/11: the bridge-read current value for a marker/gridlines style op (already present on
 * every chart read since X-O04/10 -- no new live call here), or why it cannot be compared.
 */
function styleBridgeNow(op, chart) {
  if (op.styleTarget === 'gridlines') {
    const colour = chart.gridlinesColor;
    if (isUnobservableStyle(colour)) return { unobservable: true, reason: colour };
    const key = `${op.axis}Gridlines`;
    return { unobservable: false, value: colour && typeof colour === 'object' ? colour[key] : undefined };
  }
  // X-O05/3: gapOverlap is chart-level (no seriesIndex), unlike marker/seriesOrder below.
  if (op.styleTarget === 'gapOverlap') {
    const go = chart.gapOverlap;
    if (isUnobservableStyle(go)) return { unobservable: true, reason: go };
    return { unobservable: false, value: go && typeof go === 'object' ? go : undefined };
  }
  // X-O05/3 (design review): seriesOrder is WHOLE-CHART -- every series' own order, in
  // GetAllSeries() index order, not just op.seriesIndex (there is no such field any more: the
  // whole permutation IS the op). The FIRST unobservable series' own reason is reported (named,
  // not folded into a generic one) -- classifyStyle decides what an unobservable entry means.
  if (op.styleTarget === 'seriesOrder') {
    const series = Array.isArray(chart.series) ? chart.series : [];
    // The chart.series.order.read/v1 Bridge value is the object { order: n } (that is what read_charts
    // carries per series); unwrap it to the number here. LIVE on 1.0.1-9e03483b93ef: comparing the object
    // itself never matched, so an applied order read as VERIFY_CANNOT_PROVE_TERMINAL_STATE.
    const orders = series.map(s => s && s.order).map(o => (o && typeof o === 'object' && typeof o.order === 'number' ? o.order : o));
    const badIndex = orders.findIndex(isUnobservableStyle);
    if (badIndex !== -1) return { unobservable: true, reason: orders[badIndex], index: badIndex };
    return { unobservable: false, value: orders.every(o => typeof o === 'number') ? { order: orders } : undefined };
  }
  const s = Array.isArray(chart.series) ? chart.series[op.seriesIndex] : undefined;
  const m = s && s.markerStyle;
  if (isUnobservableStyle(m)) return { unobservable: true, reason: m };
  return { unobservable: false, value: m && typeof m === 'object' ? m : undefined };
}

/** Does the bridge-read current value already match the requested fill/outline (whichever the step asked for)? */
function styleBridgeMatches(op, now) {
  if (op.styleTarget === 'gridlines') return now.value === op.styleOutline.color;
  if (op.styleTarget === 'gapOverlap') {
    const v = now.value || {};
    if (op.styleGapWidth !== null && v.gapWidth !== op.styleGapWidth) return false;
    if (op.styleOverlap !== null && v.overlap !== op.styleOverlap) return false;
    return true;
  }
  if (op.styleTarget === 'seriesOrder') {
    const now0 = (now.value || {}).order;
    return Array.isArray(now0) && now0.length === op.styleOrder.length && now0.every((o, i) => o === op.styleOrder[i]);
  }
  const v = now.value || {};
  if (op.styleFill !== null && v.fill !== op.styleFill) return false;
  if (op.styleOutline !== null && v.outline !== op.styleOutline.color) return false;
  return true;
}

/** X-O04/1+2 (1bafc2ca): the six ORIGINAL targets are WRITE-ONLY, by design -- no getter exists
 * for any of those style fields, so this NEVER returns SATISFIED for them (unlike every other
 * postCreate classifier above, which reads-and-compares first). `seriesIndex` IS bounds-checked
 * here, against `GetAllSeries()`'s own count -- an EXISTING, already-read field
 * (`fingerprint().series`), not a new capability; `dataPointIndex` has no such count anywhere in
 * this family's reads, so it is NOT bounds-checked here -- the setter's own return value (checked
 * at dispatch, `lower()`) is the only guard for it.
 *
 * X-O04/11: `marker` and `gridlines` are DIFFERENT -- verifiable through the SAME `chartReads`
 * fields X-O04/10 already attaches (`series[].markerStyle` / `gridlinesColor`), so these two DO
 * read-and-compare, exactly like every other postCreate classifier; `marker` is additionally
 * scoped to the two chart types the LIVE fixtures proved (MARKER_CHART_TYPES), refused named on
 * anything else -- never attempted on an unmeasured type. */
const CHART_STYLE_BRIDGE_WRITE_TARGETS = Object.freeze(['gapOverlap', 'seriesOrder']);
const BAR_ONLY_STYLE_TARGETS = Object.freeze(['gapOverlap']);

function classifyStyle(g, op, where, p) {
  const hits = named(g.charts, op.name);
  if (hits.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: hits.length ? 'xlsx-chart-name-ambiguous' : 'xlsx-chart-not-found', ...where, count: hits.length } };
  // X-O05/3: BEFORE touching the editor on an unsupported runtime -- the SAME discipline B05-WRITE
  // established (format.cjs's wantsAlignExtWrite gate), never relying on the write attempt's own
  // natural failure.
  if (CHART_STYLE_BRIDGE_WRITE_TARGETS.includes(op.styleTarget)) {
    const runtime = p && p.runtime ? p.runtime.identity : null;
    if (!B22_BRIDGE_WRITE_RUNTIMES.includes(runtime)) {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-bridge-write-runtime-not-proven', ...where, target: op.styleTarget, runtime, proven: B22_BRIDGE_WRITE_RUNTIMES } };
    }
  }
  // X-O05/5-C: the SAME "never touch the editor on an unsupported runtime" discipline as
  // gapOverlap/seriesOrder above -- but on B24 (/15-ONLY), never B22.
  if (CHART_SERIES_STYLE_WRITE_ONLY_TARGETS.includes(op.styleTarget)) {
    const runtime = p && p.runtime ? p.runtime.identity : null;
    if (!B24_BRIDGE_WRITE_RUNTIMES.includes(runtime)) {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-bridge-write-runtime-not-proven', ...where, target: op.styleTarget, runtime, proven: B24_BRIDGE_WRITE_RUNTIMES } };
    }
  }
  if (op.styleTarget === 'marker' && !MARKER_CHART_TYPES.includes(hits[0].type)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-marker-type-unsupported', ...where, type: hits[0].type, supported: MARKER_CHART_TYPES } };
  }
  // X-O05/5-C: markerSymbol reuses the SAME measured scope as `marker` (fill/outline) above --
  // round8/round8d's own symbol-write-and-readback measurement ran on the SAME `series.marker`
  // sub-object, scatter + lineNormal only.
  if (op.styleTarget === 'markerSymbol' && !MARKER_CHART_TYPES.includes(hits[0].type)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-marker-symbol-type-unsupported', ...where, type: hits[0].type, supported: MARKER_CHART_TYPES } };
  }
  // X-O05/5-C: lineStyle's OWN measured scope -- R8q's LIVE probe proved dash/smooth on a LINE
  // series (with and without an existing line) and a BAR series' own line; never wider.
  if (op.styleTarget === 'lineStyle' && !LINE_STYLE_CHART_TYPES.includes(hits[0].type)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-line-style-type-unsupported', ...where, type: hits[0].type, supported: LINE_STYLE_CHART_TYPES } };
  }
  // X-O05/5-C round 2 (internal tracker's LIVE finding): `smooth` is its OWN, NARROWER measured scope
  // than `dash` -- refused named, BEFORE dispatch, never left to the dispatch attempt's own
  // uncertain outcome (see LINE_STYLE_SMOOTH_CHART_TYPES's own comment above).
  if (op.styleTarget === 'lineStyle' && op.styleSmooth !== null && !LINE_STYLE_SMOOTH_CHART_TYPES.includes(hits[0].type)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-line-style-smooth-type-unsupported', ...where, type: hits[0].type, supported: LINE_STYLE_SMOOTH_CHART_TYPES } };
  }
  if (BAR_ONLY_STYLE_TARGETS.includes(op.styleTarget) && hits[0].type !== 'bar') {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-gap-overlap-type-unsupported', ...where, type: hits[0].type, supported: ['bar'] } };
  }
  if (op.styleTarget !== 'gridlines' && op.styleTarget !== 'gapOverlap' && op.styleTarget !== 'seriesOrder' && !CHART_SERIES_STYLE_WRITE_ONLY_TARGETS.includes(op.styleTarget) && STYLE_TARGETS[op.styleTarget].indexed) {
    const seriesCount = Array.isArray(hits[0].series) ? hits[0].series.length : 0;
    if (op.seriesIndex >= seriesCount) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-series-index-out-of-range', ...where, seriesIndex: op.seriesIndex, seriesCount } };
  }
  // X-O05/5-C: markerSymbol/lineStyle ARE indexed (one series each), just not through
  // STYLE_TARGETS (they are WRITE-ONLY, never added there) -- the SAME bounds check as above,
  // its own branch since `STYLE_TARGETS[op.styleTarget]` does not exist for either.
  if (CHART_SERIES_STYLE_WRITE_ONLY_TARGETS.includes(op.styleTarget)) {
    const seriesCount = Array.isArray(hits[0].series) ? hits[0].series.length : 0;
    if (op.seriesIndex >= seriesCount) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-series-index-out-of-range', ...where, seriesIndex: op.seriesIndex, seriesCount } };
  }
  // X-O05/3 (design review, reviewed and accepted): `order` names EVERY series (the whole
  // permutation), so the array's LENGTH must equal the chart's ACTUAL series count -- a mismatch
  // refuses named, BEFORE the write (never silently ignoring an extra entry or leaving a series
  // unnamed). This replaces a per-series collision check: a permutation cannot collide by
  // construction (styleProblem already proved it is 0..n-1 with no repeats), so the only thing
  // left to confirm against LIVE data is that `n` matches.
  if (op.styleTarget === 'seriesOrder') {
    const seriesCount = Array.isArray(hits[0].series) ? hits[0].series.length : 0;
    if (op.styleOrder.length !== seriesCount) {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-series-order-length-mismatch', ...where, requested: op.styleOrder.length, seriesCount } };
    }
  }
  if (op.styleTarget === 'marker' || op.styleTarget === 'gridlines' || CHART_STYLE_BRIDGE_WRITE_TARGETS.includes(op.styleTarget)) {
    const now = styleBridgeNow(op, hits[0]);
    if (now.unobservable) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-style-bridge-unobservable', ...where, target: op.styleTarget, detail: now.reason } };
    if (styleBridgeMatches(op, now)) return { state: 'SATISFIED', evidence: { ...where, target: op.styleTarget, proof: 'fresh-live-chart-style-bridge-read' } };
    return { state: 'APPLY', evidence: { ...where, target: op.styleTarget, now: now.value } };
  }
  return { state: 'APPLY', evidence: { ...where, target: op.styleTarget } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, chart: selector.chart }) };
}

/** GH #302 G6: the copy's chart-level flags (only when one is on -- a fresh chart reads all off), then every POINT label */
function copyLabelCalls(op) {
  const l = op.copyLabels;
  const chartCall = l && LABEL_KEYS.some(k => l[k]) ? `
  c.SetShowDataLabels(${LABEL_KEYS.map(k => l[k]).join(', ')});` : '';
  const points = (op.copyLabelPlan && op.copyLabelPlan.points) || [];
  return chartCall + points.map(x => `
  if (c.SetShowPointDataLabel(${x.series}, ${x.point}, ${LABEL_KEYS.map(k => x.flags[k]).join(', ')}) === false) return { ok: false, stage: 'copy-point-label-rejected', point: ${JSON.stringify(x.key)} };`).join('');
}

function lower(op) {
  const { sheet } = op.selector;
  const nameOf = 'c => { try { return c.GetName(); } catch (e) { return null; } }';
  let call;
  if (op.intent === 'set_chart') {
    const g = op.grid;
    const ref = `${sheet}!$${colName(g.c1)}$${g.r1}:$${colName(g.c2)}$${g.r2}`;
    // re-checked at dispatch: a chart of the name that appeared since the read is never doubled
    call = `if ((ws.GetAllCharts() || []).some(c => (${nameOf})(c) === ${JSON.stringify(op.name)})) return { ok: false, stage: 'name-taken' };
  const c = ws.AddChart(${JSON.stringify(ref)}, ${op.inRows ? 'true' : 'false'}, ${JSON.stringify(op.type)}, 2, 3600000, 2160000, ${g.c2}, 0, ${g.r1 - 1}, 0);
  if (!c) return { ok: false, stage: 'add-returned-nothing' };
  c.SetName(${JSON.stringify(op.name)});
  c.SetTitle(${JSON.stringify(op.title)}, 12, false);`;
  } else if (op.intent === 'set_chart_title') {
    call = `const hit = (ws.GetAllCharts() || []).filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity' };
  hit[0].SetTitle(${JSON.stringify(op.title)}, 12, false);`;
  } else if (op.intent === 'set_chart_legend') {
    // measured: an unsupported position is REJECTED WITHOUT THROWING (the setter returns false,
    // the position stays unchanged) -- the return value is checked explicitly, never assumed.
    call = `const hit = (ws.GetAllCharts() || []).filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity' };
  if (hit[0].SetLegendPos(${JSON.stringify(op.legend)}) === false) return { ok: false, stage: 'legend-rejected' };`;
  } else if (op.intent === 'set_chart_axis_title') {
    call = `const hit = (ws.GetAllCharts() || []).filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity' };
  hit[0].${AXES[op.axis]}(${JSON.stringify(op.title)});`;
  } else if (op.intent === 'move_chart') {
    // measured call shape: the p115-o06 chart size/position probe v2 ACT script --
    // chart.SetPosition(col, 0, row, 0), cell-aligned, offset always 0.
    call = `const hit = (ws.GetAllCharts() || []).filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity' };
  hit[0].SetPosition(${op.at.col}, 0, ${op.at.row}, 0);`;
  } else if (op.intent === 'set_chart_style' && op.styleTarget === 'gridlines') {
    // X-O04/11: no fill, no index -- the setter NAME itself is the requested axis. The return
    // value is NOT checked (the trusted signal is the X-O04/10 bridge read, classifyStyle/
    // PRIMITIVE below -- the SAME "no second signal in the return value" finding the header
    // comment already records for this target, now resolved by reading instead of trusting it).
    const widthEmu = styleEmu(op.styleOutline.width);
    const r = parseInt(op.styleOutline.color.slice(0, 2), 16), gC = parseInt(op.styleOutline.color.slice(2, 4), 16), b = parseInt(op.styleOutline.color.slice(4, 6), 16);
    call = `const hit = (ws.GetAllCharts() || []).filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity' };
  const styleStroke = Api.CreateStroke(${widthEmu}, Api.CreateSolidFill(Api.CreateRGBColor(${r}, ${gC}, ${b})));
  hit[0].${GRIDLINE_SETTERS[op.axis]}(styleStroke);`;
  } else if (op.intent === 'set_chart_style' && (op.styleTarget === 'gapOverlap' || op.styleTarget === 'seriesOrder')) {
    // X-O05/3: neither the six-target fill/outline setters NOR gridlines' PUBLIC setter pattern
    // applies here -- there is no public/internal setter for gapOverlap at all, and seriesOrder's
    // internal setOrder() is not public either. Both go through the DS Bridge WRITE instead; the
    // chart/sheet identity lookup that EVERY other branch above does in-body is NOT repeated here
    // -- the bridge handler resolves sheet+chart SERVER-SIDE from the args the gate already
    // carries, and returns its own named 'sheet-not-found'/'chart-not-found' on the same race
    // every other branch's `hit.length !== 1` guards against.
    const gapOverlap = op.styleTarget === 'gapOverlap';
    const args = {}; const requested = [];
    if (gapOverlap) {
      if (op.styleGapWidth !== null) { args.gapWidth = op.styleGapWidth; requested.push('gapWidth'); }
      if (op.styleOverlap !== null) { args.overlap = op.styleOverlap; requested.push('overlap'); }
    } else {
      // X-O05/3 (design review): the WHOLE permutation, one write, never a per-series arg.
      args.order = op.styleOrder; requested.push('order');
    }
    call = `const g = __eoChartStyleGate(name, ${JSON.stringify(op.name)});
  if (!g.ok) return g;
  const w = __eoChartStyleWrite(${JSON.stringify(args)}, ${JSON.stringify(requested)});
  if (!w.ok) return __eoChartStyleFail(w);`;
  } else if (op.intent === 'set_chart_style' && CHART_SERIES_STYLE_WRITE_ONLY_TARGETS.includes(op.styleTarget)) {
    // X-O05/5-C: the SAME "no in-body chart/sheet lookup, the handler resolves it server-side"
    // shape as gapOverlap/seriesOrder above -- but through the B24 (/15-ONLY) gate/write client
    // (__eoSeriesStyleGate/__eoSeriesStyleWrite/__eoSeriesStyleFail), and carrying the series
    // index EVERY call (unlike gapOverlap/seriesOrder, which address the whole chart).
    const markerSymbol = op.styleTarget === 'markerSymbol';
    const args = { series: op.seriesIndex }; const requested = [];
    if (markerSymbol) {
      if (op.styleSymbol !== null) { args.symbol = op.styleSymbol; requested.push('symbol'); }
      if (op.styleSize !== null) { args.size = op.styleSize; requested.push('size'); }
    } else {
      // the handler applies smooth BEFORE dash when both are requested (build-rev15.py's
      // LINE_STYLE_WRITE_IMPL: `if (hasSmooth) {...} if (hasDash) {...}`) -- `requested` must
      // match that order byte-for-byte, since __eoSeriesStyleWrite compares it against the
      // receipt's own `applied` array verbatim.
      if (op.styleSmooth !== null) { args.smooth = op.styleSmooth; requested.push('smooth'); }
      if (op.styleDash !== null) { args.dash = op.styleDash; requested.push('dash'); }
    }
    call = `const g = __eoSeriesStyleGate(name, ${JSON.stringify(op.name)});
  if (!g.ok) return g;
  const w = __eoSeriesStyleWrite(${JSON.stringify(args)}, ${JSON.stringify(requested)});
  if (!w.ok) return __eoSeriesStyleFail(w);`;
  } else if (op.intent === 'set_chart_style') {
    // X-O04/1+2: WRITE-ONLY, by design -- the setter's OWN return value is the ONLY evidence of
    // success (no getter exists to re-confirm it later), checked here, fail-closed, the SAME
    // pattern set_chart_legend's SetLegendPos check already uses.
    // X-O04/11: `marker` is the ONE exception in this branch -- its extraArgs differ (nMarker=0,
    // bAllMarkers=true, the LIVE-measured shape that persists) AND its return value is NOT checked
    // (X-O04/1 round 1 measured it unreliable; the trusted signal is the bridge read instead, same
    // reasoning as gridlines above).
    const t = STYLE_TARGETS[op.styleTarget];
    const marker = op.styleTarget === 'marker';
    const extraArgs = marker ? `, ${op.seriesIndex}, 0, true` : !t.indexed ? '' : op.styleTarget === 'dataPoint' ? `, ${op.seriesIndex}, ${op.dataPointIndex}, false` : `, ${op.seriesIndex}, false`;
    let fillCall = '';
    if (op.styleFill) {
      const r = parseInt(op.styleFill.slice(0, 2), 16), gC = parseInt(op.styleFill.slice(2, 4), 16), b = parseInt(op.styleFill.slice(4, 6), 16);
      fillCall = marker ? `
  const styleFill = Api.CreateSolidFill(Api.CreateRGBColor(${r}, ${gC}, ${b}));
  hit[0].${t.fill}(styleFill${extraArgs});` : `
  const styleFill = Api.CreateSolidFill(Api.CreateRGBColor(${r}, ${gC}, ${b}));
  if (hit[0].${t.fill}(styleFill${extraArgs}) === false) return { ok: false, stage: 'style-fill-rejected' };`;
    }
    let outlineCall = '';
    if (op.styleOutline) {
      const widthEmu = styleEmu(op.styleOutline.width);
      const r = parseInt(op.styleOutline.color.slice(0, 2), 16), gC = parseInt(op.styleOutline.color.slice(2, 4), 16), b = parseInt(op.styleOutline.color.slice(4, 6), 16);
      outlineCall = marker ? `
  const styleStroke = Api.CreateStroke(${widthEmu}, Api.CreateSolidFill(Api.CreateRGBColor(${r}, ${gC}, ${b})));
  hit[0].${t.outline}(styleStroke${extraArgs});` : `
  const styleStroke = Api.CreateStroke(${widthEmu}, Api.CreateSolidFill(Api.CreateRGBColor(${r}, ${gC}, ${b})));
  if (hit[0].${t.outline}(styleStroke${extraArgs}) === false) return { ok: false, stage: 'style-outline-rejected' };`;
    }
    call = `const hit = (ws.GetAllCharts() || []).filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity' };${fillCall}${outlineCall}`;
  } else if (op.intent === 'delete_chart') {
    // X-O07/3: the bridge resolves sheet+chart SERVER-SIDE from the args the gate already
    // carries (same convention gapOverlap/seriesOrder's branch above uses) -- no chart/sheet
    // identity lookup repeated in-body; the handler's own name-lookup is STRICTER (refuses
    // ambiguous rather than taking the first match), already proven X-O07/1 LIVE.
    call = `const g = __eoChartDeleteGate(name, ${JSON.stringify(op.name)});
  if (!g.ok) return g;
  const w = __eoChartDeleteWrite();
  if (!w.ok) return __eoChartDeleteFail(w);`;
  } else if (op.intent === 'copy_chart') {
    // X-O07/4b: `op.grid`/`op.inRows`/`op.type`/`op.title`/`op.legend`/`op.copyHorAxisTitle`/
    // `op.copyVerAxisTitle` were MUTATED onto this SAME operation by classifyCopy, from the
    // source's freshly-read values (see classifyCopy's own header comment) -- `ref` names the
    // SOURCE sheet (same-sheet when sourceSheet === sheet, the SAME mechanism set_chart's own
    // lower() already uses; cross-sheet is UNMEASURED on this DS engine, attempted here, decided
    // by the LIVE harness).
    const g = op.grid;
    const ref = `${op.sourceSheet}!$${colName(g.c1)}$${g.r1}:$${colName(g.c2)}$${g.r2}`;
    const legendCall = op.legend !== null && op.legend !== 'none' ? `
  if (c.SetLegendPos(${JSON.stringify(op.legend)}) === false) return { ok: false, stage: 'copy-legend-rejected' };` : '';
    const horCall = op.copyHorAxisTitle ? `
  c.SetHorAxisTitle(${JSON.stringify(op.copyHorAxisTitle)});` : '';
    const verCall = op.copyVerAxisTitle ? `
  c.SetVerAxisTitle(${JSON.stringify(op.copyVerAxisTitle)});` : '';
    call = `if ((ws.GetAllCharts() || []).some(c => (${nameOf})(c) === ${JSON.stringify(op.name)})) return { ok: false, stage: 'name-taken' };
  const c = ws.AddChart(${JSON.stringify(ref)}, ${op.inRows ? 'true' : 'false'}, ${JSON.stringify(op.type)}, 2, 3600000, 2160000, ${op.at.col}, 0, ${op.at.row}, 0);
  if (!c) return { ok: false, stage: 'add-returned-nothing' };
  c.SetName(${JSON.stringify(op.name)});
  c.SetTitle(${JSON.stringify(op.title)}, 12, false);${legendCall}${horCall}${verCall}${copyLabelCalls(op)}`;
  } else {
    call = `const all = ws.GetAllCharts() || [];
  const hit = all.filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1 || all.some(c => (${nameOf})(c) === ${JSON.stringify(op.newName)})) return { ok: false, stage: 'rename-identity' };
  hit[0].SetName(${JSON.stringify(op.newName)});`;
  }
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ${call}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    // X-O07/4b: `series` is left null for copy_chart -- expectedSeries(op) reads op.selector.sheet
    // (the DESTINATION) as the formula prefix, which would be WRONG here (the actual series
    // reference op.sourceSheet, built directly into `ref` above); this meta field is observability
    // only (never read by verifyCopy/classifyCopy), so it is omitted rather than shown misleading.
    meta: Object.freeze({ family: 'charts', action: op.intent, sheet, name: op.name, newName: op.newName, title: op.title, legend: op.legend, axis: op.axis, type: op.intent === 'set_chart' || op.intent === 'copy_chart' ? op.type : null, series: op.intent === 'set_chart' ? expectedSeries(op) : null, at: (op.intent === 'move_chart' || op.intent === 'copy_chart') ? op.at : null, sourceSheet: op.sourceSheet, sourceName: op.sourceName, copyHorAxisTitle: op.copyHorAxisTitle ?? null, copyVerAxisTitle: op.copyVerAxisTitle ?? null, copyLabels: op.copyLabels ?? null, copyPoints: (op.copyLabelPlan && op.copyLabelPlan.points) || null, target: op.styleTarget, styleFill: op.styleFill, styleOutline: op.styleOutline, seriesIndex: op.seriesIndex, dataPointIndex: op.dataPointIndex, styleGapWidth: op.styleGapWidth, styleOverlap: op.styleOverlap, styleOrder: op.styleOrder,
      // X-O05/5-C: observability only, the SAME role styleFill/styleOutline/styleGapWidth/... have
      // above (never read by verifyPrimitive/classifyStyle, which both read op.* directly).
      styleSymbol: op.styleSymbol, styleSize: op.styleSize, styleDash: op.styleDash, styleSmooth: op.styleSmooth,
      // X-O05/3: the capability ID, not a bare boolean (unlike bridgeCopyWrite/bridgeAlignExtWrite/
      // bridgePrintWrite) -- two distinct chart-style Bridge WRITEs share this one driver gate.
      bridgeChartStyleWrite: op.styleTarget === 'gapOverlap' ? GAP_OVERLAP_WRITE_ID : op.styleTarget === 'seriesOrder' ? SERIES_ORDER_WRITE_ID : null,
      // X-O07/3: a bare boolean, like bridgeCopyWrite/bridgeAlignExtWrite/bridgePrintWrite -- only
      // ONE Bridge WRITE sits behind delete_chart, unlike chart-style's two-capability lookup above.
      bridgeChartDeleteWrite: op.intent === 'delete_chart',
      // X-O05/5-C: the SAME capability-ID-in-meta shape as bridgeChartStyleWrite above, on its
      // OWN meta key (never bridgeChartStyleWrite itself) so the driver's two gates -- B22 for
      // gapOverlap/seriesOrder, B24 for markerSymbol/lineStyle -- never collide on one lookup.
      bridgeSeriesStyleWrite: op.styleTarget === 'markerSymbol' ? MARKER_SYMBOL_WRITE_ID : op.styleTarget === 'lineStyle' ? LINE_STYLE_WRITE_ID : null }),
  };
}

const chartsBefore = (op, before) => {
  const was = before && before.primary && before.primary.chartReads && before.primary.chartReads[op.selector.sheet];
  return was && Array.isArray(was.charts) ? was.charts : null;
};

const PRIMITIVE = {
  set_chart(op, charts, was) {
    const hits = named(charts, op.name);
    const grew = !was || charts.length === was.length + (named(was, op.name).length ? 0 : 1);
    return hits.length === 1 && same(fingerprint(hits[0]), wanted(op)) && grew;
  },
  // the title changed to the requested one; type and every series exactly as before. The whole-task check has no
  // before (null): there the title alone -- the series were proven by L1
  set_chart_title(op, charts, was, before) {
    const hits = named(charts, op.name);
    const old = was ? named(was, op.name) : [];
    const seriesKept = before === null || (old.length === 1 && same({ ...fingerprint(hits[0] || {}), title: null }, { ...fingerprint(old[0]), title: null }));
    return hits.length === 1 && trimTitle(hits[0].title) === op.title && seriesKept;
  },
  // X-O05: the requested axis title changed; type, title, every series, the data labels AND the OTHER axis title
  // exactly as before (nulling only the axis key being set -- same pattern as set_chart_title above)
  set_chart_axis_title(op, charts, was, before) {
    const hits = named(charts, op.name);
    if (hits.length !== 1) return false;
    const key = op.axis === 'horizontal' ? 'horAxisTitle' : 'verAxisTitle';
    const now = axisTitleOf(hits[0][key]);
    const old = was ? named(was, op.name) : [];
    const nullKey = f => ({ ...f, [key]: null });
    const kept = before === null || (old.length === 1 && same(nullKey(fingerprint(hits[0])), nullKey(fingerprint(old[0]))));
    return now === op.title && kept;
  },
  rename_chart(op, charts, was) {
    const to = named(charts, op.newName);
    const old = was ? named(was, op.name) : [];
    const kept = old.length !== 1 || (to.length === 1 && same(fingerprint(to[0]), fingerprint(old[0])));
    return named(charts, op.name).length === 0 && to.length === 1 && kept;
  },
  // X-O07/3: the named chart is gone; if we know the prior count (a live L1 check -- `was` is
  // null in a whole-task check, same escape hatch every other PRIMITIVE above uses), the total
  // dropped by exactly one (never more, never a no-op silently counted as success).
  delete_chart(op, charts, was) {
    const gone = named(charts, op.name).length === 0;
    const shrank = !was || named(was, op.name).length !== 1 || charts.length === was.length - 1;
    return gone && shrank;
  },
  // the legend changed to the requested position; type, title and every series exactly as before.
  set_chart_legend(op, charts, was, before) {
    const hits = named(charts, op.name);
    const old = was ? named(was, op.name) : [];
    const kept = before === null || old.length !== 1 || same(fingerprint(hits[0] || {}), fingerprint(old[0]));
    return hits.length === 1 && hits[0].legend === op.legend && kept;
  },
  // X-O04/1+2 (1bafc2ca): the six ORIGINAL targets are WRITE-ONLY -- identity is the ONLY thing this
  // can confirm (the style field itself has no getter). Everything ELSE this family already reads
  // (type/title/series/axis titles/labels via fingerprint(), PLUS legend -- the one field
  // fingerprint() excludes) must stay exactly as before; the setter's own return value was already
  // checked at dispatch.
  // X-O04/11: `marker`/`gridlines` ADDITIONALLY require the freshly bridge-read value to match the
  // request -- fingerprint() deliberately excludes markerStyle/gridlinesColor (so writing one never
  // looks like a preservation violation of itself), and this extra check is layered on top of the
  // SAME identity+preservation gate every target already passes.
  set_chart_style(op, charts, was, before) {
    const hits = named(charts, op.name);
    const old = was ? named(was, op.name) : [];
    const kept = before === null || old.length !== 1
      || (same(fingerprint(hits[0] || {}), fingerprint(old[0])) && (hits[0] || {}).legend === old[0].legend);
    if (hits.length !== 1 || !kept) return false;
    // X-O05/3: gapOverlap/seriesOrder are the SAME "verifiable, bridge-read-and-compare" class as
    // marker/gridlines.
    if (op.styleTarget !== 'marker' && op.styleTarget !== 'gridlines' && !CHART_STYLE_BRIDGE_WRITE_TARGETS.includes(op.styleTarget)) return true;
    const now = styleBridgeNow(op, hits[0]);
    return !now.unobservable && styleBridgeMatches(op, now);
  },
};

/** X-O06-6: move_chart's evidence is POSITION, which `p.chartReads` does not carry at all --
 * read via `drawings.observe()` instead, same source `classifyMove` already uses. */
function verifyMove(op, after) {
  const p = after && after.primary;
  const o = drawings.observe(p && p.drawingReads, p ? p.sheets.map(x => x.name) : []);
  if (!p || !o.complete) return { ok: false, evidence: { reason: 'xlsx-charts-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length } };
  const recs = (o.bySheet[op.selector.sheet] || []).filter(r => r.kind === 'chart' && r.name === op.name);
  const ok = recs.length === 1 && recs[0].fromCol === op.at.col && recs[0].fromRow === op.at.row && recs[0].colOffset === 0 && recs[0].rowOffset === 0;
  return { ok, evidence: { sheet: op.selector.sheet, chart: op.name, at: op.at, now: recs[0] || null, proof: 'fresh-live-drawing-position' } };
}

/** X-O07/4b: identity (type/title/series/axis-titles/labels + legend, read fresh from BOTH sides)
 * AND position, re-read from the FRESHEST observation each time -- correct for both the initial
 * APPLY and a later retry (which must see identity+position ALREADY matching and report SATISFIED
 * via classifyCopy, never re-dispatching). No PRIMITIVE.copy_chart entry exists -- this bypasses
 * it entirely, the SAME pattern move_chart's own dedicated verifyMove already uses. */
function verifyCopy(op, after) {
  const p = after && after.primary;
  if (!p) return { ok: false, evidence: { reason: 'xlsx-chart-copy-observation-missing' } };
  const destG = gate(p, op);
  if (!destG.ok) return { ok: false, evidence: destG };
  const destHits = named(destG.charts, op.name);
  if (destHits.length !== 1) return { ok: false, evidence: { reason: 'xlsx-chart-copy-dest-identity-missing', count: destHits.length } };
  // L1 compares against what `op` CAPTURED at classify time (the source's values AT THE MOMENT
  // OF DISPATCH, mutated onto op by classifyCopy) -- NOT a fresh re-read of the (possibly
  // since-changed) live source. Whether the source ITSELF stayed unchanged is verifyPreservation's
  // (L3) job via the generic "every OTHER chart kept" check; conflating the two here would make a
  // sabotaged SOURCE read as a DEST-identity failure instead of the L3 violation it actually is.
  const wantedCopy = { type: readbackType(op.type), title: op.title, series: expectedCopySeries({ ...op, selector: { ...op.selector, sheet: op.sourceSheet } }), horAxisTitle: op.copyHorAxisTitle, verAxisTitle: op.copyVerAxisTitle, labels: op.copyLabels };
  const pointsOk = copiedPointsMatch(destHits[0], op.copyLabelPlan);
  const identityOk = same(fingerprint(destHits[0]), wantedCopy) && destHits[0].legend === op.legend && pointsOk;
  const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  if (!o.complete) return { ok: false, evidence: { reason: 'xlsx-charts-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length } };
  const recs = (o.bySheet[op.selector.sheet] || []).filter(r => r.kind === 'chart' && r.name === op.name);
  const posOk = recs.length === 1 && recs[0].fromCol === op.at.col && recs[0].fromRow === op.at.row && recs[0].colOffset === 0 && recs[0].rowOffset === 0;
  return {
    ok: identityOk && posOk,
    evidence: { sheet: op.selector.sheet, chart: op.name, source: { sheet: op.sourceSheet, chart: op.sourceName }, at: op.at, now: recs[0] || null, destIdentity: { ...fingerprint(destHits[0]), legend: destHits[0].legend }, wanted: { ...wantedCopy, legend: op.legend }, pointLabelsOk: pointsOk, proof: 'fresh-live-chart-copy-vs-captured-source' },
  };
}

function verifyPrimitive(op, before, after) {
  if (op.intent === 'move_chart') return verifyMove(op, after);
  if (op.intent === 'copy_chart') return verifyCopy(op, after);
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  const ok = PRIMITIVE[op.intent](op, g.charts, chartsBefore(op, before), before);
  return { ok, evidence: { sheet: op.selector.sheet, chart: op.name, newName: op.newName, now: g.charts.map(c => c.name), proof: 'fresh-live-chart-getters' } };
}

/** L3: every cell, every OTHER chart of every sheet (name + identity, in order), every pivot name. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary.chartReads;
  const a = after.primary.chartReads;
  const mine = (sheet, c) => sheet === op.selector.sheet && (c.name === op.name || (op.newName !== null && c.name === op.newName));
  // `legend` is outside fingerprint() (which also feeds the set_chart identity check, where legend
  // has no creation-time counterpart) -- added here directly so an OTHER chart's legend is still
  // covered by this preservation check.
  const others = reads => (reads ? Object.entries(reads).map(([s, r]) => [s, r && Array.isArray(r.charts) ? r.charts.filter(c => !mine(s, c)).map(c => ({ name: c.name, legend: c.legend, ...fingerprint(c) })) : r, r && r.pivots]) : null);
  const chartsKept = !!b && !!a && same(others(b), others(a));
  // M7 S2b: every sheet's drawings equal; the requested chart change exempts ONLY its requested field -- set_chart adds
  // that one chart, rename_chart changes that chart's name (its anchor, size, rotation, flips must stay).
  // X-O06-6: move_chart expects the SAME chart's position to change to `op.at`, and (sizeBlind) its
  // own size is NEVER exact-compared (the measured ±1 EMU drift) -- every OTHER drawing, including
  // its own size, stays exactly compared.
  const sh = op.selector.sheet;
  const mineRec = r => r.sheet === sh && r.kind === 'chart' && r.name === op.name;
  const draw = drawings.drawingDiffs(drawings.observe(before.primary.drawingReads, before.primary.sheets.map(x => x.name)), drawings.observe(after.primary.drawingReads, after.primary.sheets.map(x => x.name)), {
    expect: r => {
      if (op.intent === 'rename_chart' && r.sheet === sh && r.kind === 'chart' && r.name === op.name) return Object.freeze({ ...r, name: op.newName });
      if (op.intent === 'move_chart' && mineRec(r)) return Object.freeze({ ...r, fromCol: op.at.col, fromRow: op.at.row, colOffset: 0, rowOffset: 0 });
      return r;
    },
    // X-O07/4b: copy_chart creates a NEW drawing record on the dest sheet, same as set_chart.
    added: r => (op.intent === 'set_chart' || op.intent === 'copy_chart') && r.sheet === sh && r.kind === 'chart' && r.name === op.name,
    // X-O07/3: the deleted chart's own drawing record is EXPECTED to disappear -- the mirror
    // image of `added` above (set_chart's new chart is expected to APPEAR).
    removed: r => op.intent === 'delete_chart' && mineRec(r),
    sizeBlind: r => op.intent === 'move_chart' && mineRec(r),
  });
  return { ok: rest.ok && chartsKept && draw.length === 0, evidence: { ...rest.evidence, otherChartsAndPivotsKept: chartsKept, drawingsKept: draw.length === 0, drawingDiffs: draw } };
}

// X-O04/1+2 (1bafc2ca): the written style field(s), named the SAME way 16fd90b9's
// COLOUR_NOT_PROVEN already does for drawings-insert.cjs -- 'chart-style-<target-lowercase>-fill'
// / '-outline', capability-specific, never folded into `ok`.
function styleNotProven(op) {
  if (op.intent !== 'set_chart_style') return [];
  // X-O04/11: marker/gridlines are verified via the bridge read (classifyStyle/PRIMITIVE), never
  // write-only -- unlike the six original targets below, their write is never semanticsNotProven.
  // X-O05/3: gapOverlap/seriesOrder join them, same reasoning.
  if (op.styleTarget === 'marker' || op.styleTarget === 'gridlines' || CHART_STYLE_BRIDGE_WRITE_TARGETS.includes(op.styleTarget)) return [];
  const t = op.styleTarget.toLowerCase();
  // X-O05/5-C: markerSymbol/lineStyle are WRITE-ONLY, like the six original targets below --
  // neither field is readable back through any capability this product has (the SAME reasoning
  // as the six-original case, just with different field names: symbol/size, dash/smooth).
  if (op.styleTarget === 'markerSymbol' || op.styleTarget === 'lineStyle') {
    const out = [];
    if (op.styleSymbol !== null) out.push(`chart-style-${t}-symbol`);
    if (op.styleSize !== null) out.push(`chart-style-${t}-size`);
    if (op.styleDash !== null) out.push(`chart-style-${t}-dash`);
    if (op.styleSmooth !== null) out.push(`chart-style-${t}-smooth`);
    return out;
  }
  const out = [];
  if (op.styleFill !== null) out.push(`chart-style-${t}-fill`);
  if (op.styleOutline !== null) out.push(`chart-style-${t}-outline`);
  return out;
}

function wholeTaskChecks(op, finalObservation) {
  // X-O07/4b: this `op` comes from adapter.verifyWholeTask's OWN fresh `planOperations(task)` call
  // (src/adapters/xlsx/adapter.cjs) -- a DIFFERENT object instance than the one classify/dispatch
  // mutated during the main run, so op.grid/type/title/... are back to null here. classify() is
  // PURE given a fresh observation (same inputs -> same mutation), so re-running it against
  // `finalObservation` repopulates this fresh op's fields before verifyCopy reads them -- no
  // change to verifyPrimitive/verifyCopy themselves, no new field threaded through the executor.
  if (op.intent === 'copy_chart' && finalObservation && finalObservation.primary) classify(finalObservation.primary, op);
  const v = verifyPrimitive(op, null, finalObservation);
  const what = op.intent === 'set_chart' ? `${op.name} over ${op.selector.range}`
    : op.intent === 'set_chart_title' ? `${op.name} titled ${JSON.stringify(op.title)}`
    : op.intent === 'set_chart_legend' ? `${op.name} legend set to ${op.legend}`
    : op.intent === 'set_chart_axis_title' ? `${op.name} ${op.axis} axis titled ${JSON.stringify(op.title)}`
    : op.intent === 'move_chart' ? `${op.name} moved to col ${op.at.col}/row ${op.at.row}`
    : op.intent === 'set_chart_style' ? `${op.name} style (${op.styleTarget}) set`
    : op.intent === 'delete_chart' ? `${op.name} deleted`
    : op.intent === 'copy_chart' ? `${op.name} copied from ${op.sourceSheet}!${op.sourceName}`
    : `${op.name} renamed to ${op.newName}`;
  // X-O07/4b: copy_chart's v1 scope NEVER attempts these fields (see classifyCopy's header
  // comment) -- named here, unconditionally, so the receipt never implies more than v1 actually
  // does (the card's own "amit a read NEM lat/nem masol at, az NEVESITETT elhagyas" requirement).
  const notProven = op.intent === 'copy_chart' ? [...COPY_CHART_ALWAYS_OMITTED, ...((op.copyLabelPlan && op.copyLabelPlan.notProven) || [])] : styleNotProven(op);
  return [{ name: `${op.selector.sheet}: chart ${what}`, ok: v.ok, detail: v.ok ? null : v.evidence, operationId: op.id, ...(notProven.length ? { notProven } : {}) }];
}

const windowsOf = op => (op.intent === 'set_chart' ? [{ sheet: op.selector.sheet, range: op.selector.range }] : []);
const liveExtras = () => ({ chartSheets: ['*'], drawingSheets: ['*'] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, B22_BRIDGE_WRITE_RUNTIMES, B23_BRIDGE_WRITE_RUNTIMES, B24_BRIDGE_WRITE_RUNTIMES };
