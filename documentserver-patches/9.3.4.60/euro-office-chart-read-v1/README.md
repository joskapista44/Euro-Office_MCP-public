# euro-office-chart-read-v1 (9.3.4.60, revision 1)

Own implementation of the nine chart / drawing READ getters the product uses (M12, owner decision D2a = C). It replaces,
for the public release, the historical `chart-public-getters` patch, whose text contains upstream implementation
fragments and therefore stays private evidence.

## PUBLIC-FIRST (measured 2026-10-12)

The unpatched Euro-Office 9.3.4.60 cell bundle (`2e5059f4…`) defines **none** of the nine getters (0 definitions each),
and the spreadsheet `ApiChart` has no `ToJSON` (LIVE: `c.ToJSON is not a function`). There is no public read path for
legend position, axis titles, chart-level data labels, series references or the drawing anchor, so a DS-side layer is
needed. Evidence: `docs/evidence/m12-chart/public-first-probe-20261012.json`.

## Contract (what the product relies on)

| getter | returns |
|---|---|
| `ApiChart.GetLegendPos()` | `"left"` / `"top"` / `"right"` / `"bottom"`; `"none"` when the chart has no legend; `null` for any other legend position code (overlay variants, `layout`, `topRight`) |
| `ApiChart.GetHorAxisTitle()` / `GetVerAxisTitle()` | the axis title's text exactly as the title's document content reports it (a set title ends in `"\r\n"`); `null` when the axis or its title is absent |
| `ApiChart.GetDataLabels()` | `{ showSerName, showCatName, showVal, showPercent }`, booleans, from the FIRST chart group's chart-level labels; all `false` when there are none |
| `ApiChartSeries.GetName()` / `GetCatFormula()` / `GetXValues()` | the reference formula with a leading `=` (it follows a sheet rename); `""` when the series has no such reference |
| `ApiChartSeries.GetValues()` | the values reference formula; THROWS (`TypeError`) when the series has no values reference -- e.g. a scatter series (yVal) -- so the product's read of that sheet stays UNOBSERVABLE, as before |
| `ApiDrawing.GetPosition()` | `{ fromCol, colOffset, fromRow, rowOffset }`: the top-left anchor, 0-based, offsets in EMU, NOT rounded (`123455.99999999999` is a valid value) |

Identity: `Api.GetEuroOfficeChartReadLayer()` -> `{ layer: "euro-office-chart-read-v1", runtime: "9.3.4.60", revision: 1 }`;
every installed getter carries `fn.euroOfficeChartRead` (the layer record), which lets a harness prove that the getters
it calls are this layer's.

## Semantic parity (LIVE, 2026-10-12)

`test/live/xlsx-m12-chart-read-ab-parity.live.cjs`: in ONE session on the served production runtime, this file's install
function is installed under a prefix next to the historical getters, and both read the same states: the five chart
types the product opens (bar, lineNormal, area, pie, doughnut), every legend position code, axis titles set / emptied /
removed, all 16 data-label combinations, a sheet rename, series without a name / category / values reference, a shape's
anchor. **409 comparisons, 0 differences.** The only difference is the TEXT of the thrown error for a series without
values (diagnostic: the product records `UNOBSERVABLE:<message>`; the class is the same). Evidence:
`docs/evidence/m12-chart/ab-parity-20261012.json`.

## Provenance (licence gate)

The block is written from this contract, the measured behaviour and the runtime's model fields; it is not derived from
the historical patch. `test/acceptance/m12-chart-read-provenance.test.cjs` (private-only, it reads the historical patch)
asserts that no string of 30+ characters is shared with any historical patch text and that every shared string of 16+
characters is an identifier path (measured longest: 24, `ApiChartSeries.prototype` / `this.Drawing.drawingBase`). Against
the unpatched upstream bundle: 0 shared 40-character windows (`docs/evidence/m12-chart/provenance-20261012.txt`).

## Build

Inserted right before the anchor `AscBuilder.Cell.init=function(` of the unpatched bundle -- the end of the cell builder
scope -- by `install_bridge.py --target cell` (plan: `../fresh-install.json`; the DS Bridge block keeps its own anchor).
Deterministic; two builds are byte-identical.

Why the end of the scope (measured on the 9.3.4.60 bundle): the cell scope REPLACES `ApiChart.prototype`
(`Object.create(ApiDrawing.prototype)`) and assigns `GetSeries` only after the Bridge anchor, so getters installed there
would sit on a discarded prototype. The series class is defined in a different builder scope and is not in scope here;
its prototype is taken from the object the runtime's own `GetSeries` constructs (the constructor only stores its two
arguments). An install at the Bridge anchor failed to load (`ReferenceError: ApiChartSeries is not defined`).
