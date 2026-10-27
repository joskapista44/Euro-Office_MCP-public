#!/usr/bin/env python3
"""
X-O04/10 (card an internal branch, kioszto: zero;
design docs/X-O04-9-CHART-STYLE-READ-BRIDGE-DESIGN-20261014.md, accepted zero comment #ref) --
OFFLINE build of the euro-office-bridge-v1 patchVersion /11 CANDIDATE.

/11 = /10 + TWO new capabilities (READ, mutation NO):
  chart.series.markerStyle.read/v1       -- a series' marker fill+outline colour
  chart.axis.majorGridlinesColor.read/v1 -- a chart's vertical/horizontal gridline colour

Both chains are X-O04/4..X-O04/8 LIVE VERIFIED (cap-live/xo046-r1.out, cap-live/xo048-r1.out):
  marker fill    : series.private_GetSeries().marker.spPr.Fill.fill.color.color.RGBA
  marker outline : series.private_GetSeries().marker.spPr.ln.Fill.fill.color.color.RGBA
  gridline (H/V) : cs.chart.plotArea.getHorizontalAxis()/getVerticalAxis().majorGridlines.ln.Fill.fill.color.color.RGBA

The chart-lookup-by-name / series-resolution prefix for markerStyle is VERBATIM the
chart.series.scatterY.read/v1 dispatch's own text (up to series-index-out-of-range) -- not
re-derived, same convention build-rev9.py's own docstring established for its own reuse of the
dataLabels precedent's entry point.

The cross-binding X-O04/8 found (SetMajorVerticalGridlines's colour lands on getHorizontalAxis(),
and vice versa) is hidden ENTIRELY inside majorGridlinesColor's own dispatch function: the response
is keyed by the PUBLIC setter's own semantics (verticalGridlines/horizontalGridlines), never the
internal getter name -- the design doc's own central decision (point 2).

marker.spPr===null (X-O04/6 MEASURED, lineNormal control) and axis.majorGridlines===null (X-O04/7-8
MEASURED, one axis of an untouched gridline control) are both PROVEN VALID model states -> null,
NOT UNOBSERVABLE. An untouched scatter series' marker.spPr being a non-null object whose own
deeper chain resolution was never measured, and the possibility that a "default" colour (e.g. the
untouched gridline control's OTHER axis, which carries a real default black RGBA) is
indistinguishable from an explicitly file-written one, are both NEM-MERT -- the response never
claims more than "the model's current value" (design doc, point 1 + point 6).

Deterministic: the /11 block is derived from the CANONICAL /10 block
(../rev10-candidate/bridge-block.js, sha pinned) by FOUR exact string substitutions, each asserted
to occur exactly once. No timestamps, no environment input.

Usage:
  build-rev11.py                         -> writes bridge-block.js, old.txt, new.txt,
                                            manifest-mechanism-proof.json next to this file
This script NEVER writes a Document Server file, and has NO --sdk mode: the real served /10
sdk-all.js is a the production host/.10-host artifact this agent (designated-uid, no an owner-approval gate access)
cannot fetch or verify -- producing a manifest.json that CLAIMS an SDK sha without having read
that file would be a fabricated measurement, not a build. The SDK-backed manifest.json (and the
apply itself) is maintainer's gate, a separate, later step -- see README.md.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V10_BLOCK = HERE.parent / "rev10-candidate" / "bridge-block.js"
V10_BLOCK_SHA = "16ad3780e9ba99eb029ea8573cabdf8ff1e41cc7341eaba8cb15e2ad65e169ed"
V10_BLOCK_SIZE = 53015

MARKER_REASONS = ("bad-args, sheet-not-found, charts-threw, chart-not-found, chart-type-not-single, "
                   "series-index-out-of-range, marker-sppr-chain-broken")
GRIDLINE_REASONS = ("bad-args, sheet-not-found, charts-threw, chart-not-found, "
                     "axis-getters-missing, axis-getter-threw, majorgridlines-chain-broken")

MARKER_STYLE_ENTRY = (
    '{ id: "chart.series.markerStyle.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one chart series (sheet, chart name, series index): the series\' own marker '
    'fill+outline colour, read through the internal model (X-O04/6+/8 LIVE VERIFIED, series.private_GetSeries().'
    'marker.spPr.Fill.fill.color.color.RGBA / .spPr.ln.Fill.fill.color.color.RGBA) -- { fill, outline }, each a '
    '6-hex-digit RGB string or null (null is the MODEL\'s own no-colour-set state -- X-O04/6 measured '
    'marker.spPr===null on an untouched lineNormal series\' marker; an untouched scatter series\' marker.spPr is '
    'a non-null object whose own further chain was NOT measured for this capability and may itself resolve to '
    'the DS\'s own compiled default rather than null -- the hex value, when present, is always the model\'s '
    'CURRENT value, never a guarantee of an explicit user/file setting), never guessed; fill and outline read '
    'INDEPENDENTLY from separate marker.spPr.Fill / marker.spPr.ln sub-objects, never a shared search root '
    '(X-O04/6\'s own false-positive lesson: a single shared byte between two different written colours must '
    'never verify either on its own); reasons: ' + MARKER_REASONS + ' (X-O04/10, /11)" }, '
)
GRIDLINES_COLOR_ENTRY = (
    '{ id: "chart.axis.majorGridlinesColor.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one chart (sheet, chart name): the chart\'s own major-gridline colour on both axes, '
    'read through the internal model (X-O04/7+/8 LIVE VERIFIED, cs.chart.plotArea.getHorizontalAxis()/'
    'getVerticalAxis().majorGridlines.ln.Fill.fill.color.color.RGBA) -- { verticalGridlines, horizontalGridlines '
    '}, each a 6-hex-digit RGB string or null, KEYED BY THE PUBLIC SETTER\'S OWN SEMANTICS, never the internal '
    'getter name (X-O04/8 LIVE finding: SetMajorVerticalGridlines\'s colour lands on getHorizontalAxis(), and '
    'vice versa -- verticalGridlines is always read via getHorizontalAxis(), horizontalGridlines via '
    'getVerticalAxis(), the cross-binding hidden entirely inside this one function); null means majorGridlines '
    'itself is null on that axis (a PROVEN valid model state -- X-O04/8 measured an untouched chart\'s two axes '
    'behave DIFFERENTLY: one axis\'s majorGridlines is null, the other already carries the DS\'s own default '
    'black RGBA, which reads as a REAL hex value here, not null); the hex value, when present, is the model\'s '
    'CURRENT value, never a guarantee of an explicit user/file setting; reasons: ' + GRIDLINE_REASONS +
    ' (X-O04/10, /11)" }, '
)

# the chart-lookup + series-resolution prefix, VERBATIM from the landed chart.series.scatterY.read/v1
# dispatch (rev10-candidate/bridge-block.js) up to series-index-out-of-range -- not re-derived.
MARKER_STYLE_IMPL = (
    '"chart.series.markerStyle.read/v1": function (args) { '
    'if (!args || typeof args.sheet !== "string" || typeof args.chart !== "string" || typeof args.series !== "number") return { ok: false, reason: "bad-args" }; '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var api = null; for (var i = 0; i < charts.length; i += 1) { var ok = false; try { ok = charts[i].GetName() === args.chart; } catch (e) { ok = false; } if (ok) { api = charts[i]; break; } } '
    'if (!api) return { ok: false, reason: "chart-not-found" }; '
    'var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; var charts2 = plot && plot.charts; '
    'if (!charts2 || charts2.length !== 1) return { ok: false, reason: "chart-type-not-single" }; '
    'var ch = charts2[0]; var series = ch.series || []; var s = series[args.series]; '
    'if (!s) return { ok: false, reason: "series-index-out-of-range" }; '
    # X-O04/6-8 proven chain: spPrLike.Fill.fill.color.color.RGBA, {R,G,B} number fields.
    'var rgbaOf = function (spPrLike) { '
    'var rgba = spPrLike && spPrLike.Fill && spPrLike.Fill.fill && spPrLike.Fill.fill.color && spPrLike.Fill.fill.color.color && spPrLike.Fill.fill.color.color.RGBA; '
    'if (!rgba || typeof rgba.R !== "number" || typeof rgba.G !== "number" || typeof rgba.B !== "number") return undefined; '
    'var h = function (n) { return ("0" + n.toString(16).toUpperCase()).slice(-2); }; '
    'return h(rgba.R) + h(rgba.G) + h(rgba.B); }; '
    'var marker = s.marker; '
    # X-O04/9 review fix: the MEASURED null-state is marker.spPr===null, not
    # marker===null -- the !marker guard stays, defensively, but is NOT itself a measured branch.
    'if (!marker || !marker.spPr) return { ok: true, value: { fill: null, outline: null } }; '
    'var fillHex = rgbaOf(marker.spPr); '
    # KULON sub-root (marker.spPr vs marker.spPr.ln) -- never a shared search, X-O04/6's false-positive lesson.
    'var outlineHex = rgbaOf(marker.spPr.ln); '
    'if (fillHex === undefined && outlineHex === undefined) return { ok: false, reason: "marker-sppr-chain-broken" }; '
    'return { ok: true, value: { fill: fillHex === undefined ? null : fillHex, outline: outlineHex === undefined ? null : outlineHex } }; }, '
)
GRIDLINES_COLOR_IMPL = (
    '"chart.axis.majorGridlinesColor.read/v1": function (args) { '
    'if (!args || typeof args.sheet !== "string" || typeof args.chart !== "string") return { ok: false, reason: "bad-args" }; '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var api = null; for (var i = 0; i < charts.length; i += 1) { var ok = false; try { ok = charts[i].GetName() === args.chart; } catch (e) { ok = false; } if (ok) { api = charts[i]; break; } } '
    'if (!api) return { ok: false, reason: "chart-not-found" }; '
    'var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; '
    'if (!plot || typeof plot.getHorizontalAxis !== "function" || typeof plot.getVerticalAxis !== "function") return { ok: false, reason: "axis-getters-missing" }; '
    'var hAxis, vAxis; try { hAxis = plot.getHorizontalAxis(); vAxis = plot.getVerticalAxis(); } catch (e) { return { ok: false, reason: "axis-getter-threw" }; } '
    'var rgbaHexOf = function (axis) { '
    'var mg = axis && axis.majorGridlines; if (!mg) return null; '
    'var ln = mg.ln; var rgba = ln && ln.Fill && ln.Fill.fill && ln.Fill.fill.color && ln.Fill.fill.color.color && ln.Fill.fill.color.color.RGBA; '
    'if (!rgba || typeof rgba.R !== "number" || typeof rgba.G !== "number" || typeof rgba.B !== "number") return undefined; '
    'var h = function (n) { return ("0" + n.toString(16).toUpperCase()).slice(-2); }; '
    'return h(rgba.R) + h(rgba.G) + h(rgba.B); }; '
    # X-O04/8 LIVE finding, the cross-binding hidden HERE, ONE line each -- nowhere else in the codebase.
    'var verticalGridlinesHex = rgbaHexOf(hAxis); '
    'var horizontalGridlinesHex = rgbaHexOf(vAxis); '
    'if (verticalGridlinesHex === undefined || horizontalGridlinesHex === undefined) return { ok: false, reason: "majorgridlines-chain-broken" }; '
    'return { ok: true, value: { verticalGridlines: verticalGridlinesHex, horizontalGridlines: horizontalGridlinesHex } }; }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the /10 block
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/10"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/11"'),
    ('(X-O03/6b, /9)" }, ], };', '(X-O03/6b, /9)" }, ' + MARKER_STYLE_ENTRY + GRIDLINES_COLOR_ENTRY + '], };'),
    # same insertion-point lesson build-rev9.py's own docstring records: the dispatch-object literal's
    # closing '};' is the ONLY safe anchor -- inserting after it would land the new entries outside
    # any object literal (a dangling "string": token, SyntaxError).
    ('out.push(isFinite(n) ? n : null); } return { ok: true, value: out }; }, }; var EO_BRIDGE_WRITES = {',
     'out.push(isFinite(n) ? n : null); } return { ok: true, value: out }; }, ' + MARKER_STYLE_IMPL + GRIDLINES_COLOR_IMPL + '}; var EO_BRIDGE_WRITES = {'),
]
MARKERS = [
    '"chart.series.markerStyle.read/v1": function', '{ id: "chart.series.markerStyle.read/v1"',
    '"chart.axis.majorGridlinesColor.read/v1": function', '{ id: "chart.axis.majorGridlinesColor.read/v1"',
]


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v10 = V10_BLOCK.read_bytes()
    if sha(v10) != V10_BLOCK_SHA or len(v10) != V10_BLOCK_SIZE:
        sys.exit(f"REFUSE: /10 block sha/size {sha(v10)}/{len(v10)} != {V10_BLOCK_SHA}/{V10_BLOCK_SIZE}; NO OUTPUT")
    t = v10.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    # (X-O04/10, /11) must appear exactly twice overall (once per new manifest entry) -- the single
    # MARKERS check above only proves >=1; assert the exact count here too.
    if t.count("(X-O04/10, /11)") != 2:
        sys.exit(f"REFUSE: (X-O04/10, /11) count {t.count('(X-O04/10, /11)')} != 2; NO OUTPUT")
    return v10, t.encode("utf-8")


def main():
    v10, v11 = build_block()
    out = {"v10_block_sha": sha(v10), "v10_block_size": len(v10), "v11_block_sha": sha(v11), "v11_block_size": len(v11)}
    (HERE / "bridge-block.js").write_bytes(v11)
    (HERE / "old.txt").write_bytes(v10)
    (HERE / "new.txt").write_bytes(v11)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v11-x-o04-10-chart-style-read", "old": v10.decode("utf-8"), "new": v11.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- a served-sdk-backed manifest.json for the real production apply is a SEPARATE, later "
                              "step, gated to the designated operator role, with the real served file in hand.",
                 original_sha256=sha(v10), patched_sha256=sha(v11), original_size=len(v10), patched_size=len(v11))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
