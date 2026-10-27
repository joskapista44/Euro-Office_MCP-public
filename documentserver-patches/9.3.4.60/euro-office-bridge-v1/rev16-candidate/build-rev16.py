#!/usr/bin/env python3
"""
GH #139 B /16 (owner GO) --
OFFLINE build of the euro-office-bridge-v1 patchVersion /16 CANDIDATE.

/16 = /15 + TWO new READ capabilities (the measured counterpart of /15's two WRITE-ONLY targets):
  chart.series.markerSymbol.read/v1  -- a series' own marker symbol + size
  chart.series.lineStyle.read/v1     -- a series' own line dash + smooth

THE MEASUREMENT this build encodes (LIVE measurement, eldobhato masolat, telepitett
1.0.1-dcafc9b3cf14, ket friss editor-session S1+reopen S2, S1==S2 minden charton, 10/10 symbol +
11/11 dash token BIJEKTIV tablaval, a mezok: `api.Chart.chart.plotArea.charts[0].series[i].marker.symbol`
/ `.marker.size` / `.spPr.ln.prstDash` / `.smooth`; evidence: the GH #139 B /16 reflection + enum probes).
No token below is invented: every one is the measurement's own written-and-read-back table.

SYMBOL_BY_TOKEN (measured, NOT the write side's TOKEN_BY_NAME inverted blindly -- token 4 ('none')
is NOT a write-side key there, since 'none' writes through setNoneMarker(), but it IS a real,
measured READ state: an untouched lineNormal series' own marker reads 4 = 'none', the engine's own
default, not a guessed value (design doc section 7, decision (a)). Token 5 is NOT in the table --
the measurement named it "5 = egyik sem" (neither write token nor 'none'; build-rev15.py's
own docstring already suspected picture/auto) -- a read of 5 is refused NAMED
(`symbol-token-unmeasured`), never guessed as one of the nine.

DASH_BY_CODE (measured, NOT DASH_ENUM's own declaration order -- round 8d/8k's internal codes are
a DIFFERENT, measured enumeration; e.g. solid=6, not 0). smooth: read directly off the series'
OWN field; a type that never carries it (bar, the measurement's own NEG control: "nem irt bar: ...
smooth UNDEFINED") reads back `typeof !== 'boolean'`, named `smooth-field-undefined` (design doc
decision (b)) -- never a guessed boolean, and never silently merged with "set to false".

size/dash NULL ("nem irt"): passed through as literal `null`, never coerced to a guessed default
(design doc decision (a): size:null / prstDash:null must never read as "solid" / a number).

Deterministic: the /16 block is derived from the CANONICAL /15 block
(../rev15-candidate/bridge-block.js, sha pinned) by THREE exact string substitutions, each
asserted to occur exactly once -- the SAME discipline build-rev15.py already used for /14->/15.

Usage:
  build-rev16.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json
                      next to this file.
This script NEVER writes a Document Server file and has NO --sdk mode (the same reason
build-rev15.py's docstring gives: this agent, designated-uid, has no an owner-approval gate access to the
served sdk-all.js). The SDK-backed manifest.json and the real DS apply are the build gate +
the LIVE gate -- see README.md.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V15_BLOCK = HERE.parent / "rev15-candidate" / "bridge-block.js"
V15_BLOCK_SHA = "de50333fce5667cd94362d74ed7f57496afb584ca0f291f41483905991c619ca"
V15_BLOCK_SIZE = 92315

# ---- SYMBOL: measured token<->name table (LIVE reflection -- write+OOXML+
# in-session-field round trip, S1==S2 after reopen). 'none' (4) IS in this table, unlike the
# WRITE side's TOKEN_BY_NAME (which routes 'none' through setNoneMarker() instead) -- here it is
# a real, measured READ state, not a write request.
SYMBOL_BY_TOKEN = {
    0: "circle", 1: "dash", 2: "diamond", 3: "dot", 4: "none",
    6: "plus", 7: "square", 8: "star", 9: "triangle", 10: "x",
}
# ---- DASH: measured code<->name table (LIVE reflection, 11/11 bijective).
DASH_BY_CODE = {
    0: "dash", 1: "dashDot", 2: "dot", 3: "lgDash", 4: "lgDashDot", 5: "lgDashDotDot",
    6: "solid", 7: "sysDash", 8: "sysDashDot", 9: "sysDashDotDot", 10: "sysDot",
}

MARKER_SYMBOL_READ_REASONS = ("bad-args, sheet-not-found, charts-threw, chart-not-found, "
                               "chart-type-not-single, series-index-out-of-range")
LINE_STYLE_READ_REASONS = MARKER_SYMBOL_READ_REASONS

MARKER_SYMBOL_READ_ENTRY = (
    '{ id: "chart.series.markerSymbol.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one chart series (sheet, chart name, series index): the series\' own marker '
    'symbol+size, read through the internal model (GH #139 B /16, LIVE MEASURED via write+session+reopen '
    'round trip, SAME series.marker object chart.series.markerStyle.read/v1 already reads '
    'fill/outline off -- symbol/size are DIFFERENT fields on it) -- { symbol, size }; symbol is one of '
    + '|'.join(sorted(SYMBOL_BY_TOKEN.values())) + ' (the model\'s own token, mapped through a table LIVE '
    'MEASURED by write-and-read-back, NOT the ECMA-376 XSD\'s declaration order -- an unmapped token '
    '(measured: token 5 exists and is neither a named symbol nor none) reads a NAMED '
    'symbol-token-unmeasured string, never guessed as one of the ten); size is the model\'s own integer '
    "(2..72, ST_MarkerSize) or null (the model's own never-written state -- never defaulted to a "
    "guessed number); an untouched series (no prior marker write) reads symbol='none' (the model's own "
    'default, NOT a guess -- MEASURED), size=null; reasons: ' + MARKER_SYMBOL_READ_REASONS
    + ' (GH #139 B /16)" }, '
)
LINE_STYLE_READ_ENTRY = (
    '{ id: "chart.series.lineStyle.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one chart series (sheet, chart name, series index): the series\' own line dash+'
    'smooth, read through the internal model (GH #139 B /16, LIVE MEASURED, SAME chain '
    'chart.series.lineStyle.write/v1 writes through -- series.spPr.ln.prstDash / series.smooth) -- '
    '{ dash, smooth }; dash is one of ' + '|'.join(sorted(DASH_BY_CODE.values())) + ' (the model\'s own '
    'code, mapped through a table LIVE MEASURED by write-and-read-back -- NOT the write side\'s own '
    'DASH_ENUM declaration order, e.g. solid is code 6, not 0) or null (no line / no explicit dash '
    "written -- the model's own unset state, never defaulted to 'solid' or any guessed name); smooth "
    'is the model\'s own boolean, or a NAMED smooth-field-undefined when the series\' own field is not a '
    'boolean (MEASURED: a bar-type series never carries this field at all, the measurement\'s own negative '
    'control -- never guessed as false, never silently merged with an explicit false write); dash and '
    'smooth are read INDEPENDENTLY (dash off spPr.ln, smooth off the series itself) -- a series with no '
    'spPr/ln still reads its own smooth field, and vice versa; reasons: ' + LINE_STYLE_READ_REASONS
    + ' (GH #139 B /16)" }, '
)

_SERIES_LOOKUP_PREFIX_READ = (
    'if (!args || typeof args.sheet !== "string" || typeof args.chart !== "string" || typeof args.series !== "number") return { ok: false, reason: "bad-args" }; '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var api = null; for (var i = 0; i < charts.length; i += 1) { var ok = false; try { ok = charts[i].GetName() === args.chart; } catch (e) { ok = false; } if (ok) { api = charts[i]; break; } } '
    'if (!api) return { ok: false, reason: "chart-not-found" }; '
    'var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; var charts2 = plot && plot.charts; '
    'if (!charts2 || charts2.length !== 1) return { ok: false, reason: "chart-type-not-single" }; '
    'var ch = charts2[0]; var series = ch.series || []; var s = series[args.series]; '
    'if (!s) return { ok: false, reason: "series-index-out-of-range" }; '
)

MARKER_SYMBOL_READ_IMPL = (
    '"chart.series.markerSymbol.read/v1": function (args) { '
    + _SERIES_LOOKUP_PREFIX_READ +
    'var SYMBOL_BY_TOKEN = ' + json.dumps({str(k): v for k, v in SYMBOL_BY_TOKEN.items()}) + '; '
    'var marker = s.marker; '
    # the pre-deploy mechanism probe: {symbol:null,size:null} with ok:true is byte-identical to a
    # written-but-unset series (a real state, LIVE measured) -- a series with NO marker object at
    # all (bar/pie) is a DIFFERENT state and must fail closed, named, not look like "not written".
    'if (!marker) return { ok: false, reason: "no-marker-object" }; '
    'var tok = marker.symbol; '
    'var symbolOut = (typeof tok === "number" && Object.prototype.hasOwnProperty.call(SYMBOL_BY_TOKEN, String(tok))) '
    '? SYMBOL_BY_TOKEN[String(tok)] : (typeof tok === "number" ? ("UNOBSERVABLE:symbol-token-unmeasured:" + tok) : null); '
    'var sizeOut = (typeof marker.size === "number") ? marker.size : null; '
    'return { ok: true, value: { symbol: symbolOut, size: sizeOut } }; }, '
)
LINE_STYLE_READ_IMPL = (
    '"chart.series.lineStyle.read/v1": function (args) { '
    + _SERIES_LOOKUP_PREFIX_READ +
    'var DASH_BY_CODE = ' + json.dumps({str(k): v for k, v in DASH_BY_CODE.items()}) + '; '
    'var ln = s.spPr && s.spPr.ln; var code = ln ? ln.prstDash : null; '
    'var dashOut = (typeof code === "number" && Object.prototype.hasOwnProperty.call(DASH_BY_CODE, String(code))) '
    '? DASH_BY_CODE[String(code)] : (typeof code === "number" ? ("UNOBSERVABLE:dash-code-unmeasured:" + code) : null); '
    'var smoothOut = (typeof s.smooth === "boolean") ? s.smooth : "UNOBSERVABLE:smooth-field-undefined"; '
    'return { ok: true, value: { dash: dashOut, smooth: smoothOut } }; }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the EVOLVING text at the time it is applied.
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/15"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/16"'),
    ('(X-O05/5-B, /15)" }, ], };', '(X-O05/5-B, /15)" }, ' + MARKER_SYMBOL_READ_ENTRY + LINE_STYLE_READ_ENTRY + '], };'),
    # the READ-handlers dispatch object's closing anchor (drawing.fillLineColor.read/v1's own tail,
    # /14's last READ addition), immediately before "var EO_BRIDGE_WRITES = {".
    ('return { ok: true, value: { fill: fill, line: line } }; }, }; var EO_BRIDGE_WRITES = {',
     'return { ok: true, value: { fill: fill, line: line } }; }, ' + MARKER_SYMBOL_READ_IMPL + LINE_STYLE_READ_IMPL + '}; var EO_BRIDGE_WRITES = {'),
]
MARKERS = [
    '"chart.series.markerSymbol.read/v1": function', '{ id: "chart.series.markerSymbol.read/v1"',
    '"chart.series.lineStyle.read/v1": function', '{ id: "chart.series.lineStyle.read/v1"',
]


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v15 = V15_BLOCK.read_bytes()
    if sha(v15) != V15_BLOCK_SHA or len(v15) != V15_BLOCK_SIZE:
        sys.exit(f"REFUSE: /15 block sha/size {sha(v15)}/{len(v15)} != {V15_BLOCK_SHA}/{V15_BLOCK_SIZE}; NO OUTPUT")
    t = v15.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("(GH #139 B /16)") != 2:
        sys.exit(f"REFUSE: (GH #139 B /16) count {t.count('(GH #139 B /16)')} != 2; NO OUTPUT")
    return v15, t.encode("utf-8")


def main():
    v15, v16 = build_block()
    out = {"v15_block_sha": sha(v15), "v15_block_size": len(v15), "v16_block_sha": sha(v16), "v16_block_size": len(v16)}
    (HERE / "bridge-block.js").write_bytes(v16)
    (HERE / "old.txt").write_bytes(v15)
    (HERE / "new.txt").write_bytes(v16)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v16-gh139-b-markersymbol-linestyle-read", "old": v15.decode("utf-8"), "new": v16.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- a served-sdk-backed manifest.json for the real production apply is a SEPARATE, later "
                              "step, gated to the designated operator role, with the real served file in hand.",
                 original_sha256=sha(v15), patched_sha256=sha(v16), original_size=len(v15), patched_size=len(v16))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
