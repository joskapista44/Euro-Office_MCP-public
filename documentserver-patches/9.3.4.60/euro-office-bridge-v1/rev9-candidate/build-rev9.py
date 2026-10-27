#!/usr/bin/env python3
"""
X-O03/6b (card an internal branch, kioszto: zero;
X-O03/5 design docs/X-O03-5-SCATTER-Y-READ-CONSUMER-DESIGN-20261014.md; X-O03/6a LIVE r6 PASS,
internal tracker) -- OFFLINE build of the euro-office-bridge-v1 patchVersion /9 CANDIDATE.

/9 = /8 + exactly ONE new capability, chart.series.scatterY.read/v1 (kind READ, mutation NO),
served by the existing read() entry. X-O03/4 (LIVE r4, FOUND_READ_PATH) measured that a scatter
chart's Y-series value is NOT observable through the PUBLIC chartApi (`GetValues()` is a FORMULA
STRING or THROWS on scatter), but IS observable through the internal model at
`series.private_GetSeries().yVal.getValues()`. X-O03/6a (LIVE r6) measured that this entry point
and the `chart.dataLabels.effective.read/v1` precedent's own entry point
(`api.Chart.chart.plotArea.charts[0].series[i]`) resolve to the SAME object (`===`, both series,
both sessions) -- so this capability reuses that ALREADY-SERVED entry-point code verbatim (same
chart-by-name lookup, same `charts2.length !== 1` single-type guard), never introducing a second
way to reach the internal model. a maintainer's decision (internal tracker): read via `yVal.getValues()`
(flat array), each element through `Number(x)` (non-finite -> null, never guessed); ceiling 300
points/series (the EXISTING `MAX_CHART_LABEL_POINTS`/row-break-style ceiling value, reused as a
placeholder -- NOT yet measured for THIS capability's own payload size, named in the manifest
entry's own semanticContract).

Deterministic: the /9 block is derived from the CANONICAL /8 block (../rev8-candidate/bridge-block.js,
sha pinned) by three exact string substitutions, each asserted to occur exactly once. No
timestamps, no environment input.

Usage:
  build-rev9.py                          -> writes bridge-block.js, old.txt, new.txt,
                                             manifest-mechanism-proof.json next to this file
This script NEVER writes a Document Server file, and (unlike build-rev8.py) has NO --sdk mode:
the real served /8 `sdk-all.js` is a the production host-host artifact this agent (designated-uid, no
an owner-approval gate access) cannot fetch or verify -- producing a manifest.json that CLAIMS an SDK sha
without having read that file would be a fabricated measurement, not a build. The SDK-backed
manifest.json (and the apply itself) is maintainer's gate, a separate, later step -- see README.md.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V8_BLOCK = HERE.parent / "rev8-candidate" / "bridge-block.js"
V8_BLOCK_SHA = "9f285b897f58a751c3164791b2416e085d627a1bd10419261622349806a96be7"
V8_BLOCK_SIZE = 49851

REASONS = ("bad-args, sheet-not-found, charts-threw, chart-not-found, chart-type-not-single, "
           "series-index-out-of-range, yVal-getValues-missing, getValues-threw, "
           "getValues-not-array-like, scatter-y-points-truncated")

READ_ENTRY = (
    '{ id: "chart.series.scatterY.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one chart series (sheet, chart name, series index): the scatter series\' own Y-value '
    'array, read through the internal model (X-O03/4 LIVE FOUND_READ_PATH, series.private_GetSeries().yVal.'
    'getValues(); X-O03/6a LIVE confirmed this is the SAME object as this capability\'s own chart.plotArea.'
    'charts[0].series[i] entry point, reused verbatim) -- a flat array of numbers or null (X-O03/6b review: '
    'null/undefined, a boolean, and an empty-or-whitespace-only string are null BEFORE Number() coercion, '
    'since each of those would otherwise coerce to a numeric 0/1 and never hit the isFinite check below; '
    'every other value through Number(x), with non-finite as null), never guessed; more than 300 points (a '
    'reused, NOT YET independently '
    'measured ceiling -- see X-O03/5 design doc point 6) truncates the WHOLE series to a named reason, never a '
    'partial list; reasons: ' + REASONS + ' (X-O03/6b, /9)" }, '
)

# One line, ES5, no arrow functions -- the same style as every other bridge entry. The chart-lookup
# and single-type-check prefix is VERBATIM the chart.dataLabels.effective.read/v1 dispatch's own
# text (X-O03/6a: same object, same entry point) up to the series-index check -- not re-derived.
READ_IMPL = (
    '"chart.series.scatterY.read/v1": function (args) { '
    'if (!args || typeof args.sheet !== "string" || typeof args.chart !== "string" || typeof args.series !== "number") return { ok: false, reason: "bad-args" }; '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var api = null; for (var i = 0; i < charts.length; i += 1) { var ok = false; try { ok = charts[i].GetName() === args.chart; } catch (e) { ok = false; } if (ok) { api = charts[i]; break; } } '
    'if (!api) return { ok: false, reason: "chart-not-found" }; '
    'var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; var charts2 = plot && plot.charts; '
    'if (!charts2 || charts2.length !== 1) return { ok: false, reason: "chart-type-not-single" }; '
    'var ch = charts2[0]; var series = ch.series || []; var s = series[args.series]; '
    'if (!s) return { ok: false, reason: "series-index-out-of-range" }; '
    'var yVal = s.yVal; if (!yVal || typeof yVal.getValues !== "function") return { ok: false, reason: "yVal-getValues-missing" }; '
    'var raw; try { raw = yVal.getValues(); } catch (e) { return { ok: false, reason: "getValues-threw" }; } '
    'if (!raw || typeof raw.length !== "number") return { ok: false, reason: "getValues-not-array-like" }; '
    'if (raw.length > 300) return { ok: false, reason: "scatter-y-points-truncated" }; '
    'var out = []; for (var p = 0; p < raw.length; p += 1) { '
    'var v = raw[p]; var n; '
    'if (v === null || v === undefined || typeof v === "boolean") { n = NaN; } '
    'else if (typeof v === "string" && v.replace(/^\\s+|\\s+$/g, "") === "") { n = NaN; } '
    'else { n = Number(v); } '
    'out.push(isFinite(n) ? n : null); } '
    'return { ok: true, value: out }; }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the /8 block
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/8"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/9"'),
    ('(B03-WRITE, /8)" }, ], };', '(B03-WRITE, /8)" }, ' + READ_ENTRY + '], };'),
    # NOTE: '}; var EO_BRIDGE_WRITES = {' alone is WRONG as an insertion point -- it already closes
    # the `var EO_BRIDGE_CAPS = {...};` STATEMENT (the leading '};'), so anything inserted between
    # that '};' and 'var EO_BRIDGE_WRITES' sits OUTSIDE any object literal, as a bare dangling
    # "string": token -- a SyntaxError ('Unexpected token :'), caught by node --check on the first
    # build attempt. The insertion must land BEFORE the EO_BRIDGE_CAPS object's own closing '}',
    # i.e. right after the LAST existing capability entry's trailing comma.
    ('}, }; var EO_BRIDGE_WRITES = {', '}, ' + READ_IMPL + '}; var EO_BRIDGE_WRITES = {'),
]
MARKERS = ['"chart.series.scatterY.read/v1": function', '{ id: "chart.series.scatterY.read/v1"', '(X-O03/6b, /9)']


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v8 = V8_BLOCK.read_bytes()
    if sha(v8) != V8_BLOCK_SHA or len(v8) != V8_BLOCK_SIZE:
        sys.exit(f"REFUSE: /8 block sha/size {sha(v8)}/{len(v8)} != {V8_BLOCK_SHA}/{V8_BLOCK_SIZE}; NO OUTPUT")
    t = v8.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v8, t.encode("utf-8")


def main():
    v8, v9 = build_block()
    out = {"v8_block_sha": sha(v8), "v8_block_size": len(v8), "v9_block_sha": sha(v9), "v9_block_size": len(v9)}
    (HERE / "bridge-block.js").write_bytes(v9)
    (HERE / "old.txt").write_bytes(v8)
    (HERE / "new.txt").write_bytes(v9)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v9-x-o03-6b-scatter-y-read", "old": v8.decode("utf-8"), "new": v9.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- a served-sdk-backed manifest.json for the real production apply is a SEPARATE, later "
                              "step, gated to the designated operator role, with the real served file in hand.",
                 original_sha256=sha(v8), patched_sha256=sha(v9), original_size=len(v8), patched_size=len(v9))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
