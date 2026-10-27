#!/usr/bin/env python3
"""
X-O05/2 + X-O05/3 (internal tracker; both reviewed and accepted) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /12 CANDIDATE, in ONE deterministic pass from the pinned /11
block.

/12 = /11 + FOUR new capabilities:
  READ  (X-O05/2, mutation NO):
    chart.series.order.read/v1     -- a chart series' own c:order index
    chart.bar.gapOverlap.read/v1   -- a bar chart's own gapWidth/overlap
  WRITE (X-O05/3, mutation TARGETED):
    chart.bar.gapOverlap.write/v1   -- the SAME field gapOverlap.read/v1 reads, written through the
                                       internal setGapWidth()/setOverlap() METHODS (NOT a direct
                                       field assignment -- that does not persist, PRE-DEPLOY LIVE)
    chart.series.order.write/v1     -- the SAME series object order.read/v1 reads, via setOrder()

X-O05/3's WRITE pair used to be a SEPARATE second-stage script (build-rev12-writes.py) that read
THIS script's own output and substituted onto it. That split was found FRAGILE (internal tracker,
reviewed): re-running this script alone silently dropped the write stage, because it wrote
bridge-block.js's FULL content, read-only, overwriting whatever a later stage had added. There is
now ONE entry point: running this file produces the complete READ+WRITE /12 candidate every time,
from the SAME pinned /11 input, in one pass -- there is no second script and no intermediate
file whose staleness could silently regress the output.

order chain is X-O05/1 LIVE VERIFIED (session + after reopen, getOrder() == the written order
index): series.private_GetSeries().getOrder() -- except the Bridge's internal chart-type
resolution already lands one level further in than the public API wrapper does (same as
chart.series.markerStyle.read/v1's own `s.marker`, no `.private_GetSeries()` call inside THIS
handler either), so the implementation below calls `.getOrder()` directly on the resolved
internal series object. seriesOrder.write/v1 writes the WHOLE permutation in one call, calling the
resolved series objects' own setOrder() once each, in the array's own order -- NOT a single
{seriesIndex, order} pair (design review, internal tracker: the product allows only one
chart-style write per task, so a per-series shape could never implement a two-series swap, since
the first write would always collide with the other series' as-yet-unmoved order). The engine's
own behavior on an INTERMEDIATE colliding order value is NOT MEASURED by this handler -- deferred
to a dedicated LIVE harness (2-series swap + 3-series cycle).

gapOverlap is a DIFFERENT shape from every other migrated axis in this file: there is no PUBLIC
getter AND no PUBLIC setter either (X-O05/1 LIVE r4's own discovery, confirmed by a per-host
positive control that would otherwise raise doubt on the NO_CANDIDATE_FOUND). The READ path is
the internal DATA field `gapWidth`/`overlap` itself, on the SAME plotArea.charts[0] host order's
own chart-type-not-single check resolves. A LIVE re-measurement (after X-O05/1 shipped) read that
field DIRECTLY, in-session right after a write AND after a fresh reopen with no further write:
BOTH read back the written value -- not an unsynchronized write-time-only mirror.

The WRITE path is NOT the same field assignment, though it reads as though it should be: a
PRE-DEPLOY LIVE probe (X-O05/3, internal tracker) found that `ch.gapWidth = v` / `ch.overlap = v`
APPEAR to apply (the SAME session's own read-back shows the new value) but do NOT survive a
reopen -- the field write bypasses the save/history path, leaving a transient, memory-only change.
The internal object DOES carry a METHOD for this (`setGapWidth()`/`setOverlap()`, NOT public),
and calling it is the SAME path X-O05/1's own original measurement used and found persisting
(OOXML round-trip 219 -> written value, confirmed across a reopen). The write handler below calls
these methods, never the field, for exactly this reason.

Measured on a BAR chart only; every other chart type is UNMEASURED for this field (both read and
write) and refuses by its own name (chart-type-not-bar), never a guessed number or null.

The chart-lookup-by-name / chart-type-resolution prefix for all four is VERBATIM the
chart.series.markerStyle.read/v1 dispatch's own text up to chart-type-not-single (not
re-derived, same convention build-rev9.py's own docstring established for its own reuse of the
dataLabels precedent's entry point).

Deterministic: the /12 block is derived from the CANONICAL /11 block
(../rev11-candidate/bridge-block.js, sha pinned) by FIVE exact string substitutions, each
asserted to occur exactly once. No timestamps, no environment input.

Usage:
  build-rev12.py                         -> writes bridge-block.js, old.txt, new.txt,
                                            manifest-mechanism-proof.json next to this file
This script NEVER writes a Document Server file, and has NO --sdk mode: the real served /11
sdk-all.js is a host artifact this agent (designated-uid, no an owner-approval gate access) cannot fetch
or verify -- producing a manifest.json that CLAIMS an SDK sha without having read that file would
be a fabricated measurement, not a build. The SDK-backed manifest.json (and the apply itself) is
a separate, later, gated step -- see README.md.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V11_BLOCK = HERE.parent / "rev11-candidate" / "bridge-block.js"
V11_BLOCK_SHA = "3325bf94a80f677d8b0c97d9acf7c8df446c67a6266d87d93253f9ea3017b216"
V11_BLOCK_SIZE = 59436

ORDER_REASONS = ("bad-args, sheet-not-found, charts-threw, chart-not-found, chart-type-not-single, "
                  "series-index-out-of-range, getorder-missing, getorder-threw, getorder-not-number")
GAP_OVERLAP_REASONS = ("bad-args, sheet-not-found, charts-threw, chart-not-found, chart-type-threw, "
                        "chart-type-not-bar, chart-type-not-single, gapwidth-overlap-not-number")
GAP_OVERLAP_WRITE_REASONS = ("bad-args, unsupported-property, no-property, invalid-gapwidth, invalid-overlap, "
                             "sheet-not-found, charts-threw, chart-not-found, chart-type-threw, chart-type-not-bar, "
                             "chart-type-not-single, setgapwidth-missing, setoverlap-missing, setgapwidth-threw, "
                             "setoverlap-threw")
SERIES_ORDER_WRITE_REASONS = ("bad-args, unsupported-property, order-empty, order-not-permutation, sheet-not-found, "
                               "charts-threw, chart-not-found, chart-type-not-single, series-count-mismatch, "
                               "setorder-missing, setorder-threw")

SERIES_ORDER_ENTRY = (
    '{ id: "chart.series.order.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one chart series (sheet, chart name, series index): the series\' own '
    'c:order index, read through the internal model (X-O05/1 LIVE VERIFIED, session AND after '
    'reopen, series.private_GetSeries().getOrder() -- the Bridge\'s own chart-type resolution '
    'already lands one level further in, so this handler calls getOrder() directly on the '
    'resolved internal series object, same convention as chart.series.markerStyle.read/v1\'s own '
    '`s.marker`) -- { order }, a non-negative integer, never guessed; getOrder() itself, never a '
    'data field, so there is no risk of reading an unsynchronized write-time-only cache. '
    'reasons: ' + ORDER_REASONS + ' (X-O05/2, /12)" }, '
)
GAP_OVERLAP_ENTRY = (
    '{ id: "chart.bar.gapOverlap.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one chart (sheet, chart name): a BAR chart\'s own gapWidth/overlap, read '
    'through the internal model\'s OWN data field (X-O05/1 LIVE VERIFIED there is no public '
    'getter and no internal method either; a LIVE re-measurement read plotArea.charts[0].gapWidth '
    'and .overlap directly and confirmed BOTH survive a fresh reopen with no further write, not '
    'an unsynchronized write-time-only mirror) -- { gapWidth, overlap }, both numbers, never '
    'guessed. Measured on a BAR chart only; every OTHER chart type is unmeasured for this field '
    'and refuses by its own name (chart-type-not-bar), never a guessed number or null. '
    'reasons: ' + GAP_OVERLAP_REASONS + ' (X-O05/2, /12)" }, '
)
GAP_OVERLAP_WRITE_ENTRY = (
    '{ id: "chart.bar.gapOverlap.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one BAR chart (sheet, chart name), at least one of gapWidth (0..500) / overlap '
    '(-100..100): writes through the internal setGapWidth()/setOverlap() METHODS -- NOT a direct '
    'assignment to the data field chart.bar.gapOverlap.read/v1 reads: a PRE-DEPLOY LIVE measurement '
    'found the field assignment APPEARS to apply in-session but does not survive a reopen; the '
    'setter is the SAME path X-O05/1\'s own measurement found persisting. No public setter exists '
    'for either. An unrequested property is never touched (preserved); every input is validated, '
    'the chart\'s identity and type are resolved, and each requested setter is confirmed to exist, '
    'ALL BEFORE the first mutation; the result is an operation receipt {applied:[written property '
    'names]} and NOT a verification -- the caller verifies with a fresh chart.bar.gapOverlap.read/v1; '
    'a setter that throws after an earlier one succeeded leaves a PARTIAL write, reported as '
    '{partial:true, applied:[...]}, never rolled back; no fallback; served only through write(); '
    'reasons: ' + GAP_OVERLAP_WRITE_REASONS + ' (X-O05/3, /12)" }, '
)
SERIES_ORDER_WRITE_ENTRY = (
    '{ id: "chart.series.order.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one chart (sheet, chart name), ONE call, the WHOLE permutation: order, an array '
    'whose length equals the chart\'s own series count, every value a distinct integer 0..length-1 -- '
    'NOT a single-series {seriesIndex, order} pair (design review, internal tracker: the product allows '
    'only one chart-style write per task, so a per-series shape could never implement a two-series swap '
    '-- the first write would always collide with the other series\' as-yet-unmoved order); applied '
    'through the internal model\'s own setOrder() method (X-O05/1 LIVE VERIFIED), called once per series '
    'in the array\'s own order; the chart\'s identity, type-is-single, and series-count-matches-order-'
    'length are all checked BEFORE the first setOrder() call; the engine\'s own behavior on an '
    'INTERMEDIATE colliding order value (e.g. the first setOrder() call of a direct two-series swap) is '
    'NOT MEASURED by this handler -- deferred to a dedicated LIVE harness (2-series swap + 3-series '
    'cycle); a setOrder() that throws after an earlier one in the SAME call succeeded fails the whole '
    'call and reports {partial:true, applied:[]} (the compound order property is never reported applied '
    'unless EVERY series\' setOrder() succeeded), never rolled back; no fallback; served only through '
    'write(); reasons: ' + SERIES_ORDER_WRITE_REASONS + ' (X-O05/3, /12)" }, '
)

# the chart-lookup + chart-type-resolution prefix, VERBATIM from the landed
# chart.series.markerStyle.read/v1 dispatch (rev11-candidate/bridge-block.js) up to
# chart-type-not-single -- not re-derived.
SERIES_ORDER_IMPL = (
    '"chart.series.order.read/v1": function (args) { '
    'if (!args || typeof args.sheet !== "string" || typeof args.chart !== "string" || typeof args.series !== "number") return { ok: false, reason: "bad-args" }; '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var api = null; for (var i = 0; i < charts.length; i += 1) { var ok = false; try { ok = charts[i].GetName() === args.chart; } catch (e) { ok = false; } if (ok) { api = charts[i]; break; } } '
    'if (!api) return { ok: false, reason: "chart-not-found" }; '
    'var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; var charts2 = plot && plot.charts; '
    'if (!charts2 || charts2.length !== 1) return { ok: false, reason: "chart-type-not-single" }; '
    'var ch = charts2[0]; var series = ch.series || []; var s = series[args.series]; '
    'if (!s) return { ok: false, reason: "series-index-out-of-range" }; '
    'if (typeof s.getOrder !== "function") return { ok: false, reason: "getorder-missing" }; '
    'var order; try { order = s.getOrder(); } catch (e) { return { ok: false, reason: "getorder-threw" }; } '
    'if (typeof order !== "number") return { ok: false, reason: "getorder-not-number" }; '
    'return { ok: true, value: { order: order } }; }, '
)
GAP_OVERLAP_IMPL = (
    '"chart.bar.gapOverlap.read/v1": function (args) { '
    'if (!args || typeof args.sheet !== "string" || typeof args.chart !== "string") return { ok: false, reason: "bad-args" }; '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var api = null; for (var i = 0; i < charts.length; i += 1) { var ok = false; try { ok = charts[i].GetName() === args.chart; } catch (e) { ok = false; } if (ok) { api = charts[i]; break; } } '
    'if (!api) return { ok: false, reason: "chart-not-found" }; '
    'var typ; try { typ = api.GetChartType(); } catch (e) { return { ok: false, reason: "chart-type-threw" }; } '
    'if (typ !== "bar") return { ok: false, reason: "chart-type-not-bar" }; '
    'var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; var charts2 = plot && plot.charts; '
    'if (!charts2 || charts2.length !== 1) return { ok: false, reason: "chart-type-not-single" }; '
    'var ch = charts2[0]; var gw = ch.gapWidth; var ov = ch.overlap; '
    'if (typeof gw !== "number" || typeof ov !== "number") return { ok: false, reason: "gapwidth-overlap-not-number" }; '
    'return { ok: true, value: { gapWidth: gw, overlap: ov } }; }, '
)
GAP_OVERLAP_WRITE_IMPL = (
    '"chart.bar.gapOverlap.write/v1": function (args) { '
    'if (!args || typeof args !== "object" || typeof args.sheet !== "string" || typeof args.chart !== "string") return { ok: false, reason: "bad-args" }; '
    'var allowed = { sheet: 1, chart: 1, gapWidth: 1, overlap: 1 }; var keys = Object.keys(args); '
    'for (var i = 0; i < keys.length; i += 1) { if (!Object.prototype.hasOwnProperty.call(allowed, keys[i])) return { ok: false, reason: "unsupported-property" }; } '
    'var hasGapWidth = Object.prototype.hasOwnProperty.call(args, "gapWidth"); var hasOverlap = Object.prototype.hasOwnProperty.call(args, "overlap"); '
    'if (!hasGapWidth && !hasOverlap) return { ok: false, reason: "no-property" }; '
    'if (hasGapWidth && !(typeof args.gapWidth === "number" && args.gapWidth >= 0 && args.gapWidth <= 500)) return { ok: false, reason: "invalid-gapwidth" }; '
    'if (hasOverlap && !(typeof args.overlap === "number" && args.overlap >= -100 && args.overlap <= 100)) return { ok: false, reason: "invalid-overlap" }; '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var api = null; for (var j = 0; j < charts.length; j += 1) { var okName = false; try { okName = charts[j].GetName() === args.chart; } catch (e) { okName = false; } if (okName) { api = charts[j]; break; } } '
    'if (!api) return { ok: false, reason: "chart-not-found" }; '
    'var typ; try { typ = api.GetChartType(); } catch (e) { return { ok: false, reason: "chart-type-threw" }; } '
    'if (typ !== "bar") return { ok: false, reason: "chart-type-not-bar" }; '
    'var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; var charts2 = plot && plot.charts; '
    'if (!charts2 || charts2.length !== 1) return { ok: false, reason: "chart-type-not-single" }; '
    'var ch = charts2[0]; '
    'if (hasGapWidth && typeof ch.setGapWidth !== "function") return { ok: false, reason: "setgapwidth-missing" }; '
    'if (hasOverlap && typeof ch.setOverlap !== "function") return { ok: false, reason: "setoverlap-missing" }; '
    'var applied = []; '
    'if (hasGapWidth) { try { ch.setGapWidth(args.gapWidth); } catch (e) { return { ok: false, reason: "setgapwidth-threw", partial: applied.length > 0, applied: applied }; } applied.push("gapWidth"); } '
    'if (hasOverlap) { try { ch.setOverlap(args.overlap); } catch (e) { return { ok: false, reason: "setoverlap-threw", partial: applied.length > 0, applied: applied }; } applied.push("overlap"); } '
    'return { ok: true, value: { applied: applied } }; }, '
)
SERIES_ORDER_WRITE_IMPL = (
    '"chart.series.order.write/v1": function (args) { '
    'if (!args || typeof args !== "object" || typeof args.sheet !== "string" || typeof args.chart !== "string" || !Array.isArray(args.order)) return { ok: false, reason: "bad-args" }; '
    'var allowed = { sheet: 1, chart: 1, order: 1 }; var keys = Object.keys(args); '
    'for (var i = 0; i < keys.length; i += 1) { if (!Object.prototype.hasOwnProperty.call(allowed, keys[i])) return { ok: false, reason: "unsupported-property" }; } '
    'var n = args.order.length; if (n === 0) return { ok: false, reason: "order-empty" }; '
    'var seen = {}; for (var i2 = 0; i2 < n; i2 += 1) { var v = args.order[i2]; '
    'if (typeof v !== "number" || Math.floor(v) !== v || v < 0 || v >= n) return { ok: false, reason: "order-not-permutation" }; '
    'if (Object.prototype.hasOwnProperty.call(seen, v)) return { ok: false, reason: "order-not-permutation" }; seen[v] = 1; } '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var api = null; for (var j = 0; j < charts.length; j += 1) { var okName = false; try { okName = charts[j].GetName() === args.chart; } catch (e) { okName = false; } if (okName) { api = charts[j]; break; } } '
    'if (!api) return { ok: false, reason: "chart-not-found" }; '
    'var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; var charts2 = plot && plot.charts; '
    'if (!charts2 || charts2.length !== 1) return { ok: false, reason: "chart-type-not-single" }; '
    'var ch = charts2[0]; var series = ch.series || []; '
    'if (series.length !== n) return { ok: false, reason: "series-count-mismatch" }; '
    'for (var k = 0; k < n; k += 1) { if (typeof series[k].setOrder !== "function") return { ok: false, reason: "setorder-missing" }; } '
    'var applied = []; '
    'for (var m = 0; m < n; m += 1) { try { series[m].setOrder(args.order[m]); } catch (e) { return { ok: false, reason: "setorder-threw", partial: applied.length > 0, applied: [] }; } applied.push(m); } '
    'return { ok: true, value: { applied: ["order"] } }; }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the EVOLVING text at the time it is applied.
# Order matters for the two manifest substitutions (the WRITE anchor is literally the READ
# substitution's own trailing text), but not between the manifest pair and the impl pair (disjoint
# regions of the block) -- kept in this order for readability, matching the capability order above.
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/11"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/12"'),
    ('(X-O04/10, /11)" }, ], };', '(X-O04/10, /11)" }, ' + SERIES_ORDER_ENTRY + GAP_OVERLAP_ENTRY + '], };'),
    # same insertion-point lesson build-rev9.py's own docstring records: the dispatch-object literal's
    # closing '};' is the ONLY safe anchor -- inserting after it would land the new entries outside
    # any object literal (a dangling "string": token, SyntaxError).
    ('(X-O05/2, /12)" }, ], };', '(X-O05/2, /12)" }, ' + GAP_OVERLAP_WRITE_ENTRY + SERIES_ORDER_WRITE_ENTRY + '], };'),
    ('horizontalGridlines: horizontalGridlinesHex } }; }, }; var EO_BRIDGE_WRITES = {',
     'horizontalGridlines: horizontalGridlinesHex } }; }, ' + SERIES_ORDER_IMPL + GAP_OVERLAP_IMPL + '}; var EO_BRIDGE_WRITES = {'),
    # the EO_BRIDGE_WRITES object's closing anchor -- the LAST write handler already in the /11
    # block (worksheet.copy/v1's own copy-verification tail) ends with this exact text,
    # immediately before "var EO_BRIDGE_HYDRATIONS = {".
    ('before: before, after: after }; }, }; var EO_BRIDGE_HYDRATIONS = {',
     'before: before, after: after }; }, ' + GAP_OVERLAP_WRITE_IMPL + SERIES_ORDER_WRITE_IMPL + '}; var EO_BRIDGE_HYDRATIONS = {'),
]
MARKERS = [
    '"chart.series.order.read/v1": function', '{ id: "chart.series.order.read/v1"',
    '"chart.bar.gapOverlap.read/v1": function', '{ id: "chart.bar.gapOverlap.read/v1"',
    '"chart.bar.gapOverlap.write/v1": function', '{ id: "chart.bar.gapOverlap.write/v1"',
    '"chart.series.order.write/v1": function', '{ id: "chart.series.order.write/v1"',
]


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v11 = V11_BLOCK.read_bytes()
    if sha(v11) != V11_BLOCK_SHA or len(v11) != V11_BLOCK_SIZE:
        sys.exit(f"REFUSE: /11 block sha/size {sha(v11)}/{len(v11)} != {V11_BLOCK_SHA}/{V11_BLOCK_SIZE}; NO OUTPUT")
    t = v11.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("(X-O05/2, /12)") != 2:
        sys.exit(f"REFUSE: (X-O05/2, /12) count {t.count('(X-O05/2, /12)')} != 2; NO OUTPUT")
    if t.count("(X-O05/3, /12)") != 2:
        sys.exit(f"REFUSE: (X-O05/3, /12) count {t.count('(X-O05/3, /12)')} != 2; NO OUTPUT")
    return v11, t.encode("utf-8")


def main():
    v11, v12 = build_block()
    out = {"v11_block_sha": sha(v11), "v11_block_size": len(v11), "v12_block_sha": sha(v12), "v12_block_size": len(v12)}
    (HERE / "bridge-block.js").write_bytes(v12)
    (HERE / "old.txt").write_bytes(v11)
    (HERE / "new.txt").write_bytes(v12)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v12-x-o05-2-series-order-gap-overlap-read-plus-x-o05-3-write", "old": v11.decode("utf-8"), "new": v12.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- a served-sdk-backed manifest.json for the real production apply is a SEPARATE, later "
                              "step, gated to the designated operator role, with the real served file in hand.",
                 original_sha256=sha(v11), patched_sha256=sha(v12), original_size=len(v11), patched_size=len(v12))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
