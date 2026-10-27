#!/usr/bin/env python3
"""
X-O07/2 (internal tracker; X-O07/1 LIVE MEASURED and reviewed) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /13 CANDIDATE, in ONE deterministic pass from the pinned /12
block (the single-entry-point convention build-rev12.py's own docstring records -- a split
build/write-stage script was found FRAGILE there and is not repeated here).

/13 = /12 + ONE new capability:
  WRITE (X-O07/2, mutation TARGETED):
    chart.drawing.delete.write/v1  -- deletes the WHOLE named chart drawing

X-O07/1's LIVE mechanism probe (DISCOVER + CANDIDATE_TEST + REAL_TARGET_TEST, on a disposable
copy, reviewed and MEASURED) found the public chart API has no delete method of its own (the
earlier X-O07 reflection probe: only series-level RemoveSeria) -- the internal call the handler
below uses is `api.Chart.deleteDrawingBase()` (no-arg): it removed the targeted chart from
`GetAllCharts()`, left the two control charts and the sheet's own A1:C4 data untouched, both
in-session and after a fresh reopen, and reproduced identically when re-applied to the card's
real target chart (not just the probe's own throwaway one).

This handler's name lookup is STRICTER than every other chart.* handler in this block: those all
take the FIRST chart whose GetName() matches. A delete is irreversible, so a name that resolves
to more than one chart refuses (chart-name-ambiguous) rather than guessing which one the caller
meant; zero matches refuses chart-not-found, same as elsewhere.

The chart-lookup-by-name prefix up to (but not including) chart-type resolution is adapted from
the same convention the other write handlers use (build-rev12.py's own docstring, reused from
chart.series.markerStyle.read/v1) -- this handler needs no chart-TYPE check (deleteDrawingBase()
is not chart-type-specific), so chart-type-not-single/chart-type-not-bar do not apply here.

Deterministic: the /13 block is derived from the CANONICAL /12 block
(../rev12-candidate/bridge-block.js, sha pinned) by TWO exact string substitutions, each
asserted to occur exactly once. No timestamps, no environment input.

Usage:
  build-rev13.py                         -> writes bridge-block.js, old.txt, new.txt,
                                            manifest-mechanism-proof.json next to this file
This script NEVER writes a Document Server file, and has NO --sdk mode: the real served /12
sdk-all.js is a host artifact this agent (designated-uid, no an owner-approval gate access) cannot fetch
or verify -- producing a manifest.json that CLAIMS an SDK sha without having read that file would
be a fabricated measurement, not a build. The SDK-backed manifest.json (and the apply itself) is
a separate, later, gated step -- see README.md.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V12_BLOCK = HERE.parent / "rev12-candidate" / "bridge-block.js"
V12_BLOCK_SHA = "ffdfe0a168296c1ec56a9856673d0ed3b5b4a94d124aba7eae7ed2cba258d884"
V12_BLOCK_SIZE = 71408

DELETE_REASONS = ("bad-args, unsupported-property, sheet-not-found, charts-threw, chart-not-found, "
                   "chart-name-ambiguous, deletedrawingbase-missing, deletedrawingbase-threw")

DELETE_ENTRY = (
    '{ id: "chart.drawing.delete.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one chart (sheet, chart name): deletes the WHOLE chart drawing -- X-O07/1 LIVE '
    'MEASURED the internal CChartSpace.deleteDrawingBase() method (api.Chart, no-arg) removes the named '
    'chart from GetAllCharts(), leaves every OTHER chart and the sheet\'s own cell data untouched, in '
    'session AND after a fresh reopen, and reproduced identically when re-applied to a second, '
    'independent target chart; the public chart API has NO delete method of its own (the X-O07 '
    'reflection probe found only series-level RemoveSeria). The name lookup requires EXACTLY one match: '
    'zero matches refuses chart-not-found, more than one refuses chart-name-ambiguous (never guesses '
    'which one the caller meant for an IRREVERSIBLE operation) -- a stricter lookup than every other '
    'chart.* handler in this block, all of which take the FIRST name match. No partial state: a single '
    'no-arg call, either it throws before anything changes or it fully removes the drawing; no fallback; '
    'served only through write(); reasons: ' + DELETE_REASONS + ' (X-O07/2, /13)" }, '
)

DELETE_IMPL = (
    '"chart.drawing.delete.write/v1": function (args) { '
    'if (!args || typeof args !== "object" || typeof args.sheet !== "string" || typeof args.chart !== "string") return { ok: false, reason: "bad-args" }; '
    'var allowed = { sheet: 1, chart: 1 }; var keys = Object.keys(args); '
    'for (var i = 0; i < keys.length; i += 1) { if (!Object.prototype.hasOwnProperty.call(allowed, keys[i])) return { ok: false, reason: "unsupported-property" }; } '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var matches = []; for (var j = 0; j < charts.length; j += 1) { var okName = false; try { okName = charts[j].GetName() === args.chart; } catch (e) { okName = false; } if (okName) matches.push(charts[j]); } '
    'if (matches.length === 0) return { ok: false, reason: "chart-not-found" }; '
    'if (matches.length > 1) return { ok: false, reason: "chart-name-ambiguous" }; '
    'var api = matches[0]; var cs = api.Chart; '
    'if (!cs || typeof cs.deleteDrawingBase !== "function") return { ok: false, reason: "deletedrawingbase-missing" }; '
    'try { cs.deleteDrawingBase(); } catch (e) { return { ok: false, reason: "deletedrawingbase-threw" }; } '
    'return { ok: true, value: { deleted: true } }; }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the EVOLVING text at the time it is applied.
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/12"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/13"'),
    ('(X-O05/3, /12)" }, ], };', '(X-O05/3, /12)" }, ' + DELETE_ENTRY + '], };'),
    # the EO_BRIDGE_WRITES object's closing anchor: the LAST write handler already in the /12
    # block (chart.series.order.write/v1's own applied:["order"] tail) ends with this exact text,
    # immediately before "var EO_BRIDGE_HYDRATIONS = {" -- same insertion-point convention
    # build-rev12.py's own docstring records (the dispatch object literal's closing '};' is the
    # ONLY safe anchor).
    ('return { ok: true, value: { applied: ["order"] } }; }, }; var EO_BRIDGE_HYDRATIONS = {',
     'return { ok: true, value: { applied: ["order"] } }; }, ' + DELETE_IMPL + '}; var EO_BRIDGE_HYDRATIONS = {'),
]
MARKERS = [
    '"chart.drawing.delete.write/v1": function', '{ id: "chart.drawing.delete.write/v1"',
]


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v12 = V12_BLOCK.read_bytes()
    if sha(v12) != V12_BLOCK_SHA or len(v12) != V12_BLOCK_SIZE:
        sys.exit(f"REFUSE: /12 block sha/size {sha(v12)}/{len(v12)} != {V12_BLOCK_SHA}/{V12_BLOCK_SIZE}; NO OUTPUT")
    t = v12.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("(X-O07/2, /13)") != 1:
        sys.exit(f"REFUSE: (X-O07/2, /13) count {t.count('(X-O07/2, /13)')} != 1; NO OUTPUT")
    return v12, t.encode("utf-8")


def main():
    v12, v13 = build_block()
    out = {"v12_block_sha": sha(v12), "v12_block_size": len(v12), "v13_block_sha": sha(v13), "v13_block_size": len(v13)}
    (HERE / "bridge-block.js").write_bytes(v13)
    (HERE / "old.txt").write_bytes(v12)
    (HERE / "new.txt").write_bytes(v13)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v13-x-o07-2-chart-drawing-delete-write", "old": v12.decode("utf-8"), "new": v13.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- a served-sdk-backed manifest.json for the real production apply is a SEPARATE, later "
                              "step, gated to the designated operator role, with the real served file in hand.",
                 original_sha256=sha(v12), patched_sha256=sha(v13), original_size=len(v12), patched_size=len(v13))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
