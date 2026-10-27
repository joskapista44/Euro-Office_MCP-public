#!/usr/bin/env python3
"""
X-O08 (zero LIVE #ref, driver-side wiring) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /14 CANDIDATE, in ONE deterministic pass from the pinned /13
block (the single-entry-point convention build-rev12.py's own docstring records, repeated by
build-rev13.py -- a split build/write-stage script was found FRAGILE and is not repeated here).

/14 = /13 + ONE new capability:
  READ (X-O08, shape fill/line colour):
    drawing.fillLineColor.read/v1  -- a shape drawing's own fill and line colour

X-O08's LIVE belso-modell DISCOVER (#ref/#ref/#ref/#ref, zero reviewed and MEASURED,
in-session AND after a fresh reopen, three shapes/three distinct colours plus a no-stroke and a
no-fill edge) found the PUBLIC GetFill()/GetLine().GetFill() wrappers expose the fill TYPE only
("solid"/"nofill", X-O08-7 WRITE_ONLY finding, confirmed again by two independent LIVE rounds),
never the RGB byte value -- the internal chain the handler below reads, ONLY when the public type
says "solid", is the public wrapper's own `.Drawing` field (the SAME access `drawing.anchor.read/v1`
/`drawing.text.read/v1` already use) then `.spPr[.ln].Fill.fill.color.color.RGBA`.

This handler never guesses: the TYPE comes from the public getter first (same style as every other
migrated axis in this block), and the internal RGBA chain is read only to fill in what the public
getter cannot see, never to override or second-guess it. A field value is one of: a 6-hex-digit RGB
string, the named reason `noFill`/`noLine` (the public type already says there is no fill/line, so
there is nothing to chain into), a named `notMeasured:<type>` for any non-solid fill type
(gradient/pattern/picture -- X-O08 never measured those, so never guessed), or an
`UNOBSERVABLE:<reason>` chain-break reason if the public type says solid but the internal chain
still breaks (a genuine anomaly, not swallowed). Unlike chart.axis.majorGridlinesColor.read/v1
(whose two axes fail the WHOLE call together), fill and line here are independent fields on an
otherwise-successful call -- a shape with a solid fill and no stroke is one valid, fully-observed
answer, not a partial failure.

The drawing-lookup prefix (`findDrawing`, `argsOk`) is the SAME convention `drawing.anchor.read/v1`
and `drawing.text.read/v1` already use in this block -- no new lookup mechanism.

Deterministic: the /14 block is derived from the CANONICAL /13 block
(../rev13-candidate/bridge-block.js, sha pinned) by THREE exact string substitutions, each
asserted to occur exactly once. No timestamps, no environment input.

Usage:
  build-rev14.py                         -> writes bridge-block.js, old.txt, new.txt,
                                            manifest-mechanism-proof.json next to this file
This script NEVER writes a Document Server file, and has NO --sdk mode: the real served /13
sdk-all.js is a host artifact this agent (designated-uid, no an owner-approval gate access) cannot fetch
or verify -- producing a manifest.json that CLAIMS an SDK sha without having read that file would
be a fabricated measurement, not a build. The SDK-backed manifest.json (and the apply itself) is
a separate, later, gated step -- see README.md.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V13_BLOCK = HERE.parent / "rev13-candidate" / "bridge-block.js"
V13_BLOCK_SHA = "4d218a2250643e2bb5a8704822f8509b02c4946d034df084fbfadb38091e750a"
V13_BLOCK_SIZE = 73986

FILL_LINE_COLOR_REASONS = ("bad-args, sheet-not-found, no-GetAllDrawings, drawings-threw, "
                           "drawings-not-a-list, drawing-not-found")

FILL_LINE_COLOR_ENTRY = (
    '{ id: "drawing.fillLineColor.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one shape drawing (sheet, drawing name): the shape\'s own fill and line colour '
    '-- X-O08 LIVE VERIFIED session+reopen on three shapes/three distinct colours plus a no-stroke and '
    'a no-fill edge, reading d.Drawing.spPr[.ln].Fill.fill.color.color.RGBA, the SAME `.Drawing` access '
    '`drawing.anchor.read/v1`/`drawing.text.read/v1` already use. {{ fill, line }}, each one EITHER a '
    '6-hex-digit RGB string, the named reason noFill/noLine (the public GetFill()/GetLine().GetFill() '
    'type already says there is none), a named notMeasured:<type> for any non-solid fill type '
    '(gradient/pattern/picture -- never guessed, X-O08 never measured those), OR an UNOBSERVABLE '
    'chain-break reason if the public type says solid but the internal chain still breaks (a genuine '
    'anomaly, not swallowed) -- unlike chart.axis.majorGridlinesColor.read/v1, fill and line are '
    'INDEPENDENT fields on one otherwise-successful call, never a whole-call failure. The PUBLIC '
    'GetFill()/GetLine().GetFill() give the TYPE only (X-O08-7 WRITE_ONLY finding, confirmed by two '
    'independent LIVE rounds), never the RGB byte value -- read through the public getter FIRST, never '
    'guessed; the internal chain is read only when that type says solid. reasons: ' + FILL_LINE_COLOR_REASONS + ' (X-O08, /14)" }, '
)

FILL_LINE_COLOR_IMPL = (
    '"drawing.fillLineColor.read/v1": function (args) { '
    'if (!argsOk(args, ["sheet", "drawing"])) return { ok: false, reason: "bad-args" }; '
    'var r = findDrawing(args.sheet, args.drawing); if (r.err) return r.err; '
    'var d = r.drawing; '
    'var rgbaHexOf = function (uniFill) { '
    'var rgba = uniFill && uniFill.fill && uniFill.fill.color && uniFill.fill.color.color && uniFill.fill.color.color.RGBA; '
    'if (!rgba || typeof rgba.R !== "number" || typeof rgba.G !== "number" || typeof rgba.B !== "number") return void 0; '
    'var h = function (n) { return ("0" + n.toString(16).toUpperCase()).slice(-2); }; '
    'return h(rgba.R) + h(rgba.G) + h(rgba.B); }; '
    'var valueOf = function (getPublicFill, internalFillGetter, noneReason) { '
    'if (typeof getPublicFill !== "function") return "UNOBSERVABLE:getter-missing"; '
    'var fillApi; try { fillApi = getPublicFill(); } catch (e) { return "UNOBSERVABLE:getter-threw"; } '
    'if (!fillApi || typeof fillApi.GetType !== "function") return "UNOBSERVABLE:type-getter-missing"; '
    'var type; try { type = fillApi.GetType(); } catch (e) { return "UNOBSERVABLE:type-getter-threw"; } '
    'if (type === "nofill") return noneReason; '
    'if (type !== "solid") return "notMeasured:" + String(type); '
    'var internalFill; try { internalFill = internalFillGetter(); } catch (e) { return "UNOBSERVABLE:internal-chain-threw"; } '
    'var hex = rgbaHexOf(internalFill); return hex === void 0 ? "UNOBSERVABLE:internal-chain-broken" : hex; }; '
    'var fill = valueOf( '
    'typeof d.GetFill === "function" ? function () { return d.GetFill(); } : void 0, '
    'function () { return d.Drawing && d.Drawing.spPr && d.Drawing.spPr.Fill; }, '
    '"noFill"); '
    'var line = valueOf( '
    '(function () { '
    'var lineApi; try { lineApi = typeof d.GetLine === "function" ? d.GetLine() : void 0; } catch (e) { lineApi = void 0; } '
    'return lineApi && typeof lineApi.GetFill === "function" ? function () { return lineApi.GetFill(); } : void 0; '
    '})(), '
    'function () { return d.Drawing && d.Drawing.spPr && d.Drawing.spPr.ln && d.Drawing.spPr.ln.Fill; }, '
    '"noLine"); '
    'return { ok: true, value: { fill: fill, line: line } }; }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the EVOLVING text at the time it is applied.
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/13"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/14"'),
    ('(X-O07/2, /13)" }, ], };', '(X-O07/2, /13)" }, ' + FILL_LINE_COLOR_ENTRY + '], };'),
    # the EO_BRIDGE_CAPS (the READ dispatch object)'s closing anchor: the LAST read handler already
    # in the /13 block (chart.bar.gapOverlap.read/v1's own gapWidth/overlap return) ends with this
    # exact text, immediately before "var EO_BRIDGE_WRITES = {" -- same insertion-point convention
    # build-rev13.py's own docstring records (the dispatch object literal's closing '};' is the
    # ONLY safe anchor), here on the READ object rather than the WRITE object.
    ('return { ok: true, value: { gapWidth: gw, overlap: ov } }; }, }; var EO_BRIDGE_WRITES = {',
     'return { ok: true, value: { gapWidth: gw, overlap: ov } }; }, ' + FILL_LINE_COLOR_IMPL + '}; var EO_BRIDGE_WRITES = {'),
]
MARKERS = [
    '"drawing.fillLineColor.read/v1": function', '{ id: "drawing.fillLineColor.read/v1"',
]


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v13 = V13_BLOCK.read_bytes()
    if sha(v13) != V13_BLOCK_SHA or len(v13) != V13_BLOCK_SIZE:
        sys.exit(f"REFUSE: /13 block sha/size {sha(v13)}/{len(v13)} != {V13_BLOCK_SHA}/{V13_BLOCK_SIZE}; NO OUTPUT")
    t = v13.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("(X-O08, /14)") != 1:
        sys.exit(f"REFUSE: (X-O08, /14) count {t.count('(X-O08, /14)')} != 1; NO OUTPUT")
    return v13, t.encode("utf-8")


def main():
    v13, v14 = build_block()
    out = {"v13_block_sha": sha(v13), "v13_block_size": len(v13), "v14_block_sha": sha(v14), "v14_block_size": len(v14)}
    (HERE / "bridge-block.js").write_bytes(v14)
    (HERE / "old.txt").write_bytes(v13)
    (HERE / "new.txt").write_bytes(v14)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v14-x-o08-drawing-fill-line-color-read", "old": v13.decode("utf-8"), "new": v14.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- a served-sdk-backed manifest.json for the real production apply is a SEPARATE, later "
                              "step, gated to the designated operator role, with the real served file in hand.",
                 original_sha256=sha(v13), patched_sha256=sha(v14), original_size=len(v13), patched_size=len(v14))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
