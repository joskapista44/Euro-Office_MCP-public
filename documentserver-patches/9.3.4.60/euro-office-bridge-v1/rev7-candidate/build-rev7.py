#!/usr/bin/env python3
"""
M11.5b.1 (owner: M11.5 continuous mandate, owner channel, internal tracker) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /7 CANDIDATE.

/7 = /6 + exactly ONE new capability, sheet.printSettings.write/v1 (kind WRITE, mutation TARGETED), served by the
existing write() entry. PUBLIC_FIRST found no public setter for any of the eight B10 print/page settings (the 9.3.4
ApiWorksheet has only margins / orientation / gridlines / headings). Every /6 capability (manifest entry and
implementation) stays byte-identical.

The WRITE calls the SAME engine setters, in the SAME order and with the SAME history transaction, as the product's B10
direct-internal path (src/adapters/xlsx/families/print.cjs), one setting per call. It validates every input and checks
every setter the setting needs BEFORE the first mutation; the result is an operation receipt {applied:[...]} named by
the sheet.printSettings.read/v1 fields, NOT a verification; a later setter that throws after an earlier one succeeded
is reported as {partial:true, applied:[...]}, never rolled back; no fallback of any kind.

Deterministic: the /7 block is derived from the CANONICAL /6 block (../rev6-candidate/bridge-block.js, sha pinned) by
three exact string substitutions, each asserted to occur exactly once. No timestamps, no environment input.

Usage:
  build-rev7.py                          -> writes bridge-block.js, old.txt, new.txt next to this file
  build-rev7.py --sdk <served /6 sdk-all.js>
                                         -> also writes manifest.json (whole-block replace, /6 output sha -> /7 sha)
                                            and manifest-mechanism-proof.json; the sdk file must be the /6 output
This script never writes a Document Server file.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V6_BLOCK = HERE.parent / "rev6-candidate" / "bridge-block.js"
V6_BLOCK_SHA = "4faf6031ab62e715880e995ae184759638e95c93419051f2a807dd303ef29325"   # = /6 new.txt
V6_SDK_SHA = "af42c2327de11bd2537f7532187de33b13811aef67949d4a6b2f7c2ff3916b63"     # = /6 served (production)
V6_SDK_SIZE = 20955549

CAP = "sheet.printSettings.write/v1"
REASONS = ("bad-args, unsupported-property, unknown-setting, invalid-value, sheet-not-found, model-missing, "
           "history-missing, <setter>-missing, <setter>-threw")

WRITE_ENTRY = (
    '{ id: "sheet.printSettings.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one sheet (by name), ONE print/page setting per call: fit {fitToWidth, fitToHeight 0..100; '
    '0+0 = fit off, scale 100} | scale {scale 10..400, fit off} | paper {width, height: one of the editor\'s paper '
    'sizes in mm} | firstPage {firstPage 1..32767 or null} | center {axis horizontal|vertical, value boolean} | '
    'headerFooter {slot odd/first/even Header/Footer, value 1..255 chars or null} | headerFooterFlag {flag '
    'differentFirst|differentOddEven, value boolean} | pageBreak {axis row|column, id, remove boolean; manual break}; '
    'through the engine\'s own print setters in the product\'s order (header/footer inside one history transaction); '
    'every input is validated and every needed setter is checked BEFORE the first mutation; an unrequested setting is '
    'never touched; the result is an operation receipt {applied:[read fields]} and NOT a verification -- the caller '
    'verifies with a fresh sheet.printSettings.read/v1; a setter that throws after an earlier one succeeded leaves a '
    'PARTIAL write, reported as {partial:true, applied:[...]}, never rolled back; no fallback; served only through '
    'write(); reasons: ' + REASONS + ' (B10-WRITE, /7)" }, '
)

# One line, ES5, no arrow functions -- the same style as every other bridge entry.
WRITE_IMPL = (
    '"sheet.printSettings.write/v1": function (args) { '
    'var own = function (o, k) { return Object.prototype.hasOwnProperty.call(o, k); }; '
    'var isInt = function (v, lo, hi) { return typeof v === "number" && Math.floor(v) === v && v >= lo && v <= hi; }; '
    'if (!args || typeof args !== "object" || typeof args.sheet !== "string" || args.sheet === "" || typeof args.setting !== "string") return { ok: false, reason: "bad-args" }; '
    'var FIELDS = { fit: ["fitToWidth", "fitToHeight"], scale: ["scale"], paper: ["width", "height"], firstPage: ["firstPage"], center: ["axis", "value"], headerFooter: ["slot", "value"], headerFooterFlag: ["flag", "value"], pageBreak: ["axis", "id", "remove"] }; '
    'if (!own(FIELDS, args.setting)) return { ok: false, reason: "unknown-setting" }; '
    'var want = FIELDS[args.setting]; var keys = Object.keys(args); '
    'for (var i = 0; i < keys.length; i += 1) { if (keys[i] !== "sheet" && keys[i] !== "setting" && want.indexOf(keys[i]) < 0) return { ok: false, reason: "unsupported-property" }; } '
    'for (var j = 0; j < want.length; j += 1) { if (!own(args, want[j])) return { ok: false, reason: "bad-args" }; } '
    'var PAPER = [[215.9, 279.4], [215.9, 355.6], [210, 297], [148, 210], [176, 250], [104.8, 241.3], [110, 220], [279.4, 431.8], [297, 420], [296.9, 457.2], [196.8, 273], [120, 235], [305, 487]]; '
    'var SLOTS = ["oddHeader", "oddFooter", "firstHeader", "firstFooter", "evenHeader", "evenFooter"]; '
    'var s = args.setting; var a = args; var valid; '
    'if (s === "fit") valid = isInt(a.fitToWidth, 0, 100) && isInt(a.fitToHeight, 0, 100); '
    'else if (s === "scale") valid = isInt(a.scale, 10, 400); '
    'else if (s === "paper") { valid = false; for (var p = 0; p < PAPER.length; p += 1) { if (a.width === PAPER[p][0] && a.height === PAPER[p][1]) valid = true; } } '
    'else if (s === "firstPage") valid = a.firstPage === null || isInt(a.firstPage, 1, 32767); '
    'else if (s === "center") valid = (a.axis === "horizontal" || a.axis === "vertical") && typeof a.value === "boolean"; '
    'else if (s === "headerFooter") valid = SLOTS.indexOf(a.slot) >= 0 && (a.value === null || (typeof a.value === "string" && a.value.length >= 1 && a.value.length <= 255)); '
    'else if (s === "headerFooterFlag") valid = (a.flag === "differentFirst" || a.flag === "differentOddEven") && typeof a.value === "boolean"; '
    'else valid = (a.axis === "row" ? isInt(a.id, 1, 1048575) : a.axis === "column" ? isInt(a.id, 1, 16383) : false) && typeof a.remove === "boolean"; '
    'if (!valid) return { ok: false, reason: "invalid-value" }; '
    'var ws = Api.GetSheet(a.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var m; try { m = ws.worksheet; } catch (e) { return { ok: false, reason: "model-missing" }; } '
    'if (!m || !m.PagePrintOptions) return { ok: false, reason: "model-missing" }; '
    'var ppo = m.PagePrintOptions; if (typeof ppo.asc_getPageSetup !== "function") return { ok: false, reason: "model-missing" }; '
    'var ps = ppo.asc_getPageSetup(); if (!ps) return { ok: false, reason: "model-missing" }; '
    'var hf = m.headerFooter; var H = typeof AscCommon !== "undefined" && AscCommon ? AscCommon.History : null; '
    'var off = s === "fit" && a.fitToWidth === 0 && a.fitToHeight === 0; '
    'var steps; '
    'if (s === "fit") { steps = [[ps, "asc_setFitToWidth", [a.fitToWidth], "fitW"], [ps, "asc_setFitToHeight", [a.fitToHeight], "fitH"], [m, "setFitToPage", [!off], "fitFlag"]]; if (off) steps.push([ps, "asc_setScale", [100], "scale"]); } '
    'else if (s === "scale") steps = [[ps, "asc_setFitToWidth", [0], "fitW"], [ps, "asc_setFitToHeight", [0], "fitH"], [m, "setFitToPage", [false], "fitFlag"], [ps, "asc_setScale", [a.scale], "scale"]]; '
    'else if (s === "paper") steps = [[ps, "asc_setWidth", [a.width], "width"], [ps, "asc_setHeight", [a.height], "height"]]; '
    'else if (s === "firstPage") steps = [[ps, "asc_setFirstPageNumber", [a.firstPage], "firstPage"]]; '
    'else if (s === "center") steps = a.axis === "horizontal" ? [[ppo, "asc_setHorizontalCentered", [a.value], "hCenter"]] : [[ppo, "asc_setVerticalCentered", [a.value], "vCenter"]]; '
    'else if (s === "headerFooter") steps = [[hf, "set" + a.slot.charAt(0).toUpperCase() + a.slot.slice(1), [a.value], a.slot]]; '
    'else if (s === "headerFooterFlag") steps = a.flag === "differentFirst" ? [[hf, "setDifferentFirst", [a.value], "diffFirst"]] : [[hf, "setDifferentOddEven", [a.value], "diffOddEven"]]; '
    'else { if (typeof m.getPrintAreaRangeByRowCol !== "function") return { ok: false, reason: "getPrintAreaRangeByRowCol-missing" }; '
    'var col = a.axis === "column"; var rng; try { rng = col ? m.getPrintAreaRangeByRowCol(null, a.id) : m.getPrintAreaRangeByRowCol(a.id, null); } catch (e) { return { ok: false, reason: "getPrintAreaRangeByRowCol-threw", partial: false, applied: [] }; } '
    'steps = [[m, "changeRowColBreaks", a.remove ? [a.id, null, rng, col, true] : [null, a.id, rng, col, true], col ? "colBreaks" : "rowBreaks"]]; } '
    'for (var n = 0; n < steps.length; n += 1) { if (!steps[n][0] || typeof steps[n][0][steps[n][1]] !== "function") return { ok: false, reason: steps[n][1] + "-missing" }; } '
    'var tx = s === "headerFooter" || s === "headerFooterFlag"; '
    'if (tx && !(H && typeof H.Create_NewPoint === "function" && typeof H.StartTransaction === "function" && typeof H.EndTransaction === "function")) return { ok: false, reason: "history-missing" }; '
    'var applied = []; '
    'var run = function () { for (var q = 0; q < steps.length; q += 1) { var st = steps[q]; try { st[0][st[1]].apply(st[0], st[2]); } catch (e) { return { ok: false, reason: st[1] + "-threw", partial: applied.length > 0, applied: applied.slice() }; } applied.push(st[3]); } return null; }; '
    'var bad; if (tx) { H.Create_NewPoint(); H.StartTransaction(); try { bad = run(); } finally { H.EndTransaction(); } } else bad = run(); '
    'if (bad) return bad; return { ok: true, value: { applied: applied } }; }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the /6 block
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/6"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/7"'),
    ('(B05-WRITE, /6)" }, ], };', '(B05-WRITE, /6)" }, ' + WRITE_ENTRY + '], };'),
    ('return { ok: true, value: { applied: applied } }; }, }; var EO_BRIDGE_HYDRATIONS = {',
     'return { ok: true, value: { applied: applied } }; }, ' + WRITE_IMPL + '}; var EO_BRIDGE_HYDRATIONS = {'),
]
MARKERS = ['"sheet.printSettings.write/v1": function', '{ id: "sheet.printSettings.write/v1"', '(B10-WRITE, /7)']


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v6 = V6_BLOCK.read_bytes()
    if sha(v6) != V6_BLOCK_SHA: sys.exit(f"REFUSE: /6 block sha {sha(v6)} != {V6_BLOCK_SHA}; NO OUTPUT")
    t = v6.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v6, t.encode("utf-8")


def main():
    v6, v7 = build_block()
    out = {"v6_block_sha": sha(v6), "v6_block_size": len(v6), "v7_block_sha": sha(v7), "v7_block_size": len(v7)}
    sdk = None
    if "--sdk" in sys.argv:
        sdk = Path(sys.argv[sys.argv.index("--sdk") + 1]).read_bytes()
        if sha(sdk) != V6_SDK_SHA or len(sdk) != V6_SDK_SIZE:
            sys.exit(f"REFUSE: sdk input sha {sha(sdk)} / {len(sdk)} is not the /6 output; NO OUTPUT")
        text = sdk.decode("utf-8")
        old, new = v6.decode("utf-8"), v7.decode("utf-8")
        if text.count(old) != 1: sys.exit("REFUSE: /6 block not exactly once in the sdk input; NO OUTPUT")
        if any(text.count(m) for m in MARKERS): sys.exit("REFUSE: a /7 marker already present in the input; NO OUTPUT")
        patched = text.replace(old, new, 1).encode("utf-8")
        out.update({"v6_sdk_sha": sha(sdk), "v6_sdk_size": len(sdk), "v7_sdk_sha": sha(patched), "v7_sdk_size": len(patched)})
    (HERE / "bridge-block.js").write_bytes(v7)
    (HERE / "old.txt").write_bytes(v6)
    (HERE / "new.txt").write_bytes(v7)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v7-b10w-sheet-printsettings-write", "old": v6.decode("utf-8"), "new": v7.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt blocks, not the served bundle.",
                 original_sha256=sha(v6), patched_sha256=sha(v7), original_size=len(v6), patched_size=len(v7))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    if sdk is not None:
        man = dict(base, original_sha256=out["v6_sdk_sha"], patched_sha256=out["v7_sdk_sha"], original_size=out["v6_sdk_size"], patched_size=out["v7_sdk_size"],
                   _patched_sha256_NOTE="computed by build-rev7.py from the SERVED /6 cell sdk-all.js (sha = production /6); CANDIDATE until the M11.5b DS apply.")
        (HERE / "manifest.json").write_text(json.dumps(man, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
