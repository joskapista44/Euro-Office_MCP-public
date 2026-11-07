#!/usr/bin/env python3
"""
GH #356 / #357 (Bridge /23 READ, DESIGN docs/DS-9.3.5-BRIDGE23-PRINT-PAGE-COUNT-READ-DESIGN.md) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /23 READ CANDIDATE for DS 9.3.5, in ONE deterministic pass from the /22 candidate block
(../rev22-candidate/bridge-block.js, sha pinned -- the block installed on the production DS 2026-10-31 21:44).
/23 = /22 + one READ capability (handler-read-v23.txt), NO installer, NO write:
  sheet.printPageCount.read/v1: { sheet } -> { pages } -- the editor's own rendered print page count of ONE sheet under its stored
  print settings (WorkbookView.calcPagesPrint, ActiveSheets [index]). The live model is never written: the sheet's PagePrintOptions
  is swapped for its clone and History is off for the call (the first calcPagesPrint on a fit-to-page sheet otherwise rewrites
  `scale` and adds a History point -- measured, gh356 M1/M2 r1); everything is restored in finally. A hidden sheet, a count at the
  engine's c_kMaxPrintPages cap and every missing engine piece are named refusals -- never a 0 or a guess.
Same pattern as build-rev22.py: exact substitutions, each asserted once, no timestamps, no environment input, NEVER writes a
Document Server file, no --sdk mode.
MEASURED BEFORE THE PATCH (the same logic in the page context, production pair, docs/evidence/gh357-clone-pagecount-probe-live-r1-20261101.json):
  no History point, the real scale / options identity / zoom / active sheet kept, the stored pageSetup unchanged, counts 1 / 9 = a
  direct calcPagesPrint. NOT MEASURED: the block itself on the engine (design gates G1-G3).
Usage:
  build-rev23.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""

import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V22_BLOCK = HERE.parent / "rev22-candidate" / "bridge-block.js"
V22_BLOCK_SHA = "10d59f5d4bd5e7effd61d4c25f2e2bd047a995e9f31c192b3045f3358e1fde8a"
V22_BLOCK_SIZE = 150351
HANDLER_SRC = HERE / "handler-read-v23.txt"
R_REASONS = ("bad-args, sheet-not-found, model-missing, sheet-hidden, print-calc-unavailable, history-missing, options-clone-threw, "
             "print-calc-threw, print-calc-unexpected-shape, page-count-capped")
ENTRIES = (
    '{ id: "sheet.printPageCount.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one sheet (by name): {pages} -- the integer number of pages the editor lays the sheet out on for printing '
    'under its OWN stored print settings (print area, fit, scale, margins, breaks, paper, orientation), the editor\'s own '
    'WorkbookView.calcPagesPrint for that sheet alone (ActiveSheets); the active sheet is not switched; the live model is never '
    'written: the sheet\'s print options are swapped for their clone and History is off for the call, then restored (the fit scale '
    'the engine computes stays on the clone); 0 is a legitimate count; a hidden sheet is refused (sheet-hidden), never 0; a count at '
    'the engine\'s c_kMaxPrintPages cap is refused (page-count-capped), never truncated; served only through read(); reasons: '
    + R_REASONS + ' (GH #356, #357, /23)" }, '
)
MANIFEST_END = '(GH #340, #341, /22)" }, ], };'
CAPS_START = 'var EO_BRIDGE_CAPS = { '
BLOCK_ENTRY = 'Api.GetEuroOfficeBridge = function () {'
MARKERS = ['"sheet.printPageCount.read/v1": function', '{ id: "sheet.printPageCount.read/v1"']
FORBIDDEN = ("Object.keys", "Object.values", "Object.entries", "getOwnPropertyNames", "for (var k in", "for (var p in", " in args", "//",
             "hashValue", "saltValue", "spinCount", "algorithmName", "password", "temporaryPassword", "PasswordXL", "asc_setSheet",
             "asc_setPassword")


def sha(b):
    return hashlib.sha256(b).hexdigest()


def one_line(path):
    lines = [ln.strip() for ln in path.read_text(encoding="utf-8").split("\n")]
    text = " ".join(ln for ln in lines if ln) + " "
    for word in FORBIDDEN:
        if word in text:
            sys.exit(f"REFUSE: {path.name} contains {word!r} -- named fields only, no credential field, no line comments; NO OUTPUT")
    return text


def build_block():
    v22 = V22_BLOCK.read_bytes()
    if sha(v22) != V22_BLOCK_SHA or len(v22) != V22_BLOCK_SIZE:
        sys.exit(f"REFUSE: /22 block sha/size {sha(v22)}/{len(v22)} != {V22_BLOCK_SHA}/{V22_BLOCK_SIZE}; NO OUTPUT")
    impl = one_line(HANDLER_SRC)
    subs = [
        ('patchVersion: "euro-office-bridge-v1/9.3.5.0/22"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/23"'),
        (MANIFEST_END, '(GH #340, #341, /22)" }, ' + ENTRIES + '], };'),
        (CAPS_START, CAPS_START + impl),
    ]
    t = v22.decode("utf-8")
    if t.count(BLOCK_ENTRY) != 1:
        sys.exit("REFUSE: the /22 block does not carry the bridge entrypoint exactly once; NO OUTPUT")
    for old, new in subs:
        n = t.count(old)
        if n != 1:
            sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1:
            sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v22, t.encode("utf-8")


def main():
    v22, v23 = build_block()
    out = {"v22_block_sha": sha(v22), "v22_block_size": len(v22), "v23_block_sha": sha(v23), "v23_block_size": len(v23)}
    (HERE / "bridge-block.js").write_bytes(v23)
    (HERE / "old.txt").write_bytes(v22)
    (HERE / "new.txt").write_bytes(v23)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v23-gh356-gh357-print-page-count-read", "old": v22.decode("utf-8"), "new": v23.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (maintainer; the disposable "
                              "DS first, the live DS needs its own owner GO).",
                 original_sha256=sha(v22), patched_sha256=sha(v23), original_size=len(v22), patched_size=len(v23))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
