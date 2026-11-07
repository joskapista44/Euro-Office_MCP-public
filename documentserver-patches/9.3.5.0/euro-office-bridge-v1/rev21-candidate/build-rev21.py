#!/usr/bin/env python3
"""
GH #354 / #355 (Bridge /21 WRITE, DESIGN docs/DS-9.3.5-BRIDGE20-PAGE-SETUP-AND-READS-DESIGN.md, gate G0 =) -- OFFLINE build
of the euro-office-bridge-v1 patchVersion /21 WRITE CANDIDATE for DS 9.3.5, in ONE deterministic pass from the /20 READ candidate block
(../rev20-candidate/bridge-block.js, sha pinned).
/21 = /20 + two things:
  1. a LOAD-TIME installer (installer-v21.txt, an expression placed BEFORE `Api.GetEuroOfficeBridge = ...` in the same comma sequence):
     AscCH.historyitem_Layout_HeaderMargin / FooterMargin / PageOrder = 16 / 17 / 18, and AscCommonExcel.UndoRedoRedoLayout.prototype.UndoRedo
     wrapped for exactly those three types (every other type goes to the ORIGINAL function unchanged). The factory Layout class knows only
     types 1-15 and its switch has no default: a 16-18 change on an UNPATCHED applier is a SILENT no-op (design, "The engine finding").
     It runs when sdk-all.js loads -- in the editor AND in the server-side converter, IF the converter loads this file (G0: derived, not
     behaviourally proven; gate G1 proves it).
  2. sheet.printSettings.write/v2 (handler-write-v21.txt): ONE of headerMargin | footerMargin (mm, 0..500 -- a product-side bound; the
     engine's own bound is NOT MEASURED) | pageOrder (downThenOver | overThenDown -> ST_PageOrder 0 | 1) per call; the target value already
     held is SATISFIED with no history; otherwise one history point + transaction: set, then History.Add(g_oUndoRedoLayout, 16|17|18,
     sheetId, null, UndoRedoData_Layout(old, new)) -- the factory asc_setOrientation pattern; read-after-write, an unchanged value is no-effect.
Same pattern as build-rev20.py: exact substitutions, each asserted once, no timestamps, no environment input, NEVER writes a Document
Server file, no --sdk mode.
NOT MEASURED: everything on the real engine (G1-G3): that the installer runs in the converter, that a 16-18 change survives co-editing,
save and a fresh reopen, the mixed-version and rollback behaviour.
Usage:
  build-rev21.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""

import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V20_BLOCK = HERE.parent / "rev20-candidate" / "bridge-block.js"
V20_BLOCK_SHA = "c9d46ada6e7e47992502c2917ae4d90dcd3176bc424686d265e494e4e4fd014a"
V20_BLOCK_SIZE = 134465
INSTALLER_SRC = HERE / "installer-v21.txt"
HANDLER_SRC = HERE / "handler-write-v21.txt"
WRITE_REASONS = ("bad-args, unknown-setting, invalid-value, sheet-not-found, model-missing, layout-history-missing, layout-types-not-installed, "
                 "history-missing, history-off, asc_getHeader-missing, asc_getFooter-missing, asc_setHeader-missing, asc_setFooter-missing, "
                 "setter-threw, history-add-threw, no-effect")
ENTRY = (
    '{ id: "sheet.printSettings.write/v2", version: 2, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one sheet (by name), ONE of headerMargin | footerMargin {value: mm 0..500} | pageOrder {value: downThenOver | overThenDown} '
    'per call -- the three settings the factory Layout history cannot carry; this block installs history change types 16 / 17 / 18 at load '
    '(AscCH.historyitem_Layout_HeaderMargin / FooterMargin / PageOrder, UndoRedoRedoLayout.UndoRedo wrapped for exactly those types); a target '
    'value already held is {applied: [], satisfied: true} with NO history; otherwise one history point + transaction: the value is set and '
    'History.Add(g_oUndoRedoLayout, type, sheetId, null, UndoRedoData_Layout(old, new)) is recorded, a setter that succeeded before a failed Add '
    'is restored; read-after-write, an unchanged value is no-effect; the result is a receipt {applied: [setting]}, the caller verifies with '
    'sheet.printSettings.read/v2; v1 stays served; served only through write(); reasons: ' + WRITE_REASONS + ' (GH #354, #355, /21)" }, '
)
MANIFEST_END = '(GH #346, /20)" }, ], };'
WRITES_START = 'var EO_BRIDGE_WRITES = { '
BLOCK_START = 'Api.GetEuroOfficeBridge = function () {'
MARKERS = ['"sheet.printSettings.write/v2": function', '{ id: "sheet.printSettings.write/v2"', 'P.__eoLayoutV21 = true;']
FORBIDDEN = ("Object.keys", "Object.values", "Object.entries", "getOwnPropertyNames", "for (var k in", "for (var p in", " in args", "//")


def sha(b):
    return hashlib.sha256(b).hexdigest()


def one_line(path):
    lines = [ln.strip() for ln in path.read_text(encoding="utf-8").split("\n")]
    text = " ".join(ln for ln in lines if ln) + " "
    for word in FORBIDDEN:
        if word in text:
            sys.exit(f"REFUSE: {path.name} contains {word!r} -- named fields only, no line comments; NO OUTPUT")
    return text


def build_block():
    v20 = V20_BLOCK.read_bytes()
    if sha(v20) != V20_BLOCK_SHA or len(v20) != V20_BLOCK_SIZE:
        sys.exit(f"REFUSE: /20 block sha/size {sha(v20)}/{len(v20)} != {V20_BLOCK_SHA}/{V20_BLOCK_SIZE}; NO OUTPUT")
    installer, impl = one_line(INSTALLER_SRC), one_line(HANDLER_SRC)
    subs = [
        ('patchVersion: "euro-office-bridge-v1/9.3.5.0/20"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/21"'),
        (MANIFEST_END, '(GH #346, /20)" }, ' + ENTRY + '], };'),
        (WRITES_START, WRITES_START + impl),
    ]
    t = v20.decode("utf-8")
    if not t.startswith(BLOCK_START):
        sys.exit("REFUSE: the /20 block does not start with the bridge entrypoint; NO OUTPUT")
    for old, new in subs:
        n = t.count(old)
        if n != 1:
            sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    t = installer + t
    for m in MARKERS:
        if t.count(m) != 1:
            sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v20, t.encode("utf-8")


def main():
    v20, v21 = build_block()
    out = {"v20_block_sha": sha(v20), "v20_block_size": len(v20), "v21_block_sha": sha(v21), "v21_block_size": len(v21)}
    (HERE / "bridge-block.js").write_bytes(v21)
    (HERE / "old.txt").write_bytes(v20)
    (HERE / "new.txt").write_bytes(v21)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v21-gh354-gh355-layout-history-write", "old": v20.decode("utf-8"), "new": v21.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (maintainer; the disposable "
                              "DS first, the live DS needs its own owner GO).",
                 original_sha256=sha(v20), patched_sha256=sha(v21), original_size=len(v20), patched_size=len(v21))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
