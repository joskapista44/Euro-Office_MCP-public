#!/usr/bin/env python3
"""
GH #338-#342 (Bridge /22 WRITE, DESIGN docs/DS-9.3.5-BRIDGE22-SHEET-PROTECTION-WRITE-DESIGN.md) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /22 WRITE CANDIDATE for DS 9.3.5, in ONE deterministic pass from the /21 candidate block
(../rev21-candidate/bridge-block.js, sha pinned -- the block installed on the production DS 2026-10-31).
/22 = /21 + two WRITE capabilities (handler-write-v22.txt), NO installer (the factory engine already records every protection
field -- historyitem_Protected_Set* -- and the cell locked / hidden change -- historyitem_Cell_SetLocked / SetHidden):
  1. sheet.protection.write/v1: { sheet, protect: true, forbids?: {<flag>: boolean, ...} } | { sheet, protect: false } --
     PASSWORDLESS protect / unprotect through the engine model ws.setProtectedSheet(props, true); props from setDefaultInterface()
     + setSheet + the 15 plain flag setters (never asc_setSheet, whose first argument is a password). `forbids` uses the exact
     sheet.protection.read/v1 keys and meaning (true = the protection forbids it), keys in the read order. A sheet whose current
     protection carries a credential (the engine's own asc_isPassword()) is refused: credential-protected.
  2. sheet.cellProtection.write/v1: { sheet, range, locked?, hidden? } on an UNPROTECTED sheet, <= 10000 cells, no whole
     row / column range -- Range.setLocked / setHiddenFormulas, read back per cell through getCell3().getXfs().
Both: a target state already held is {applied: [], satisfied: true} with NO history; otherwise one history point + transaction,
read-after-write, an unchanged state is no-effect. The handler text names no credential field and enumerates nothing (the /17 rule).
Same pattern as build-rev21.py: exact substitutions, each asserted once, no timestamps, no environment input, NEVER writes a
Document Server file, no --sdk mode.
NOT MEASURED: everything on the real engine (design gates G1-G3) -- above all that the protected-sheet change reaches the
server-side save (its applier is guarded by window.Asc.editor.wb).
Usage:
  build-rev22.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""

import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V21_BLOCK = HERE.parent / "rev21-candidate" / "bridge-block.js"
V21_BLOCK_SHA = "04e39b24f7441ae1d6408f442e8d2552f973811dcf487c393d60e668cc358e0e"
V21_BLOCK_SIZE = 140668
HANDLER_SRC = HERE / "handler-write-v22.txt"
P_REASONS = ("bad-args, flag-not-boolean:<flag>, forbids-not-canonical, sheet-not-found, model-missing, protection-api-missing, "
             "history-missing, history-off, credential-protected, protection-unreadable, setter-threw, no-effect")
C_REASONS = ("bad-args, invalid-value, range-shape-not-supported, sheet-not-found, model-missing, sheet-protected, range-invalid, "
             "range-over-cap:10000, history-missing, history-off, cells-walk-threw, setter-threw, no-effect")
ENTRIES = (
    '{ id: "sheet.protection.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one sheet (by name): {protect: true, forbids?: {flag: boolean}} protects it PASSWORDLESS, each named flag of the '
    '15 sheet.protection.read/v1 keys (read order, same meaning: true = forbidden) set, every other flag the Excel default; {protect: false} '
    'unprotects it; a sheet whose current protection carries a credential is refused, never changed; props from setDefaultInterface + '
    'setSheet + the plain flag setters, ws.setProtectedSheet(props, true) in one history point + transaction; a target state already held '
    'is {applied: [], satisfied: true} with NO history; read-after-write, an unchanged state is no-effect; the result is a receipt '
    '{applied: [protect | unprotect]}, the caller verifies with sheet.protection.read/v1; served only through write(); reasons: '
    + P_REASONS + ' (GH #338, #339, #342, /22)" }, '
    '{ id: "sheet.cellProtection.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one A1 range (at most 10000 cells, no whole row / column) of an UNPROTECTED sheet: locked and / or hidden (the '
    'cell xf protection) set through Range.setLocked / setHiddenFormulas in one history point + transaction; a target state already held '
    'is {applied: [], satisfied: true} with NO history; read-after-write per cell (an empty cell reads as the Excel default: locked, not '
    'hidden), any cell not holding it is no-effect; the result is a receipt {applied: [locked | hidden]}, the caller verifies with '
    'sheet.protection.read/v1 (range); served only through write(); reasons: ' + C_REASONS + ' (GH #340, #341, /22)" }, '
)
MANIFEST_END = '(GH #354, #355, /21)" }, ], };'
WRITES_START = 'var EO_BRIDGE_WRITES = { '
BLOCK_ENTRY = 'Api.GetEuroOfficeBridge = function () {'
MARKERS = ['"sheet.protection.write/v1": function', '"sheet.cellProtection.write/v1": function',
           '{ id: "sheet.protection.write/v1"', '{ id: "sheet.cellProtection.write/v1"']
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
    v21 = V21_BLOCK.read_bytes()
    if sha(v21) != V21_BLOCK_SHA or len(v21) != V21_BLOCK_SIZE:
        sys.exit(f"REFUSE: /21 block sha/size {sha(v21)}/{len(v21)} != {V21_BLOCK_SHA}/{V21_BLOCK_SIZE}; NO OUTPUT")
    impl = one_line(HANDLER_SRC)
    subs = [
        ('patchVersion: "euro-office-bridge-v1/9.3.5.0/21"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/22"'),
        (MANIFEST_END, '(GH #354, #355, /21)" }, ' + ENTRIES + '], };'),
        (WRITES_START, WRITES_START + impl),
    ]
    t = v21.decode("utf-8")
    if t.count(BLOCK_ENTRY) != 1:
        sys.exit("REFUSE: the /21 block does not carry the bridge entrypoint exactly once; NO OUTPUT")
    for old, new in subs:
        n = t.count(old)
        if n != 1:
            sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1:
            sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v21, t.encode("utf-8")


def main():
    v21, v22 = build_block()
    out = {"v21_block_sha": sha(v21), "v21_block_size": len(v21), "v22_block_sha": sha(v22), "v22_block_size": len(v22)}
    (HERE / "bridge-block.js").write_bytes(v22)
    (HERE / "old.txt").write_bytes(v21)
    (HERE / "new.txt").write_bytes(v22)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v22-gh338-gh342-protection-write", "old": v21.decode("utf-8"), "new": v22.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (maintainer; the disposable "
                              "DS first, the live DS needs its own owner GO).",
                 original_sha256=sha(v21), patched_sha256=sha(v22), original_size=len(v21), patched_size=len(v22))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
