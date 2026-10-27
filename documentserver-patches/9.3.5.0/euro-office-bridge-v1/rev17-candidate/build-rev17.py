#!/usr/bin/env python3
"""
GH #326 / #327 (owner GO, owner channel, answer to 41211: path A with conditions) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /17 CANDIDATE for DS 9.3.5, in ONE deterministic pass from the shipped
9.3.5.0 /16 block (../bridge-block.js, sha pinned).

/17 = /16 + ONE new capability, READ ONLY (the owner condition: "kizarolag olvasasi kepesseg"):
  sheet.protection.read/v1  -- one sheet: is it protected, which operations the protection forbids, and (for an
                               optional range) which non-empty cells are unlocked / have a hidden formula

Owner conditions this handler is built to (the product-side ones -- #326 task-level guard, #327 redaction, the
race rule -- live in the product, not here):
  - read only: the handler calls no setter and no editor action
  - never returns a password, a password hash or any other credential: the engine's sheetProtection object DOES
    carry algorithmName / hashValue / saltValue / spinCount / password / temporaryPassword and a _ws
    back-reference (LIVE measured, names and types only: docs/evidence/gh326-bridge17-engine-names-run1-20261025.json
    on the research branch). The handler reads ONLY the 16 named boolean fields below, by name -- it never
    enumerates the object, so a field it does not name can never reach the output
  - only the protection metadata the guards need, no general internal-API access: the arguments are a sheet name
    and an optional A1 range, the output is booleans and A1 references, nothing else

Engine access (LIVE measured on DS 9.3.5, same evidence): ws = Api.GetSheet(name).worksheet; ws.sheetProtection
(null on an unprotected sheet); ws.getRange3(r1,c1,r2,c2)._foreachNoEmpty(cell => cell.xfs) -- cell.xfs null is
the Excel default (locked, not hidden), which agrees with getCell3().getXfs() on the measured cells. Only
NON-EMPTY cells are walked: an empty cell is reported nowhere, i.e. it reads as the default LOCKED cell -- the
fail-closed direction for a write guard. Column / row level protection styles are NOT read (NOT_MEASURED).

Deterministic: the /17 block is derived from the CANONICAL 9.3.5.0 /16 block by THREE exact string
substitutions, each asserted to occur exactly once. No timestamps, no environment input. Like build-rev14.py it
NEVER writes a Document Server file and has NO --sdk mode: the served sdk-all.js and its apply are a separate,
gated operator step (maintainer, disposable DS first; the live DS needs a separate owner GO).

Usage:
  build-rev17.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V16_BLOCK = HERE.parent / "bridge-block.js"
V16_BLOCK_SHA = "2686705c45e49ae3c40b8a5b4127f9242abb5484b7efd5ed623ae492b21b59f6"
V16_BLOCK_SIZE = 98252

FLAGS = ("objects", "scenarios", "formatCells", "formatColumns", "formatRows", "insertColumns", "insertRows",
         "insertHyperlinks", "deleteColumns", "deleteRows", "selectLockedCells", "sort", "autoFilter", "pivotTables",
         "selectUnlockedCells")
MAX_CELLS = 10000

PROTECTION_REASONS = ("bad-args, sheet-not-found, worksheet-missing, protection-threw, protection-not-an-object, "
                      "protection-flag-not-boolean:<name>, range-invalid, cells-walk-threw, cell-flag-not-boolean, "
                      "cells-over-cap:" + str(MAX_CELLS))

PROTECTION_ENTRY = (
    '{ id: "sheet.protection.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one sheet (sheet name, optional A1 range): {protected, forbids, cells} -- protected is the '
    'sheet protection state; forbids holds the 15 named Excel protection flags (true = the protection forbids that '
    'operation; null when the sheet is not protected); cells (only with a range) lists the NON-EMPTY cells of the '
    'range that are unlocked and those whose formula is hidden, as A1 references -- an empty cell is the Excel '
    'default (locked, not hidden) and is listed nowhere; column/row level styles are not read. READ ONLY; never '
    'returns a password, a password hash, a salt or a spin count: only the named boolean fields are read, the '
    'protection object is never enumerated. reasons: ' + PROTECTION_REASONS + ' (GH #326/#327, /17)" }, '
)

PROTECTION_IMPL = (
    '"sheet.protection.read/v1": function (args) { '
    'if (!args || typeof args.sheet !== "string" || (args.range !== void 0 && typeof args.range !== "string")) return { ok: false, reason: "bad-args" }; '
    'var w = Api.GetSheet(args.sheet); if (!w) return { ok: false, reason: "sheet-not-found" }; '
    'var m; try { m = w.worksheet; } catch (e) { return { ok: false, reason: "worksheet-missing" }; } '
    'if (!m) return { ok: false, reason: "worksheet-missing" }; '
    'var sp; try { sp = m.sheetProtection; } catch (e) { return { ok: false, reason: "protection-threw" }; } '
    'var FLAGS = ' + json.dumps(list(FLAGS)) + '; '
    'var value = { protected: false, forbids: null, cells: null }; '
    'if (sp !== null && sp !== void 0) { '
    'if (typeof sp !== "object") return { ok: false, reason: "protection-not-an-object" }; '
    'var on = sp.sheet; if (typeof on !== "boolean") return { ok: false, reason: "protection-flag-not-boolean:sheet" }; '
    'if (on) { var f = {}; for (var i = 0; i < FLAGS.length; i += 1) { var v = sp[FLAGS[i]]; '
    'if (typeof v !== "boolean") return { ok: false, reason: "protection-flag-not-boolean:" + FLAGS[i] }; f[FLAGS[i]] = v; } '
    'value.protected = true; value.forbids = f; } } '
    'if (args.range !== void 0) { '
    'var b; try { var rr = w.GetRange(args.range); b = rr && rr.range && rr.range.bbox; } catch (e) { b = null; } '
    'if (!b || typeof b.r1 !== "number" || typeof b.c1 !== "number" || typeof b.r2 !== "number" || typeof b.c2 !== "number") return { ok: false, reason: "range-invalid" }; '
    'var col = function (c) { var s = ""; var n = c + 1; while (n > 0) { var k = (n - 1) % 26; s = String.fromCharCode(65 + k) + s; n = Math.floor((n - 1) / 26); } return s; }; '
    'var unlocked = []; var hidden = []; var seen = 0; var bad = null; '
    'try { m.getRange3(b.r1, b.c1, b.r2, b.c2)._foreachNoEmpty(function (cell) { '
    'if (bad) return; seen += 1; if (seen > ' + str(MAX_CELLS) + ') { bad = "cells-over-cap:' + str(MAX_CELLS) + '"; return; } '
    'var x = cell.xfs; var lk = x ? x.locked : null; var hd = x ? x.hidden : null; '
    'if ((lk !== null && lk !== void 0 && typeof lk !== "boolean") || (hd !== null && hd !== void 0 && typeof hd !== "boolean")) { bad = "cell-flag-not-boolean"; return; } '
    'var ref = col(cell.nCol) + (cell.nRow + 1); '
    'if (lk === false) unlocked.push(ref); if (hd === true) hidden.push(ref); }); '
    '} catch (e) { return { ok: false, reason: "cells-walk-threw" }; } '
    'if (bad) return { ok: false, reason: bad }; '
    'value.cells = { range: args.range, nonEmpty: seen, unlocked: unlocked, hidden: hidden }; } '
    'return { ok: true, value: value }; }, '
)

READ_END = ('return { ok: true, value: { dash: dashOut, smooth: smoothOut } }; }, }; var EO_BRIDGE_WRITES = {')
MANIFEST_END = '(GH #139 B /16)" }, ], };'

# (old, new) -- each `old` must occur EXACTLY once in the EVOLVING text at the time it is applied.
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.5.0/16"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/17"'),
    (MANIFEST_END, '(GH #139 B /16)" }, ' + PROTECTION_ENTRY + '], };'),
    # the READ dispatch object's closing anchor: the LAST read handler already in the /16 block ends with this
    # exact text immediately before "var EO_BRIDGE_WRITES = {" (the same insertion-point convention as
    # build-rev14.py).
    (READ_END, READ_END.replace('}; var EO_BRIDGE_WRITES = {', PROTECTION_IMPL + '}; var EO_BRIDGE_WRITES = {')),
]
MARKERS = ['"sheet.protection.read/v1": function', '{ id: "sheet.protection.read/v1"']
FORBIDDEN_IN_HANDLER = ("hashValue", "saltValue", "spinCount", "algorithmName", "password", "Object.keys", "for (var k in")


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v16 = V16_BLOCK.read_bytes()
    if sha(v16) != V16_BLOCK_SHA or len(v16) != V16_BLOCK_SIZE:
        sys.exit(f"REFUSE: /16 block sha/size {sha(v16)}/{len(v16)} != {V16_BLOCK_SHA}/{V16_BLOCK_SIZE}; NO OUTPUT")
    for word in FORBIDDEN_IN_HANDLER:
        if word in PROTECTION_IMPL:
            sys.exit(f"REFUSE: the handler text names {word!r} -- it must read the named flags only; NO OUTPUT")
    t = v16.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("(GH #326/#327, /17)") != 1:
        sys.exit("REFUSE: (GH #326/#327, /17) count != 1; NO OUTPUT")
    return v16, t.encode("utf-8")


def main():
    v16, v17 = build_block()
    out = {"v16_block_sha": sha(v16), "v16_block_size": len(v16), "v17_block_sha": sha(v17), "v17_block_size": len(v17)}
    (HERE / "bridge-block.js").write_bytes(v17)
    (HERE / "old.txt").write_bytes(v16)
    (HERE / "new.txt").write_bytes(v17)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v17-gh326-sheet-protection-read", "old": v16.decode("utf-8"), "new": v17.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (maintainer; disposable DS "
                              "first, the live DS needs its own owner GO).",
                 original_sha256=sha(v16), patched_sha256=sha(v17), original_size=len(v16), patched_size=len(v17))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
