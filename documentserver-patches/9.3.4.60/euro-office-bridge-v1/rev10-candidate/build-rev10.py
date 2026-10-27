#!/usr/bin/env python3
"""
GH #92 (owner request via alice, card an internal branch) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /10 CANDIDATE.

/10 = /9 with ONE behavioural change inside the existing range.appearance.read/v1 capability (no new capability):
a border side the engine returns as `null` reads as `{ s: 0, c: 0 }` (no border) instead of refusing the whole read
with `border-unexpected-shape`.

Why (LIVE probe test/live/xlsx-gh92-merge-border-read-probe.live.cjs, disposable copy, internal tracker): right after
`ApiRange.Merge(false)`, in the SAME session, `getBorderFull()` returns `null` for every member's INTERIOR sides and
for the diagonal (R3: 11/12 members refused). After a forceSave + reopen the same sides read `{ s: 0, c: <rgb 0> }`
-- byte-equal to the original merged state (R0 = R4). So the session-local `null` is the engine's own "no border"
form; mapping it to the reopened form makes the read equal on both sides of a reopen. `undefined` (a missing key)
is NOT mapped: it still refuses (`border-unexpected-shape`), so a genuinely unexpected shape is never guessed.
The side() helper is shared with the row-2 / column-B neighbour fixup; there a mapped null side has s 0, which the
fixup's own `n && n.s` guard already treats as "no neighbour edge" -- unchanged behaviour.

Deterministic: derived from the CANONICAL /9 block (../rev9-candidate/bridge-block.js, sha pinned) by three exact
string substitutions, each asserted to occur exactly once. No timestamps, no environment input. NEVER writes a
Document Server file; no --sdk mode (same reason as build-rev9.py: the served sdk-all.js is a host artifact this
agent does not read -- the served-sdk manifest and the apply are maintainer's gate).
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V9_BLOCK = HERE.parent / "rev9-candidate" / "bridge-block.js"
V9_BLOCK_SHA = "fdc0724d1300360607b6a83059430a3cc3977e629dec0b6f272ade0266943e0c"
V9_BLOCK_SIZE = 52792

OLD_SIDE = 'var side = function (b) { return b ? { s: b.s, c: colour(b.c) } : void 0; };'
NEW_SIDE = 'var side = function (b) { if (b === null) return { s: 0, c: 0 }; return b ? { s: b.s, c: colour(b.c) } : void 0; };'
OLD_CONTRACT = "(the engine's own getBorderFull gap); reason if either read fails\" }"
NEW_CONTRACT = ("(the engine's own getBorderFull gap); a side the engine returns as null (the session-local form right "
                "after Merge, GH #92) reads {s:0,c:0}, the same side's form after a reopen; a missing side still refuses; "
                "reason if either read fails (GH #92, /10)\" }")
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/9"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/10"'),
    (OLD_SIDE, NEW_SIDE),
    (OLD_CONTRACT, NEW_CONTRACT),
]
MARKERS = ['if (b === null) return { s: 0, c: 0 };', '(GH #92, /10)', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/10"']


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v9 = V9_BLOCK.read_bytes()
    if sha(v9) != V9_BLOCK_SHA or len(v9) != V9_BLOCK_SIZE:
        sys.exit(f"REFUSE: /9 block sha/size {sha(v9)}/{len(v9)} != {V9_BLOCK_SHA}/{V9_BLOCK_SIZE}; NO OUTPUT")
    t = v9.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v9, t.encode("utf-8")


def main():
    v9, v10 = build_block()
    out = {"v9_block_sha": sha(v9), "v9_block_size": len(v9), "v10_block_sha": sha(v10), "v10_block_size": len(v10)}
    (HERE / "bridge-block.js").write_bytes(v10)
    (HERE / "old.txt").write_bytes(v9)
    (HERE / "new.txt").write_bytes(v10)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v10-gh92-merged-null-border-side", "old": v9.decode("utf-8"), "new": v10.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- a served-sdk-backed manifest.json for the real production apply is a SEPARATE, later "
                              "step, gated to the designated operator role, with the real served file in hand.",
                 original_sha256=sha(v9), patched_sha256=sha(v10), original_size=len(v9), patched_size=len(v10))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
