#!/usr/bin/env python3
"""
M11.5a.1 (owner GO: the owner, owner channel) -- OFFLINE build of the euro-office-bridge-v1 patchVersion /6 CANDIDATE.

/6 = /5 + exactly ONE new capability, range.alignmentExt.write/v1 (kind WRITE), served by a NEW write() entry next to
read() / hydrate(). Every /5 capability (manifest entry and implementation) stays byte-identical; read() and hydrate()
are untouched, so they refuse the WRITE id (capability-not-advertised) and write() refuses every READ / HYDRATION id.

Deterministic: the /6 block is derived from the CANONICAL /5 block (../bridge-block.js, sha pinned below) by four exact
string substitutions, each asserted to occur exactly once. No timestamps, no environment input.

Usage:
  build-rev6.py                         -> writes bridge-block.js, old.txt, new.txt next to this file
  build-rev6.py --sdk <served /5 sdk-all.js>
                                        -> also writes manifest.json (whole-block replace, /5 output sha -> /6 output sha)
                                           and manifest-mechanism-proof.json; the sdk file must be the /5 output
This script never writes a Document Server file.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V5_BLOCK = HERE.parent / "bridge-block.js"
V5_BLOCK_SHA = "dc879526624950fd66b1c520d13521a315b493d8d5006be1aa0e972ae157c388"   # = /5 new.txt (mechanism proof)
V5_SDK_SHA = "f1462d87f211c6d6cdfe999465c02b7a0f551139ab4738eefe02657b71d053ee"     # = /5 manifest patched_sha256
V5_SDK_SIZE = 20952345

WRITE_ENTRY = (
    '{ id: "range.alignmentExt.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one cell (sheet name + A1 of a SINGLE cell): sets ONLY the requested alignment properties -- '
    'indent (integer 0..15) and/or shrink (boolean) -- through the engine\'s own cell setters; an unrequested property '
    'is never touched (preserved); every input is validated and every requested setter is checked BEFORE the first '
    'mutation; the result is an operation receipt {applied:[...]} and NOT a verification -- the caller verifies with a '
    'fresh range.alignmentExt.read/v1; a setter that throws after an earlier one succeeded leaves a PARTIAL write, '
    'reported as {partial:true, applied:[...]}, never rolled back; no fallback of any kind; served only through '
    'write(), never through read() or hydrate(); reasons: bad-args, unsupported-property, no-property, '
    'invalid-indent, invalid-shrink, invalid-target, sheet-not-found, cell-not-found, range-missing, '
    'setIndent-missing, setShrinkToFit-missing, setIndent-threw, setShrinkToFit-threw (B05-WRITE, /6)" }, '
)

WRITE_IMPL = (
    'var EO_BRIDGE_WRITES = { "range.alignmentExt.write/v1": function (args) { '
    'if (!args || typeof args !== "object" || typeof args.sheet !== "string" || typeof args.cell !== "string") return { ok: false, reason: "bad-args" }; '
    'var allowed = { sheet: 1, cell: 1, indent: 1, shrink: 1 }; var keys = Object.keys(args); '
    'for (var i = 0; i < keys.length; i += 1) { if (!Object.prototype.hasOwnProperty.call(allowed, keys[i])) return { ok: false, reason: "unsupported-property" }; } '
    'var hasIndent = Object.prototype.hasOwnProperty.call(args, "indent"); var hasShrink = Object.prototype.hasOwnProperty.call(args, "shrink"); '
    'if (!hasIndent && !hasShrink) return { ok: false, reason: "no-property" }; '
    'if (hasIndent && !(typeof args.indent === "number" && Math.floor(args.indent) === args.indent && args.indent >= 0 && args.indent <= 15)) return { ok: false, reason: "invalid-indent" }; '
    'if (hasShrink && typeof args.shrink !== "boolean") return { ok: false, reason: "invalid-shrink" }; '
    'if (!/^[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(args.cell)) return { ok: false, reason: "invalid-target" }; '
    'var r = resolveCell(args.sheet, args.cell); if (r.err) return r.err; var inner = r.inner; '
    'if (hasIndent && typeof inner.setIndent !== "function") return { ok: false, reason: "setIndent-missing" }; '
    'if (hasShrink && typeof inner.setShrinkToFit !== "function") return { ok: false, reason: "setShrinkToFit-missing" }; '
    'var applied = []; '
    'if (hasIndent) { try { inner.setIndent(args.indent); } catch (e) { return { ok: false, reason: "setIndent-threw", partial: applied.length > 0, applied: applied }; } applied.push("indent"); } '
    'if (hasShrink) { try { inner.setShrinkToFit(args.shrink); } catch (e) { return { ok: false, reason: "setShrinkToFit-threw", partial: applied.length > 0, applied: applied }; } applied.push("shrink"); } '
    'return { ok: true, value: { applied: applied } }; }, }; '
)

WRITE_ENTRYPOINT = (
    'write: function (capabilityId, args) { try { if (!Object.prototype.hasOwnProperty.call(EO_BRIDGE_WRITES, capabilityId)) '
    'return { ok: false, reason: "capability-not-advertised" }; return EO_BRIDGE_WRITES[capabilityId](args); } '
    'catch (e) { return { ok: false, reason: "internal-threw" }; } }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the /5 block
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/5"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/6"'),
    ('no-view (B15, /5)" }, ], };', 'no-view (B15, /5)" }, ' + WRITE_ENTRY + '], };'),
    ('var EO_BRIDGE_HYDRATIONS = {', WRITE_IMPL + 'var EO_BRIDGE_HYDRATIONS = {'),
    ('return { ok: false, reason: "internal-threw" }; } }, }; },',
     'return { ok: false, reason: "internal-threw" }; } }, ' + WRITE_ENTRYPOINT + '}; },'),
]
MARKERS = ['"range.alignmentExt.write/v1": function', 'var EO_BRIDGE_WRITES = ', 'write: function (capabilityId, args)']


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v5 = V5_BLOCK.read_bytes()
    if sha(v5) != V5_BLOCK_SHA: sys.exit(f"REFUSE: /5 block sha {sha(v5)} != {V5_BLOCK_SHA}; NO OUTPUT")
    t = v5.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v5, t.encode("utf-8")


def main():
    v5, v6 = build_block()
    out = {"v5_block_sha": sha(v5), "v5_block_size": len(v5), "v6_block_sha": sha(v6), "v6_block_size": len(v6)}
    sdk = None
    if "--sdk" in sys.argv:
        sdk = Path(sys.argv[sys.argv.index("--sdk") + 1]).read_bytes()
        if sha(sdk) != V5_SDK_SHA or len(sdk) != V5_SDK_SIZE:
            sys.exit(f"REFUSE: sdk input sha {sha(sdk)} / {len(sdk)} is not the /5 output; NO OUTPUT")
        text = sdk.decode("utf-8")
        old, new = v5.decode("utf-8"), v6.decode("utf-8")
        if text.count(old) != 1: sys.exit("REFUSE: /5 block not exactly once in the sdk input; NO OUTPUT")
        if any(text.count(m) for m in MARKERS): sys.exit("REFUSE: a /6 marker already present in the input; NO OUTPUT")
        patched = text.replace(old, new, 1).encode("utf-8")
        out.update({"v5_sdk_sha": sha(sdk), "v5_sdk_size": len(sdk), "v6_sdk_sha": sha(patched), "v6_sdk_size": len(patched)})
    (HERE / "bridge-block.js").write_bytes(v6)
    (HERE / "old.txt").write_bytes(v5)
    (HERE / "new.txt").write_bytes(v6)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v6-b05w-range-alignmentext-write", "old": v5.decode("utf-8"), "new": v6.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt blocks, not the served bundle.",
                 original_sha256=sha(v5), patched_sha256=sha(v6), original_size=len(v5), patched_size=len(v6))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    if sdk is not None:
        man = dict(base, original_sha256=out["v5_sdk_sha"], patched_sha256=out["v6_sdk_sha"], original_size=out["v5_sdk_size"], patched_size=out["v6_sdk_size"],
                   _patched_sha256_NOTE="computed by build-rev6.py from the SERVED /5 cell sdk-all.js (sha = /5 patched_sha256); CANDIDATE, NOT applied to any Document Server (DS_PATCH_APPLY=HOLD).")
        (HERE / "manifest.json").write_text(json.dumps(man, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
