#!/usr/bin/env python3
"""
M11.5d.1 (owner: M11.5 continuous mandate, owner channel, internal tracker) -- OFFLINE build of the
euro-office-bridge-word-v1 patchVersion /2 CANDIDATE.

/2 = /1 + a write() entry + exactly ONE new capability, docx.paragraph.numbering.remove/v1 (kind WRITE, mutation
TARGETED). PUBLIC_FIRST: no public Builder call removes numbering (2026-09-24 probes 1-3; re-measured 2026-10-11 on
the served /1 runtime: SetBullet(null) and ParaPr.SetBullet(null) no effect, the internal RemoveNumPr control removes,
card #ref). The capability wraps the editor's own CParagraph.RemoveNumPr behind the Bridge, with the SAME target rule
and precondition as the product's B18 path (src/adapters/docx/families/list.cjs commandBody): the target paragraph by
its EXACT text (Search, MatchCase, GetText({Numbering:false}) === text, exactly one hit), its live numbering level
re-checked against the caller's expected level immediately before the call; success = the paragraph has NO numbering
right after the call. Anything else is ok:false with a reason, never retried inside. Every /1 capability (manifest
entry and implementation) stays byte-identical.

Deterministic: the /2 block is derived from the CANONICAL /1 block (../bridge-block.js without its trailing comma, sha
pinned; it occurs exactly once in the served /1 word sdk-all.js) by four exact substitutions, each asserted once.

Usage:
  build-rev2.py                          -> writes bridge-block.js, old.txt, new.txt next to this file
  build-rev2.py --sdk <served /1 word sdk-all.js>
                                         -> also writes manifest.json (whole-block replace, /1 output sha -> /2 sha)
                                            and manifest-mechanism-proof.json; the sdk file must be the /1 output
This script never writes a Document Server file.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V1_FILE = HERE.parent / "bridge-block.js"
V1_FILE_SHA = "a85af48d9397d7be92021e14afec96e732f330548a64961eb72ba9c67cac079e"
V1_SDK_SHA = "8362e3a9fe07173e7ce9ece91e73d0829bdd7bd6ecf2b13a28330ae0e08b4699"     # = /1 manifest patched_sha256 (served)
V1_SDK_SIZE = 18548802

REASONS = ("bad-args, unsupported-property, invalid-level, document-unavailable, search-threw, target-not-found, "
           "target-ambiguous, paragraph-unreadable, no-GetNumbering, not-numbered, level-mismatch, RemoveNumPr-missing, "
           "remove-threw, remove-not-observed")

WRITE_ENTRY = (
    '{ id: "docx.paragraph.numbering.remove/v1", version: 1, kind: "WRITE", format: "DOCX", mutation: "TARGETED", '
    'semanticContract: "one paragraph per call: {text: its exact text (search MatchCase, text read without numbering, '
    'exactly one hit), level: the caller\'s expected live numbering level 0..8}; removes the paragraph\'s numbering through '
    'the editor\'s own paragraph routine (RemoveNumPr) -- no public Builder call removes numbering; the target, its '
    'numbering and its live level are re-checked immediately before the call; success = the paragraph has NO numbering '
    'right after the call; the result is an operation receipt {applied:[numbering-removed], before:{level, numId}} and NOT '
    'a verification -- the caller verifies with a fresh read; a paragraph without numbering is not-numbered (no mutation), '
    'never a silent success; no fallback; served only through write(); reasons: ' + REASONS + ' (B18-WRITE, word /2)" }, '
)

WRITE_IMPL = (
    'var EO_BRIDGE_WRITES = { "docx.paragraph.numbering.remove/v1": function (args) { '
    'if (!args || typeof args !== "object" || typeof args.text !== "string" || args.text === "") return { ok: false, reason: "bad-args" }; '
    'var keys = Object.keys(args); for (var i = 0; i < keys.length; i += 1) { if (keys[i] !== "text" && keys[i] !== "level") return { ok: false, reason: "unsupported-property" }; } '
    'if (!(typeof args.level === "number" && Math.floor(args.level) === args.level && args.level >= 0 && args.level <= 8)) return { ok: false, reason: "invalid-level" }; '
    'var doc; try { doc = Api.GetDocument(); } catch (e) { return { ok: false, reason: "document-unavailable" }; } '
    'if (!doc || typeof doc.Search !== "function") return { ok: false, reason: "document-unavailable" }; '
    'var hits; try { hits = (doc.Search(args.text, { MatchCase: true }) || []).filter(function (r) { return String(r.GetText({ Numbering: false }) || "") === args.text; }); } catch (e) { return { ok: false, reason: "search-threw" }; } '
    'if (hits.length === 0) return { ok: false, reason: "target-not-found" }; if (hits.length !== 1) return { ok: false, reason: "target-ambiguous", count: hits.length }; '
    'var p; try { p = hits[0].GetParagraph(0); } catch (e) { p = null; } if (!p) return { ok: false, reason: "paragraph-unreadable" }; '
    'if (typeof p.GetNumbering !== "function") return { ok: false, reason: "no-GetNumbering" }; '
    'var cur = p.GetNumbering(); if (!cur) return { ok: false, reason: "not-numbered" }; '
    'var lvl = typeof cur.GetLevelIndex === "function" ? cur.GetLevelIndex() : null; if (lvl !== args.level) return { ok: false, reason: "level-mismatch", level: Number.isInteger(lvl) ? lvl : null }; '
    'var inner = p.Paragraph; if (!inner || typeof inner.RemoveNumPr !== "function") return { ok: false, reason: "RemoveNumPr-missing" }; '
    'var numId = null; try { var np = inner.Pr ? inner.Pr.NumPr : null; if (np && np.NumId !== undefined && np.NumId !== null) numId = String(np.NumId); } catch (e) { numId = null; } '
    'try { inner.RemoveNumPr(); } catch (e) { return { ok: false, reason: "remove-threw", error: String(e).slice(0, 200) }; } '
    'var after; try { after = p.GetNumbering(); } catch (e) { after = undefined; } '
    'if (after) return { ok: false, reason: "remove-not-observed", before: { level: lvl, numId: numId } }; '
    'return { ok: true, value: { applied: ["numbering-removed"], before: { level: lvl, numId: numId } } }; }, }; '
)

WRITE_ENTRYPOINT = (
    'write: function (capabilityId, args) { try { if (!Object.prototype.hasOwnProperty.call(EO_BRIDGE_WRITES, capabilityId)) '
    'return { ok: false, reason: "capability-not-advertised" }; return EO_BRIDGE_WRITES[capabilityId](args); } '
    'catch (e) { return { ok: false, reason: "internal-threw" }; } }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the /1 block
SUBS = [
    ('patchVersion: "euro-office-bridge-word-v1/9.3.4.60/1"', 'patchVersion: "euro-office-bridge-word-v1/9.3.4.60/2"'),
    ('null field" }, ], };', 'null field" }, ' + WRITE_ENTRY + '], };'),
    ('return { manifest: function', WRITE_IMPL + 'return { manifest: function'),
    ('return { ok: false, reason: "internal-threw" }; } }, }; }',
     'return { ok: false, reason: "internal-threw" }; } }, ' + WRITE_ENTRYPOINT + '}; }'),
]
MARKERS = ['"docx.paragraph.numbering.remove/v1": function', 'write: function (capabilityId, args)']


def sha(b):
    return hashlib.sha256(b).hexdigest()


def build_block():
    raw = V1_FILE.read_bytes()
    if sha(raw) != V1_FILE_SHA: sys.exit(f"REFUSE: /1 block file sha {sha(raw)} != pinned; NO OUTPUT")
    v1 = raw.decode("utf-8").rstrip()
    if not v1.endswith(","): sys.exit("REFUSE: /1 block file does not end with the separator comma; NO OUTPUT")
    v1 = v1[:-1]
    t = v1
    if any(t.count(m) for m in MARKERS): sys.exit("REFUSE: a /2 marker already in the /1 block; NO OUTPUT")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v1.encode("utf-8"), t.encode("utf-8")


def main():
    v1, v2 = build_block()
    out = {"v1_block_sha": sha(v1), "v1_block_size": len(v1), "v2_block_sha": sha(v2), "v2_block_size": len(v2)}
    sdk = None
    if "--sdk" in sys.argv:
        sdk = Path(sys.argv[sys.argv.index("--sdk") + 1]).read_bytes()
        if sha(sdk) != V1_SDK_SHA or len(sdk) != V1_SDK_SIZE:
            sys.exit(f"REFUSE: sdk input sha {sha(sdk)} / {len(sdk)} is not the /1 output; NO OUTPUT")
        text = sdk.decode("utf-8")
        old, new = v1.decode("utf-8"), v2.decode("utf-8")
        if text.count(old) != 1: sys.exit("REFUSE: /1 block not exactly once in the sdk input; NO OUTPUT")
        if any(text.count(m) for m in MARKERS): sys.exit("REFUSE: a /2 marker already present in the input; NO OUTPUT")
        patched = text.replace(old, new, 1).encode("utf-8")
        out.update({"v1_sdk_sha": sha(sdk), "v1_sdk_size": len(sdk), "v2_sdk_sha": sha(patched), "v2_sdk_size": len(patched)})
        if "--out" in sys.argv: Path(sys.argv[sys.argv.index("--out") + 1]).write_bytes(patched)
    (HERE / "bridge-block.js").write_bytes(v2 + b",")   # the /1 file form: block + separator comma, no newline
    (HERE / "old.txt").write_bytes(v1)
    (HERE / "new.txt").write_bytes(v2)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-word-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-word-v2-b18w-paragraph-numbering-remove", "old": v1.decode("utf-8"), "new": v2.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt blocks, not the served bundle.",
                 original_sha256=sha(v1), patched_sha256=sha(v2), original_size=len(v1), patched_size=len(v2))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    if sdk is not None:
        man = dict(base, original_sha256=out["v1_sdk_sha"], patched_sha256=out["v2_sdk_sha"], original_size=out["v1_sdk_size"], patched_size=out["v2_sdk_size"],
                   _patched_sha256_NOTE="computed by build-rev2.py from the SERVED /1 word sdk-all.js (sha = /1 patched_sha256); CANDIDATE, NOT applied to any Document Server (DS_PATCH_APPLY=HOLD).")
        (HERE / "manifest.json").write_text(json.dumps(man, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
