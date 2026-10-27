#!/usr/bin/env python3
"""
GH #292 S4 (owner GO, 2026-10-22: a narrow DS-Bridge capability for the next / linked style, only the measured setter) -- OFFLINE build
of the euro-office-bridge-word-v1 patchVersion /4 CANDIDATE on DS 9.3.5.

/4 = /3 + exactly ONE new capability, docx.style.nextLink.write/v1 (kind WRITE, mutation TARGETED). PUBLIC_FIRST: ApiStyle has no
SetNext / SetLink on DS 9.3.5 (docs/evidence/gh292-style-definition-probe-20261021.json). Measured behind it
(test/live/docx-gh292-next-link-internal-probe.live.cjs; docs/evidence/gh292s4-next-link-internal-probe-20261022.json and
gh292s4b-next-link-internal-probe-run2-20261022.json): the internal style object (ApiStyle.Style) has Set_Next / Set_Link / Get_Next /
Get_Link / GetId, they take a style id, the public ToJSON then names next / link and the saved styles.xml has w:next / w:link, also after a
reopen; a character style's own Set_Link back to the paragraph style persists too. The editor does NOT validate: an unknown id and a
next on a character style are stored as given -- so the handler resolves both styles by NAME (GetName() === name, never the Normal
fallback), checks their types (next: paragraph -> paragraph; link: a paragraph and a character style, both sides set), refuses a style
already linked elsewhere, and re-reads the getter after the call. Clearing (no next / no link) is NOT offered -- not measured.
Every /3 capability (3 READ + the B18 WRITE) stays byte-identical.

Deterministic: the /4 block is derived from the shipped /3 block (../bridge-block.js, sha pinned) by THREE exact substitutions, each
asserted once.

Usage:
  build-rev4.py                          -> writes bridge-block.js, old.txt, new.txt next to this file
  build-rev4.py --sdk <served /3 word sdk-all.js> [--out <file>]
                                         -> also writes manifest.json (whole-block replace, /3 output sha -> /4 sha)
This script never writes a Document Server file.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V3_FILE = HERE.parent / "bridge-block.js"
V3_FILE_SHA = "55795f6c07d09ba290094c1b2e426cdfad9ed4b31b5417bc755898ef291eb8b0"   # = fresh-install.json word block sha
V3_SDK_SHA = "4afc854de1356d62be533e9e4d4fe766dd4ef4b3f9f32915d163fd15ce9a10ca"    # = fresh-install.json word output_sha256 (served /3)
V3_SDK_SIZE = 18559873

REASONS = ("bad-args, unsupported-property, unsupported-field, document-unavailable, style-not-found, target-not-found, type-threw, "
           "type-mismatch, internal-missing, getter-threw, no-change, already-linked-elsewhere, set-threw, write-not-observed")

WRITE_ENTRY = (
    '{ id: "docx.style.nextLink.write/v1", version: 1, kind: "WRITE", format: "DOCX", mutation: "TARGETED", '
    'semanticContract: "one style per call: {style: its exact name, field: next|link, target: the exact name of the style it points to}; '
    'both styles are resolved by NAME (doc.GetStyle(name) whose GetName() is that name -- never the Normal fallback); next: both must be '
    'paragraph styles, the style\'s next is set to the target; link: one paragraph and one character style, BOTH sides are set (Word\'s '
    'linked pair), and a style already linked to a third style is refused, never re-pointed -- no public Builder call sets either (ApiStyle has '
    'no SetNext / SetLink on DS 9.3.5); the editor\'s own internal style setters Set_Next / Set_Link take the target\'s id; the value is '
    're-read with Get_Next / Get_Link right after the call; success = the getter answers the target\'s id; the result is an operation '
    'receipt {applied:[next-set|link-set], before:{value}} and NOT a verification -- the caller verifies with a fresh read; an already '
    'equal value is no-change (no mutation), never a silent success; clearing a next / link is not offered; no fallback; served only through '
    'write(); reasons: ' + REASONS + ' (GH #292 S4, word /4)" }, '
)

WRITE_IMPL = (
    '"docx.style.nextLink.write/v1": function (args) { '
    'if (!args || typeof args !== "object" || typeof args.style !== "string" || args.style === "" || typeof args.target !== "string" || args.target === "") return { ok: false, reason: "bad-args" }; '
    'var keys = Object.keys(args); for (var i = 0; i < keys.length; i += 1) { if (keys[i] !== "style" && keys[i] !== "field" && keys[i] !== "target") return { ok: false, reason: "unsupported-property" }; } '
    'if (args.field !== "next" && args.field !== "link") return { ok: false, reason: "unsupported-field" }; '
    'var doc; try { doc = Api.GetDocument(); } catch (e) { return { ok: false, reason: "document-unavailable" }; } '
    'if (!doc || typeof doc.GetStyle !== "function") return { ok: false, reason: "document-unavailable" }; '
    'function byName(name) { var st; try { st = doc.GetStyle(name); } catch (e) { return null; } if (!st || typeof st.GetName !== "function") return null; var nm; try { nm = st.GetName(); } catch (e) { return null; } return nm === name ? st : null; } '
    'var s = byName(args.style); if (!s) return { ok: false, reason: "style-not-found" }; '
    'var t = byName(args.target); if (!t) return { ok: false, reason: "target-not-found" }; '
    'var sType, tType; try { sType = s.GetType(); tType = t.GetType(); } catch (e) { return { ok: false, reason: "type-threw" }; } '
    'if (args.field === "next" && !(sType === "paragraph" && tType === "paragraph")) return { ok: false, reason: "type-mismatch", styleType: String(sType), targetType: String(tType) }; '
    'if (args.field === "link" && !((sType === "paragraph" && tType === "run") || (sType === "run" && tType === "paragraph"))) return { ok: false, reason: "type-mismatch", styleType: String(sType), targetType: String(tType) }; '
    'var si = s.Style, ti = t.Style; '
    'var setter = args.field === "next" ? "Set_Next" : "Set_Link", getter = args.field === "next" ? "Get_Next" : "Get_Link"; '
    'if (!si || !ti || typeof si.GetId !== "function" || typeof ti.GetId !== "function" || typeof si[setter] !== "function" || typeof si[getter] !== "function" || (args.field === "link" && (typeof ti.Set_Link !== "function" || typeof ti.Get_Link !== "function"))) return { ok: false, reason: "internal-missing" }; '
    'function val(v) { return v === undefined || v === null || v === "" ? null : String(v); } '
    'var sid, tid, before, tBefore = null; try { sid = String(si.GetId()); tid = String(ti.GetId()); before = val(si[getter]()); if (args.field === "link") tBefore = val(ti.Get_Link()); } catch (e) { return { ok: false, reason: "getter-threw" }; } '
    'if (args.field === "next" && before === tid) return { ok: false, reason: "no-change", before: { value: before } }; '
    'if (args.field === "link") { if (before === tid && tBefore === sid) return { ok: false, reason: "no-change", before: { value: before } }; if ((before !== null && before !== tid) || (tBefore !== null && tBefore !== sid)) return { ok: false, reason: "already-linked-elsewhere", before: { value: before, targetValue: tBefore } }; } '
    'try { si[setter](tid); if (args.field === "link") ti.Set_Link(sid); } catch (e) { return { ok: false, reason: "set-threw", error: String(e).slice(0, 200), before: { value: before } }; } '
    'var after, tAfter = sid; try { after = val(si[getter]()); if (args.field === "link") tAfter = val(ti.Get_Link()); } catch (e) { return { ok: false, reason: "write-not-observed", before: { value: before } }; } '
    'if (after !== tid || tAfter !== sid) return { ok: false, reason: "write-not-observed", before: { value: before } }; '
    'return { ok: true, value: { applied: [args.field === "next" ? "next-set" : "link-set"], before: { value: before } } }; '
    '}, '
)

# (old, new) -- each `old` must occur EXACTLY once in the /3 block
SUBS = [
    ('patchVersion: "euro-office-bridge-word-v1/9.3.5.0/3"', 'patchVersion: "euro-office-bridge-word-v1/9.3.5.0/4"'),
    ('(B21-READ, word /3)" }, ], };', '(B21-READ, word /3)" }, ' + WRITE_ENTRY + '], };'),
    ('before: { level: lvl, numId: numId } } }; }, }; return {', 'before: { level: lvl, numId: numId } } }; }, ' + WRITE_IMPL + '}; return {'),
]
MARKERS = ['"docx.style.nextLink.write/v1": function', 'Set_Next', 'Set_Link']


def sha(b):
    return hashlib.sha256(b).hexdigest()


def build_block():
    raw = V3_FILE.read_bytes()
    if sha(raw) != V3_FILE_SHA: sys.exit(f"REFUSE: /3 block file sha {sha(raw)} != pinned; NO OUTPUT")
    v3 = raw.decode("utf-8").rstrip()
    if not v3.endswith(","): sys.exit("REFUSE: /3 block file does not end with the separator comma; NO OUTPUT")
    v3 = v3[:-1]
    t = v3
    if any(t.count(m) for m in MARKERS): sys.exit("REFUSE: a /4 marker already in the /3 block; NO OUTPUT")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    if t.count(MARKERS[0]) != 1: sys.exit(f"REFUSE: marker {MARKERS[0]!r} count {t.count(MARKERS[0])}; NO OUTPUT")
    return v3.encode("utf-8"), t.encode("utf-8")


def main():
    v3, v4 = build_block()
    out = {"v3_block_sha": sha(v3), "v3_block_size": len(v3), "v4_block_sha": sha(v4), "v4_block_size": len(v4)}
    sdk = None
    if "--sdk" in sys.argv:
        sdk = Path(sys.argv[sys.argv.index("--sdk") + 1]).read_bytes()
        if sha(sdk) != V3_SDK_SHA or len(sdk) != V3_SDK_SIZE:
            sys.exit(f"REFUSE: sdk input sha {sha(sdk)} / {len(sdk)} is not the /3 output; NO OUTPUT")
        text = sdk.decode("utf-8")
        old, new = v3.decode("utf-8"), v4.decode("utf-8")
        if text.count(old) != 1: sys.exit("REFUSE: /3 block not exactly once in the sdk input; NO OUTPUT")
        if text.count(MARKERS[0]): sys.exit("REFUSE: a /4 marker already present in the input; NO OUTPUT")
        patched = text.replace(old, new, 1).encode("utf-8")
        out.update({"v3_sdk_sha": sha(sdk), "v3_sdk_size": len(sdk), "v4_sdk_sha": sha(patched), "v4_sdk_size": len(patched)})
        if "--out" in sys.argv: Path(sys.argv[sys.argv.index("--out") + 1]).write_bytes(patched)
    (HERE / "bridge-block.js").write_bytes(v4 + b",")   # the /3 file form: block + separator comma, no newline
    (HERE / "old.txt").write_bytes(v3)
    (HERE / "new.txt").write_bytes(v4)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-word-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS[:1],
            "operations": [{"id": "impl-bridge-entrypoint-word-v4-gh292-style-next-link-write", "old": v3.decode("utf-8"), "new": v4.decode("utf-8")}]}
    if sdk is not None:
        man = dict(base, original_sha256=out["v3_sdk_sha"], patched_sha256=out["v4_sdk_sha"], original_size=out["v3_sdk_size"], patched_size=out["v4_sdk_size"],
                   _patched_sha256_NOTE="computed by build-rev4.py from the SERVED /3 word sdk-all.js (sha = fresh-install.json word output_sha256); CANDIDATE, NOT applied to any Document Server (DS_PATCH_APPLY=HOLD).")
        (HERE / "manifest.json").write_text(json.dumps(man, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
