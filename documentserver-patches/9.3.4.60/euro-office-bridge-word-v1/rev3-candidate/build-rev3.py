#!/usr/bin/env python3
"""
M11.5e.1 (owner: M11.5 continuous mandate, owner channel, internal tracker; design an internal branch
#ref, elfogadva ket korrekcioval) -- OFFLINE build of the euro-office-bridge-word-v1
patchVersion /3 CANDIDATE.

/3 = /2 + exactly ONE new capability, docx.section.variantState.read/v1 (kind READ, mutation NO). PUBLIC_FIRST:
ApiSection has no public getter for either flag -- SetTitlePage exists, GetTitlePage/IsTitlePage do not (0
occurrences in the served /2 word sdk-all.js, internal tracker); even/odd is document-level and has no
public getter at all. The capability reads both flags from the internal section object
(ApiSection.Section.Get_TitlePage / IsEvenAndOdd) in ONE call, fail-closed on any defect (missing reader,
throw, non-boolean) -- never a partial {titlePage, evenAndOdd} value. Every /2 capability (2 READ + 1 WRITE)
stays byte-identical.

Deterministic: the /3 block is derived from the CANDIDATE /2 block (../rev2-candidate/bridge-block.js without
its trailing comma, sha pinned; it is the EXACT output of build-rev2.py, verified byte-identical to the served
/2 word sdk-all.js's embedded block -- internal tracker measured the served sdk sha against the checked-in /2
manifest.json and it matches) by THREE exact substitutions, each asserted once (fewer than rev2's four: a READ
capability dispatches through the EXISTING read()/EO_BRIDGE_CAPS machinery, no new entrypoint is added).

Usage:
  build-rev3.py                          -> writes bridge-block.js, old.txt, new.txt next to this file
  build-rev3.py --sdk <served /2 word sdk-all.js>
                                         -> also writes manifest.json (whole-block replace, /2 output sha -> /3 sha)
                                            and manifest-mechanism-proof.json; the sdk file must be the /2 output
This script never writes a Document Server file.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V2_FILE = HERE.parent / "rev2-candidate" / "bridge-block.js"
V2_FILE_SHA = "1f396f38c30bb45b3fa948a967d853a837cccb49b47277c67c673633fa1ceac2"
V2_SDK_SHA = "acd8fa14a65341ef2aa8a5c3adcf5c4e0ffb43627300732219b2e417410be924"    # = /2 manifest patched_sha256 (served)
V2_SDK_SIZE = 18552617

REASONS = ("bad-args, document-unavailable, sections-threw, sections-not-a-list, section-out-of-range, "
           "section-not-found, reader-missing, reader-threw, not-boolean")

READ_ENTRY = (
    '{ id: "docx.section.variantState.read/v1", version: 1, kind: "READ", format: "DOCX", mutation: "NO", '
    'semanticContract: "one section by 0-based GetSections() index: {titlePage: boolean, evenAndOdd: boolean} '
    '-- no public getter exposes either flag (ApiSection has SetTitlePage but no GetTitlePage/IsTitlePage; '
    'even/odd is document-level and has no public getter at all); both flags are read from the internal section '
    'object (ApiSection.Section.Get_TitlePage / IsEvenAndOdd) in ONE call; if either reader is unavailable, '
    'throws, or returns a non-boolean, the capability fails closed rather than returning a partial value; '
    'reasons: ' + REASONS + ' (B21-READ, word /3)" }, '
)

READ_IMPL = (
    '"docx.section.variantState.read/v1": function (args) { '
    'if (!args || !argsOk("section", args.section)) return { ok: false, reason: "bad-args" }; '
    'var doc; try { doc = Api.GetDocument(); } catch (e) { return { ok: false, reason: "document-unavailable" }; } '
    'if (!doc || typeof doc.GetSections !== "function") return { ok: false, reason: "document-unavailable" }; '
    'var secs; try { secs = doc.GetSections() || []; } catch (e) { return { ok: false, reason: "sections-threw" }; } '
    'if (!Array.isArray(secs)) return { ok: false, reason: "sections-not-a-list" }; '
    'if (args.section >= secs.length) return { ok: false, reason: "section-out-of-range" }; '
    'var sec = secs[args.section]; if (!sec) return { ok: false, reason: "section-not-found" }; '
    'var inner; try { inner = sec.Section; } catch (e) { inner = null; } '
    'if (!inner || typeof inner.Get_TitlePage !== "function" || typeof inner.IsEvenAndOdd !== "function") return { ok: false, reason: "reader-missing" }; '
    'var tp; try { tp = inner.Get_TitlePage(); } catch (e) { return { ok: false, reason: "reader-threw" }; } '
    'var eo; try { eo = inner.IsEvenAndOdd(); } catch (e) { return { ok: false, reason: "reader-threw" }; } '
    'if (typeof tp !== "boolean" || typeof eo !== "boolean") return { ok: false, reason: "not-boolean" }; '
    'return { ok: true, value: { titlePage: tp, evenAndOdd: eo } }; '
    '}, '
)

# (old, new) -- each `old` must occur EXACTLY once in the /2 block
SUBS = [
    ('patchVersion: "euro-office-bridge-word-v1/9.3.4.60/2"', 'patchVersion: "euro-office-bridge-word-v1/9.3.4.60/3"'),
    ('(B18-WRITE, word /2)" }, ], };', '(B18-WRITE, word /2)" }, ' + READ_ENTRY + '], };'),
    ('} }; }, }; var EO_BRIDGE_WRITES = ', '} }; }, ' + READ_IMPL + '}; var EO_BRIDGE_WRITES = '),
]
MARKERS = ['"docx.section.variantState.read/v1": function', 'Get_TitlePage', 'IsEvenAndOdd']


def sha(b):
    return hashlib.sha256(b).hexdigest()


def build_block():
    raw = V2_FILE.read_bytes()
    if sha(raw) != V2_FILE_SHA: sys.exit(f"REFUSE: /2 block file sha {sha(raw)} != pinned; NO OUTPUT")
    v2 = raw.decode("utf-8").rstrip()
    if not v2.endswith(","): sys.exit("REFUSE: /2 block file does not end with the separator comma; NO OUTPUT")
    v2 = v2[:-1]
    t = v2
    if any(t.count(m) for m in MARKERS): sys.exit("REFUSE: a /3 marker already in the /2 block; NO OUTPUT")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    if t.count(MARKERS[0]) != 1: sys.exit(f"REFUSE: marker {MARKERS[0]!r} count {t.count(MARKERS[0])}; NO OUTPUT")
    return v2.encode("utf-8"), t.encode("utf-8")


def main():
    v2, v3 = build_block()
    out = {"v2_block_sha": sha(v2), "v2_block_size": len(v2), "v3_block_sha": sha(v3), "v3_block_size": len(v3)}
    sdk = None
    if "--sdk" in sys.argv:
        sdk = Path(sys.argv[sys.argv.index("--sdk") + 1]).read_bytes()
        if sha(sdk) != V2_SDK_SHA or len(sdk) != V2_SDK_SIZE:
            sys.exit(f"REFUSE: sdk input sha {sha(sdk)} / {len(sdk)} is not the /2 output; NO OUTPUT")
        text = sdk.decode("utf-8")
        old, new = v2.decode("utf-8"), v3.decode("utf-8")
        if text.count(old) != 1: sys.exit("REFUSE: /2 block not exactly once in the sdk input; NO OUTPUT")
        if text.count(MARKERS[0]): sys.exit("REFUSE: a /3 marker already present in the input; NO OUTPUT")
        patched = text.replace(old, new, 1).encode("utf-8")
        out.update({"v2_sdk_sha": sha(sdk), "v2_sdk_size": len(sdk), "v3_sdk_sha": sha(patched), "v3_sdk_size": len(patched)})
        if "--out" in sys.argv: Path(sys.argv[sys.argv.index("--out") + 1]).write_bytes(patched)
    (HERE / "bridge-block.js").write_bytes(v3 + b",")   # the /2 file form: block + separator comma, no newline
    (HERE / "old.txt").write_bytes(v2)
    (HERE / "new.txt").write_bytes(v3)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-word-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS[:1],
            "operations": [{"id": "impl-bridge-entrypoint-word-v3-b21-section-variant-state-read", "old": v2.decode("utf-8"), "new": v3.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt blocks, not the served bundle.",
                 original_sha256=sha(v2), patched_sha256=sha(v3), original_size=len(v2), patched_size=len(v3))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    if sdk is not None:
        man = dict(base, original_sha256=out["v2_sdk_sha"], patched_sha256=out["v3_sdk_sha"], original_size=out["v2_sdk_size"], patched_size=out["v3_sdk_size"],
                   _patched_sha256_NOTE="computed by build-rev3.py from the SERVED /2 word sdk-all.js (sha = /2 patched_sha256); CANDIDATE, NOT applied to any Document Server (DS_PATCH_APPLY=HOLD).")
        (HERE / "manifest.json").write_text(json.dumps(man, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
