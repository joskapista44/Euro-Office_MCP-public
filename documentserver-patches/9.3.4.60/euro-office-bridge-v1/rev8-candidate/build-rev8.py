#!/usr/bin/env python3
"""
M11.5c.1 (owner: M11.5 continuous mandate, owner channel, internal tracker) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /8 CANDIDATE.

/8 = /7 + exactly ONE new capability, worksheet.copy/v1 (kind WRITE, mutation TARGETED), served by the existing
write() entry. PUBLIC_FIRST found no public worksheet copy (9.3.4 ApiWorksheet has Move only; Api.AddSheet is an empty
sheet). The M11.5c.0 probe (docs/M11-5C-B03-SHEET-COPY-BRIDGE-WRITE-DESIGN.md) measured that the editor's own copy
routine asc_copyWorksheet, called from the Bridge scope inside callCommand, gives a faithful, persisted copy appended
LAST (the product's B03 semantics), while the model routine Workbook.copyWorksheet reorders the sheets on save -- so the
capability wraps the editor routine, with the SAME preconditions and success rule as the product's B03 path
(src/drivers/eurooffice/xlsx-driver.cjs copyInFrame): case-insensitive name collision and source identity re-checked
immediately before the call, canEdit / global lock / protected workbook, asc_onError captured, success = exactly ONE new
sheet of the name, observed SYNCHRONOUSLY, and no error. Anything else is ok:false with a reason, never retried inside.
Every /7 capability (manifest entry and implementation) stays byte-identical.

Deterministic: the /8 block is derived from the CANONICAL /7 block (../rev7-candidate/bridge-block.js, sha pinned) by
three exact string substitutions, each asserted to occur exactly once. No timestamps, no environment input.

Usage:
  build-rev8.py                          -> writes bridge-block.js, old.txt, new.txt next to this file
  build-rev8.py --sdk <served /7 sdk-all.js>
                                         -> also writes manifest.json (whole-block replace, /7 output sha -> /8 sha)
                                            and manifest-mechanism-proof.json; the sdk file must be the /7 output
This script never writes a Document Server file.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V7_BLOCK = HERE.parent / "rev7-candidate" / "bridge-block.js"
V7_BLOCK_SHA = "74ecfd1539fa3b158f72d240110162834641ae337293d8b256fb6f149454937d"   # = /7 new.txt
V7_SDK_SHA = "3a73042b6d4314925ea83907fec9d1ef48246facb2dcbf4b35e5741fee070b2e"     # = /7 served (production)
V7_SDK_SIZE = 20962662

REASONS = ("bad-args, unsupported-property, invalid-name, editor-missing, edit-not-allowed, workbook-locked, "
           "workbook-protected, name-taken, source-not-found, source-ambiguous, <api>-missing, copy-threw, "
           "copy-error, copy-not-observed, copy-unexpected")

WRITE_ENTRY = (
    '{ id: "worksheet.copy/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one worksheet copy per call: {sheet: source name, name: new name 1..31 chars, none of []:*?/\\\\ '
    'and not starting or ending with an apostrophe}; through the editor\'s own copy routine (asc_copyWorksheet), the copy '
    'appended LAST; the name collision (case-insensitive) and the source identity are re-checked immediately before the '
    'call, and the edit permission, the global lock and the workbook protection are checked BEFORE it; success = exactly '
    'ONE new sheet of the name observed right after the call and no editor error; the result is an operation receipt '
    '{applied:[copy], name, index} and NOT a verification -- the caller verifies with a fresh read; any other outcome is '
    'ok:false with the observed sheet lists, never retried or rolled back inside; no fallback; served only through '
    'write(); reasons: ' + REASONS + ' (B03-WRITE, /8)" }, '
)

# One line, ES5, no arrow functions -- the same style as every other bridge entry.
WRITE_IMPL = (
    '"worksheet.copy/v1": function (args) { '
    'if (!args || typeof args !== "object" || typeof args.sheet !== "string" || args.sheet === "" || typeof args.name !== "string") return { ok: false, reason: "bad-args" }; '
    'var keys = Object.keys(args); for (var i = 0; i < keys.length; i += 1) { if (keys[i] !== "sheet" && keys[i] !== "name") return { ok: false, reason: "unsupported-property" }; } '
    'var nm = args.name; if (nm.length < 1 || nm.length > 31 || /[\\[\\]:*?\\/\\\\]/.test(nm) || nm.charAt(0) === "\'" || nm.charAt(nm.length - 1) === "\'") return { ok: false, reason: "invalid-name" }; '
    'var e; try { e = Asc.editor; } catch (x) { e = undefined; } if (!e) return { ok: false, reason: "editor-missing" }; '
    'var API = ["asc_copyWorksheet", "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_isProtectedWorkbook", "asc_registerCallback", "asc_unregisterCallback", "canEdit"]; '
    'for (var j = 0; j < API.length; j += 1) { if (typeof e[API[j]] !== "function") return { ok: false, reason: API[j] + "-missing" }; } '
    'var names = function () { var n = e.asc_getWorksheetsCount(); var o = []; for (var q = 0; q < n; q += 1) o.push(e.asc_getWorksheetName(q)); return o; }; '
    'if (e.canEdit() !== true) return { ok: false, reason: "edit-not-allowed" }; '
    'var ce = e.collaborativeEditing; if (ce && typeof ce.getGlobalLock === "function" && ce.getGlobalLock()) return { ok: false, reason: "workbook-locked" }; '
    'if (e.asc_isProtectedWorkbook() !== false) return { ok: false, reason: "workbook-protected" }; '
    'var before = names(); var low = nm.toLowerCase(); '
    'for (var k = 0; k < before.length; k += 1) { if (String(before[k]).toLowerCase() === low) return { ok: false, reason: "name-taken", before: before }; } '
    'var idx = before.indexOf(args.sheet); if (idx < 0) return { ok: false, reason: "source-not-found", before: before }; '
    'if (before.indexOf(args.sheet, idx + 1) >= 0) return { ok: false, reason: "source-ambiguous", before: before }; '
    'var errors = []; var onError = function (id, level) { errors.push({ id: id, level: level }); }; '
    'try { e.asc_registerCallback("asc_onError", onError); } catch (x) { errors.push({ id: "callback-not-registered" }); } '
    'var threw = null; try { e.asc_copyWorksheet(-1, [nm], [idx]); } catch (x) { threw = String(x).slice(0, 200); } '
    'try { e.asc_unregisterCallback("asc_onError", onError); } catch (x) { } '
    'var after = names(); var hits = 0; for (var h = 0; h < after.length; h += 1) { if (after[h] === nm) hits += 1; } '
    'if (threw !== null) return { ok: false, reason: "copy-threw", error: threw, before: before, after: after }; '
    'if (errors.length) return { ok: false, reason: "copy-error", errors: errors, before: before, after: after }; '
    'if (hits === 1 && after.length === before.length + 1 && after[after.length - 1] === nm) return { ok: true, value: { applied: ["copy"], name: nm, index: after.length - 1 } }; '
    'return { ok: false, reason: after.length === before.length && hits === 0 ? "copy-not-observed" : "copy-unexpected", before: before, after: after }; }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the /7 block
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/7"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/8"'),
    ('(B10-WRITE, /7)" }, ], };', '(B10-WRITE, /7)" }, ' + WRITE_ENTRY + '], };'),
    ('if (bad) return bad; return { ok: true, value: { applied: applied } }; }, }; var EO_BRIDGE_HYDRATIONS = {',
     'if (bad) return bad; return { ok: true, value: { applied: applied } }; }, ' + WRITE_IMPL + '}; var EO_BRIDGE_HYDRATIONS = {'),
]
MARKERS = ['"worksheet.copy/v1": function', '{ id: "worksheet.copy/v1"', '(B03-WRITE, /8)']


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v7 = V7_BLOCK.read_bytes()
    if sha(v7) != V7_BLOCK_SHA: sys.exit(f"REFUSE: /7 block sha {sha(v7)} != {V7_BLOCK_SHA}; NO OUTPUT")
    t = v7.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v7, t.encode("utf-8")


def main():
    v7, v8 = build_block()
    out = {"v7_block_sha": sha(v7), "v7_block_size": len(v7), "v8_block_sha": sha(v8), "v8_block_size": len(v8)}
    sdk = None
    if "--sdk" in sys.argv:
        sdk = Path(sys.argv[sys.argv.index("--sdk") + 1]).read_bytes()
        if sha(sdk) != V7_SDK_SHA or len(sdk) != V7_SDK_SIZE:
            sys.exit(f"REFUSE: sdk input sha {sha(sdk)} / {len(sdk)} is not the /7 output; NO OUTPUT")
        text = sdk.decode("utf-8")
        old, new = v7.decode("utf-8"), v8.decode("utf-8")
        if text.count(old) != 1: sys.exit("REFUSE: /7 block not exactly once in the sdk input; NO OUTPUT")
        if any(text.count(m) for m in MARKERS): sys.exit("REFUSE: a /8 marker already present in the input; NO OUTPUT")
        patched = text.replace(old, new, 1).encode("utf-8")
        out.update({"v7_sdk_sha": sha(sdk), "v7_sdk_size": len(sdk), "v8_sdk_sha": sha(patched), "v8_sdk_size": len(patched)})
    (HERE / "bridge-block.js").write_bytes(v8)
    (HERE / "old.txt").write_bytes(v7)
    (HERE / "new.txt").write_bytes(v8)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v8-b03w-worksheet-copy", "old": v7.decode("utf-8"), "new": v8.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt blocks, not the served bundle.",
                 original_sha256=sha(v7), patched_sha256=sha(v8), original_size=len(v7), patched_size=len(v8))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    if sdk is not None:
        man = dict(base, original_sha256=out["v7_sdk_sha"], patched_sha256=out["v8_sdk_sha"], original_size=out["v7_sdk_size"], patched_size=out["v8_sdk_size"],
                   _patched_sha256_NOTE="computed by build-rev8.py from the SERVED /7 cell sdk-all.js (sha = production /7); CANDIDATE until the M11.5c DS apply.")
        (HERE / "manifest.json").write_text(json.dumps(man, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
