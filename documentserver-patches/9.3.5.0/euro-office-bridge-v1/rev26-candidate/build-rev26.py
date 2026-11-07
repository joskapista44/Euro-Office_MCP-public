#!/usr/bin/env python3
"""
f0451f47 (internal tracker, measured through a live incident on the the production host /25 window, tracker card for this unit is
its own history) -- OFFLINE build of the euro-office-bridge-v1 patchVersion /26 CANDIDATE for DS 9.3.5, in ONE
deterministic pass from the /25 candidate block (../rev25-candidate/bridge-block.js, sha pinned).

/26 = /25 + EXACTLY TWO handler-body changes, to worksheet.copy/v1 and sheet.table.write/v1 -- no new capability,
no capability removed, no other handler touched:

THE LELET (f0451f47 card, a maintainer's the production host /25 window measurement, DS card #ref): on a workbook carrying a bare,
attribute-less <workbookProtection/> element, copy_sheet refused DISPATCH_REFUSED_NOT_PROVEN_PRE_DISPATCH while the
other four sheet-structure operations APPLIED. Code reading (not itself a live measurement) found the cause: both
worksheet.copy/v1 and sheet.table.write/v1 guard on `e.asc_isProtectedWorkbook() !== false` -- a SEPARATE engine
API from the one workbook.protection.read/v1 (the /25 fix target) already reads. asc_isProtectedWorkbook() is not
measured here to return the same three-valued shape (true/false/null) that wbModel.getWorkbookProtection returns;
it is a distinct code path the /25 fix never touched, so the SAME "empty <workbookProtection/>" shape that /25
clears for the read capability can still make these two WRITE capabilities answer workbook-protected (or refuse
for an unrelated, unnamed reason) on a workbook that is not actually locked.

THE FIX, here, and ONLY in these two handlers: replace the asc_isProtectedWorkbook() guard with the SAME
getter path workbook.protection.read/v1 already uses (wbModel.getWorkbookProtection(c_oAscWorkbookProtectType.
lockStructure)) -- not a new API, the one the /25 read handler already measures and trusts. The three-valued
semantics are now consistent in BOTH write handlers: lockStructure === true -> workbook-protected, refused (the
ONLY state that was ever meant to be refused); null (the OOXML default for a missing lockStructure attribute, an
empty <workbookProtection/> element) or false -> not locked, the guard passes and the handler proceeds exactly as
before; any other shape (undefined, a number, a string, a plain object, or a thrown read) -> fail-closed, named
workbook-protection-unreadable, UNCHANGED in spirit from the read handler's own protection-flag-not-boolean: an
unobservable protection state is still refused, never treated as "not locked".

asc_isProtectedWorkbook is correspondingly dropped from each handler's own API-presence checklist (the <api>-
missing reasons entry is a generic placeholder already covering it; it is not re-listed by name anywhere in the
two semanticContract reasons lists, so dropping it from the internal array changes no externally observable
reason name) -- an editor that lacks asc_isProtectedWorkbook but exposes wbModel.getWorkbookProtection no longer
fails closed on a method this unit no longer calls.

Each handler's semanticContract reasons list gains exactly one new name, workbook-protection-unreadable, next to
the retained workbook-protected; nothing else in either contract string changes.

No other handler, and no OTHER part of either of these two handlers (the copy/table-write logic itself, the
other guards, the error/receipt shapes) is touched. workbook.protection.read/v1 and workbook.protection.write/v1
are untouched by this unit -- their /25 fix already stands.

The four substitution anchors per handler (API array, protection guard, reasons list, plus the single shared
patchVersion bump) are each asserted to occur exactly once in the /25 block, the same discipline as build-rev25.py.
No handler source .txt file is used here (unlike rev25's full-handler replace) because each change is a small,
anchored fragment inside an otherwise-unmodified handler body -- writing it as a separate one-line file would not
narrow the diff, it would just relocate it.

NOT MEASURED in this unit (offline build + static identity + handler tests only -- no LIVE run):
  - that the fixed handlers, once installed, actually let copy_sheet (and the sheet.table.write/v1 capability)
    through on a real DS carrying the empty <workbookProtection/> shape -- the LIVE run that proves it is a
    separate, gated step (a maintainer's lane)
  - whether asc_isProtectedWorkbook ever disagrees with wbModel.getWorkbookProtection on a DS build where BOTH
    are present and well-formed (e.g. a workbook locked through a mechanism the latter getter does not see) --
    this fix answers the SHAPE actually measured (the empty-element null); it does not audit every other state
    the two APIs might theoretically diverge on

Usage:
  build-rev26.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""

import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V25_BLOCK = HERE.parent / "rev25-candidate" / "bridge-block.js"
V25_BLOCK_SHA = "7e62088d0ae10894d42dda7e990e0b26f0449ce9c99b505c5d049d2a908a68fa"
V25_BLOCK_SIZE = 161466

FORBIDDEN = ("Object.keys", "Object.values", "Object.entries", "getOwnPropertyNames", "for (var k in", "for (var p in", " in args", "//",
             "hashValue", "saltValue", "spinCount", "algorithmName", "password", "temporaryPassword", "PasswordXL", "asc_setSheet",
             "asc_setPassword")

PATCHVER_OLD = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/25"'
PATCHVER_NEW = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/26"'

API_COPY_OLD = (
    '["asc_copyWorksheet", "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_isProtectedWorkbook", '
    '"asc_registerCallback", "asc_unregisterCallback", "canEdit"]'
)
API_COPY_NEW = (
    '["asc_copyWorksheet", "asc_getWorksheetsCount", "asc_getWorksheetName", '
    '"asc_registerCallback", "asc_unregisterCallback", "canEdit"]'
)

GUARD_COPY_OLD = (
    'if (e.asc_isProtectedWorkbook() !== false) return { ok: false, reason: "workbook-protected" }; '
    'var before = names();'
)
GUARD_COPY_NEW = (
    'var wpWb; try { wpWb = e.wbModel; } catch (x) { wpWb = undefined; } '
    'if (!wpWb || typeof wpWb.getWorkbookProtection !== "function") return { ok: false, reason: "workbook-protection-unreadable" }; '
    'var wpType; try { wpType = Asc.c_oAscWorkbookProtectType; } catch (x) { wpType = undefined; } '
    'if (!wpType || typeof wpType.lockStructure !== "number") return { ok: false, reason: "workbook-protection-unreadable" }; '
    'var wpVal; try { wpVal = wpWb.getWorkbookProtection(wpType.lockStructure); } catch (x) { return { ok: false, reason: "workbook-protection-unreadable" }; } '
    'if (wpVal === true) return { ok: false, reason: "workbook-protected" }; '
    'if (wpVal !== null && wpVal !== false) return { ok: false, reason: "workbook-protection-unreadable" }; '
    'var before = names();'
)

REASONS_COPY_OLD = (
    'reasons: bad-args, unsupported-property, invalid-name, editor-missing, edit-not-allowed, workbook-locked, '
    'workbook-protected, name-taken, source-not-found, source-ambiguous, <api>-missing, copy-threw, copy-error, '
    'copy-not-observed, copy-unexpected (B03-WRITE, /8)" },'
)
REASONS_COPY_NEW = (
    'reasons: bad-args, unsupported-property, invalid-name, editor-missing, edit-not-allowed, workbook-locked, '
    'workbook-protected, workbook-protection-unreadable, name-taken, source-not-found, source-ambiguous, '
    '<api>-missing, copy-threw, copy-error, copy-not-observed, copy-unexpected (B03-WRITE, /8, f0451f47/26)" },'
)

API_TABLE_OLD = (
    '[CALLS[prop], "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_showWorksheet", "asc_isProtectedWorkbook", '
    '"asc_registerCallback", "asc_unregisterCallback", "canEdit"]'
)
API_TABLE_NEW = (
    '[CALLS[prop], "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_showWorksheet", '
    '"asc_registerCallback", "asc_unregisterCallback", "canEdit"]'
)

GUARD_TABLE_OLD = (
    'if (e.asc_isProtectedWorkbook() !== false) return { ok: false, reason: "workbook-protected" }; '
    'var sheetNames = function () {'
)
GUARD_TABLE_NEW = (
    'var wpWb; try { wpWb = e.wbModel; } catch (x) { wpWb = undefined; } '
    'if (!wpWb || typeof wpWb.getWorkbookProtection !== "function") return { ok: false, reason: "workbook-protection-unreadable" }; '
    'var wpType; try { wpType = Asc.c_oAscWorkbookProtectType; } catch (x) { wpType = undefined; } '
    'if (!wpType || typeof wpType.lockStructure !== "number") return { ok: false, reason: "workbook-protection-unreadable" }; '
    'var wpVal; try { wpVal = wpWb.getWorkbookProtection(wpType.lockStructure); } catch (x) { return { ok: false, reason: "workbook-protection-unreadable" }; } '
    'if (wpVal === true) return { ok: false, reason: "workbook-protected" }; '
    'if (wpVal !== null && wpVal !== false) return { ok: false, reason: "workbook-protection-unreadable" }; '
    'var sheetNames = function () {'
)

REASONS_TABLE_OLD = (
    'reasons: bad-args, unsupported-property, style-unknown, flag-unknown, value-not-boolean, name-invalid, '
    'ref-unparsable, editor-missing, <api>-missing, c_oAscChangeFilterOptions-missing, '
    'c_oAscChangeTableStyleInfo-missing, edit-not-allowed, workbook-locked, workbook-protected, sheet-not-found, '
    'sheet-ambiguous, tables-read-failed:<reason>, table-not-found, table-ambiguous, already-at-value, '
    'table-style-info-missing, header-row-missing, name-taken, range-unreadable, ref-moves-origin, ref-too-small, '
    'ref-overlaps, <call>-threw, <call>-error, table-count-changed, other-table-changed, no-effect, '
    '<property>-unexpected, target-collateral-change (GH #332-#335, /19)" },'
)
REASONS_TABLE_NEW = (
    'reasons: bad-args, unsupported-property, style-unknown, flag-unknown, value-not-boolean, name-invalid, '
    'ref-unparsable, editor-missing, <api>-missing, c_oAscChangeFilterOptions-missing, '
    'c_oAscChangeTableStyleInfo-missing, edit-not-allowed, workbook-locked, workbook-protected, '
    'workbook-protection-unreadable, sheet-not-found, sheet-ambiguous, tables-read-failed:<reason>, '
    'table-not-found, table-ambiguous, already-at-value, table-style-info-missing, header-row-missing, name-taken, '
    'range-unreadable, ref-moves-origin, ref-too-small, ref-overlaps, <call>-threw, <call>-error, '
    'table-count-changed, other-table-changed, no-effect, <property>-unexpected, target-collateral-change '
    '(GH #332-#335, /19, f0451f47/26)" },'
)

BLOCK_ENTRY = 'Api.GetEuroOfficeBridge = function () {'
MARKERS = [
    '"workbook.protection.read/v1": function', '"workbook.protection.write/v1": function',
    '{ id: "workbook.protection.read/v1"', '{ id: "workbook.protection.write/v1"',
    '"worksheet.copy/v1": function', '{ id: "worksheet.copy/v1"',
    '"sheet.table.write/v1": function', '{ id: "sheet.table.write/v1"',
]


def sha(b):
    return hashlib.sha256(b).hexdigest()


def build_block():
    v25 = V25_BLOCK.read_bytes()
    if sha(v25) != V25_BLOCK_SHA or len(v25) != V25_BLOCK_SIZE:
        sys.exit(f"REFUSE: /25 block sha/size {sha(v25)}/{len(v25)} != {V25_BLOCK_SHA}/{V25_BLOCK_SIZE}; NO OUTPUT")
    t = v25.decode("utf-8")
    if t.count(BLOCK_ENTRY) != 1:
        sys.exit("REFUSE: the /25 block does not carry the bridge entrypoint exactly once; NO OUTPUT")
    for new in (GUARD_COPY_NEW, GUARD_TABLE_NEW):
        for word in FORBIDDEN:
            if word in new:
                sys.exit(f"REFUSE: new guard fragment contains {word!r}; NO OUTPUT")
    subs = [
        (PATCHVER_OLD, PATCHVER_NEW),
        (API_COPY_OLD, API_COPY_NEW),
        (GUARD_COPY_OLD, GUARD_COPY_NEW),
        (REASONS_COPY_OLD, REASONS_COPY_NEW),
        (API_TABLE_OLD, API_TABLE_NEW),
        (GUARD_TABLE_OLD, GUARD_TABLE_NEW),
        (REASONS_TABLE_OLD, REASONS_TABLE_NEW),
    ]
    for old, new in subs:
        n = t.count(old)
        if n != 1:
            sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:70]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1:
            sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("asc_isProtectedWorkbook") != 0:
        sys.exit("REFUSE: asc_isProtectedWorkbook still referenced after substitution; NO OUTPUT")
    if t.count("workbook-protection-unreadable") != 10:
        sys.exit(f"REFUSE: workbook-protection-unreadable count {t.count('workbook-protection-unreadable')} (want 10 -- 4 per guard x 2 handlers + 1 reasons mention x 2); NO OUTPUT")
    return v25, t.encode("utf-8")


def main():
    v25, v26 = build_block()
    out = {"v25_block_sha": sha(v25), "v25_block_size": len(v25), "v26_block_sha": sha(v26), "v26_block_size": len(v26)}
    (HERE / "bridge-block.js").write_bytes(v26)
    (HERE / "old.txt").write_bytes(v25)
    (HERE / "new.txt").write_bytes(v26)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v26-f0451f47-worksheet-copy-table-write-protection-guard",
                             "old": v25.decode("utf-8"), "new": v26.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (the disposable DS "
                              "first, the live DS needs its own owner GO).",
                 original_sha256=sha(v25), patched_sha256=sha(v26), original_size=len(v25), patched_size=len(v26))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
