#!/usr/bin/env python3
"""
GH #338-#342 remainder (Bridge /24 WRITE+READ, DESIGN docs/DS-9.3.5-BRIDGE24-WORKBOOK-PROTECTION-DESIGN.md, W1+W3) + GH #335
remainder (sheet.table.write/v1 totalsRow value:false, internal tracker) -- OFFLINE build of the euro-office-bridge-v1
patchVersion /24 CANDIDATE for DS 9.3.5, in ONE deterministic pass from the /23 candidate block
(../rev23-candidate/bridge-block.js, sha pinned -- the block installed on the production DS 2026-11-01/02).

/24 = /23 + TWO new capabilities (handler-wb-protection-read-v24.txt, handler-wb-protection-write-v24.txt) + ONE existing
handler's totalsRow branch widened (sheet.table.write/v1, introduced at /19, carried byte-identical through /21-/23):
  1. workbook.protection.read/v1  -- no arguments: {lockStructure} -- the workbook own structure-protection flag
     (wbModel.getWorkbookProtection(c_oAscWorkbookProtectType.lockStructure)), workbook-wide, not per sheet (design W1).
  2. workbook.protection.write/v1 -- {protect: boolean} -- PASSWORDLESS lockStructure on/off through
     Asc.CWorkbookProtection's default constructor + asc_setLockStructure(protect) only, wbModel.setProtectedWorkbook(
     props, true) (design W3, the /22 setProtectedSheet pattern transposed to the workbook).
     FAIL-CLOSED review fix 1: {protect: false} on a workbook CURRENTLY locked is refused credential-unobservable,
     before any history point or setter call -- the product must never remove an existing lock it cannot verify is
     unprotected, the same principle as /22's own never-overwrite-a-password rule.
     FAIL-CLOSED review fix 2 (internal tracker LIVE run 2 -- corrects this module's OWN earlier
     "no credential member touched" claim and the NEM-MERT note below): asc_setLockStructure(true) is NOT
     credential-free by itself -- the engine fills the stored workbookAlgorithmName/SaltValue/SpinCount/HashValue
     members as a side effect, and asc_isPassword() reads true afterward. On {protect: true} this handler therefore
     WRITES those four members back to null immediately after (set-only, the value is never read), then reads
     wb.workbookProtection.asc_isPassword() back; if still true, the SAME transaction reverts the lock and the
     handler returns credential-not-cleared -- zero persisted credential-bearing lock, either way.
  3. sheet.table.write/v1's totalsRow property now accepts value:false (not just true): the SAME asc_changeFormatTableInfo
     call, direction-aware already-at-value / header-row-missing / verify checks, the receipt observed field carries the
     ref shrinking back (GH #335, measured LIVE through the editor API, not yet the Bridge: internal tracker LIVE run 1,
     docs/evidence/gh335-totals-header-live-1-20261102.json). This is the ONLY change to an already-shipped handler in
     this build; every other /23 handler, and the layout installer, stays byte-identical (proven by the companion static
     test, named exception sheet.table.write/v1).

RESOLVED (was NEM-MERT in the first candidate cut): whether the engine exposes a workbook-level credential detector
the same shape as sheet.protection.write/v1's own asc_isPassword() (design engine finding 4). The internal tracker
LIVE measurement above confirms it DOES exist (wb.workbookProtection.asc_isPassword()) -- this candidate now uses it as
the post-clear check in fix 2. workbook.protection.write/v1 still names no credential-protected refusal on the
REMOVE side ({protect: false}) -- that axis remains covered by credential-unobservable (fail-closed on "currently
locked", not on "carries a credential") -- only the {protect: true} side gained the detector-backed check, because
that is the side the second G1 finding showed actually needed it.

The handler sources are the readable, multi-line handler-wb-protection-read-v24.txt / handler-wb-protection-write-v24.txt
files next to this file; this builder joins each file's lines with single spaces (the block is a single-line fragment)
and inserts them, plus the manifest entries and the totalsRow substitutions, with exact string substitutions, each
asserted to occur exactly once -- the same pattern as build-rev23.py. No timestamps, no environment input. It NEVER
writes a Document Server file and has NO --sdk mode: the served sdk-all.js and its apply are a separate, gated operator
step (maintainer; the disposable DS first; the live DS needs its own owner GO).

NOT MEASURED in this unit (design gates G1-G3; the card's own scope: offline build + static identity + handler tests
only -- no LIVE run, no product-side runtime.cjs wiring, no W2 guard, no installation):
  - that workbook.protection.write/v1's lockStructure change reaches the server-side save and survives a fresh reopen
    (engine finding 3, G1's own job)
  - the credential-protected gap above
  - sheet.table.write/v1's totalsRow:false Bridge path on the real engine (the editor-API LIVE run already measured
    the engine side cleanly; this handler routes the SAME asc_changeFormatTableInfo call through the Bridge, not yet
    run live)

Usage:
  build-rev24.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""

import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V23_BLOCK = HERE.parent / "rev23-candidate" / "bridge-block.js"
V23_BLOCK_SHA = "9f46f0370d5b20175fc88d869b05937842c3e7f94277c584ab686e06a3308ab0"
V23_BLOCK_SIZE = 154010
READ_HANDLER_SRC = HERE / "handler-wb-protection-read-v24.txt"
WRITE_HANDLER_SRC = HERE / "handler-wb-protection-write-v24.txt"

WB_R_REASONS = "bad-args, model-missing, protection-api-missing, protection-flag-not-boolean"
WB_W_REASONS = ("bad-args, model-missing, protection-api-missing, credential-unobservable, credential-not-cleared, "
                "history-missing, history-off, setter-threw, no-effect")

WB_R_ENTRY = (
    '{ id: "workbook.protection.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "no arguments (an empty object is also accepted): {lockStructure} -- the workbook own structure '
    'protection flag (wbModel.getWorkbookProtection(c_oAscWorkbookProtectType.lockStructure)), workbook-wide, not per '
    'sheet; lockWindows and lockRevisions are not read, out of scope. READ ONLY; never returns a password, a password '
    'hash, a salt or a spin count: only the one named boolean field is read, no object is ever enumerated; served only '
    'through read(); reasons: ' + WB_R_REASONS + ' (GH #338-#342, /24)" }, '
)
WB_W_ENTRY = (
    '{ id: "workbook.protection.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "{protect: true} locks the workbook structure PASSWORDLESS, {protect: false} unlocks it -- props '
    'from the Asc.CWorkbookProtection default constructor plus asc_setLockStructure(protect) only; a LIVE run measured that '
    'this ONE call is not credential-free by itself (it fills the stored algorithm/salt/spin/hash members as a side '
    'effect), so on {protect: true} the handler immediately WRITES those four members back to null (set-only, never '
    'read, never returned) and then verifies wb.workbookProtection.asc_isPassword() is false before returning -- '
    'credential members are CLEARED to null, never read, not merely untouched; '
    'wbModel.setProtectedWorkbook(props, true) in one history point plus transaction; a target state already '
    'held is {applied: [], satisfied: true} with NO history; read-after-write through getWorkbookProtection(lockStructure), '
    'a mismatch is no-effect; the result is a receipt {applied: [lockStructure]}, the caller verifies with '
    'workbook.protection.read/v1; the caller must send this op as the SOLE step of its task, task admission at the '
    'product layer, not this handler; whether the engine exposes a FURTHER, narrower workbook-level credential '
    'detector (the same shape as sheet.protection.write/v1 own asc_isPassword(), beyond the one now used for the '
    'post-clear check above) is NOT sourced -- this handler therefore names no credential-protected refusal on the '
    'REMOVE side, but FAILS CLOSED on the one case that detector would gate: {protect: false} on a '
    'workbook CURRENTLY locked is refused credential-unobservable, before any history point or setter call, zero '
    'engine mutation -- the same never-remove-a-lock-you-cannot-verify principle sheet.protection.write/v1 own '
    'credential-protected enforces for the sheet axis, reworded because this axis has no detector to name; if the '
    'post-clear check on {protect: true} still finds a password, the SAME transaction reverts the lock before '
    'returning credential-not-cleared -- zero persisted credential-bearing lock; '
    '{protect: true} on an unlocked workbook proceeds normally, there being no existing lock to preserve; served only '
    'through write(); reasons: ' + WB_W_REASONS + ' (GH #338-#342, /24)" }, '
)

OLD_TOTALS_VALIDATE = ('else if (prop === "totalsRow") { if (typeof args.value !== "boolean") return { ok: false, '
                        'reason: "value-not-boolean" }; if (args.value !== true) return { ok: false, reason: '
                        '"totals-off-not-measured" }; } else if (prop === "name")')
NEW_TOTALS_VALIDATE = ('else if (prop === "totalsRow") { if (typeof args.value !== "boolean") return { ok: false, '
                        'reason: "value-not-boolean" }; } else if (prop === "name")')

OLD_TOTALS_CLASSIFY = ('else if (prop === "totalsRow") { if (t0.headerRow !== true) return { ok: false, reason: '
                        '"header-row-missing" }; if (t0.totalsRow === true) return { ok: false, reason: '
                        '"already-at-value" }; callArgs = [args.table, optConst, true]; } else if (prop === ')
NEW_TOTALS_CLASSIFY = ('else if (prop === "totalsRow") { if (args.value === true && t0.headerRow !== true) return '
                        '{ ok: false, reason: "header-row-missing" }; if (t0.totalsRow === args.value) return { ok: '
                        'false, reason: "already-at-value" }; callArgs = [args.table, optConst, args.value]; } else '
                        'if (prop === ')

OLD_TOTALS_VERIFY = ('else if (prop === "totalsRow") { observed = t1.totalsRow; hit = observed === true; same = '
                      'observed === t0.totalsRow; } else if (prop === "name")')
NEW_TOTALS_VERIFY = ('else if (prop === "totalsRow") { observed = t1.totalsRow; hit = observed === args.value; same = '
                      'observed === t0.totalsRow; } else if (prop === "name")')

OLD_TABLE_REASONS = "value-not-boolean, totals-off-not-measured, name-invalid"
NEW_TABLE_REASONS = "value-not-boolean, name-invalid"

OLD_TOTALS_CONTRACT = ('property totalsRow {value: true only; OFF is not measured and is refused} '
                        '(asc_changeFormatTableInfo, the engine own default functions; the engine INSERTS the totals '
                        'row, it does not overwrite);')
NEW_TOTALS_CONTRACT = ('property totalsRow {value: boolean} (asc_changeFormatTableInfo, the engine own default '
                        'functions; true INSERTS the totals row, it does not overwrite; false REMOVES it, measured '
                        'through the editor API, internal tracker);')

MANIFEST_END = '(GH #356, #357, /23)" }, ], };'
CAPS_START = 'var EO_BRIDGE_CAPS = { '
WRITES_START = 'var EO_BRIDGE_WRITES = { '
BLOCK_ENTRY = 'Api.GetEuroOfficeBridge = function () {'
MARKERS = ['"workbook.protection.read/v1": function', '"workbook.protection.write/v1": function',
           '{ id: "workbook.protection.read/v1"', '{ id: "workbook.protection.write/v1"']
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
    v23 = V23_BLOCK.read_bytes()
    if sha(v23) != V23_BLOCK_SHA or len(v23) != V23_BLOCK_SIZE:
        sys.exit(f"REFUSE: /23 block sha/size {sha(v23)}/{len(v23)} != {V23_BLOCK_SHA}/{V23_BLOCK_SIZE}; NO OUTPUT")
    impl_r = one_line(READ_HANDLER_SRC)
    impl_w = one_line(WRITE_HANDLER_SRC)
    t = v23.decode("utf-8")
    if t.count(BLOCK_ENTRY) != 1:
        sys.exit("REFUSE: the /23 block does not carry the bridge entrypoint exactly once; NO OUTPUT")
    subs = [
        ('patchVersion: "euro-office-bridge-v1/9.3.5.0/23"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/24"'),
        (MANIFEST_END, '(GH #356, #357, /23)" }, ' + WB_R_ENTRY + WB_W_ENTRY + '], };'),
        (CAPS_START, CAPS_START + impl_r),
        (WRITES_START, WRITES_START + impl_w),
        (OLD_TOTALS_VALIDATE, NEW_TOTALS_VALIDATE),
        (OLD_TOTALS_CLASSIFY, NEW_TOTALS_CLASSIFY),
        (OLD_TOTALS_VERIFY, NEW_TOTALS_VERIFY),
        (OLD_TABLE_REASONS, NEW_TABLE_REASONS),
        (OLD_TOTALS_CONTRACT, NEW_TOTALS_CONTRACT),
    ]
    for old, new in subs:
        n = t.count(old)
        if n != 1:
            sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:70]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1:
            sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("totals-off-not-measured") != 0:
        sys.exit("REFUSE: the retired reason name 'totals-off-not-measured' still present; NO OUTPUT")
    return v23, t.encode("utf-8")


def main():
    v23, v24 = build_block()
    out = {"v23_block_sha": sha(v23), "v23_block_size": len(v23), "v24_block_sha": sha(v24), "v24_block_size": len(v24)}
    (HERE / "bridge-block.js").write_bytes(v24)
    (HERE / "old.txt").write_bytes(v23)
    (HERE / "new.txt").write_bytes(v24)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v24-gh338-gh342-workbook-protection-plus-gh335-totals-off",
                             "old": v23.decode("utf-8"), "new": v24.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (maintainer; the disposable "
                              "DS first, the live DS needs its own owner GO).",
                 original_sha256=sha(v23), patched_sha256=sha(v24), original_size=len(v23), patched_size=len(v24))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
