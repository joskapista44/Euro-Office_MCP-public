#!/usr/bin/env python3
"""
W2 regression fix (internal tracker, measured through two real production incidents, tracker card for this unit is
its own history) -- OFFLINE build of the euro-office-bridge-v1 patchVersion /25 CANDIDATE for DS 9.3.5, in ONE
deterministic pass from the /24 candidate block (../rev24-candidate/bridge-block.js, sha pinned -- the block
installed on production and then ROLLED BACK after the regression below).

/25 = /24 + EXACTLY ONE handler-body change, to workbook.protection.read/v1 -- no new capability, no capability
removed, no other handler touched:

MEASURED ENGINE SEMANTICS (the served sdk-all.js, sha 979067ee7aab = the DS build this fix targets):
  function CWorkbookProtection(wb){ this.lockStructure = null, ... }   // constructor default
  asc_getLockStructure(){ return this.lockStructure }
  getWorkbookProtection(type){ var res = false; if (this.workbookProtection) ... res = this.workbookProtection.asc_getLockStructure(); return res; }
So: workbookProtection object absent -> getWorkbookProtection returns false (boolean). workbookProtection object
PRESENT but its lockStructure attribute absent from the OOXML (an empty <workbookProtection/> element, the OOXML
default for a missing attribute) -> the constructor default null flows straight through asc_getLockStructure,
UNCHANGED by getWorkbookProtection's own `if (this.workbookProtection)` branch (that branch only decides WHETHER
to call the getter, not what it returns) -- the Bridge sees `v === null`, a THIRD, legitimate value the /24
handler never named. OOXML's own semantics for a present element with an absent lockStructure attribute: false
(ECMA-376 default). The engine's null is therefore "not locked", not "unknown".

THE REGRESSION (/24, 2026-11-03, installed 09:22, rolled back ~09:5x after alice's live report): the /24 handler's
`if (typeof v !== "boolean") return { ok:false, reason:"protection-flag-not-boolean" }` caught null too, because
null fails the same typeof check as every OTHER non-boolean shape (undefined, a number, a string, an object).
The W2 workbookGate (product side, unchanged by this unit, correctly fail-closed on any truly unobservable read)
then refused EVERY one of its five guarded sheet-structure operations on any workbook carrying a bare, attribute-
less <workbookProtection/> element -- a common, entirely unprotected shape -- because the Bridge collapsed "not
locked, unattributed" and "I cannot tell" into the identical wire shape.

THE FIX, here, and ONLY here: the handler now tests for `v === null` BEFORE the boolean check, and answers it as
the workbook's own semantics do -- {ok:true, value:{lockStructure:false}}. Every OTHER non-boolean value
(undefined if the engine ever returns it, a number, a string, a plain object) still answers protection-flag-not-
boolean, fail-closed, UNCHANGED from /24 -- this fix narrows the failure surface by exactly the one shape the
engine's own constructor default proves is not actually a failure; it does not touch the unobservable-on-anything-
else guarantee W2 depends on.

workbook.protection.write/v1 is NOT touched: the write handler's own asc_setLockStructure(protect) call sets a
real boolean, so the write side never produces a null read-back in the same way; this unit's scope is the READ
regression only, per its own card.

The handler source is the readable, multi-line handler-wb-protection-read-v25.txt file next to this file; this
builder joins its lines with single spaces (the block is a single-line fragment) and performs the manifest text
and handler-body substitutions with exact string matches, each asserted to occur exactly once -- the same pattern
as build-rev24.py. No timestamps, no environment input. It NEVER writes a Document Server file and has NO --sdk
mode: the served sdk-all.js and its apply are a separate, gated operator step (the the production host disposable DS first, the
live DS needs its own owner GO -- the regression above happened on the live DS directly, which is why this fix's
own rollout, unlike the candidate build itself, goes through the disposable DS window again before anything live).

NOT MEASURED in this unit (offline build + static identity + handler tests only -- no LIVE run):
  - that the fixed handler, once installed, actually clears the production incident on a real DS (the G1/G1b
    harness extension in this same unit adds the fixture; the LIVE run that proves it is a separate, gated step)
  - whether the engine's getWorkbookProtection ever returns `undefined` rather than `null` for the same "attribute
    absent" case on some other DS build -- this fix answers the SHAPE actually measured (null); an undefined
    shape, if it exists elsewhere, still falls through to protection-flag-not-boolean, fail-closed, unreviewed

Usage:
  build-rev25.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""

import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V24_BLOCK = HERE.parent / "rev24-candidate" / "bridge-block.js"
V24_BLOCK_SHA = "815da491c38d6ed237553398accb188c4b5125d06dc16c9807268477304bbe5d"
V24_BLOCK_SIZE = 161009
READ_HANDLER_SRC = HERE / "handler-wb-protection-read-v25.txt"

OLD_READ_HANDLER = (
    '"workbook.protection.read/v1": function (args) { if (args !== undefined && args !== null && typeof args !== '
    '"object") return { ok: false, reason: "bad-args" }; var wb; try { wb = (Asc.editor || {}).wbModel; } catch '
    '(e) { wb = undefined; } if (!wb || typeof wb.getWorkbookProtection !== "function") return { ok: false, '
    'reason: "model-missing" }; var T; try { T = Asc.c_oAscWorkbookProtectType; } catch (e) { T = undefined; } if '
    '(!T || typeof T.lockStructure !== "number") return { ok: false, reason: "protection-api-missing" }; var v; '
    'try { v = wb.getWorkbookProtection(T.lockStructure); } catch (e) { return { ok: false, reason: '
    '"protection-api-missing" }; } if (typeof v !== "boolean") return { ok: false, reason: '
    '"protection-flag-not-boolean" }; return { ok: true, value: { lockStructure: v } }; },'
)

OLD_WB_R_CONTRACT = (
    'semanticContract: "no arguments (an empty object is also accepted): {lockStructure} -- the workbook own '
    'structure protection flag (wbModel.getWorkbookProtection(c_oAscWorkbookProtectType.lockStructure)), '
    'workbook-wide, not per sheet; lockWindows and lockRevisions are not read, out of scope. READ ONLY; never '
    'returns a password, a password hash, a salt or a spin count: only the one named boolean field is read, no '
    'object is ever enumerated; served only through read(); reasons: bad-args, model-missing, '
    'protection-api-missing, protection-flag-not-boolean (GH #338-#342, /24)" }, '
)
NEW_WB_R_CONTRACT = (
    'semanticContract: "no arguments (an empty object is also accepted): {lockStructure} -- the workbook own '
    'structure protection flag (wbModel.getWorkbookProtection(c_oAscWorkbookProtectType.lockStructure)), '
    'workbook-wide, not per sheet; lockWindows and lockRevisions are not read, out of scope. A null engine value '
    '(the constructor default when the workbookProtection object exists but the OOXML lockStructure attribute is '
    'absent -- an empty <workbookProtection/> element) answers {lockStructure:false}, the OOXML default for a '
    'missing attribute, NOT a refusal; every OTHER non-boolean value (undefined, a number, a string, a plain '
    'object) is unchanged, fail-closed. READ ONLY; never returns a password, a password hash, a salt or a spin '
    'count: only the one named boolean field is read, no object is ever enumerated; served only through read(); '
    'reasons: bad-args, model-missing, protection-api-missing, protection-flag-not-boolean (GH #338-#342 regression '
    'fix, /25)" }, '
)

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
    v24 = V24_BLOCK.read_bytes()
    if sha(v24) != V24_BLOCK_SHA or len(v24) != V24_BLOCK_SIZE:
        sys.exit(f"REFUSE: /24 block sha/size {sha(v24)}/{len(v24)} != {V24_BLOCK_SHA}/{V24_BLOCK_SIZE}; NO OUTPUT")
    new_read = one_line(READ_HANDLER_SRC)
    if new_read.strip().rstrip(',').strip() == '':
        sys.exit("REFUSE: empty read handler; NO OUTPUT")
    t = v24.decode("utf-8")
    if t.count(BLOCK_ENTRY) != 1:
        sys.exit("REFUSE: the /24 block does not carry the bridge entrypoint exactly once; NO OUTPUT")
    subs = [
        ('patchVersion: "euro-office-bridge-v1/9.3.5.0/24"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/25"'),
        (OLD_READ_HANDLER, new_read.strip()),
        (OLD_WB_R_CONTRACT, NEW_WB_R_CONTRACT),
    ]
    for old, new in subs:
        n = t.count(old)
        if n != 1:
            sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:70]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1:
            sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if "protection-flag-not-boolean" not in t:
        sys.exit("REFUSE: the retained reason name 'protection-flag-not-boolean' missing; NO OUTPUT")
    if "GH #338-#342 regression fix, /25" not in t:
        sys.exit("REFUSE: the NEW /25 read-contract tag did not land; NO OUTPUT")
    return v24, t.encode("utf-8")


def main():
    v24, v25 = build_block()
    out = {"v24_block_sha": sha(v24), "v24_block_size": len(v24), "v25_block_sha": sha(v25), "v25_block_size": len(v25)}
    (HERE / "bridge-block.js").write_bytes(v25)
    (HERE / "old.txt").write_bytes(v24)
    (HERE / "new.txt").write_bytes(v25)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v25-w2-regression-fix-workbook-protection-read-null",
                             "old": v24.decode("utf-8"), "new": v25.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (the disposable DS "
                              "first, the live DS needs its own owner GO).",
                 original_sha256=sha(v24), patched_sha256=sha(v25), original_size=len(v24), patched_size=len(v25))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
