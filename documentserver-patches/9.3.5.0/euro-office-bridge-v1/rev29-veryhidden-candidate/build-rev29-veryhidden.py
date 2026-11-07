#!/usr/bin/env python3
"""
GH #365 option C, 1st unit (U1 -- card an internal branch, kioszto zero, OWNER-GO: the owner "C GO" owner channel, rogzitve 2026-11-06 07:59, valasz 49439) -- OFFLINE build of the euro-office-bridge-v1
patchVersion /29 CANDIDATE for DS 9.3.5, in ONE deterministic pass from the /28 candidate block
(../rev28-table-autofilter-criteria-candidate/bridge-block.js, sha pinned).

/29 = /28 + ONE new READ capability -- no handler touched, no capability removed, nothing else
added. Companion to this unit's own 3-step ENGINE REPLACE (see plan-rev29-veryhidden.py, old1..3.txt/
new1..3.txt next to this script): the bridge-block ADDITION here is independent plumbing -- the
new capability only returns a sensible, non-always-false boolean once the engine REPLACE steps
land (ws.worksheet.bEoVeryHidden is a field the REPLACE steps themselves introduce). This build
does not require the REPLACE steps to have been applied to validate itself: the handler reads
`m.bEoVeryHidden` defensively (undefined on an unpatched engine reads as not-true, never throws).

CAPABILITY: "sheet.veryHidden.read/v1" (GH #365 option C).

THE MEASURED FACT THIS CAPABILITY STARTS FROM (docs/GH365-VERYHIDDEN-SHEET-STATE-SOURCE-CONTRACT-
20261105.md "3. EGYSEG", prod code-read of sdk-all.js /27, e39658a4): the binary-reader-level loss
that collapses 'hidden' and 'veryHidden' into the SAME `oWorksheet.bHidden=!0` happens because
`EVisibleType.visibleVeryHidden` shares a `case` with `EVisibleType.visibleHidden`, setting no
other field. This unit's own 3-step REPLACE (R1 in the binary reader, R2 in the binary writer, R3
in `setHidden`) introduces a SECOND, sibling field on the SAME worksheet object -- `bEoVeryHidden`
-- that this handler reads. `bHidden` alone is UNCHANGED semantics (still true for either hidden
state); the handler's OWN job is only to report the combination, never to re-derive visibility
itself.

    "sheet.veryHidden.read/v1": { sheet: <sheet name> } -> { ok: true, value: { veryHidden: bool } }

veryHidden is true ONLY when BOTH `ws.worksheet.bHidden === true` AND `ws.worksheet.bEoVeryHidden
=== true` -- a visible sheet (bHidden false/undefined) or a plain-hidden sheet (bHidden true,
bEoVeryHidden not true) both read `veryHidden: false`, never a guess, never a third state: the
card's own stated contract is a single boolean, nothing else attempted (no range, no "reason" on
the happy path -- only on a precondition failure, matching every other READ handler's idiom in
this file: bad-args / sheet-not-found / worksheet-missing / threw).

    build-rev29-veryhidden.py [--check]

Run with --check to only verify the anchors are unique and print the would-be output_sha256,
without writing bridge-block.js. Without --check it writes bridge-block.js and
manifest-mechanism-proof.json next to this script.
"""
import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REV28_BLOCK = HERE.parent / "rev28-table-autofilter-criteria-candidate" / "bridge-block.js"
OUT_BLOCK = HERE / "bridge-block.js"
OUT_PROOF = HERE / "manifest-mechanism-proof.json"

NEW_CAPABILITY_ID = "sheet.veryHidden.read/v1"
OLD_PATCH_VERSION = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/28"'
NEW_PATCH_VERSION = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/29"'

# Inserted right before the manifest's `capabilities` array closes -- the SAME "plain append, last
# entry" convention build-rev28-table-autofilter-criteria.py already uses for this anchor.
MANIFEST_ANCHOR = '], }; function argsOk(args, keys) {'
MANIFEST_ENTRY = (
    ' { id: "sheet.veryHidden.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one sheet name: whether it is specifically veryHidden (not merely hidden) -- '
    'true only when the worksheet model carries BOTH bHidden===true AND bEoVeryHidden===true, the '
    'second field introduced by this unit\'s own 3-step engine REPLACE (binary reader/writer + '
    'setHidden) that stops folding EVisibleType.visibleVeryHidden into the same case as visibleHidden. '
    'A visible or plain-hidden sheet reads veryHidden:false, never a guess; on an engine where the '
    'REPLACE steps have not been applied, bEoVeryHidden is simply undefined and this reads false too '
    '(never throws). GH #365 option C, docs/GH365-VERYHIDDEN-SHEET-STATE-SOURCE-CONTRACT-20261105.md '
    '\\"3. EGYSEG\\"." },'
)

# Right after the previous /28 build's own handler-append point: the literal text immediately
# before `var EO_BRIDGE_WRITES = {` on the ACTUAL /28 output (verified by this script's own
# transform() below, not assumed) -- NOT the pre-insertion `}, }; var EO_BRIDGE_WRITES = {` anchor
# build-rev28-table-autofilter-criteria.py used (that exact text no longer occurs once two handlers
# were spliced into the gap: the output is `...value: out }; },}; var EO_BRIDGE_WRITES = {`, no
# space between the two closing tokens). Re-deriving a stale anchor from an EARLIER candidate is
# exactly the bug the /28 CAPS_END_ANCHOR fix (a maintainer, 2026-11-06, see that script's own comment) was
# about -- so this script's own --check mode fails loudly (ANCHOR_NOT_UNIQUE) rather than silently
# re-landing outside the object if the real text ever differs from what's hardcoded here.
#
# The anchor is split in two: CAPS_END_PREFIX (`},` -- the PREVIOUS handler's own closing `}` for
# its function body, plus the trailing comma after that whole entry) and CAPS_END_SUFFIX (`};
# var EO_BRIDGE_WRITES = {` -- EO_BRIDGE_CAPS's own closing brace onward). The new handler is
# spliced BETWEEN the two, never before the prefix -- a first version of this script matched the
# WHOLE anchor and prepended the new handler in front of it, which lands the handler BEFORE the
# previous entry's own closing `}`, i.e. INSIDE that handler's function body (caught by this
# script's own syntax_check(): "Unexpected token ':'", the SAME class of bug as the /28
# CAPS_END_ANCHOR fix, just one token to the left of it).
CAPS_END_PREFIX = '},'
CAPS_END_SUFFIX = '}; var EO_BRIDGE_WRITES = {'
CAPS_END_ANCHOR = CAPS_END_PREFIX + CAPS_END_SUFFIX

HANDLER_ENTRY = (
    ' "sheet.veryHidden.read/v1": function (args) {'
    ' if (!args || typeof args.sheet !== "string") return { ok: false, reason: "bad-args" };'
    ' var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" };'
    ' var m; try { m = ws.worksheet; } catch (e) { return { ok: false, reason: "worksheet-missing" }; }'
    ' if (!m) return { ok: false, reason: "worksheet-missing" };'
    ' var h; try { h = m.bHidden; } catch (e) { return { ok: false, reason: "threw" }; }'
    ' var vh; try { vh = m.bEoVeryHidden; } catch (e) { vh = null; }'
    ' return { ok: true, value: { veryHidden: h === true && vh === true } };'
    ' },'
)


def sha(b):
    return hashlib.sha256(b).hexdigest()


def transform(data):
    if OLD_PATCH_VERSION not in data:
        return None, "PATCH_VERSION_MARKER_NOT_FOUND"
    if data.count(MANIFEST_ANCHOR) != 1:
        return None, f"MANIFEST_ANCHOR_NOT_UNIQUE count={data.count(MANIFEST_ANCHOR)}"
    if data.count(CAPS_END_ANCHOR) != 1:
        return None, f"CAPS_END_ANCHOR_NOT_UNIQUE count={data.count(CAPS_END_ANCHOR)}"
    out = data.replace(OLD_PATCH_VERSION, NEW_PATCH_VERSION, 1)
    out = out.replace(MANIFEST_ANCHOR, MANIFEST_ENTRY + MANIFEST_ANCHOR, 1)
    # the new handler goes BETWEEN the prior entry's trailing `,` (CAPS_END_PREFIX) and
    # EO_BRIDGE_CAPS's own closing `};` (CAPS_END_SUFFIX) -- NOT before the whole anchor (that
    # would land the handler inside the PREVIOUS entry's own function body; see this script's
    # comment above the anchor definition).
    out = out.replace(CAPS_END_ANCHOR, CAPS_END_PREFIX + HANDLER_ENTRY + CAPS_END_SUFFIX, 1)
    if out.count(f'"{NEW_CAPABILITY_ID}"') < 2:  # once in the manifest id, once as the handler key
        return None, "NEW_CAPABILITY_NOT_PRESENT_AFTER_TRANSFORM"
    return out, None


def syntax_check(js_text):
    """Same composed-block syntax gate as build-rev28-table-autofilter-criteria.py's own
    syntax_check() (array-wrapped `node --check`, since the fragment is a trailing-comma
    expression sequence, not standalone JS) -- duplicated rather than imported, matching that
    script's own precedent of a small, self-contained build script."""
    try:
        r = subprocess.run(["node", "--check"], input=f"[{js_text}]", capture_output=True, text=True, timeout=30)
    except (OSError, subprocess.SubprocessError) as e:
        return False, f"NODE_CHECK_COULD_NOT_RUN: {e}"
    if r.returncode != 0:
        err_line = next((ln for ln in r.stderr.splitlines() if "Error" in ln), None)
        return False, err_line or (r.stderr.strip().splitlines()[-1] if r.stderr.strip() else f"node --check exited {r.returncode}")
    return True, None


def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument("--check", action="store_true", help="verify anchors + print output_sha256, write nothing")
    a = p.parse_args(argv)
    data = REV28_BLOCK.read_bytes().decode("utf-8")
    out, err = transform(data)
    if err:
        print(f"REFUSE: {err}; NOTHING WRITTEN", file=sys.stderr)
        return 2
    syntax_ok, syntax_detail = syntax_check(out)
    if not syntax_ok:
        print(f"REFUSE: COMPOSED_BLOCK_SYNTAX_ERROR: {syntax_detail}; NOTHING WRITTEN", file=sys.stderr)
        return 2
    input_sha256 = sha(data.encode("utf-8"))
    output_sha256 = sha(out.encode("utf-8"))
    if a.check:
        print(json.dumps({"input_sha256": input_sha256, "output_sha256": output_sha256, "would_write": str(OUT_BLOCK)}))
        return 0
    OUT_BLOCK.write_text(out, encoding="utf-8")
    proof = {
        "format": "euro-office-bridge-v1-candidate-insert",
        "name": "euro-office-bridge-v1",
        "runtime": "9.3.5.0",
        "base": "rev28-table-autofilter-criteria-candidate/bridge-block.js",
        "base_sha256": input_sha256,
        "output_sha256": output_sha256,
        "new_capabilities": [NEW_CAPABILITY_ID],
        "getters": [
            f'"{NEW_CAPABILITY_ID}": function',
            f'{{ id: "{NEW_CAPABILITY_ID}"',
        ],
        "operations": [
            {"id": "patch-version-marker", "old": OLD_PATCH_VERSION, "new": NEW_PATCH_VERSION},
            {"id": "manifest-capability-entry", "anchor": MANIFEST_ANCHOR, "inserted": MANIFEST_ENTRY},
            {"id": "handler-entry", "anchor": CAPS_END_ANCHOR, "inserted": HANDLER_ENTRY},
        ],
        "_comment": (
            "GH #365 option C U1 candidate, derived by build-rev29-veryhidden.py from the real "
            "rev28-table-autofilter-criteria-candidate/bridge-block.js (base_sha256 pinned above). "
            "NOT installed anywhere. The companion 3-step engine REPLACE (R1/R2/R3) that makes "
            "bEoVeryHidden a real, non-undefined field lives in plan-rev29-veryhidden.py "
            "(old1..3.txt/new1..3.txt), combined with THIS bridge-block in one install plan -- "
            "this script alone only adds the READ plumbing."
        ),
    }
    OUT_PROOF.write_text(json.dumps(proof, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"block": str(OUT_BLOCK), "proof": str(OUT_PROOF), "input_sha256": input_sha256, "output_sha256": output_sha256}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
