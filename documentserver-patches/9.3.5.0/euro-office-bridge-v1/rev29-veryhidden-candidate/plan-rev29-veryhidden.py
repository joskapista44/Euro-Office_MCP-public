#!/usr/bin/env python3
"""
GH #365 option C, U1 (card an internal branch, kioszto zero, OWNER-GO: the owner "C GO" owner channel,
rogzitve 2026-11-06 07:59, valasz 49439) -- derive the euro-office-bridge-v1 /29 fresh-install PLAN
for DS 9.3.5, EXTENDING the /28 Bridge block (this directory's own bridge-block.js, built by the
sibling build-rev29-veryhidden.py) with THREE engine REPLACE steps, in the SAME plan (ONE
input_sha256 -> output_sha256 chain, ONE install_bridge.py --apply) -- same "UT 2" precedent as
plan-rev27-xlnm-fix.py (docs/GH362-PRINT-TITLES-PREFIX-RESEARCH-20261105.md point 3).

    plan-rev29-veryhidden.py --input <unpatched sdk-all.js> [--out fresh-install-rev29-veryhidden.json]

The fix (docs/GH365-VERYHIDDEN-SHEET-STATE-SOURCE-CONTRACT-20261105.md "3. EGYSEG"): the binary
reader's `case EVisibleType.visibleHidden:case EVisibleType.visibleVeryHidden:` shares ONE branch,
folding both OOXML states into the SAME `oWorksheet.bHidden=!0` -- the distinction is lost going
INTO the JS engine. Three REPLACE steps, run in order against the SAME buffer (R2's old text does
not occur inside R1's span, and vice versa -- independently placed call sites in the binary
reader/writer and in Worksheet.prototype.setHidden):

  R1 (binary reader, c_oSerWorksheetPropTypes_State): the shared case splits in two; ONLY
     visibleVeryHidden also sets the new `oWorksheet.bEoVeryHidden=!0` sibling field -- the plain
     hidden case is otherwise byte-identical to before.
  R2 (binary writer): on save, `1==ws.bHidden` now branches a second time on `ws.bEoVeryHidden` to
     choose which byte to write back -- visibleVeryHidden only if the sibling field is still true,
     else the old visibleHidden/visibleVisible choice, UNCHANGED.
  R3 (Worksheet.prototype.setHidden): un-hiding a sheet (`hidden` falsy) clears the sibling field --
     without this, a later re-hide through the SAME runtime session would write visibleVeryHidden
     for a sheet that was only ever plain-hidden through the public API (the sibling field would
     still read true from an EARLIER veryHidden state the sheet no longer carries).

`old1..3.txt` / `new1..3.txt` (next to this script) are read verbatim, never retyped inline -- the
SAME convention plan-rev27-xlnm-fix.py's old.txt/new.txt use. Each `old` is asserted to occur
EXACTLY ONCE in the real, unpatched 9.3.5.0 cell bundle, same as install_bridge.py's own
REPLACE_OLD_NOT_UNIQUE gate at apply time; this script does not trust its own transform() alone --
output_sha256 is cross-checked against install_bridge.py's OWN plan_result() (imported directly,
same process) before anything is written.

It starts from the shipped 9.3.5.0 plan (../../fresh-install.json, the /16 one), swaps in THIS
directory's own bridge-block.js (built by build-rev29-veryhidden.py, sha pinned), prepends the
three replace steps (replace steps run FIRST, in plan order, per install_bridge.py's own module
docstring), and writes the plan with the resulting output_sha256. The plan is then used unchanged
by the existing tool:

    install_bridge.py --target cell --plan fresh-install-rev29-veryhidden.json --input <sdk-all.js> --output <path> --check|--apply

Nothing here writes a Document Server file. Like every earlier candidate in this family (/27-xlnm-
fix, /28-table-autofilter-criteria), this agent has no access to the real unpatched sdk-all.js --
output_sha256 for the REAL engine is never claimed here; only the mechanism (transform() agreeing
with install_bridge.py's plan_result()) is proven, against whatever --input the eventual runner
supplies.
"""
import argparse
import hashlib
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE_PLAN = HERE.parent.parent / "fresh-install.json"
REV29_BLOCK = "euro-office-bridge-v1/rev29-veryhidden-candidate/bridge-block.js"
# same convention as plan-rev28-xlnm-fix.py's own LAYER: the CLEAN "/NN" marker, matching the
# bridge-block's own embedded patchVersion string (build-rev29-veryhidden.py's NEW_PATCH_VERSION),
# not the directory's own descriptive suffix (a maintainer's review, internal tracker, named in that
# build script's own comment).
LAYER = "euro-office-bridge-v1/9.3.5.0/29"
INSTALL_BRIDGE = HERE.parent.parent.parent / "9.3.4.60" / "install_bridge.py"

# the card's own MI A KESZ item 3: this plan's replace chain is "/29 blokk + a #362 REPLACE +
# R1-R3" -- the GH #362 Print_Titles save-branch fix (plan-rev27-xlnm-fix.py's own old.txt/new.txt,
# a SIBLING candidate forked from rev27-candidate, never folded into the rev28/rev29 bridge-block
# lineage) runs FIRST, then R1/R2/R3 in order. Read verbatim from that candidate's own files, never
# retyped here -- the SAME single-source-of-truth discipline old1..3.txt/new1..3.txt already use.
XLNM_FIX_DIR = HERE.parent / "rev27-xlnm-print-titles-candidate"

REPLACE_STEPS = ["1", "2", "3"]  # old1.txt/new1.txt .. old3.txt/new3.txt, applied AFTER the xlnm fix


def sha(b):
    return hashlib.sha256(b).hexdigest()


def transform(data, replace_pairs, default_anchor, insert_entries, read_block):
    """The SAME replace-then-insert two-pass transform install_bridge.py's plan_result() performs
    (and plan-rev27-xlnm-fix.py's own transform() already duplicates for a single replace step) --
    here generalised to an ORDERED LIST of replace steps, each asserted unique against the
    THEN-CURRENT buffer (so R2 sees R1's result, R3 sees R2's), matching install_bridge.py's own
    per-step REPLACE_OLD_NOT_UNIQUE rule exactly. `insert_entries` is the FULL list of insert-step
    plan entries (not just the one this script modified) -- every existing block in the base plan
    must still be inserted."""
    out = data
    for old_text, new_text in replace_pairs:
        if out.count(old_text) != 1:
            return None, f"REPLACE_OLD_NOT_UNIQUE count={out.count(old_text)}"
        out = out.replace(old_text, new_text)
    groups = {}
    for p in insert_entries:
        groups.setdefault(p.get("anchor", default_anchor).encode("utf-8"), []).append(read_block(p["file"]))
    if any(out.count(anchor) != 1 for anchor in groups):
        return None, "ANCHOR_NOT_UNIQUE"
    for anchor in sorted(groups, key=out.index, reverse=True):
        at = out.index(anchor)
        out = out[:at] + b"".join(groups[anchor]) + out[at:]
    return out, None


def derive(base, data, replace_pairs, read_block):
    cell = json.loads(json.dumps(base["cell"]))
    if sha(data) != cell["input_sha256"]:
        return None, f"UNKNOWN_INPUT {sha(data)} (want the unpatched {cell['input_sha256']})"
    hits = [b for b in cell["blocks"] if b["file"].startswith("euro-office-bridge-v1/")]
    if len(hits) != 1:
        return None, f"BRIDGE_BLOCK_COUNT {len(hits)}"
    hits[0]["file"], hits[0]["layer"] = REV29_BLOCK, LAYER
    hits[0]["sha256"] = sha(read_block(REV29_BLOCK))
    cell["patchVersion"] = LAYER
    # the FOUR replace steps are FIRST in plan order (install_bridge.py runs every replace step
    # before any insert, and in plan order among themselves): the GH #362 Print_Titles fix first
    # (the card's own MI A KESZ item 3 states "/29 blokk + a #362 REPLACE + R1-R3" in this order),
    # then R1/R2/R3. Text is read verbatim from the xlnm-fix candidate's own old.txt/new.txt and
    # this directory's own old1..3.txt/new1..3.txt, never retyped here.
    replace_blocks = [{"old": o.decode("utf-8"), "new": n.decode("utf-8")} for o, n in replace_pairs]
    cell["blocks"] = replace_blocks + cell["blocks"]
    cell["_comment"] = ("GH #365 option C U1 /29-veryhidden candidate, derived by plan-rev29-veryhidden.py "
                         "from the real unpatched input; the Bridge block is rev29-veryhidden-candidate/"
                         "bridge-block.js (/28 + sheet.veryHidden.read/v1), PLUS FOUR engine replace steps, "
                         "per the card's own MI A KESZ item 3 (\"/29 blokk + a #362 REPLACE + R1-R3\"): the "
                         "GH #362 Print_Titles save-branch fix (plan-rev27-xlnm-fix.py's own old.txt/new.txt, "
                         "read from that sibling candidate, not duplicated), then the GH #365 binary reader "
                         "(R1), binary writer (R2) and Worksheet.prototype.setHidden (R3) -- docs/GH365-"
                         "VERYHIDDEN-SHEET-STATE-SOURCE-CONTRACT-20261105.md \"3. EGYSEG\". output_sha256 is "
                         "CROSS-CHECKED against install_bridge.py's own plan_result() before being written "
                         "(see main()). The XML reader path (CT_Sheet, \"veryHidden\"===val) is deliberately "
                         "NOT touched -- whether the x2t server-side converter even reaches it on open is "
                         "unmeasured; the card names this as a the production host-probe decision, U3 if needed.")
    insert_entries = [b for b in cell["blocks"] if "file" in b]
    out, err = transform(data, replace_pairs, cell["anchor"], insert_entries, read_block)
    if err:
        return None, err
    return {"cell": cell}, out


def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True)
    p.add_argument("--out", default=str(HERE.parent.parent / "fresh-install-rev29-veryhidden.json"))
    a = p.parse_args(argv)
    base = json.loads(BASE_PLAN.read_text(encoding="utf-8"))
    data = Path(a.input).read_bytes()
    read_block = lambda f: (BASE_PLAN.parent / f).read_bytes()
    xlnm_pair = (
        (XLNM_FIX_DIR / "old.txt").read_bytes(),
        (XLNM_FIX_DIR / "new.txt").read_bytes(),
    )
    replace_pairs = [xlnm_pair] + [
        ((HERE / f"old{i}.txt").read_bytes(), (HERE / f"new{i}.txt").read_bytes())
        for i in REPLACE_STEPS
    ]
    plan, out_or_err = derive(base, data, replace_pairs, read_block)
    if plan is None:
        print(f"REFUSE: {out_or_err}; NOTHING WRITTEN", file=sys.stderr)
        return 2
    plan["cell"]["output_sha256"] = sha(out_or_err)

    # CROSS-CHECK: install_bridge.py's OWN plan_result(), imported directly (same process, no
    # subprocess ambiguity), must independently reach PATCHABLE with the EXACT same bytes this
    # script's own transform() produced. A divergence here means this script's transform() and the
    # installer's plan_result() disagree, and NOTHING is written.
    sys.path.insert(0, str(INSTALL_BRIDGE.parent))
    from install_bridge import plan_result  # noqa: E402  (deliberate late import, path set just above)
    insert_blocks = [read_block(b["file"]) for b in plan["cell"]["blocks"] if "file" in b]
    status, out2 = plan_result(plan["cell"], data, insert_blocks)
    if status != "PATCHABLE" or out2 != out_or_err:
        print(f"REFUSE: install_bridge.py's plan_result() disagrees (status={status}); NOTHING WRITTEN", file=sys.stderr)
        return 2

    Path(a.out).write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"plan": a.out, "input_sha256": plan["cell"]["input_sha256"], "output_sha256": plan["cell"]["output_sha256"], "patchVersion": LAYER}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
