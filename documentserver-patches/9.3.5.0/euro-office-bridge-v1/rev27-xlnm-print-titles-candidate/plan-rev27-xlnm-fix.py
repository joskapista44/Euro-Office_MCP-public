#!/usr/bin/env python3
"""
GH #362 4th unit: derive the euro-office-bridge-v1 /27 fresh-install PLAN for DS 9.3.5, EXTENDED with the
`_xlnm.Print_Titles` save-fix (docs/GH362-PRINT-TITLES-PREFIX-RESEARCH-20261105.md, point 2) -- the SAME /27
Bridge block as plan-rev27.py's own candidate, PLUS one REPLACE step on the engine's own InitSaveManager save
branch, in the SAME plan (ONE input_sha256 -> output_sha256 chain, ONE install_bridge.py --apply, per that
document's point 3 "UT 2" decision -- no separate, independently-pinned patch tool).

    plan-rev27-xlnm-fix.py --input <unpatched sdk-all.js> [--out fresh-install-rev27-xlnm-fix.json]

The fix: `InitSaveManager.prototype.prepare`'s save branch names ONLY "Print_Area" when deciding whether to
re-add the `_xlnm.` prefix to a built-in defined name -- "Print_Titles" is not mentioned at all (the research
doc's own offset-verified finding: the literal "Print_Titles" string occurs 6x in the file, NONE inside this
branch). The replace step widens the ONE condition to also match "Print_Titles", byte-identical otherwise
(old.txt / new.txt next to this script, read verbatim -- never retyped inline, so a single source of truth
feeds both this script and its own offline test). `old` occurs EXACTLY ONCE in the real, unpatched 9.3.5.0
cell bundle (verified here AND by install_bridge.py's own REPLACE_OLD_NOT_UNIQUE gate at apply time -- this
script does not trust its own count or its own transform alone: it computes output_sha256 from its OWN
transform() first, then CROSS-CHECKS the result against install_bridge.py's OWN plan_result(), imported
directly (same process) -- only once BOTH independently agree is the plan written. A divergence between this
script's understanding of the plan and the installer's therefore can never produce a plan the installer itself
would then refuse at apply time).

It starts from the shipped 9.3.5.0 plan (../../fresh-install.json, the /16 one -- the SAME base plan-rev27.py
itself starts from), swaps in the SAME /27 Bridge block plan-rev27.py uses (this directory's own
../rev27-candidate/bridge-block.js -- read from there, not copied, so the two candidates can never drift), ADDS
the replace step, and writes the plan with the resulting output_sha256. The plan is then used unchanged by the
existing tool:

    install_bridge.py --target cell --plan fresh-install-rev27-xlnm-fix.json --input <sdk-all.js> --output <path> --check|--apply

Nothing here writes a Document Server file.
"""
import argparse
import hashlib
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE_PLAN = HERE.parent.parent / "fresh-install.json"
REV27_BLOCK = "euro-office-bridge-v1/rev27-candidate/bridge-block.js"
LAYER = "euro-office-bridge-v1/9.3.5.0/27-xlnm-print-titles-fix"
INSTALL_BRIDGE = HERE.parent.parent.parent / "9.3.4.60" / "install_bridge.py"


def sha(b):
    return hashlib.sha256(b).hexdigest()


def transform(data, old_text, new_text, default_anchor, insert_entries, read_block):
    """The SAME two-pass transform install_bridge.py's plan_result() performs (replace, THEN group-and-insert
    by anchor, same anchor-sharing/plan-order rules), duplicated here ONLY to produce `out` bytes without yet
    knowing output_sha256 (plan_result() needs the target hash up front, chicken-and-egg for a script that is
    DERIVING that hash) -- `main()` below then confirms this matches install_bridge.py's OWN, independently-
    computed result before writing anything. `insert_entries` is the FULL list of insert-step plan entries
    (e.g. BOTH the chart-read block at its own anchor AND the bridge block at the default anchor), not just the
    one this script itself modified -- every existing block in the base plan must still be inserted."""
    if data.count(old_text) != 1:
        return None, f"REPLACE_OLD_NOT_UNIQUE count={data.count(old_text)}"
    out = data.replace(old_text, new_text)
    groups = {}
    for p in insert_entries:
        groups.setdefault(p.get("anchor", default_anchor).encode("utf-8"), []).append(read_block(p["file"]))
    if any(out.count(anchor) != 1 for anchor in groups):
        return None, "ANCHOR_NOT_UNIQUE"
    for anchor in sorted(groups, key=out.index, reverse=True):
        at = out.index(anchor)
        out = out[:at] + b"".join(groups[anchor]) + out[at:]
    return out, None


def derive(base, data, old_text, new_text, read_block):
    cell = json.loads(json.dumps(base["cell"]))
    if sha(data) != cell["input_sha256"]:
        return None, f"UNKNOWN_INPUT {sha(data)} (want the unpatched {cell['input_sha256']})"
    hits = [b for b in cell["blocks"] if b["file"].startswith("euro-office-bridge-v1/")]
    if len(hits) != 1:
        return None, f"BRIDGE_BLOCK_COUNT {len(hits)}"
    hits[0]["file"], hits[0]["layer"] = REV27_BLOCK, LAYER
    hits[0]["sha256"] = sha(read_block(REV27_BLOCK))
    cell["patchVersion"] = LAYER
    # the replace step is FIRST in plan order (install_bridge.py runs every replace step before any insert, see
    # its own module docstring) -- its text is read verbatim from old.txt/new.txt, never retyped here.
    cell["blocks"] = [{"old": old_text.decode("utf-8"), "new": new_text.decode("utf-8")}] + cell["blocks"]
    cell["_comment"] = ("GH #362 4th unit /27-xlnm-print-titles-fix candidate, derived by plan-rev27-xlnm-fix.py "
                         "from the real unpatched input; the /27 Bridge block is UNCHANGED from plan-rev27.py's "
                         "own candidate, PLUS one Print_Titles save-branch replace (docs/GH362-PRINT-TITLES-"
                         "PREFIX-RESEARCH-20261105.md point 2). output_sha256 is CROSS-CHECKED against "
                         "install_bridge.py's own plan_result() before being written (see main()).")
    insert_entries = [b for b in cell["blocks"] if "file" in b]
    out, err = transform(data, old_text, new_text, cell["anchor"], insert_entries, read_block)
    if err:
        return None, err
    return {"cell": cell}, out


def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True)
    p.add_argument("--out", default=str(HERE.parent.parent / "fresh-install-rev27-xlnm-fix.json"))
    a = p.parse_args(argv)
    base = json.loads(BASE_PLAN.read_text(encoding="utf-8"))
    data = Path(a.input).read_bytes()
    old_text = (HERE / "old.txt").read_bytes()
    new_text = (HERE / "new.txt").read_bytes()
    # plan "file" entries are relative to BASE_PLAN's own directory (9.3.5.0/), e.g.
    # "euro-office-bridge-v1/rev27-candidate/bridge-block.js" -- the SAME convention plan-rev27.py uses.
    read_block = lambda f: (BASE_PLAN.parent / f).read_bytes()
    plan, out_or_err = derive(base, data, old_text, new_text, read_block)
    if plan is None:
        print(f"REFUSE: {out_or_err}; NOTHING WRITTEN", file=sys.stderr)
        return 2
    plan["cell"]["output_sha256"] = sha(out_or_err)

    # CROSS-CHECK: install_bridge.py's OWN plan_result(), imported directly (same process, no subprocess
    # ambiguity), must independently reach PATCHABLE with the EXACT same bytes this script's own transform()
    # produced -- now that output_sha256 is set, plan_result()'s own final-hash gate can pass. A divergence
    # here means this script's transform() and the installer's plan_result() disagree, and NOTHING is written.
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
