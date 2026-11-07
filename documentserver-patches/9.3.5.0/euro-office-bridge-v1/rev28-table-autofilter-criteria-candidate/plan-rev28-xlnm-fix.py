#!/usr/bin/env python3
"""
GH #361+#363 (route probe route (c), extended) + GH #362 4th unit, COMBINED -- derive the euro-office-
bridge-v1 fresh-install PLAN that carries the /28 Bridge candidate (rev28-table-autofilter-criteria-candidate,
sheet.table.autoFilter.criteria.read/v1 + sheet.cells.compiledStyle.read/v1) AND the `_xlnm.Print_Titles`
save-fix (rev27-xlnm-print-titles-candidate/old.txt, new.txt) in ONE plan, so a single DS outage installs all
three (internal tracker): today the #362 fix lives ONLY in the /27 plan (plan-rev27-xlnm-fix.py) -- installing /28
as-is would drop it.

    plan-rev28-xlnm-fix.py --input <unpatched sdk-all.js> [--out fresh-install-rev28-xlnm-fix.json]

SAME two source scripts this one is derived from, read for comparison, never copied byte-for-byte:
  plan-rev28-table-autofilter-criteria.py  -- the /28 bridge-block swap (file, sha256, layer, patchVersion)
  plan-rev27-xlnm-fix.py                   -- the ONE REPLACE step + the plan_result() cross-check discipline

The replace step's old.txt/new.txt are read from their ORIGINAL location (../rev27-xlnm-print-titles-
candidate/, a SIBLING directory of this script, not this directory) -- never copied or retyped, so the two
candidates' own fix text can never drift apart. `old` must occur EXACTLY ONCE in the real, unpatched 9.3.5.0
cell bundle (REPLACE_OLD_NOT_UNIQUE otherwise, install_bridge.py's own rule). Same cross-check discipline as
plan-rev27-xlnm-fix.py: this script computes output_sha256 from its OWN transform() first, then CROSS-CHECKS
the result against install_bridge.py's OWN plan_result(), imported directly (same process) -- only once BOTH
independently agree is the plan written.

It starts from the shipped 9.3.5.0 plan (../../fresh-install.json, the /16 one -- the SAME base every other
plan-rev*.py script in this tree starts from), swaps in the /28 bridge block (this directory's OWN
bridge-block.js, read from here -- not copied), ADDS the ONE replace step, and writes the plan with the
resulting output_sha256. The plan is then used unchanged by the existing tool:

    install_bridge.py --target cell --plan fresh-install-rev28-xlnm-fix.json --input <sdk-all.js> --output <path> --check|--apply

Nothing here writes a Document Server file. This agent has no access to the real unpatched sdk-all.js, so the
REAL output_sha256 is never asserted or claimed from here -- the operator computes it when running this script
against the real file, the SAME discipline plan-rev27-xlnm-fix.py and plan-rev28-table-autofilter-criteria.py
already state about themselves (see e.g. internal trackers own commit message, daf63261...). Installing /28 and
the #362 fix together remains OUT OF SCOPE for the agent that wrote this plan tool (owner-GO and maintainer, same as
every prior rev in this directory).
"""
import argparse
import hashlib
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE_PLAN = HERE.parent.parent / "fresh-install.json"
REV28_BLOCK = "euro-office-bridge-v1/rev28-table-autofilter-criteria-candidate/bridge-block.js"
XLNM_DIR = HERE.parent / "rev27-xlnm-print-titles-candidate"
LAYER = "euro-office-bridge-v1/9.3.5.0/28"
INSTALL_BRIDGE = HERE.parent.parent.parent / "9.3.4.60" / "install_bridge.py"


def sha(b):
    return hashlib.sha256(b).hexdigest()


def transform(data, old_text, new_text, default_anchor, insert_entries, read_block):
    """The SAME two-pass transform install_bridge.py's plan_result() performs (replace, THEN group-and-insert
    by anchor, same anchor-sharing/plan-order rules), duplicated here ONLY to produce `out` bytes without yet
    knowing output_sha256 -- main() below then confirms this matches install_bridge.py's OWN, independently-
    computed result before writing anything. `insert_entries` is the FULL list of insert-step plan entries
    (every existing block in the base plan, PLUS the swapped-in /28 bridge block), not just the one this
    script itself modified."""
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
    hits[0]["file"], hits[0]["layer"] = REV28_BLOCK, LAYER
    hits[0]["sha256"] = sha(read_block(REV28_BLOCK))
    cell["patchVersion"] = LAYER
    # the replace step is FIRST in plan order (install_bridge.py runs every replace step before any insert) --
    # its text is read verbatim from old.txt/new.txt, never retyped here.
    cell["blocks"] = [{"old": old_text.decode("utf-8"), "new": new_text.decode("utf-8")}] + cell["blocks"]
    cell["_comment"] = ("GH #361+#363 /28 candidate PLUS GH #362 4th unit's Print_Titles save-fix, in ONE "
                         "input_sha256 -> output_sha256 chain, derived by plan-rev28-xlnm-fix.py from the real "
                         "unpatched input. The bridge block is the /28 candidate (rev28-table-autofilter-"
                         "criteria-candidate/bridge-block.js, unchanged from plan-rev28-table-autofilter-"
                         "criteria.py's own); the replace step is UNCHANGED from plan-rev27-xlnm-fix.py's own "
                         "(old.txt/new.txt, read from rev27-xlnm-print-titles-candidate/, not copied). "
                         "output_sha256 is CROSS-CHECKED against install_bridge.py's own plan_result() before "
                         "being written (see main()).")
    insert_entries = [b for b in cell["blocks"] if "file" in b]
    out, err = transform(data, old_text, new_text, cell["anchor"], insert_entries, read_block)
    if err:
        return None, err
    return {"cell": cell}, out


def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True)
    p.add_argument("--out", default=str(HERE.parent.parent / "fresh-install-rev28-xlnm-fix.json"))
    a = p.parse_args(argv)
    base = json.loads(BASE_PLAN.read_text(encoding="utf-8"))
    data = Path(a.input).read_bytes()
    old_text = (XLNM_DIR / "old.txt").read_bytes()
    new_text = (XLNM_DIR / "new.txt").read_bytes()
    # plan "file" entries are relative to BASE_PLAN's own directory (9.3.5.0/), e.g.
    # "euro-office-bridge-v1/rev28-table-autofilter-criteria-candidate/bridge-block.js" -- the SAME convention
    # every other plan-rev*.py script in this tree uses.
    read_block = lambda f: (BASE_PLAN.parent / f).read_bytes()
    plan, out_or_err = derive(base, data, old_text, new_text, read_block)
    if plan is None:
        print(f"REFUSE: {out_or_err}; NOTHING WRITTEN", file=sys.stderr)
        return 2
    plan["cell"]["output_sha256"] = sha(out_or_err)

    # CROSS-CHECK: install_bridge.py's OWN plan_result(), imported directly (same process, no subprocess
    # ambiguity), must independently reach PATCHABLE with the EXACT same bytes this script's own transform()
    # produced -- a divergence here means this script's transform() and the installer's plan_result() disagree,
    # and NOTHING is written.
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
