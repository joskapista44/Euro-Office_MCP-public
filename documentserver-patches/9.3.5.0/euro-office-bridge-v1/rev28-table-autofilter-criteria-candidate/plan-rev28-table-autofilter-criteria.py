#!/usr/bin/env python3
"""
GH #361 route probe, route (c) -- derive the euro-office-bridge-v1 /28 fresh-install PLAN for DS 9.3.5 from the
REAL unpatched sdk-all.js. Run by the operator who holds that file; this agent cannot read it, so the plan's
output hash is never written from here (it would be a fabricated measurement) -- the SAME constraint
plan-rev27.py states about itself.

    plan-rev28-table-autofilter-criteria.py --input <unpatched sdk-all.js> [--out fresh-install-rev28.json]

It starts from the shipped 9.3.5.0 plan (../../fresh-install.json, the /16 one -- the SAME base plan-rev27.py
starts from), swaps ONLY the bridge block for this directory's /28 candidate (file, sha256, layer,
patchVersion), refuses unless the input is that plan's known UNPATCHED input (input_sha256), inserts the
blocks exactly the way install_bridge.py's plan_result() does and writes the plan with the resulting
output_sha256. The plan is then used unchanged by the existing tool:

    install_bridge.py --target cell --plan fresh-install-rev28.json --input <sdk-all.js> --output <path> --check|--apply

Nothing here writes a Document Server file. Installing /28 is explicitly OUT OF SCOPE for the agent that wrote
this candidate (owner-GO and maintainer, same as every prior rev in this directory) -- this script exists so that,
WHEN that GO is given, the plan derivation follows the SAME audited mechanism every other rev already uses,
not a one-off hand-edit.
"""
import argparse, hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BASE_PLAN = HERE.parent.parent / "fresh-install.json"
REV28 = "euro-office-bridge-v1/rev28-table-autofilter-criteria-candidate/bridge-block.js"
# a maintainer's review (internal tracker, komment #ref): the directory name (rev28-table-autofilter-criteria-
# candidate) stays -- only the patchVersion marker itself is now the clean "/NN" form every earlier rev uses.
LAYER = "euro-office-bridge-v1/9.3.5.0/28"


def sha(b): return hashlib.sha256(b).hexdigest()


def derive(base, data, read_block):
    cell = json.loads(json.dumps(base["cell"]))
    if sha(data) != cell["input_sha256"]:
        return None, f"UNKNOWN_INPUT {sha(data)} (want the unpatched {cell['input_sha256']})"
    hits = [b for b in cell["blocks"] if b["file"].startswith("euro-office-bridge-v1/")]
    if len(hits) != 1:
        return None, f"BRIDGE_BLOCK_COUNT {len(hits)}"
    hits[0]["file"], hits[0]["layer"] = REV28, LAYER
    hits[0]["sha256"] = sha(read_block(REV28))
    cell["patchVersion"] = LAYER
    groups = {}
    for b in cell["blocks"]:
        groups.setdefault(b.get("anchor", cell["anchor"]).encode("utf-8"), []).append(read_block(b["file"]))
    if any(data.count(a) != 1 for a in groups):
        return None, "ANCHOR_NOT_UNIQUE"
    out = data
    for a in sorted(groups, key=data.index, reverse=True):
        at = data.index(a)
        out = out[:at] + b"".join(groups[a]) + out[at:]
    cell["output_sha256"] = sha(out)
    cell["_comment"] = ("GH #361 route probe route (c) /28 candidate plan, derived by "
                         "plan-rev28-table-autofilter-criteria.py from the real unpatched input; only the "
                         "bridge block differs from the /16 plan.")
    return {"cell": cell}, None


def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True)
    p.add_argument("--out", default=str(HERE.parent.parent / "fresh-install-rev28.json"))
    a = p.parse_args(argv)
    base = json.loads(BASE_PLAN.read_text(encoding="utf-8"))
    plan, err = derive(base, Path(a.input).read_bytes(), lambda f: (BASE_PLAN.parent / f).read_bytes())
    if err:
        print(f"REFUSE: {err}; NOTHING WRITTEN", file=sys.stderr)
        return 2
    Path(a.out).write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"plan": a.out, "input_sha256": plan["cell"]["input_sha256"], "output_sha256": plan["cell"]["output_sha256"], "patchVersion": LAYER}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
