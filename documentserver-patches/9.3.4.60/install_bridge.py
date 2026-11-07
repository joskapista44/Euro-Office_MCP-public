#!/usr/bin/env python3
"""
Fresh install of the Euro-Office DS Bridge on an unpatched Euro-Office Document Server 9.3.4.60 bundle (M12).

    install_bridge.py --target cell|word --input <sdk-all.js> --output <path> (--check | --apply)

The plan (fresh-install.json, next to this file) names, per target: the exact input bundle hash (the UNPATCHED
Euro-Office 9.3.4.60 bundle), a default anchor, the blocks (each with its hash, optionally with its own anchor) and the
exact output hash. The script inserts each block once right before its anchor (blocks sharing an anchor in plan order)
and accepts the result only if its hash is the plan's output hash. Anything else is refused and nothing is written:

    input already == output        ALREADY_PATCHED   (exit 0, nothing written)
    input hash unknown             UNKNOWN_INPUT     (exit 2) -- a different DS build is never patched blindly
    block file hash differs        BLOCK_MISMATCH    (exit 2)
    any anchor count != 1          ANCHOR_NOT_UNIQUE (exit 2)
    result hash != output_sha256   RESULT_MISMATCH   (exit 2)

GH #362 4th unit: a plan's "blocks" list may ALSO carry REPLACE steps -- {"old": <text>, "new": <text>}, no
"file" -- alongside the existing INSERT steps ({"file": <path>, "sha256": <hash>, "anchor": <optional>}). The
two step shapes are told apart by which keys they carry (an entry has either "file" or "old", never both); a
plan with ZERO replace steps produces BYTE-IDENTICAL output to before this change (the replace loop below is a
no-op on an empty list, exactly as untested as the rest of the file was). Replace steps run FIRST, in plan
order, each against the THEN-CURRENT buffer (so a later replace can see an earlier one's result): its "old"
text must occur EXACTLY ONCE, else refused, named REPLACE_OLD_NOT_UNIQUE (this single status covers BOTH "old"
missing entirely [0 occurrences] and "old" ambiguous [2 or more] -- a chosen text that could match more than
one place, or no place, is equally un-actionable, same fail-closed spirit as ANCHOR_NOT_UNIQUE). Only once
every replace step has applied do the existing anchor-insertion blocks run, on the POST-replace buffer -- an
insertion anchor is never expected to overlap a replace step's own text in practice, but running replacements
first keeps the single input_sha256 -> output_sha256 chain meaningful end to end, the same discipline the
existing insertion-only chain already has.

    replace "old" not found, or found more than once    REPLACE_OLD_NOT_UNIQUE (exit 2)

--check never writes; --apply writes to --output atomically (temporary file + rename).
"""

import argparse
import hashlib
import json
import os
from pathlib import Path
import sys
import tempfile

HERE = Path(__file__).resolve().parent


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def plan_result(plan, data, blocks):
    """(status, output bytes or None). Pure: no file system access. `blocks` carries the read bytes of the plan's
    INSERT steps only (those with "file"), in the same relative order as those entries appear in plan["blocks"] --
    REPLACE steps ("old"/"new", no "file") carry their own text inline in the plan, nothing to read from disk, and
    are NOT represented in `blocks`. Each insert block is placed right before its anchor (the block's own "anchor",
    else the target's); blocks sharing an anchor keep plan order (the anchor ends up after the last one). Replace
    steps run FIRST, in plan order, each against the then-current buffer; see the module docstring for the full
    REPLACE_OLD_NOT_UNIQUE rule. A plan with no replace steps behaves exactly as before this function grew the
    replace pass (the loop below is then a no-op, `out` starts as `data` unchanged)."""
    if sha256(data) == plan["output_sha256"]:
        return "ALREADY_PATCHED", None
    if sha256(data) != plan["input_sha256"]:
        return "UNKNOWN_INPUT", None
    insert_steps = [p for p in plan["blocks"] if "file" in p]
    replace_steps = [p for p in plan["blocks"] if "old" in p]
    if len(blocks) != len(insert_steps) or any(sha256(b) != p["sha256"] for b, p in zip(blocks, insert_steps)):
        return "BLOCK_MISMATCH", None
    out = data
    for r in replace_steps:
        old, new = r["old"].encode("utf-8"), r["new"].encode("utf-8")
        if out.count(old) != 1:
            return "REPLACE_OLD_NOT_UNIQUE", None
        out = out.replace(old, new)
    groups = {}
    for b, p in zip(blocks, insert_steps):
        groups.setdefault(p.get("anchor", plan["anchor"]).encode("utf-8"), []).append(b)
    if any(out.count(anchor) != 1 for anchor in groups):
        return "ANCHOR_NOT_UNIQUE", None
    for anchor in sorted(groups, key=out.index, reverse=True):
        at = out.index(anchor)
        out = out[:at] + b"".join(groups[anchor]) + out[at:]
    if sha256(out) != plan["output_sha256"]:
        return "RESULT_MISMATCH", None
    return "PATCHABLE", out


def main(argv=None):
    p = argparse.ArgumentParser(description="Fresh install of the Euro-Office DS Bridge (9.3.4.60).")
    p.add_argument("--target", required=True, choices=["cell", "word"])
    p.add_argument("--input", required=True)
    p.add_argument("--output")
    p.add_argument("--plan", default=str(HERE / "fresh-install.json"))
    mode = p.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true")
    mode.add_argument("--apply", action="store_true")
    a = p.parse_args(argv)

    plan = json.loads(Path(a.plan).read_text(encoding="utf-8"))[a.target]
    data = Path(a.input).read_bytes()
    base = Path(a.plan).resolve().parent
    blocks = [(base / b["file"]).read_bytes() for b in plan["blocks"] if "file" in b]
    status, out = plan_result(plan, data, blocks)
    print(f"TARGET={a.target} INPUT_SHA256={sha256(data)} STATUS={status} PATCH_VERSION={plan['patchVersion']}")
    if status == "ALREADY_PATCHED":
        return 0
    if status != "PATCHABLE":
        return 2
    if a.check:
        return 0
    if not a.output:
        print("--apply needs --output", file=sys.stderr)
        return 64
    dest = Path(a.output)
    fd, tmp = tempfile.mkstemp(prefix=".install-bridge-", dir=str(dest.parent))
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(out)
        if sha256(Path(tmp).read_bytes()) != plan["output_sha256"]:
            print("STATUS=WRITE_VERIFY_FAILED", file=sys.stderr)
            return 2
        os.replace(tmp, dest)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    print(f"WRITTEN={dest} OUTPUT_SHA256={plan['output_sha256']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
