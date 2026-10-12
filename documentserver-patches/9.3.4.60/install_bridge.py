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
    """(status, output bytes or None). Pure: no file system access. Each block is inserted right before its anchor
    (the block's own "anchor", else the target's); blocks sharing an anchor keep plan order (the anchor ends up after
    the last one)."""
    if sha256(data) == plan["output_sha256"]:
        return "ALREADY_PATCHED", None
    if sha256(data) != plan["input_sha256"]:
        return "UNKNOWN_INPUT", None
    if len(blocks) != len(plan["blocks"]) or any(sha256(b) != p["sha256"] for b, p in zip(blocks, plan["blocks"])):
        return "BLOCK_MISMATCH", None
    groups = {}
    for b, p in zip(blocks, plan["blocks"]):
        groups.setdefault(p.get("anchor", plan["anchor"]).encode("utf-8"), []).append(b)
    if any(data.count(anchor) != 1 for anchor in groups):
        return "ANCHOR_NOT_UNIQUE", None
    out = data
    for anchor in sorted(groups, key=data.index, reverse=True):
        at = data.index(anchor)
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
    blocks = [(base / b["file"]).read_bytes() for b in plan["blocks"]]
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
