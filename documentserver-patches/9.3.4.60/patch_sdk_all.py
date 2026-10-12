#!/usr/bin/env python3

"""
Generalized from the chart-public-getters family's own patcher (a maintainer's dispatch via internal tracker/#ref: "a meglevo patch_sdk_all.py ALTALANOSITASA -- csaladonkenti manifest bemenet;
NE legyen masodik patcher"). ONE canonical patcher, shared by every 9.3.4.60 patch family
(chart-public-getters/, euro-office-bridge-v1/, ...); each family supplies its OWN
manifest.json via --manifest. The patching/classification LOGIC is unchanged from the
chart-getter family's own script -- only the manifest's SOURCE moved from a hardcoded,
same-directory path to an explicit CLI argument. Re-run against the chart-getter family's own
known inputs (original/patched/unrelated) to confirm the SAME classification as before this
change -- see the family's own acceptance record for that regression check.
"""

import argparse
import hashlib
import json
import os
from pathlib import Path
import sys
import tempfile


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


def load_manifest(manifest_path):
    return json.loads(Path(manifest_path).read_text(encoding="utf-8"))


def getter_inventory(text, manifest):
    return {
        getter: text.count(getter)
        for getter in manifest["getters"]
    }


def print_inventory(inventory, expected_count):
    for getter, count in inventory.items():
        status = "PASS" if count == expected_count else "FAIL"
        print(
            f"{status:4} {getter} "
            f"count={count} expected={expected_count}"
        )


def classify(data, manifest):
    digest = sha256_bytes(data)
    text = data.decode("utf-8")
    inventory = getter_inventory(text, manifest)

    if digest == manifest["patched_sha256"]:
        if all(count == 1 for count in inventory.values()):
            return "ALREADY_PATCHED", digest, inventory
        return "INCOMPATIBLE", digest, inventory

    if digest == manifest["original_sha256"]:
        if all(count == 0 for count in inventory.values()):
            return "PATCHABLE", digest, inventory
        return "INCOMPATIBLE", digest, inventory

    if all(count == 1 for count in inventory.values()):
        return "NATIVE_OR_UNKNOWN_SUPPORT", digest, inventory

    return "INCOMPATIBLE", digest, inventory


def build_patched(data, manifest):
    text = data.decode("utf-8")

    for operation in manifest["operations"]:
        old = operation["old"]
        new = operation["new"]

        count = text.count(old)

        if count != 1:
            raise RuntimeError(
                f'{operation["id"]}: expected old block exactly once, got {count}'
            )

        text = text.replace(old, new, 1)

    output = text.encode("utf-8")
    digest = sha256_bytes(output)

    if digest != manifest["patched_sha256"]:
        raise RuntimeError(
            "patched SHA mismatch: "
            f"expected {manifest['patched_sha256']}, got {digest}"
        )

    inventory = getter_inventory(text, manifest)

    bad = {
        getter: count
        for getter, count in inventory.items()
        if count != 1
    }

    if bad:
        raise RuntimeError(f"getter inventory mismatch: {bad}")

    return output


def atomic_write(path, data):
    path = Path(path)
    parent = path.parent

    original_stat = path.stat() if path.exists() else None

    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{path.name}.",
        suffix=".tmp",
        dir=parent,
    )

    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())

        if original_stat is not None:
            os.chmod(tmp_name, original_stat.st_mode)

            try:
                os.chown(
                    tmp_name,
                    original_stat.st_uid,
                    original_stat.st_gid,
                )
            except PermissionError:
                tmp_stat = os.stat(tmp_name)

                if (
                    tmp_stat.st_uid != original_stat.st_uid
                    or tmp_stat.st_gid != original_stat.st_gid
                ):
                    raise RuntimeError(
                        "cannot preserve target owner/group"
                    )

        os.replace(tmp_name, path)

        dir_fd = os.open(parent, os.O_RDONLY)

        try:
            os.fsync(dir_fd)
        finally:
            os.close(dir_fd)

    except Exception:
        try:
            os.unlink(tmp_name)
        except FileNotFoundError:
            pass
        raise


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Fail-closed EuroOffice DocumentServer 9.3.4.60 "
            "public getter / bridge patcher (generalized, "
            "family-manifest driven)"
        )
    )

    parser.add_argument(
        "target",
        type=Path,
        help="sdk-all.js file to inspect or patch",
    )

    parser.add_argument(
        "--manifest",
        type=Path,
        required=True,
        help=(
            "the patch family's own manifest.json "
            "(e.g. chart-public-getters/manifest.json or "
            "euro-office-bridge-v1/manifest.json)"
        ),
    )

    mode = parser.add_mutually_exclusive_group(required=True)

    mode.add_argument(
        "--check",
        action="store_true",
        help="classify only; never write",
    )

    mode.add_argument(
        "--apply",
        action="store_true",
        help="apply only to the exact known original SHA",
    )

    parser.add_argument(
        "--output",
        type=Path,
        help=(
            "write patched result here instead of replacing target; "
            "recommended for testing"
        ),
    )

    args = parser.parse_args()

    manifest = load_manifest(args.manifest)
    data = args.target.read_bytes()

    state, digest, inventory = classify(data, manifest)

    print("MANIFEST:", args.manifest)
    print("STATE:", state)
    print("SHA256:", digest)
    print("SIZE:", len(data))
    print()

    expected_count = 0 if state == "PATCHABLE" else 1
    print_inventory(inventory, expected_count)

    if args.check:
        if state in {
            "PATCHABLE",
            "ALREADY_PATCHED",
            "NATIVE_OR_UNKNOWN_SUPPORT",
        }:
            return 0
        return 3

    if state == "ALREADY_PATCHED":
        print()
        print("NO WRITE: exact accepted patched state already present")
        return 0

    if state == "NATIVE_OR_UNKNOWN_SUPPORT":
        print()
        print(
            "NO WRITE: all public getters exist, but this is an unknown SHA; "
            "live semantic acceptance is required"
        )
        return 2

    if state != "PATCHABLE":
        print()
        print("NO WRITE: unknown/incompatible input; fail-closed")
        return 3

    try:
        patched = build_patched(data, manifest)
    except Exception as exc:
        print()
        print("PATCH BUILD FAILED:", exc)
        return 4

    destination = args.output if args.output else args.target

    atomic_write(destination, patched)

    written = destination.read_bytes()
    written_sha = sha256_bytes(written)

    if written_sha != manifest["patched_sha256"]:
        print()
        print("WRITE VERIFY FAILED")
        return 5

    print()
    print("PATCHED:", destination)
    print("PATCHED SHA256:", written_sha)
    print("STATE: PATCHED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
