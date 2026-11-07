#!/usr/bin/env python3
"""
GH #332-#335 (Bridge /19, DESIGN docs/DS-9.3.5-BRIDGE19-TABLE-WRITE-DESIGN.md) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /19 CANDIDATE for DS 9.3.5, in ONE deterministic pass from the /18 CANDIDATE
block (../rev18-candidate/bridge-block.js, sha pinned).

/19 = /18 + ONE new capability, WRITE:
  sheet.table.write/v1  -- ONE property of ONE Excel Table per call, through the editor's own table actions
                           (the spreadsheet_api calls the UI uses): style | flag | totalsRow | name | ref.

The handler source is the readable, multi-line handler-table-write.txt next to this file; this builder joins its lines
with single spaces (the block is a single-line fragment) and inserts it, plus the manifest entry, with exact string
substitutions, each asserted to occur exactly once -- the same pattern as build-rev18.py. No timestamps, no
environment input. It NEVER writes a Document Server file and has NO --sdk mode: the served sdk-all.js and its
apply are a separate, gated operator step (maintainer; the disposable DS first; the live DS needs its own owner GO).

NOT MEASURED (the design's first gate, on the disposable DS): that an editor table action called from inside a
Bridge write() lands in history and persists. The handler's read-after-write (through the /18 tables read) turns
every outcome into a named refusal or a verified success, but it cannot prove persistence across save+reopen --
the caller's fresh read does.

Usage:
  build-rev19.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V18_BLOCK = HERE.parent / "rev18-candidate" / "bridge-block.js"
V18_BLOCK_SHA = "3ed6e041f423b0b23663a4ab25cf21fab989e254469d4edf0afd768f85b21694"
V18_BLOCK_SIZE = 110803
HANDLER_SRC = HERE / "handler-table-write.txt"

REASONS = ("bad-args, unsupported-property, style-unknown, flag-unknown, value-not-boolean, totals-off-not-measured, "
           "name-invalid, ref-unparsable, editor-missing, <api>-missing, c_oAscChangeFilterOptions-missing, "
           "c_oAscChangeTableStyleInfo-missing, edit-not-allowed, workbook-locked, workbook-protected, sheet-not-found, "
           "sheet-ambiguous, tables-read-failed:<reason>, table-not-found, table-ambiguous, already-at-value, "
           "table-style-info-missing, header-row-missing, name-taken, range-unreadable, ref-moves-origin, ref-too-small, "
           "ref-overlaps, <call>-threw, <call>-error, table-count-changed, other-table-changed, no-effect, "
           "<property>-unexpected, target-collateral-change")

TABLE_WRITE_ENTRY = (
    '{ id: "sheet.table.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one Excel Table (sheet name + table DisplayName), ONE property per call, through the editor own '
    'table actions: property style {name: one of TableStyleLight1..21, TableStyleMedium1..28, TableStyleDark1..11} '
    '(asc_changeAutoFilter); property flag {flag: rowStripes, columnStripes, firstColumn or lastColumn, value: boolean} '
    '(asc_changeFormatTableInfo); property totalsRow {value: true only; OFF is not measured and is refused} '
    '(asc_changeFormatTableInfo, the engine own default functions; the engine INSERTS the totals row, it does not '
    'overwrite); property name {name: the new DisplayName} (asc_changeDisplayNameTable; a structured reference to the '
    'table is NOT rewritten by the engine -- the CALLER must refuse while one exists); property ref {ref: the new A1 '
    'range, same top-left cell, header row kept, no overlap with another table of the sheet} (asc_changeTableRange; the '
    'stored values of dependent structured references can stay stale -- the CALLER must refuse while one exists). The '
    'header row ON/OFF is NOT a property (OFF deletes the header cell content, measured). Every argument is type-gated, '
    'no coercion, no default; only the named fields are read, no object is ever enumerated; the edit permission, the '
    'global lock and the workbook protection are checked BEFORE the call; the table is read back through the sheet.tables.'
    'read/v1 path right after the call: success = the requested value observed, every other table unchanged and no other '
    'field of the target changed (beyond the one the property implies) and no editor error; an unchanged value is the named '
    'refusal no-effect, a value already at the target is already-at-value (no call is made); the result is an operation '
    'receipt {applied:[property], table, observed} and the caller verifies with a fresh read after save; never retried or '
    'rolled back inside; no fallback; served only through write(); reasons: ' + REASONS + ' (GH #332-#335, /19)" }, '
)

MANIFEST_END = '(GH #298 D, /18)" }, ], };'
WRITES_END = '}; var EO_BRIDGE_HYDRATIONS = {'
MARKERS = ['"sheet.table.write/v1": function', '{ id: "sheet.table.write/v1"']
# the "no general internal API" / "named fields only" discipline: the handler text must never enumerate an object
# generically, and it must carry no line comment (the source is joined onto one line).
FORBIDDEN_IN_HANDLER = ("Object.keys", "Object.values", "Object.entries", "getOwnPropertyNames", "for (var k in", "for (var p in", " in args", "//")


def sha(b): return hashlib.sha256(b).hexdigest()


def handler_text():
    lines = [ln.strip() for ln in HANDLER_SRC.read_text(encoding="utf-8").split("\n")]
    text = " ".join(ln for ln in lines if ln) + " "
    for word in FORBIDDEN_IN_HANDLER:
        if word in text:
            sys.exit(f"REFUSE: the handler text contains {word!r} -- named fields only, no line comments; NO OUTPUT")
    return text


def build_block():
    v18 = V18_BLOCK.read_bytes()
    if sha(v18) != V18_BLOCK_SHA or len(v18) != V18_BLOCK_SIZE:
        sys.exit(f"REFUSE: /18 block sha/size {sha(v18)}/{len(v18)} != {V18_BLOCK_SHA}/{V18_BLOCK_SIZE}; NO OUTPUT")
    impl = handler_text()
    subs = [
        ('patchVersion: "euro-office-bridge-v1/9.3.5.0/18"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/19"'),
        (MANIFEST_END, '(GH #298 D, /18)" }, ' + TABLE_WRITE_ENTRY + '], };'),
        (WRITES_END, impl + WRITES_END),
    ]
    t = v18.decode("utf-8")
    for old, new in subs:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("(GH #332-#335, /19)") != 1:
        sys.exit("REFUSE: (GH #332-#335, /19) count != 1; NO OUTPUT")
    return v18, t.encode("utf-8"), impl


def main():
    v18, v19, impl = build_block()
    out = {"v18_block_sha": sha(v18), "v18_block_size": len(v18), "v19_block_sha": sha(v19), "v19_block_size": len(v19)}
    (HERE / "bridge-block.js").write_bytes(v19)
    (HERE / "old.txt").write_bytes(v18)
    (HERE / "new.txt").write_bytes(v19)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v19-gh332-gh335-table-write", "old": v18.decode("utf-8"), "new": v19.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (maintainer; the disposable "
                              "DS first, the live DS needs its own owner GO).",
                 original_sha256=sha(v18), patched_sha256=sha(v19), original_size=len(v18), patched_size=len(v19))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
