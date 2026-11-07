#!/usr/bin/env python3
"""
GH #300 L1 / #298 D (owner GO, the owner 2026-10-27, owner channel: "2: A ... 3: A" -- full READ/WRITE symmetry,
#298 D order A: read_tables first) -- OFFLINE build of the euro-office-bridge-v1 patchVersion /18 CANDIDATE for
DS 9.3.5, in ONE deterministic pass from the /17 CANDIDATE block (../rev17-candidate/bridge-block.js, sha pinned).

/18 = /17 + TWO new capabilities, READ ONLY:
  sheet.hyperlinks.read/v1  -- every hyperlink on one sheet: target address, internal location, screen-tip, kind
  sheet.tables.read/v1      -- every Excel Table on one sheet: name, range, header/totals rows, column names,
                               style flags, autofilter presence

Engine access (LIVE measured on DS 9.3.5, docs/evidence/gh300-gh298-engine-access-probe-20261027.json, v3+v4 --
the probe measured BOTH right after writing the seed and after a fresh close+reopen, byte-identical):
  hyperlinks: ws.worksheet.hyperlinkManager.getAll() -> an array of { bbox, data }; bbox.getAbsName() is the A1
              cell reference (with "$", stripped here); data (the Hyperlink payload) carries .Hyperlink (target
              URL), .Location (internal location/sub-address), .Tooltip (screen tip), and a zero-arg
              .getHyperlinkType() (1 measured as external). The DISPLAYED text is NOT part of this payload -- it
              is the cell's own value (read_range covers that already).
  tables:     ws.worksheet.TableParts -- an array of table objects: .DisplayName, .Ref (a range object with its
              own .getAbsName()), .isHeaderRow()/.isTotalsRow(), .TableColumns (array of {.Name}),
              .TableStyleInfo (.Name, .ShowRowStripes, .ShowColumnStripes, .ShowFirstColumn, .ShowLastColumn) or
              absent, .AutoFilter (presence only is read here -- coerced to a boolean, its own Ref/FilterColumns
              are NOT read: out of scope for this capability).

Owner conditions this handler is built to (same discipline as /17's credential guard, generalised): READ ONLY (no
setter, no editor action); CSAK NEVESÍTETT mezők -- every field is read by name, never by enumerating the object
(so an unnamed field -- e.g. a worksheet back-reference, the raw DataIntervalTree, FilterColumns -- can never
reach the output); a STRING-or-null field of the wrong type, or a boolean field of the wrong type, is a named
refusal, never a guessed value; a cap on link/table/column counts.

Deterministic: the /18 block is derived from the CANONICAL /17 candidate by THREE exact string substitutions,
each asserted to occur exactly once -- same pattern as build-rev17.py. No timestamps, no environment input. It
NEVER writes a Document Server file and has NO --sdk mode: the served sdk-all.js and its apply are a separate,
gated operator step (maintainer, disposable DS first; the live DS needs a separate owner GO).

Usage:
  build-rev18.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V17_BLOCK = HERE.parent / "rev17-candidate" / "bridge-block.js"
V17_BLOCK_SHA = "072bf10a3b9c2231959a4bcf26c10c116c584ca93134f00fbb4f2eb8e47933d6"
V17_BLOCK_SIZE = 101924

MAX_LINKS = 10000
MAX_TABLES = 1000
MAX_COLUMNS = 16384
STYLE_FLAGS = ("ShowRowStripes", "ShowColumnStripes", "ShowFirstColumn", "ShowLastColumn")

HYPERLINKS_REASONS = ("bad-args, sheet-not-found, worksheet-missing, hyperlink-manager-missing, hyperlinks-threw, "
                      "hyperlink-field-not-string:<name>, hyperlink-kind-not-number, range-unreadable, links-over-cap:"
                      + str(MAX_LINKS))

TABLES_REASONS = ("bad-args, sheet-not-found, worksheet-missing, table-parts-missing, tables-threw, "
                   "table-field-not-string:<name>, table-field-not-boolean:<name>, range-unreadable, tables-over-cap:"
                   + str(MAX_TABLES) + ", columns-over-cap:" + str(MAX_COLUMNS))

HYPERLINKS_ENTRY = (
    '{ id: "sheet.hyperlinks.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one sheet (sheet name): {links} -- links lists every hyperlink on the sheet, '
    'row-then-column order, each as {ref, address, location, tooltip, kind}; ref is the A1 cell reference '
    'without a dollar sign; address/location/tooltip are the target URL / internal location / screen tip '
    '(string or null); kind is external, internal, or unknown:<n>, from the engine own hyperlink-type code '
    '(1 measured as external). The displayed text is NOT part of this output -- it is the cell own value. '
    'READ ONLY; reads only the named fields, never enumerates the hyperlink payload object. reasons: '
    + HYPERLINKS_REASONS + ' (GH #300 L1, /18)" }, '
)

TABLES_ENTRY = (
    '{ id: "sheet.tables.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one sheet (sheet name): {tables} -- tables lists every Excel Table on the sheet, '
    'document order, each as {name, ref, headerRow, totalsRow, columns, style, autoFilter}; name is the '
    'table DisplayName; ref is the A1 range without a dollar sign; headerRow/totalsRow are booleans; columns '
    'is the ordered list of column names; style is {name, rowStripes, columnStripes, firstColumn, lastColumn} '
    '(name string or null, the rest booleans), or null when the table carries no style info; autoFilter is a '
    'boolean PRESENCE flag only -- its own range/filter detail is not read. READ ONLY; reads only the named '
    'fields, never enumerates a table object. reasons: ' + TABLES_REASONS + ' (GH #298 D, /18)" }, '
)

HYPERLINKS_IMPL = (
    '"sheet.hyperlinks.read/v1": function (args) { '
    'if (!args || typeof args.sheet !== "string") return { ok: false, reason: "bad-args" }; '
    'var w = Api.GetSheet(args.sheet); if (!w) return { ok: false, reason: "sheet-not-found" }; '
    'var m; try { m = w.worksheet; } catch (e) { return { ok: false, reason: "worksheet-missing" }; } '
    'if (!m) return { ok: false, reason: "worksheet-missing" }; '
    'var hm; try { hm = m.hyperlinkManager; } catch (e) { return { ok: false, reason: "hyperlink-manager-missing" }; } '
    'if (!hm || typeof hm.getAll !== "function") return { ok: false, reason: "hyperlink-manager-missing" }; '
    'var all; var allLen; try { all = hm.getAll(); allLen = all ? all.length : null; } catch (e) { return { ok: false, reason: "hyperlinks-threw" }; } '
    'if (!all || typeof allLen !== "number") return { ok: false, reason: "hyperlinks-threw" }; '
    'if (allLen > ' + str(MAX_LINKS) + ') return { ok: false, reason: "links-over-cap:' + str(MAX_LINKS) + '" }; '
    'var colNum = function (s) { var n = 0; for (var i = 0; i < s.length; i += 1) n = n * 26 + (s.charCodeAt(i) - 64); return n; }; '
    'var parseRef = function (r) { var first = r.indexOf(":") === -1 ? r : r.slice(0, r.indexOf(":")); var mm = /^([A-Z]+)(\\d+)$/.exec(first); return mm ? { col: mm[1], row: Number(mm[2]) } : { col: "", row: 0 }; }; '
    'var out = []; '
    'for (var i = 0; i < allLen; i += 1) { '
    'var el = all[i]; '
    'var bb; try { bb = el && el.bbox; } catch (e) { return { ok: false, reason: "range-unreadable" }; } '
    'var ref; try { ref = bb && typeof bb.getAbsName === "function" ? bb.getAbsName() : null; } catch (e) { return { ok: false, reason: "range-unreadable" }; } '
    'if (typeof ref !== "string") return { ok: false, reason: "range-unreadable" }; '
    'ref = ref.replace(/\\$/g, ""); '
    'var d; try { d = el && el.data; } catch (e) { return { ok: false, reason: "hyperlinks-threw" }; } '
    'if (!d) return { ok: false, reason: "hyperlinks-threw" }; '
    'var address; try { address = d.Hyperlink; } catch (e) { return { ok: false, reason: "hyperlinks-threw" }; } if (address !== null && address !== void 0 && typeof address !== "string") return { ok: false, reason: "hyperlink-field-not-string:address" }; '
    'var location; try { location = d.Location; } catch (e) { return { ok: false, reason: "hyperlinks-threw" }; } if (location !== null && location !== void 0 && typeof location !== "string") return { ok: false, reason: "hyperlink-field-not-string:location" }; '
    'var tooltip; try { tooltip = d.Tooltip; } catch (e) { return { ok: false, reason: "hyperlinks-threw" }; } if (tooltip !== null && tooltip !== void 0 && typeof tooltip !== "string") return { ok: false, reason: "hyperlink-field-not-string:tooltip" }; '
    'var kraw; try { kraw = typeof d.getHyperlinkType === "function" ? d.getHyperlinkType() : null; } catch (e) { return { ok: false, reason: "hyperlinks-threw" }; } '
    'if (typeof kraw !== "number") return { ok: false, reason: "hyperlink-kind-not-number" }; '
    'var kind = kraw === 1 ? "external" : (kraw === 2 ? "internal" : ("unknown:" + kraw)); '
    'out.push({ ref: ref, address: address === void 0 ? null : address, location: location === void 0 ? null : location, tooltip: tooltip === void 0 ? null : tooltip, kind: kind, _row: parseRef(ref).row, _col: colNum(parseRef(ref).col) }); '
    '} '
    'out.sort(function (a, b) { return a._row !== b._row ? a._row - b._row : a._col - b._col; }); '
    'for (var j = 0; j < out.length; j += 1) { delete out[j]._row; delete out[j]._col; } '
    'return { ok: true, value: { links: out } }; '
    '}, '
)

TABLES_IMPL = (
    '"sheet.tables.read/v1": function (args) { '
    'if (!args || typeof args.sheet !== "string") return { ok: false, reason: "bad-args" }; '
    'var w = Api.GetSheet(args.sheet); if (!w) return { ok: false, reason: "sheet-not-found" }; '
    'var m; try { m = w.worksheet; } catch (e) { return { ok: false, reason: "worksheet-missing" }; } '
    'if (!m) return { ok: false, reason: "worksheet-missing" }; '
    'var tp; try { tp = m.TableParts; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'var tpLen; try { tpLen = tp ? tp.length : null; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'if (!tp || typeof tpLen !== "number") return { ok: false, reason: "table-parts-missing" }; '
    'if (tpLen > ' + str(MAX_TABLES) + ') return { ok: false, reason: "tables-over-cap:' + str(MAX_TABLES) + '" }; '
    'var STYLE_FLAGS = ' + json.dumps(list(STYLE_FLAGS)) + '; '
    'var out = []; '
    'for (var i = 0; i < tpLen; i += 1) { '
    'var t = tp[i]; '
    'var name; try { name = t.DisplayName; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'if (typeof name !== "string") return { ok: false, reason: "table-field-not-string:name" }; '
    'var ref; try { var rg = t.Ref; ref = rg && typeof rg.getAbsName === "function" ? rg.getAbsName() : null; } catch (e) { return { ok: false, reason: "range-unreadable" }; } '
    'if (typeof ref !== "string") return { ok: false, reason: "range-unreadable" }; '
    'ref = ref.replace(/\\$/g, ""); '
    'var headerRow; try { headerRow = typeof t.isHeaderRow === "function" ? t.isHeaderRow() : null; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'if (typeof headerRow !== "boolean") return { ok: false, reason: "table-field-not-boolean:headerRow" }; '
    'var totalsRow; try { totalsRow = typeof t.isTotalsRow === "function" ? t.isTotalsRow() : null; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'if (typeof totalsRow !== "boolean") return { ok: false, reason: "table-field-not-boolean:totalsRow" }; '
    'var cols; var colsLen; try { cols = t.TableColumns; colsLen = cols ? cols.length : null; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'if (!cols || typeof colsLen !== "number") return { ok: false, reason: "tables-threw" }; '
    'if (colsLen > ' + str(MAX_COLUMNS) + ') return { ok: false, reason: "columns-over-cap:' + str(MAX_COLUMNS) + '" }; '
    'var columns = []; '
    'for (var j = 0; j < colsLen; j += 1) { '
    'var cn; try { cn = cols[j].Name; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'if (typeof cn !== "string") return { ok: false, reason: "table-field-not-string:columns[].Name" }; '
    'columns.push(cn); '
    '} '
    'var style = null; '
    'var tsi; try { tsi = t.TableStyleInfo; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'if (tsi) { '
    'var sname; try { sname = tsi.Name; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'if (sname !== null && sname !== void 0 && typeof sname !== "string") return { ok: false, reason: "table-field-not-string:style.name" }; '
    'var sv = {}; '
    'for (var k = 0; k < STYLE_FLAGS.length; k += 1) { '
    'var fv; try { fv = tsi[STYLE_FLAGS[k]]; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'if (typeof fv !== "boolean") return { ok: false, reason: "table-field-not-boolean:style." + STYLE_FLAGS[k] }; '
    'sv[STYLE_FLAGS[k]] = fv; '
    '} '
    'style = { name: sname === void 0 ? null : sname, rowStripes: sv.ShowRowStripes, columnStripes: sv.ShowColumnStripes, firstColumn: sv.ShowFirstColumn, lastColumn: sv.ShowLastColumn }; '
    '} '
    'var af; try { af = !!t.AutoFilter; } catch (e) { return { ok: false, reason: "tables-threw" }; } '
    'out.push({ name: name, ref: ref, headerRow: headerRow, totalsRow: totalsRow, columns: columns, style: style, autoFilter: af }); '
    '} '
    'return { ok: true, value: { tables: out } }; '
    '}, '
)

READ_END = '}; var EO_BRIDGE_WRITES = {'
MANIFEST_END = '(GH #326/#327, /17)" }, ], };'

# (old, new) -- each `old` must occur EXACTLY once in the EVOLVING text at the time it is applied.
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.5.0/17"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/18"'),
    (MANIFEST_END, '(GH #326/#327, /17)" }, ' + HYPERLINKS_ENTRY + TABLES_ENTRY + '], };'),
    # the READ dispatch object's closing anchor, immediately before "var EO_BRIDGE_WRITES = {" -- same
    # insertion-point convention as build-rev17.py / build-rev14.py.
    (READ_END, HYPERLINKS_IMPL + TABLES_IMPL + READ_END),
]
MARKERS = ['"sheet.hyperlinks.read/v1": function', '{ id: "sheet.hyperlinks.read/v1"',
           '"sheet.tables.read/v1": function', '{ id: "sheet.tables.read/v1"']
# the "no general internal API" / "named fields only" discipline, generalised from /17's credential check: the
# handler text must never enumerate an object generically (it would then be able to surface an unnamed field).
FORBIDDEN_IN_HANDLER = ("Object.keys", "Object.values", "Object.entries", "getOwnPropertyNames", "for (var k in", "for (var p in")


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v17 = V17_BLOCK.read_bytes()
    if sha(v17) != V17_BLOCK_SHA or len(v17) != V17_BLOCK_SIZE:
        sys.exit(f"REFUSE: /17 block sha/size {sha(v17)}/{len(v17)} != {V17_BLOCK_SHA}/{V17_BLOCK_SIZE}; NO OUTPUT")
    for word in FORBIDDEN_IN_HANDLER:
        if word in HYPERLINKS_IMPL or word in TABLES_IMPL:
            sys.exit(f"REFUSE: a handler text names/enumerates {word!r} -- it must read named fields only; NO OUTPUT")
    t = v17.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("(GH #300 L1, /18)") != 1:
        sys.exit("REFUSE: (GH #300 L1, /18) count != 1; NO OUTPUT")
    if t.count("(GH #298 D, /18)") != 1:
        sys.exit("REFUSE: (GH #298 D, /18) count != 1; NO OUTPUT")
    return v17, t.encode("utf-8")


def main():
    v17, v18 = build_block()
    out = {"v17_block_sha": sha(v17), "v17_block_size": len(v17), "v18_block_sha": sha(v18), "v18_block_size": len(v18)}
    (HERE / "bridge-block.js").write_bytes(v18)
    (HERE / "old.txt").write_bytes(v17)
    (HERE / "new.txt").write_bytes(v18)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v18-gh300-gh298-hyperlinks-tables-read", "old": v17.decode("utf-8"), "new": v18.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (maintainer; disposable DS "
                              "first, the live DS needs its own owner GO).",
                 original_sha256=sha(v17), patched_sha256=sha(v18), original_size=len(v17), patched_size=len(v18))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
