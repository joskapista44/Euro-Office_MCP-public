#!/usr/bin/env python3
"""
GH #361 (route probe's route (c), docs/GH361-TABLE-AUTOFILTER-CRITERIA-ROUTE-20261105.md) -- OFFLINE build of
the euro-office-bridge-v1 patchVersion /28 CANDIDATE for DS 9.3.5, in ONE deterministic pass from the /27
candidate block (../rev27-candidate/bridge-block.js, sha pinned).

/28 = /27 + TWO new capabilities -- no handler touched, no capability removed, nothing else added. NOT
installed anywhere (this is the candidate; the plan file next to this script is for whoever later runs it
against the real unpatched sdk-all.js, per plan-rev27.py's own precedent -- this agent has no access to that
file, so output_sha256 for the REAL engine is never claimed here).

CAPABILITY 1: "sheet.table.autoFilter.criteria.read/v1" (GH #361).

THE MEASURED FACT THIS CAPABILITY STARTS FROM (the LANDED sheet.tables.read/v1 handler, byte-identical on /28,
verified by this build's own warrant test): `var af; try { af = !!t.AutoFilter; } catch (e) { ... }` -- a
TablePart's own `.AutoFilter` property is the internal object a table's own AutoFilter mode lives on.

*** PER-COLUMN CRITERIA ARE NOW MEASURED, NOT A GUESS -- a maintainer's LIVE runs against an OOXML-preseeded fixture
with a real table-owned AutoFilter (two columns: one value filter, one custom filter), card an internal branch
gh361-table-autofilter-criteria-bridge-candidate-20261105, comments #ref (first LIVE shape) and #ref (the
targeted, bounded dump this handler's field reads were derived from). *** The internal `.AutoFilter` object's
`.Ref` (same grounded-by-analogy hypothesis as the prior version: the SAME TablePart object already exposes
`.Ref` for its own cell range) and `.FilterColumns` (an array, each element: `ColId`, `Filters.{Blank,Values,
Dates}` for a value filter, `CustomFiltersObj.{And,CustomFilters}` for a custom filter, `ColorFilter`/
`DynamicFilter`/`Top10` for filter kinds never seen in either measured column) are now read field-by-field,
every access through try/catch, NEVER a bare property read that could throw unguarded.

*** WHAT IS STILL NOT MEASURED, AND NAMED AS SUCH -- NEVER FABRICATED: *** only ONE value-filter column and
ONE custom-filter column (with ONE condition) have ever been measured. Specifically unverified, each gets its
own named gap in the handler's own `unverified` array on the affected column, rather than a guessed value or a
silently-dropped field: `Filters.Dates` non-empty (date-filter semantics), `Filters.Blank` truthy (blank-value
semantics), `ColorFilter`/`DynamicFilter`/`Top10` truthy (entirely different, never-seen filter kinds),
`CustomFiltersObj.And === true` (how multiple custom conditions combine -- only `And: false` with a single
condition has been measured). The custom-filter `Operator` code's NUMBER is read (measured: `2` behaves as
"greater than" in the one case observed), but its NAME is never guessed from this codebase's own knowledge --
the handler tries, at RUNTIME, to read a real enum off the editor's own global (`Asc.c_oAscCustomAutoFilter`,
the ONE candidate zero named; this agent has no DS access to confirm it exists or holds the names) and only
emits `operatorName` when that lookup actually succeeds; otherwise `operatorName: null`, raw code still
returned, and the gap is the ABSENCE of a name, never an invented one.

CAPABILITY 2 (added, same /28 candidate -- EXTENDED, not a new rev29): "sheet.cells.compiledStyle.read/v1"
(GH #363 2nd unit, card an internal branch, kioszto zero).

THE MEASURED FACT THIS CAPABILITY STARTS FROM (card ALAP, LIVE on main 5954de37, docs/GH363-CF-EFFECTIVE-
READBACK-MEASUREMENT-20261105.md "2. egyseg", docs/evidence/gh363-cf-effective-readback-live-20261105/
internal-route-probe-v2-measured.json): `ws.worksheet.getCompiledStyle(row, col)` (0-based) returns the
conditional-format-EVALUATED style of a cell -- a firing cell (B7, row 6 col 1) measured `{ style: 'object',
fillColor: 16711680 }` (== the rule's own `#FF0000`); a non-firing cell (B2, row 1 col 1) measured
`{ style: 'null' }`. The row/col <- A1 conversion below is the SAME formula the probe used (verified against
those exact pairs), not a new derivation.

*** SCOPE, NAMED: ONLY the fill is read (`style.getFill()` -> `getColor()` or `getSolidFill()`, the probe's
own measured fallback order -- a maintainer's own suggestion, #ref/#ref). Every other compiled-style field (font,
border, number format, alignment, ...) is NOT attempted -- never guessed, never returned. A cell whose address
does not parse, or whose own read throws, carries its own named `reason`, never silently folded into a null
fill. The request is capped at 200 cells (`args.cells.length > 200`, matching read-tools.cjs's own pre-
dispatch `checkEffectiveFillBound`) -- defense in depth, not the only bound.

    build-rev28-table-autofilter-criteria.py [--check]

Run with --check to only verify the anchors are unique and print the would-be output_sha256, without writing
bridge-block.js. Without --check it writes bridge-block.js and manifest-mechanism-proof.json next to this
script.
"""
import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REV27_BLOCK = HERE.parent / "rev27-candidate" / "bridge-block.js"
OUT_BLOCK = HERE / "bridge-block.js"
OUT_PROOF = HERE / "manifest-mechanism-proof.json"

NEW_CAPABILITY_ID = "sheet.table.autoFilter.criteria.read/v1"
NEW_CAPABILITY_ID_2 = "sheet.cells.compiledStyle.read/v1"
OLD_PATCH_VERSION = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/27"'
# a maintainer's review (internal tracker, komment #ref): every earlier candidate (/25, /26, /27) advertises a CLEAN
# "/NN" patchVersion, even while its own directory is named "revNN-candidate" -- the directory suffix was
# never baked into the marker itself. This candidate's directory name stays
# rev28-table-autofilter-criteria-candidate; only the marker is now consistent with precedent.
NEW_PATCH_VERSION = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/28"'

# Inserted right before the manifest's `capabilities` array closes (`], }; function argsOk(args, keys) {`),
# i.e. as the LAST entry of the array -- a trailing comma on the PRIOR entry already exists (every entry ends
# with `},`), so this is a plain append.
MANIFEST_ENTRY = (
    ' { id: "sheet.table.autoFilter.criteria.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one Excel Table (sheet name + table DisplayName): whether it carries its own AutoFilter (present, from the SAME internal TablePart.AutoFilter object sheet.tables.read/v1 already reads truthily); if present, its stored range (Ref, same grounded-by-analogy hypothesis sheet.tables.read/v1 already uses for a TablePart\'s own range) and, per column, a fields array -- each entry a measured colId plus EITHER values (an array of selected value-filter keys) OR custom (and, conditions -- each condition an operator as a raw code, an operatorName ONLY when a runtime enum resolves it else null, and a value); a column of a kind never observed (color/dynamic/top10 filter) or whose observed-but-unverified sub-fields are non-default (Dates non-empty, Blank truthy, And:true) carries a named unverified array on that column instead of a guessed meaning. present:false = no AutoFilter on the table; range:null + reason:internal-shape-unverified:ref = the Ref hypothesis failed. GH #361 route probe, route (c); measured against a real, OOXML-preseeded fixture with a table-owned AutoFilter (one value filter, one custom filter)." },'
)

# GH #363 2nd unit: a second manifest entry, appended the SAME way (plain append, trailing comma already there).
MANIFEST_ENTRY_2 = (
    ' { id: "sheet.cells.compiledStyle.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "a named list of cells (sheet name + up to 200 A1 addresses): each cell\'s conditional-format-EVALUATED fill colour, read through the worksheet\'s own getCompiledStyle(row, col) getter (0-based, derived from the A1 address) -- ONLY the fill is read (style.getFill() -> getColor() or getSolidFill()); every other compiled-style field (font, border, number format, ...) is NOT attempted. A cell whose compiled style is null (the engine never produced one -- e.g. no rule evaluates to anything on it) reads compiled:null, never a guessed fill; a cell whose own address does not parse, or whose own read throws, carries its own named reason. GH #363 2nd unit, LIVE measured (docs/GH363-CF-EFFECTIVE-READBACK-MEASUREMENT-20261105.md): a firing cell\'s compiled fill equals the rule\'s own fill; a non-firing cell\'s compiled style is null." },'
)

# Inserted as the LAST entry of EO_BRIDGE_CAPS, before that object's own closing `};` -- the object's
# trailing comma pattern (`..., };`) makes this a plain append too.
HANDLER_ENTRY = (
    ' "sheet.table.autoFilter.criteria.read/v1": function (args) {'
    ' if (!args || typeof args.sheet !== "string" || typeof args.table !== "string") return { ok: false, reason: "bad-args" };'
    ' var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" };'
    ' var m; try { m = ws.worksheet; } catch (e) { return { ok: false, reason: "worksheet-missing" }; }'
    ' if (!m) return { ok: false, reason: "worksheet-missing" };'
    ' var tp; try { tp = m.TableParts; } catch (e) { return { ok: false, reason: "tables-threw" }; }'
    ' var tpLen; try { tpLen = tp ? tp.length : null; } catch (e) { return { ok: false, reason: "tables-threw" }; }'
    ' if (!tp || typeof tpLen !== "number") return { ok: false, reason: "table-parts-missing" };'
    ' var target = null;'
    ' for (var i = 0; i < tpLen; i += 1) {'
    '   var nm; try { nm = tp[i].DisplayName; } catch (e) { return { ok: false, reason: "tables-threw" }; }'
    '   if (nm === args.table) { target = tp[i]; break; }'
    ' }'
    ' if (!target) return { ok: false, reason: "table-not-found" };'
    ' var af; try { af = target.AutoFilter; } catch (e) { return { ok: false, reason: "tables-threw" }; }'
    ' if (!af) return { ok: true, value: { present: false, range: null, fields: null } };'
    ' var ref; try { var rg = af.Ref; ref = rg && typeof rg.getAbsName === "function" ? rg.getAbsName() : null; } catch (e) { ref = null; }'
    ' if (typeof ref !== "string") return { ok: true, value: { present: true, range: null, fields: null, reason: "internal-shape-unverified:ref" } };'
    ' var fc; try { fc = af.FilterColumns; } catch (e) { fc = null; }'
    ' var fcLen; try { fcLen = fc ? fc.length : null; } catch (e) { fcLen = null; }'
    ' if (!fc || typeof fcLen !== "number") return { ok: true, value: { present: true, range: ref.replace(/\\$/g, ""), fields: [] } };'
    ' var opNames = null;'
    ' try { var enumObj = (typeof Asc !== "undefined" && Asc && Asc.c_oAscCustomAutoFilter && typeof Asc.c_oAscCustomAutoFilter === "object") ? Asc.c_oAscCustomAutoFilter : null; if (enumObj) { opNames = {}; for (var ek in enumObj) { if (Object.prototype.hasOwnProperty.call(enumObj, ek)) opNames[enumObj[ek]] = ek; } } } catch (e) { opNames = null; }'
    ' var fields = [];'
    ' for (var ci = 0; ci < fcLen; ci += 1) {'
    '   var col = fc[ci];'
    '   var colId; try { colId = col.ColId; } catch (e) { colId = null; }'
    '   var entry = { colId: colId };'
    '   var unverified = [];'
    '   var colorF; try { colorF = col.ColorFilter; } catch (e) { colorF = null; }'
    '   var dynF; try { dynF = col.DynamicFilter; } catch (e) { dynF = null; }'
    '   var topF; try { topF = col.Top10; } catch (e) { topF = null; }'
    '   if (colorF) unverified.push("internal-shape-unverified:colorFilter");'
    '   if (dynF) unverified.push("internal-shape-unverified:dynamicFilter");'
    '   if (topF) unverified.push("internal-shape-unverified:top10");'
    '   var filtersObj; try { filtersObj = col.Filters; } catch (e) { filtersObj = null; }'
    '   var customObj; try { customObj = col.CustomFiltersObj; } catch (e) { customObj = null; }'
    '   if (filtersObj) {'
    '     var valuesObj; try { valuesObj = filtersObj.Values; } catch (e) { valuesObj = null; }'
    '     var blankVal; try { blankVal = filtersObj.Blank; } catch (e) { blankVal = null; }'
    '     var datesVal; try { datesVal = filtersObj.Dates; } catch (e) { datesVal = null; }'
    '     if (valuesObj) { var vKeys; try { vKeys = Object.keys(valuesObj); } catch (e) { vKeys = []; } entry.values = vKeys; }'
    '     if (blankVal) unverified.push("internal-shape-unverified:blank");'
    '     if (datesVal && typeof datesVal.length === "number" && datesVal.length > 0) unverified.push("internal-shape-unverified:dates");'
    '   } else if (customObj) {'
    '     var andVal; try { andVal = customObj.And; } catch (e) { andVal = null; }'
    '     var cf2; try { cf2 = customObj.CustomFilters; } catch (e) { cf2 = null; }'
    '     var cfLen; try { cfLen = cf2 ? cf2.length : null; } catch (e) { cfLen = null; }'
    '     var conditions = [];'
    '     if (cf2 && typeof cfLen === "number") {'
    '       for (var fi = 0; fi < cfLen; fi += 1) {'
    '         var opCode; try { opCode = cf2[fi].Operator; } catch (e) { opCode = null; }'
    '         var opVal; try { opVal = cf2[fi].Val; } catch (e) { opVal = null; }'
    '         var opName = (opNames && Object.prototype.hasOwnProperty.call(opNames, opCode)) ? opNames[opCode] : null;'
    '         conditions.push({ operator: opCode, operatorName: opName, value: opVal });'
    '       }'
    '     }'
    '     entry.custom = { and: andVal, conditions: conditions };'
    '     if (andVal === true) unverified.push("internal-shape-unverified:and-true");'
    '   } else {'
    '     unverified.push("internal-shape-unverified:column-kind");'
    '   }'
    '   if (unverified.length > 0) entry.unverified = unverified;'
    '   fields.push(entry);'
    ' }'
    ' return { ok: true, value: { present: true, range: ref.replace(/\\$/g, ""), fields: fields } };'
    ' },'
)

# GH #363 2nd unit: the second handler, appended the SAME way (plain append, trailing comma already there).
# Every internal-model read goes through try/catch (NEVER a bare property read that could throw unguarded) --
# the SAME discipline HANDLER_ENTRY above already uses. eoCompiledStyleA1ToRowCol/eoCompiledStyleColourOf/
# eoCompiledStyleFillOfStyle mirror, field-for-field, the already-LIVE-PROVEN logic of the 2nd unit's own
# offline probe (test/live/xlsx-gh363-cf-effective-internal-route-probe.live.cjs: a1ToRowCol0/colourOf/
# fillColorOfCompiledResult) -- not a fresh derivation. NAMED WITH AN eoCompiledStyle PREFIX, DELIBERATELY,
# NOT the probe's own short names: a first attempt using the short names (colourOf in particular) collided
# with an EXISTING, UNRELATED call site inside sheet.autoFilter.colorFilter.read/v1's own handler (that
# handler calls something ALSO named `colourOf(...)`, which the warrant test's digestOf() resolves by a
# file-wide `function <name>(` search, not a handler-scoped one) -- the collision corrupted that OTHER,
# untouched handler's digest, caught by the warrant test's own byte-identical check. Verbose, prefixed names
# avoid this: `grep -c eoCompiledStyle rev27-candidate/bridge-block.js` -> 0 before this change.
HANDLER_ENTRY_2 = (
    ' "sheet.cells.compiledStyle.read/v1": function (args) {'
    ' if (!args || typeof args.sheet !== "string" || !args.cells || typeof args.cells.length !== "number") return { ok: false, reason: "bad-args" };'
    ' if (args.cells.length > 200) return { ok: false, reason: "too-many-cells" };'
    ' var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" };'
    ' var m; try { m = ws.worksheet; } catch (e) { return { ok: false, reason: "worksheet-missing" }; }'
    ' if (!m) return { ok: false, reason: "worksheet-missing" };'
    ' if (typeof m.getCompiledStyle !== "function") return { ok: false, reason: "getCompiledStyle-not-a-function" };'
    ' function eoCompiledStyleA1ToRowCol(a1) {'
    '   var mm = /^([A-Z]+)(\\d+)$/.exec(a1);'
    '   if (!mm) return null;'
    '   var letters = mm[1]; var col = 0;'
    '   for (var li = 0; li < letters.length; li += 1) { col = col * 26 + (letters.charCodeAt(li) - 64); }'
    '   col -= 1;'
    '   var row = parseInt(mm[2], 10) - 1;'
    '   return { row: row, col: col };'
    ' }'
    ' function eoCompiledStyleColourOf(c) {'
    '   if (c === null || c === undefined) return null;'
    '   if (typeof c.getRgb === "function") { try { var v = c.getRgb(); return typeof v === "number" ? v : null; } catch (e) { return null; } }'
    '   if (typeof c.rgb === "number") return c.rgb;'
    '   return null;'
    ' }'
    ' function eoCompiledStyleFillOfStyle(style) {'
    '   if (!style || typeof style.getFill !== "function") return null;'
    '   var fillObj; try { fillObj = style.getFill(); } catch (e) { return null; }'
    '   if (!fillObj) return null;'
    '   if (typeof fillObj.getColor === "function") {'
    '     var c1; try { c1 = fillObj.getColor(); } catch (e) { c1 = null; }'
    '     var v1 = eoCompiledStyleColourOf(c1);'
    '     if (typeof v1 === "number") return v1;'
    '   }'
    '   if (typeof fillObj.getSolidFill === "function") {'
    '     var c2; try { c2 = fillObj.getSolidFill(); } catch (e) { c2 = null; }'
    '     var v2 = eoCompiledStyleColourOf(c2);'
    '     if (typeof v2 === "number") return v2;'
    '   }'
    '   return null;'
    ' }'
    ' var out = [];'
    ' for (var i = 0; i < args.cells.length; i += 1) {'
    '   var a1 = args.cells[i];'
    '   if (typeof a1 !== "string") { out.push({ cell: a1, compiled: null, reason: "bad-cell-address" }); continue; }'
    '   var rc = eoCompiledStyleA1ToRowCol(a1);'
    '   if (!rc) { out.push({ cell: a1, compiled: null, reason: "bad-cell-address" }); continue; }'
    '   var style; try { style = m.getCompiledStyle(rc.row, rc.col); } catch (e) { out.push({ cell: a1, compiled: null, reason: "threw" }); continue; }'
    '   if (!style) { out.push({ cell: a1, compiled: null }); continue; }'
    '   out.push({ cell: a1, compiled: { fill: eoCompiledStyleFillOfStyle(style) } });'
    ' }'
    ' return { ok: true, value: out };'
    ' },'
)

MANIFEST_ANCHOR = '], }; function argsOk(args, keys) {'
# SURGOS JAVITAS (zero lelete, 2026-11-06): a RÉGI anchor (' var EO_BRIDGE_WRITES = {') a
# EO_BRIDGE_CAPS objektum SAJÁT záró `};`-je UTÁN kezdődött -- a beszúrás tehát a handler-
# bejegyzéseket az objektumon KÍVÜLRE, a `};` és a `var EO_BRIDGE_WRITES` közé tette:
#   ...}, };  "sheet.table...": function (args) {...}, "sheet.cells...": function (args) {...}, var EO_BRIDGE_WRITES = {
# ami ÉRVÉNYTELEN JS (egy string literál colon-nal nem lehet standalone statement/label) --
# `node --check` a MEGLÉVŐ, commitolt bridge-block.js-en: "SyntaxError: Unexpected token ':'".
# A HELYES anchor a `};`-t IS MAGÁBAN FOGLALJA, és a beszúrás a `},` UTÁN, a `};` ELŐTT megy --
# ugyanaz a "plain append, utolsó bejegyzésként" minta, amit a MANIFEST_ANCHOR (ami a `]`-t is
# magában foglalja, tehát a tömb LEZÁRÁSA ELŐTT szúr be) már helyesen csinál.
CAPS_END_ANCHOR = '}, }; var EO_BRIDGE_WRITES = {'


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
    out = out.replace(MANIFEST_ANCHOR, MANIFEST_ENTRY + MANIFEST_ENTRY_2 + MANIFEST_ANCHOR, 1)
    # the new handlers go BETWEEN the prior entry's trailing `,` and EO_BRIDGE_CAPS's own closing `};`
    # -- NOT before the whole anchor (that would land them AFTER `};`, outside the object; see the
    # anchor's own comment above for the bug this replaces).
    out = out.replace(CAPS_END_ANCHOR, '}, ' + HANDLER_ENTRY + HANDLER_ENTRY_2 + '}; var EO_BRIDGE_WRITES = {', 1)
    if out.count(f'"{NEW_CAPABILITY_ID}"') < 2:  # once in the manifest id, once as the handler key
        return None, "NEW_CAPABILITY_NOT_PRESENT_AFTER_TRANSFORM"
    if out.count(f'"{NEW_CAPABILITY_ID_2}"') < 2:
        return None, "NEW_CAPABILITY_2_NOT_PRESENT_AFTER_TRANSFORM"
    return out, None


def syntax_check(js_text):
    """Composed-block syntax gate (a maintainer, 2026-11-06): the warrant test (ds935-bridge28-table-
    autofilter-criteria-write-handler-identity.test.cjs) extracts and checks individual handler
    spans by name -- it never parses the WHOLE composed file, so an anchor landing outside its
    intended object (exactly the CAPS_END_ANCHOR bug this script just fixed) passes every content
    check while the file itself fails to load in the editor. `node --check` parses the real file
    as JS (no execution, so the DS-only globals it references are never dereferenced) -- the
    cheapest real parse available, same engine family as the one that will actually load this
    block.

    bridge-block.js is NOT standalone JS: it is a FRAGMENT install_bridge.py splices verbatim
    into the real sdk-all.js right before an anchor (plan_result(), ~line 80) -- its own top level
    is a comma-expression sequence with a TRAILING comma (`(function(){...})(), Api....=function(){...},`),
    by design meant to continue into whatever follows at the real splice point. Checked standalone,
    even the UNMODIFIED, already-landed rev27-candidate/bridge-block.js fails with "Unexpected end
    of input" -- that is the fragment's OWN shape, not a defect. A comma-separated expression list
    is exactly what an ARRAY LITERAL's elements are (trailing comma included), so wrapping in `[`/`]`
    before parsing neutralizes the "fragment" shape without masking a real defect: RED/GREEN verified
    (2026-11-06) -- the unmodified rev27 base and this script's own fixed output both check CLEAN
    wrapped; the ORIGINAL BUGGY rev28 (committed before this fix, CAPS_END_ANCHOR landing handlers
    outside EO_BRIDGE_CAPS) still fails wrapped, with the SAME "Unexpected token ':'" node gave on
    the raw file -- the wrap does not hide a structural break, only the fragment's own trailing comma.

    Returns (ok, detail): ok=False on ANY failure to parse OR to even run node (fail closed, never
    silently skipped). `detail` is the FIRST line starting with a parser error name, never node's
    leading `[stdin]:N` location line or the source-context line/caret under it."""
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
    data = REV27_BLOCK.read_bytes().decode("utf-8")
    out, err = transform(data)
    if err:
        print(f"REFUSE: {err}; NOTHING WRITTEN", file=sys.stderr)
        return 2
    # composed-block syntax gate (a maintainer, 2026-11-06) -- BEFORE --check's own report and before any
    # write, so a bad anchor placement refuses here, the same way a content check already does.
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
        "base": "rev27-candidate/bridge-block.js",
        "base_sha256": input_sha256,
        "output_sha256": output_sha256,
        "new_capabilities": [NEW_CAPABILITY_ID, NEW_CAPABILITY_ID_2],
        "getters": [
            f'"{NEW_CAPABILITY_ID}": function',
            f'{{ id: "{NEW_CAPABILITY_ID}"',
            f'"{NEW_CAPABILITY_ID_2}": function',
            f'{{ id: "{NEW_CAPABILITY_ID_2}"',
        ],
        "operations": [
            {"id": "patch-version-marker", "old": OLD_PATCH_VERSION, "new": NEW_PATCH_VERSION},
            {"id": "manifest-capability-entry", "anchor": MANIFEST_ANCHOR, "inserted": MANIFEST_ENTRY},
            {"id": "manifest-capability-entry-2", "anchor": MANIFEST_ANCHOR, "inserted": MANIFEST_ENTRY_2},
            {"id": "handler-entry", "anchor": CAPS_END_ANCHOR, "inserted": HANDLER_ENTRY},
            {"id": "handler-entry-2", "anchor": CAPS_END_ANCHOR, "inserted": HANDLER_ENTRY_2},
        ],
        "_comment": (
            "GH #361 route probe route (c) candidate, derived by build-rev28-table-autofilter-criteria.py from "
            "the real rev27-candidate/bridge-block.js (base_sha256 pinned above). NOT installed anywhere; the "
            "internal .AutoFilter object's own shape beyond truthy existence is an UNVERIFIED HYPOTHESIS (see "
            "the build script's own docstring) -- per-column criteria are deliberately NOT attempted. EXTENDED "
            "for GH #363 2nd unit with sheet.cells.compiledStyle.read/v1 (LIVE measured, not installed either) "
            "-- the candidate carries BOTH capabilities on the SAME patchVersion marker (bovitve, not renamed)."
        ),
    }
    OUT_PROOF.write_text(json.dumps(proof, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"block": str(OUT_BLOCK), "proof": str(OUT_PROOF), "input_sha256": input_sha256, "output_sha256": output_sha256}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
