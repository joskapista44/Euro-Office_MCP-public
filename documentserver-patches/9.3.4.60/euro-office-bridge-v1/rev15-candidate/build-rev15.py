#!/usr/bin/env python3
"""
X-O05/5-B (card an internal branch, kioszto zero;
owner-GO the owner, owner channel/36664, relayed comment #ref: "B: DS Bridge iras, UJ
DS-revizio") -- OFFLINE build of the euro-office-bridge-v1 patchVersion /15 CANDIDATE.

/15 = /14 + TWO new capabilities:
  WRITE (X-O05/5-B, mutation TARGETED):
    chart.series.marker.write/v1     -- a series' own marker symbol + size
    chart.series.lineStyle.write/v1  -- a series' own line dash + smooth

the owner's decision (X-O05/5, comment #ref) was B for marker type/size and dash/smoothing; the
THIRD open axis of that same card (secondary-axis placement) was decided NAMED LIMIT (A) and is
NOT part of this build -- see docs/CAPABILITY-COMPLETION-STATUS-20261012.md's own X-O05 row.

THE MECHANISM, end to end, is EIGHT LIVE rounds deep (round 1 reflection -> round 4/5 direct-model
NO_EFFECT -> the owner's "try co-editing" -> round 6/6b/6c/6d coedit NO candidate found on the PUBLIC
ApiChartSeries/ApiChart wrapper objects -> round 7/7b: the correct object is the INTERNAL series,
`api.Chart.chart.plotArea.charts[0].series[i]` -- the SAME chain `chart.bar.gapOverlap.write/v1`
(/12) and `chart.series.markerStyle.read/v1` (/11) already use and already persist through --
where setMarker/setSmooth/setSpPr/setNoneMarker on the series, and setSymbol/setSize/setSpPr on
its own `marker` sub-object, ALL exist (round 7b LIVE, case-insensitive fix after round 7's own
blind uppercase-only filter) -> round 8/8b/8c/8d: three of THIS agent's own bugs (a too-small
scatter seed range; `SetSeriesOutLine` called on the series instead of the chart wrapper; reading
`.marker`/`.spPr` off `GetAllSeries()[0]`, the PUBLIC wrapper, which never carries either field on
ANY chart type) -> round 8d LIVE: MEASURED, via the SAME internal chain, byte-identical to round
7b's own navigation. No constructor or method name below is invented: every one was LIVE-checked
for existence (`typeof === 'function'`) before this build encodes it, and every one that could
not be reached is a NAMED fail-closed reason in the write handler, never a guess.

SYMBOL: `AscFormat.GetTypeMarkerByIndex(i)` for i = 0..9 is NOT an identity map (round 2's own
`MARKER_SYMBOL_TYPE` array, `[2,7,9,10,8,0,6,3,1]`, already showed this) -- round 8d wrote each
token through the PROVEN `marker.setSymbol(token)`, saved, reopened, and read the SAVED OOXML's
own `<c:marker>/<c:symbol val="...">` directly (ground truth, no translation). The resulting
TOKEN_BY_NAME table below is that measured mapping, not a guess at the ECMA-376 XSD's own
declaration order (which this build never assumes). 'none' and token 4 are NOT in this table:
round 7b measured an untouched series' own marker at token 4 = the file's own "none" state, and
the model's own `setNoneMarker()` (round 7b LIVE: exists) is the clean way to request it, not a
`setSymbol()` call with a guessed token. Token 5 (round 8d's own suspicion: picture/auto) is
likewise excluded -- the ECMA-376 `ST_MarkerStyle` enum this build accepts is the 9 named values
the table covers plus "none" via the separate path; "picture"/"auto" are refused by the schema
itself (unsupported-symbol), never silently mapped to an unverified token.

DASH (REVISED, round 8g/8j/8k/8l/8m LIVE -- supersedes the round 8d-era design this docstring
originally described): round 8d found `ln.setPrstDash(ln.GetDashCode(name))` on the EXISTING
`spPr.ln` (populated as a side effect of `SetSeriesOutLine`) -- but round 8g/8j LIVE MEASURED
that an in-place `setPrstDash` call on that EXISTING `ln` does NOT persist through a save+reopen
(the field changes in-session, round 8l's own in-session readback confirmed it, but the SAVED
FILE keeps the line's PRE-write dash). The mechanism that DOES persist (round 8k LIVE MEASURED,
all 11 ECMA-376 dash names round-tripped byte-for-byte; round 8m LIVE MEASURED with an EXISTING
red/width line, all three axes survived): build a BRAND-NEW, standalone `AscFormat.CLn()`, set
its width/fill/dash on THAT object, and only THEN attach it via `s.setSpPr(new CSpPr())` +
`spPr.setLn(lnNew)` -- never mutate the series' existing `ln` in place. To avoid silently
discarding an EXISTING line's own colour/width, the new `CLn` is seeded from the OLD one's own
FIELDS when present (`lnOld.w`, a number; `lnOld.Fill`, the INTERNAL UniFill object, reused
DIRECTLY -- round 8m LIVE: the public `Api.CreateSolidFill()` wrapper's return value must NEVER
be passed straight to `setFill`, only `.UniFill` off it or `AscFormat.CreateSolidFillRGB` --
round 8j LIVE found the wrapper-fill path corrupts and DELETES the whole sheet's drawing part on
save, a safety finding this write must never repeat). If nothing was readable on the old line
(no existing `ln`, or neither field present), the write falls back to a NAMED product default
(`AscFormat.CreateSolidFillRGB(0x11,0x22,0x33)` / `.UniFill` fallback, 12700 EMU width) -- never
silently indistinguishable from an intentional colour. `GetDashCode`/`GetDashByCode` (the SAME
naming convention as the ALREADY PRESENT `GetCapCode`/`GetCapByCode`, `GetAlgnCode`/
`GetAlgnByCode`, `GetCmpdCode`/`GetCmpdByCode` round 8d's own raw dump shows sitting right next to
it) is unchanged: `ln.setPrstDash(ln.GetDashCode(ecmaName))`, Code: name-in, ByCode: name-out.
This write NO LONGER refuses when the series has no `spPr`/`ln` yet (the round 8d-era
`sppr-missing`/`ln-missing` limitation is GONE -- the new CLn is built regardless) -- it refuses
only on a genuinely missing CONSTRUCTOR (`cln-missing`, `cspPr-or-setsppr-missing`), never on an
absent prior line.

R8o (a maintainer's review of round 8n, card comment #ref): the round 8n attach step UNCONDITIONALLY
built a brand-new, EMPTY `CSpPr` and called `s.setSpPr(spPrNew)` even when the series already HAD
an spPr -- destroying that spPr's `Fill` (the series' own fill colour, distinct from the LINE's
Fill) and any other spPr fields (effects etc.) on every chart type, not just line/scatter. The
fix is NOT a chart-type gate (zero offered that as a fallback, but no such gate is provable from
the OOXML model here): when an EXISTING spPr is present, its OWN `setLn` is called directly --
the spPr object itself is NEVER replaced, so nothing on it besides `ln` can be touched. A brand new
`CSpPr` is constructed (as round 8n did) ONLY when the series currently has no spPr at all.

R8q (a maintainer's review of the round 8p LIVE probe, card comment #ref): the L2 subject (a LINE
series the probe's own setup never gave an explicit `ln`) still read back width 28575 / a null
RGB colour after the dash write -- NOT the probe's assumed fallback (12700 / 0x112233). The
engine's own `AddChart` already seeds a line-type series with a default theme line (28575 EMU =
2.25pt, a scheme-coloured fill, never an `srgbClr`) -- so `spPrOld`/`lnOld` were truthy, and this
write correctly took the PRESERVE branch, not the fallback one; the probe's harness could not
distinguish the two from the saved file alone. The receipt now carries `widthSource`/`fillSource`
(each "preserved" or "fallback") precisely so a caller never has to guess which branch fired. `series.setSmooth(bool)` and `marker.setSize(n)` are both
round-7b-proven, unconditional setters once their own host object is confirmed to exist.

The chart-lookup-by-name / INTERNAL-series-resolution prefix for both new capabilities STARTS from
the SAME text `chart.series.markerStyle.read/v1`'s own dispatch already uses (rev11-candidate/
bridge-block.js) -- but the chart-NAME lookup itself is STRICTER than that read-only convention:
a maintainer's own request (round 8m's handoff) was to resolve chart-by-name the SAME way
`chart.drawing.delete.write/v1` (/13) already does for an IRREVERSIBLE write -- EXACTLY one name
match required, zero matches refuses `chart-not-found`, MORE than one refuses the NAMED
`chart-name-ambiguous` (never silently takes the first, the way every READ-only chart.* handler
in this block still does) -- these are WRITE capabilities, and guessing which of several
same-named charts the caller meant is the SAME class of risk the delete handler's own docstring
already names. series-index-out-of-range onward is unchanged.

Deterministic: the /15 block is derived from the CANONICAL /14 block
(../rev14-candidate/bridge-block.js, sha pinned) by THREE exact string substitutions, each
asserted to occur exactly once. No timestamps, no environment input.

Usage:
  build-rev15.py                         -> writes bridge-block.js, old.txt, new.txt,
                                            manifest-mechanism-proof.json next to this file
This script NEVER writes a Document Server file, and has NO --sdk mode: the real served /14
sdk-all.js is a the production host/.10-host artifact this agent (designated-uid, no an owner-approval gate access)
cannot fetch or verify -- producing a manifest.json that CLAIMS an SDK sha without having read
that file would be a fabricated measurement, not a build. The SDK-backed manifest.json (and the
apply itself) is maintainer's gate, a separate, later step -- see README.md.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V14_BLOCK = HERE.parent / "rev14-candidate" / "bridge-block.js"
V14_BLOCK_SHA = "a95b78fc02134c5375715390b3e6d63e5442060a496fd2319bb72bf4422aed1b"
V14_BLOCK_SIZE = 77446

# ---- SYMBOL: measured token<->ECMA-376 name table (round 8d LIVE, OOXML readback -- not the
# XSD's own declaration order, never guessed). 'none' goes through setNoneMarker(), not this map.
TOKEN_BY_NAME = {
    "circle": 0, "dash": 1, "diamond": 2, "dot": 3, "plus": 6,
    "square": 7, "star": 8, "triangle": 9, "x": 10,
}
MARKER_SYMBOL_ENUM = sorted(list(TOKEN_BY_NAME.keys()) + ["none"])  # 10 names, ECMA-376 ST_MarkerStyle minus auto/picture

# ---- DASH: ECMA-376 ST_PresetLineDashVal, the full 11-value enum (a maintainer's plan review, #ref point 2).
DASH_ENUM = ["solid", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot",
             "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"]

MARKER_WRITE_REASONS = ("bad-args, unsupported-property, no-property, invalid-symbol, invalid-size, "
                        "sheet-not-found, charts-threw, chart-not-found, chart-name-ambiguous, "
                        "chart-type-not-single, series-index-out-of-range, marker-not-found, "
                        "setnonemarker-missing, setsymbol-missing, setsize-missing")
# R8n (a maintainer's handoff after round 8m LIVE MEASURED): sppr-missing/ln-missing are GONE -- the dash
# write no longer requires a pre-existing spPr/ln, it builds one. cln-missing/
# cspPr-or-setsppr-missing are the new, build-side refusals; setw-threw/setfill-threw/
# setsppr-threw/setln-threw/setprstdash-threw are the new standalone-CLn write's own throw paths
# (setsmooth-threw was always reachable in the implementation but missing from this list before).
# R8o (a maintainer's review #ref of round 8n): cspPr-or-setsppr-missing/setsppr-threw/setln-missing now
# fire ONLY on the no-existing-spPr path (a brand-new CSpPr is built there, as before).
# existing-sppr-setln-missing is NEW: it fires when the series already HAS an spPr but that spPr
# object itself lacks `setLn` -- refused rather than silently replacing the spPr (which would
# re-introduce the exact data-loss round 8o exists to fix).
LINE_STYLE_WRITE_REASONS = ("bad-args, unsupported-property, no-property, invalid-dash, invalid-smooth, "
                            "sheet-not-found, charts-threw, chart-not-found, chart-name-ambiguous, "
                            "chart-type-not-single, series-index-out-of-range, setsmooth-missing, "
                            "setsmooth-threw, cln-missing, setw-threw, setfill-threw, "
                            "setprstdash-missing, getdashcode-missing, setprstdash-threw, "
                            "existing-sppr-setln-missing, cspPr-or-setsppr-missing, setsppr-threw, "
                            "setln-missing, setln-threw")

MARKER_WRITE_ENTRY = (
    '{ id: "chart.series.marker.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one chart series (sheet, chart name EXACTLY one match -- chart-name-ambiguous if '
    'more, the SAME strict convention chart.drawing.delete.write/v1 /13 already uses -- series index), at '
    'least one of symbol '
    '(' + '|'.join(MARKER_SYMBOL_ENUM) + ', ECMA-376 ST_MarkerStyle minus auto/picture -- unmeasured, '
    'never guessed) / size (integer 2..72, ST_MarkerSize): writes through the internal model\'s OWN '
    'setMarker/setNoneMarker/marker.setSymbol/marker.setSize METHODS (X-O05/5-B round 7b+8d LIVE VERIFIED; '
    'round 4\'s own direct FIELD assignment -- series.marker = new CMarker() -- APPEARED to apply '
    'in-session but did NOT survive a reopen, the SAME lesson chart.bar.gapOverlap.write/v1 already '
    "encodes); symbol 'none' goes through setNoneMarker(), never a guessed token; every OTHER symbol's "
    'internal token is read from a table MEASURED via round 8d\'s own write+OOXML-readback (AscFormat.'
    'GetTypeMarkerByIndex is NOT an identity map), never the ECMA-376 XSD\'s own declaration order; an '
    'unrequested field is never touched (preserved); every input is validated, the series\' identity and '
    'chart-type-is-single are resolved, and each requested setter is confirmed to exist, ALL BEFORE the '
    'first mutation; the result is an operation receipt {applied:[written field names]} and NOT a '
    'verification -- the caller verifies with a fresh chart.series.markerStyle.read/v1-shaped read of its '
    'own (this capability does not itself read back symbol/size); a setter that throws after an earlier '
    'one succeeded leaves a PARTIAL write, reported as {partial:true, applied:[...]}, never rolled back; '
    'no fallback; served only through write(); '
    'reasons: ' + MARKER_WRITE_REASONS + ' (X-O05/5-B, /15)" }, '
)
LINE_STYLE_WRITE_ENTRY = (
    '{ id: "chart.series.lineStyle.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one chart series (sheet, chart name EXACTLY one match -- chart-name-ambiguous if more, '
    'the SAME strict convention chart.drawing.delete.write/v1 /13 already uses -- series index), at least one '
    'of dash (' + '|'.join(DASH_ENUM) + ', ECMA-376 ST_PresetLineDashVal) / smooth (boolean): smooth writes '
    'through series.setSmooth() unconditionally once the series is resolved (round 7b+8d LIVE VERIFIED); dash '
    'builds a BRAND-NEW, standalone internal line object (AscFormat.CLn()) and ATTACHES it via '
    'series.setSpPr(new CSpPr()) + spPr.setLn(lnNew) -- round 8d\'s own in-place setPrstDash on the EXISTING '
    'line APPEARED to apply in-session but did NOT survive a reopen (round 8g/8j/8l LIVE MEASURED, the SAME '
    "lesson round 4's own direct ln.prstDash='dash' FIELD assignment and chart.bar.gapOverlap.write/v1 already "
    'encode); the new line\'s width/fill are SEEDED from the EXISTING line\'s own fields when present '
    '(ln.w / ln.Fill, the INTERNAL UniFill object reused DIRECTLY, round 8m LIVE MEASURED: preserves an '
    "existing red/width line through the dash write) -- the PUBLIC Api.CreateSolidFill() wrapper's own return "
    'value is NEVER passed straight to setFill (round 8j LIVE: that shape corrupts and DELETES the whole '
    'sheet\'s drawing part on save, only .UniFill off it or AscFormat.CreateSolidFillRGB are used), and if '
    'nothing was readable on the existing line the write falls back to a NAMED product default (RGB '
    '0x11,0x22,0x33 / 12700 EMU width), never silently indistinguishable from an intentional colour; dash '
    'itself writes via lnNew.setPrstDash(lnNew.GetDashCode(name)) (GetDashCode/GetDashByCode: the SAME '
    'Code/ByCode-named pair already present alongside GetCapCode/GetCapByCode, GetAlgnCode/GetAlgnByCode, '
    'GetCmpdCode/GetCmpdByCode, round 8d\'s own raw reflection dump); UNLIKE the original /15 design this '
    "capability no longer refuses when the series has no spPr/ln yet -- it builds one regardless, so the "
    'former sppr-missing/ln-missing limitation does not apply; an unrequested field is never touched '
    '(preserved); every input is validated, the series\' identity and chart-type-is-single are resolved, and '
    'the dash construction chain (CLn/CSpPr constructors, each setter, GetDashCode) is confirmed to exist at '
    'each step, immediately before that step\'s own mutation; the result is an operation receipt '
    '{applied:[written field names]} and NOT a verification -- the caller verifies with a fresh '
    'chart.series.lineStyle.read/v1-shaped read of its own (not yet built, a separate later capability); when '
    'dash is applied the receipt ALSO carries {widthSource, fillSource}, each preserved (seeded from the '
    'existing line\'s own field) or fallback (the NAMED product default), so the caller can tell the two '
    'apart WITHOUT its own read (R8q, internal tracker: the round 8p LIVE probe\'s harness could not tell '
    'an existing line\'s own scheme-coloured/theme-default width from the named RGB/EMU fallback by reading '
    'the saved file alone); a '
    'setter that throws after an earlier one succeeded leaves a PARTIAL write, reported as {partial:true, '
    'applied:[...]}, never rolled back; no fallback on the dash CONSTRUCTION itself (the colour/width seeding '
    'above is the only fallback, and it is always named); served only through write(); '
    'reasons: ' + LINE_STYLE_WRITE_REASONS + ' (X-O05/5-B, /15)" }, '
)

# the chart-lookup + INTERNAL-series-resolution prefix. series-index-out-of-range onward is
# VERBATIM from the landed chart.series.markerStyle.read/v1 dispatch (rev11-candidate/
# bridge-block.js) -- not re-derived. The chart-NAME lookup itself is the STRICTER,
# exactly-one-match convention chart.drawing.delete.write/v1 (/13) already established for an
# irreversible write -- a maintainer's own request (a maintainer's handoff after round 8m LIVE, X-O05/5-B):
# zero matches -> chart-not-found, MORE than one -> the NAMED chart-name-ambiguous (never
# silently the first match, unlike the read-only convention this prefix otherwise starts from).
_SERIES_LOOKUP_PREFIX = (
    'if (!args || typeof args !== "object" || typeof args.sheet !== "string" || typeof args.chart !== "string" || typeof args.series !== "number") return { ok: false, reason: "bad-args" }; '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var matches = []; for (var i = 0; i < charts.length; i += 1) { var okName = false; try { okName = charts[i].GetName() === args.chart; } catch (e) { okName = false; } if (okName) matches.push(charts[i]); } '
    'if (matches.length === 0) return { ok: false, reason: "chart-not-found" }; '
    'if (matches.length > 1) return { ok: false, reason: "chart-name-ambiguous" }; '
    'var api = matches[0]; '
    'var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; var charts2 = plot && plot.charts; '
    'if (!charts2 || charts2.length !== 1) return { ok: false, reason: "chart-type-not-single" }; '
    'var ch = charts2[0]; var series = ch.series || []; var s = series[args.series]; '
    'if (!s) return { ok: false, reason: "series-index-out-of-range" }; '
)

MARKER_WRITE_IMPL = (
    '"chart.series.marker.write/v1": function (args) { '
    + _SERIES_LOOKUP_PREFIX +
    'var allowed = { sheet: 1, chart: 1, series: 1, symbol: 1, size: 1 }; var keys = Object.keys(args); '
    'for (var k = 0; k < keys.length; k += 1) { if (!Object.prototype.hasOwnProperty.call(allowed, keys[k])) return { ok: false, reason: "unsupported-property" }; } '
    'var hasSymbol = Object.prototype.hasOwnProperty.call(args, "symbol"); var hasSize = Object.prototype.hasOwnProperty.call(args, "size"); '
    'if (!hasSymbol && !hasSize) return { ok: false, reason: "no-property" }; '
    'var SYMBOL_ENUM = ' + json.dumps(MARKER_SYMBOL_ENUM) + '; '
    'if (hasSymbol && SYMBOL_ENUM.indexOf(args.symbol) === -1) return { ok: false, reason: "invalid-symbol" }; '
    'if (hasSize && !(typeof args.size === "number" && args.size >= 2 && args.size <= 72 && Math.floor(args.size) === args.size)) return { ok: false, reason: "invalid-size" }; '
    'var applied = []; '
    'if (hasSymbol) { '
    'if (args.symbol === "none") { '
    'if (typeof s.setNoneMarker !== "function") return { ok: false, reason: "setnonemarker-missing", partial: applied.length > 0, applied: applied }; '
    'try { s.setNoneMarker(); } catch (e) { return { ok: false, reason: "setnonemarker-threw", partial: applied.length > 0, applied: applied }; } applied.push("symbol"); '
    '} else { '
    'var marker = s.marker; if (!marker) return { ok: false, reason: "marker-not-found", partial: applied.length > 0, applied: applied }; '
    'if (typeof marker.setSymbol !== "function") return { ok: false, reason: "setsymbol-missing", partial: applied.length > 0, applied: applied }; '
    'var TOKEN_BY_NAME = ' + json.dumps(TOKEN_BY_NAME) + '; '
    'try { marker.setSymbol(TOKEN_BY_NAME[args.symbol]); } catch (e) { return { ok: false, reason: "setsymbol-threw", partial: applied.length > 0, applied: applied }; } applied.push("symbol"); '
    '} } '
    'if (hasSize) { '
    'var marker2 = s.marker; if (!marker2) return { ok: false, reason: "marker-not-found", partial: applied.length > 0, applied: applied }; '
    'if (typeof marker2.setSize !== "function") return { ok: false, reason: "setsize-missing", partial: applied.length > 0, applied: applied }; '
    'try { marker2.setSize(args.size); } catch (e) { return { ok: false, reason: "setsize-threw", partial: applied.length > 0, applied: applied }; } applied.push("size"); '
    '} '
    'return { ok: true, value: { applied: applied } }; }, '
)
# R8n (a maintainer's handoff after round 8m LIVE MEASURED, X-O05/5-B): the dash write -- a BRAND-NEW,
# standalone CLn (fill -> width -> dash, attach LAST via setSpPr+setLn), seeded from the EXISTING
# ln's own FIELDS (w, Fill) when present, falling back to a NAMED product default otherwise. The
# in-place setPrstDash-on-the-existing-ln mechanism this build ORIGINALLY used (round 8d-era) did
# NOT survive a reopen (round 8g/8j/8l LIVE) -- this is round 8k/8m's own MEASURED replacement.
_LINE_STYLE_DASH_WRITE = (
    'if (hasDash) { '
    'var spPrOld = s.spPr; var lnOld = spPrOld ? spPrOld.ln : null; '
    'var readWidth = (lnOld && typeof lnOld.w === "number") ? lnOld.w : null; '
    'var readFill = (lnOld && lnOld.Fill) ? lnOld.Fill : null; '
    'var widthSource = readWidth !== null ? "preserved" : "fallback"; '
    'var fillSource = readFill ? "preserved" : "fallback"; '
    'var hasCLn = typeof (AscFormat && AscFormat.CLn) === "function"; '
    'if (!hasCLn) return { ok: false, reason: "cln-missing", partial: applied.length > 0, applied: applied }; '
    'var lnNew = new AscFormat.CLn(); '
    'var fillBuilt = readFill; '
    'if (!fillBuilt) { '
    'var hasCreateSolidFillRGB = typeof (AscFormat && AscFormat.CreateSolidFillRGB) === "function"; '
    'if (hasCreateSolidFillRGB) { try { fillBuilt = AscFormat.CreateSolidFillRGB(0x11, 0x22, 0x33); } catch (e) { fillBuilt = null; } } '
    'if (!fillBuilt && typeof (Api && Api.CreateSolidFill) === "function") { '
    'try { var wrapped = Api.CreateSolidFill(Api.CreateRGBColor(0x11, 0x22, 0x33)); fillBuilt = wrapped ? wrapped.UniFill : null; } catch (e) { fillBuilt = null; } '
    '} '
    '} '
    'var widthToUse = readWidth !== null ? readWidth : 12700; '
    'if (typeof lnNew.setW === "function") { try { lnNew.setW(widthToUse); } catch (e) { return { ok: false, reason: "setw-threw", partial: applied.length > 0, applied: applied }; } } '
    'if (fillBuilt && typeof lnNew.setFill === "function") { try { lnNew.setFill(fillBuilt); } catch (e) { return { ok: false, reason: "setfill-threw", partial: applied.length > 0, applied: applied }; } } '
    'if (typeof lnNew.setPrstDash !== "function") return { ok: false, reason: "setprstdash-missing", partial: applied.length > 0, applied: applied }; '
    'if (typeof lnNew.GetDashCode !== "function") return { ok: false, reason: "getdashcode-missing", partial: applied.length > 0, applied: applied }; '
    'try { lnNew.setPrstDash(lnNew.GetDashCode(args.dash)); } catch (e) { return { ok: false, reason: "setprstdash-threw", partial: applied.length > 0, applied: applied }; } '
    'if (spPrOld) { '
    'if (typeof spPrOld.setLn !== "function") return { ok: false, reason: "existing-sppr-setln-missing", partial: applied.length > 0, applied: applied }; '
    'try { spPrOld.setLn(lnNew); } catch (e) { return { ok: false, reason: "setln-threw", partial: applied.length > 0, applied: applied }; } applied.push("dash"); '
    '} else { '
    'var hasCSpPr = typeof (AscFormat && AscFormat.CSpPr) === "function"; '
    'if (!hasCSpPr || typeof s.setSpPr !== "function") return { ok: false, reason: "cspPr-or-setsppr-missing", partial: applied.length > 0, applied: applied }; '
    'var spPrNew = new AscFormat.CSpPr(); '
    'try { s.setSpPr(spPrNew); } catch (e) { return { ok: false, reason: "setsppr-threw", partial: applied.length > 0, applied: applied }; } '
    'var spPrNewAfter = s.spPr; '
    'if (!spPrNewAfter || typeof spPrNewAfter.setLn !== "function") return { ok: false, reason: "setln-missing", partial: applied.length > 0, applied: applied }; '
    'try { spPrNewAfter.setLn(lnNew); } catch (e) { return { ok: false, reason: "setln-threw", partial: applied.length > 0, applied: applied }; } applied.push("dash"); '
    '} '
    '} '
)

LINE_STYLE_WRITE_IMPL = (
    '"chart.series.lineStyle.write/v1": function (args) { '
    + _SERIES_LOOKUP_PREFIX +
    'var allowed = { sheet: 1, chart: 1, series: 1, dash: 1, smooth: 1 }; var keys = Object.keys(args); '
    'for (var k = 0; k < keys.length; k += 1) { if (!Object.prototype.hasOwnProperty.call(allowed, keys[k])) return { ok: false, reason: "unsupported-property" }; } '
    'var hasDash = Object.prototype.hasOwnProperty.call(args, "dash"); var hasSmooth = Object.prototype.hasOwnProperty.call(args, "smooth"); '
    'if (!hasDash && !hasSmooth) return { ok: false, reason: "no-property" }; '
    'var DASH_ENUM = ' + json.dumps(DASH_ENUM) + '; '
    'if (hasDash && DASH_ENUM.indexOf(args.dash) === -1) return { ok: false, reason: "invalid-dash" }; '
    'if (hasSmooth && typeof args.smooth !== "boolean") return { ok: false, reason: "invalid-smooth" }; '
    'var applied = []; '
    'if (hasSmooth) { '
    'if (typeof s.setSmooth !== "function") return { ok: false, reason: "setsmooth-missing", partial: applied.length > 0, applied: applied }; '
    'try { s.setSmooth(args.smooth); } catch (e) { return { ok: false, reason: "setsmooth-threw", partial: applied.length > 0, applied: applied }; } applied.push("smooth"); '
    '} '
    + _LINE_STYLE_DASH_WRITE +
    'return { ok: true, value: hasDash ? { applied: applied, widthSource: widthSource, fillSource: fillSource } : { applied: applied } }; }, '
)

# (old, new) -- each `old` must occur EXACTLY once in the EVOLVING text at the time it is applied.
SUBS = [
    ('patchVersion: "euro-office-bridge-v1/9.3.4.60/14"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/15"'),
    ('(X-O08, /14)" }, ], };', '(X-O08, /14)" }, ' + MARKER_WRITE_ENTRY + LINE_STYLE_WRITE_ENTRY + '], };'),
    # the EO_BRIDGE_WRITES object's closing anchor -- the LAST write handler already in the /14
    # block (chart.drawing.delete.write/v1's own tail) ends with this exact text, immediately
    # before "var EO_BRIDGE_HYDRATIONS = {".
    ('return { ok: true, value: { deleted: true } }; }, }; var EO_BRIDGE_HYDRATIONS = {',
     'return { ok: true, value: { deleted: true } }; }, ' + MARKER_WRITE_IMPL + LINE_STYLE_WRITE_IMPL + '}; var EO_BRIDGE_HYDRATIONS = {'),
]
MARKERS = [
    '"chart.series.marker.write/v1": function', '{ id: "chart.series.marker.write/v1"',
    '"chart.series.lineStyle.write/v1": function', '{ id: "chart.series.lineStyle.write/v1"',
]


def sha(b): return hashlib.sha256(b).hexdigest()


def build_block():
    v14 = V14_BLOCK.read_bytes()
    if sha(v14) != V14_BLOCK_SHA or len(v14) != V14_BLOCK_SIZE:
        sys.exit(f"REFUSE: /14 block sha/size {sha(v14)}/{len(v14)} != {V14_BLOCK_SHA}/{V14_BLOCK_SIZE}; NO OUTPUT")
    t = v14.decode("utf-8")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("(X-O05/5-B, /15)") != 2:
        sys.exit(f"REFUSE: (X-O05/5-B, /15) count {t.count('(X-O05/5-B, /15)')} != 2; NO OUTPUT")
    return v14, t.encode("utf-8")


def main():
    v14, v15 = build_block()
    out = {"v14_block_sha": sha(v14), "v14_block_size": len(v14), "v15_block_sha": sha(v15), "v15_block_size": len(v15)}
    (HERE / "bridge-block.js").write_bytes(v15)
    (HERE / "old.txt").write_bytes(v14)
    (HERE / "new.txt").write_bytes(v15)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.4.60", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v15-x-o05-5-b-marker-linestyle-write", "old": v14.decode("utf-8"), "new": v15.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- a served-sdk-backed manifest.json for the real production apply is a SEPARATE, later "
                              "step, gated to the designated operator role, with the real served file in hand.",
                 original_sha256=sha(v14), patched_sha256=sha(v15), original_size=len(v14), patched_size=len(v15))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
