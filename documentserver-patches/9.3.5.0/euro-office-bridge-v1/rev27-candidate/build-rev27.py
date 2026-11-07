#!/usr/bin/env python3
"""
GH #346 AUTO-REPAIR, megvalositas 1+2 -- OFFLINE build of the euro-office-bridge-v1 patchVersion /27 CANDIDATE
for DS 9.3.5, in ONE deterministic pass
from the /26 candidate block (../rev26-candidate/bridge-block.js, sha pinned).

megvalositas 1 (READ, LANDOLT):  /27 = /26 + EXACTLY ONE handler-body change, to chart.series.style.read/v1's
dataPoint branch -- no new capability, no capability removed, no other handler touched.

megvalositas 2 (WRITE, THIS BUILD): /27 ALSO gains ONE new capability, chart.series.dataPoint.style.write/v1 --
a maintainer's decision (a), docs/GH346-DPT-AUTO-REPAIR-DESIGN-20261104.md's explicit "nem nyitunk /28-at": the
normalising write lands in the SAME /27 candidate as the READ fix (one the production host window, one prod GO, not two).
See "MEGVALOSITAS 2" below the READ docstring for the write capability's own design, scope and the ONE
UNIMPLEMENTED sub-path it STOPS on rather than guesses.

THE LELET (GH #346, docs/GH346-DPT-DUPLICATE-PRECEDENCE-RESEARCH-20261104.md, LIVE calibrated, prod DS cell,
commit 25ffb0fc): the engine's own public setter (CChartSpace.SetDataPointFill/SetDataPointOutLine) addresses
oSeries.dPt BY ARRAY POSITION, not by idx -- a cross-session write on a sparse dPt list can create a second
<c:dPt> at the SAME idx. The renderer resolves this the same way in every chart type: brush and pen walk the
array independently and each takes the FIRST idx match (sdk-all.js :7:4099122 brush, :7:4100364 pen, further
:7:4101455 / :7:4102953 / :7:4108254 the same pattern) -- document-order first entry wins, for brush and pen
SEPARATELY (an entry carrying only a line colour never inherits a second entry's fill, and vice versa). The
existing /20-26 chart.series.style.read/v1 handler, on finding more than one dPt entry at the requested idx,
answers datapoint-ambiguous and refuses the WHOLE read -- it never tells the caller WHICH entry the renderer
actually draws, nor that a second entry even exists.

THE FIX, here, and ONLY in this one handler's dataPoint branch:
  1. the `hits.length > 1` refusal is removed -- `holder` is set to the FIRST hit (hits[0], document order),
     mirroring the renderer's own tie-break exactly (not a new, untried heuristic);
  2. the fill/outline/outlineWidthEmu extraction (the `h2`/`paint` helpers and their call sites) is hoisted
     ABOVE the two early-return shortcuts (the former "absent" and "present+spPr-null" returns), unifying three
     previously-separate return statements into one -- the two early returns produced BYTE-IDENTICAL output to
     what the unified path now computes for the same input shape (sp undefined -> fillV/outV/w all null, the
     same values the old sp-null shortcut hardcoded), so no existing non-dataPoint or non-duplicate caller
     observes any change;
  3. for the dataPoint target ONLY (`onPoint` true), the response gains effective:{brush,pen} -- an alias of
     fill/outline under the renderer's own names, present on every dataPoint read, duplicate or not -- and,
     ONLY when more than one entry is stored at the idx, duplicated:{count,order:[{fill,line},...]}, where
     each order entry's fill/line is read INDEPENDENTLY off that entry's OWN spPr (via the SAME paint()
     helper, never falling back to another entry's value), document order;
  4. `datapoint-ambiguous` is dropped from the handler's dPt branch AND from the semanticContract reasons list
     (the handler can no longer produce it); the semanticContract body text is extended to describe the new
     effective/duplicated fields and the first-match-wins rule, citing the GH #346 research doc and the /27
     unit tag.

No other handler, and no OTHER part of this handler (the chart/plotArea/title/legend/series branches, the
dPt-unexpected-shape / series-index-out-of-range / chart-type-not-single guards, the sppr-chain-broken check)
is touched.

NOT MEASURED in this unit (offline build + static identity + handler tests only -- no LIVE run):
  - that the fixed handler, once installed, actually answers effective/duplicated correctly on a real DS --
    the LIVE run that proves it is a separate, gated step
  - the dataPoint MCP-target release (megvalositas 3, write-tools.cjs allowlist) -- out of scope here

===========================================================================================================
MEGVALOSITAS 2 -- chart.series.dataPoint.style.write/v1 (WRITE, mutation TARGETED)
===========================================================================================================

args: { sheet, chart, series, point, fill?, outline?, outlineWidthEmu? } -- the card's own shorthand dropped
`sheet`, but every OTHER capability in this bridge (read and write alike) resolves the chart through
Api.GetSheet(args.sheet) first; there is no precedent anywhere in this family for a chart/series-addressing
capability WITHOUT it. Treated as a card shorthand, not a deliberate omission -- `sheet` is REQUIRED.

Design doc Dontesi pont 2 (Ag A), 1-5. pont, and the card's two decisions on top of it:
  1. idx-based hits collection: the SAME scan the /27 READ branch already does (s.dPt[d].idx === args.point),
     but keeping the ARRAY POSITION of each hit (removeDPt needs the position, not the idx);
  2. hits.length > 1: the FIRST (document order) is kept, the REST removed via s.removeDPt(pos), in
     DECREASING position order (so an earlier removal never shifts a later one still to be removed) --
     generalised to N>2, not just the one live-measured 2-duplicate case;
  3. hits.length === 0 (gap) + explicit fill/outline given -- IMPLEMENTED in v2 (a maintainer's review #ref, point 1):
     the gap was previously NOT implemented because no addDPt/CDataPoint precedent existed anywhere in this
     bridge family to confirm the constructor's calling shape; zero supplied the exact engine source
     (sdk-all.js :7:4151589, CChartSpace.prototype.SetDataPointFill's own gap branch) verbatim:
     `if (!(nDataPoint >= oSeries.getNumPts().length || nDataPoint < 0) && ((oDataPoint = oSeries.dPt[nDataPoint])
     || ((oDataPoint = new AscFormat.CDPt).setIdx(nDataPoint), oSeries.addDPt(oDataPoint), oDataPoint.spPr ||
     (oDataPoint.setSpPr(new AscFormat.CSpPr), oDataPoint.spPr.setParent(oDataPoint))), oDataPoint.spPr.setFill
     (oFill), ...)`. ONLY the creating half is ported (new AscFormat.CDPt() -> setIdx(args.point) ->
     s.addDPt(p) -> setSpPr(new AscFormat.CSpPr()) -> spPr.setParent(p)) -- the engine's own POSITION-based
     `dPt[nDataPoint]` lookup (the public setter's OWN bug, GH #346's root cause) is deliberately NOT ported;
     our write always resolves hits by idx (the SAME scan the /27 READ branch and point 1 above already use),
     never by array position, so a gap-fill NEVER touches another entry's slot;
  3b. point-out-of-range -- NEW in v2 (a maintainer's review, point 1's own boundary half): the SAME engine source
     guards `nDataPoint >= oSeries.getNumPts().length` before doing anything else; this build now does too
     (`s.getNumPts().length`, a point-only check, independent of the series' dPt ARRAY length -- most points
     carry no dPt entry at all) -- added on BOTH branches (gap-fill AND the already-has-an-entry branch),
     since the real guard in the cited source runs before the dPt[nDataPoint] lookup, i.e. before either path
     is known yet. Reason: point-out-of-range, 0 mutation, checked before hits are even collected;
  4. CSAK Fill/ln duplikatumot von ossze automatikusan -- a maintainer's decision (b), OVERRIDING the design doc's
     tentative "KIVEVE ha explicit formazas van" exception: a discarded (non-first) hit carrying ANY other
     dPt property (marker, dLbl, explosion, invertIfNegative, bubble3D, pictureOptions) refuses NAMED
     ("xlsx-datapoint-duplicate-unresolvable:<prop>"), 0 mutation, UNCONDITIONALLY -- explicit formatting on
     the call does NOT override this (the card is explicit: "explicit uj formazassal IS" elutasit). Checked
     BEFORE any removeDPt call (validate every discarded hit first, mutate only after all pass);
  4b. ELOSZOR ELLENORIZ, AZTAN MUTAL -- v2 fix (a maintainer's review, point 2): v1 validated Fill/CLn constructors
     (AscFormat.CreateSolidFillRGB / CSpPr / CLn / setFill / setLn) ONLY while building them, AFTER the
     removeDPt dedup loop had already run -- a missing constructor on a duplicate+explicit-format call left
     the series PARTIALLY normalized (the duplicate already removed) while answering ok:false. v2 moves every
     one of these existence checks (plus the gap-fill's CDPt/setIdx/setSpPr/addDPt existence, and the
     Fill/CLn objects themselves) into a single PREFLIGHT block that runs BEFORE removeDPt -- the handler's
     first real document mutation. Nothing above this preflight mutates the live document (constructing a
     detached CSpPr/CLn/CDPt, or calling setFill/setW on one, is not a document mutation until it is attached
     via setSpPr/setLn/addDPt, all of which happen strictly AFTER the preflight passes in full);
  5. explicit fill/outline/outlineWidthEmu, once exactly one entry survives (dedup, gap-fill, or already-
     singular): written onto that SURVIVING entry's spPr only -- Fill via AscFormat.CreateSolidFillRGB (the
     ONE caller-supplied-colour precedent in this engine family, build-rev15.py/9.3.4.60's line-style write) +
     spPr.setFill (named by direct symmetry with the ALREADY-established spPr.ln / spPr.setLn pair this same
     engine's chart.series.lineStyle.write/v1 uses -- a CT_SpPr container holding EITHER a fill or a line
     the same way, for a dPt exactly as for a series). outline -- CHANGED in v2 (a maintainer's review, point 3):
     v1 always built a FRESH AscFormat.CLn, seeding width/fill from the existing line by hand -- which dropped
     every OTHER line property the existing CLn carried (dash pattern, cap, join, ...). v2 confirmed
     CLn.prototype.createDuplicate exists (a maintainer, sdk-all.js :7:2220807) and uses it: when a line already
     exists, the SURVIVING line is cloned via its own createDuplicate() and ONLY the explicitly-given
     Fill/w are overwritten on the clone -- every other property (e.g. prstDash) rides along untouched. A
     fresh AscFormat.CLn() is built ONLY when there is no existing line to clone (gap-fill, or an entry whose
     line was never set) -- this still honours build-rev15.py's own "never mutate a shared object in place"
     rule: the clone is a NEW, detached object until setLn attaches it, same as the old fresh-CLn path was;
  9. idempotency: if the net result after dedup is exactly 1 entry AND (no explicit formatting given OR the
     explicit value already equals the stored one, per field), the handler returns BEFORE constructing any
     Fill/CLn/CSpPr/CDPt object and BEFORE calling any setFill/setLn/setSpPr/setIdx/addDPt -- zero mutation,
     zero History point. For a gap (hits.length===0) this is simply "no explicit formatting given" (there is
     nothing to compare an explicit value against); for an existing entry it is the v1 per-field comparison;
  6. response shape: { outcome: NORMALIZED|WRITTEN|SATISFIED, removed, before:{count,effective},
     after:{count,effective} } -- `effective` mirrors the /27 READ branch's own {brush,pen} shape (never a
     fresh encoding), `removed` is the dedup count independent of outcome, `outcome` is WRITTEN whenever an
     explicit format value actually changed something (a gap-fill is always WRITTEN, since there is nothing
     to have "already been the same" against -- even if dedup ALSO happened), else NORMALIZED when dedup
     happened with no format change, else SATISFIED (a true no-op on every axis).

IMPLEMENTED IN v2 (see points 3/3b/4b/5 above, a maintainer's review #ref): the hits.length===0 + explicit-
formatting gap-fill path (new CDPt + setIdx + addDPt + setSpPr(new CSpPr)), the point-out-of-range boundary
guard on both branches, the preflight-before-mutate ordering fix, and outline-rewrite preservation via
CLn.prototype.createDuplicate. The HEZAGOS "write lands at the correct idx on a sparse list" acceptance case
from the card's own MI A KESZ list IS satisfied by this build; the public, position-based setter
(oSeries.dPt[nDataPoint]) is NOT ported and is never called anywhere on this write path -- every lookup,
on every branch, resolves by idx.

Usage:
  build-rev27.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""

import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V26_BLOCK = HERE.parent / "rev26-candidate" / "bridge-block.js"
V26_BLOCK_SHA = "da60be8e21aab3f89b7c9acf8be0b3c103deb59de7abff6247f1f9be97cbf796"
V26_BLOCK_SIZE = 162792

FORBIDDEN = ("Object.keys", "Object.values", "Object.entries", "getOwnPropertyNames", "for (var k in", "for (var p in", " in args", "//",
             "hashValue", "saltValue", "spinCount", "algorithmName", "password", "temporaryPassword", "PasswordXL", "asc_setSheet",
             "asc_setPassword")

PATCHVER_OLD = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/26"'
PATCHVER_NEW = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/27"'

HITS_OLD = (
    'if (!Array.isArray(s.dPt)) return { ok: false, reason: "dPt-unexpected-shape" }; var hits = []; '
    'for (var d = 0; d < s.dPt.length; d += 1) { if (s.dPt[d] && s.dPt[d].idx === args.point) hits.push(s.dPt[d]); } '
    'if (hits.length > 1) return { ok: false, reason: "datapoint-ambiguous" }; holder = hits.length === 1 ? hits[0] : null;'
)
HITS_NEW = (
    'if (!Array.isArray(s.dPt)) return { ok: false, reason: "dPt-unexpected-shape" }; var hits = []; '
    'for (var d = 0; d < s.dPt.length; d += 1) { if (s.dPt[d] && s.dPt[d].idx === args.point) hits.push(s.dPt[d]); } '
    'holder = hits.length >= 1 ? hits[0] : null;'
)

TAIL_OLD = (
    'if (holder === null || holder === void 0) return { ok: true, value: { element: "absent", fill: null, '
    'outline: null, outlineWidthEmu: null } }; var sp = holder.spPr; if (sp === null || sp === void 0) return '
    '{ ok: true, value: { element: "present", fill: null, outline: null, outlineWidthEmu: null } }; '
    'var h2 = function (n) { return ("0" + n.toString(16).toUpperCase()).slice(-2); }; '
    'var paint = function (uf) { if (uf === null || uf === void 0) return null; var f = uf.fill; '
    'if (f === null || f === void 0) return null; var uc = f.color; if (uc === null || uc === void 0) return '
    '{ form: "other" }; if (uc.Mods && uc.Mods.Mods && uc.Mods.Mods.length > 0) return void 0; var cc = uc.color; '
    'if (!cc || typeof cc !== "object") return void 0; if (typeof cc.id === "number") return { form: "scheme", '
    'id: cc.id }; var rgba = cc.RGBA; if (cc.h === void 0 || !rgba || typeof rgba.R !== "number" || '
    'typeof rgba.G !== "number" || typeof rgba.B !== "number") return void 0; return { form: "rgb", '
    'hex: h2(rgba.R) + h2(rgba.G) + h2(rgba.B) }; }; var fillV = paint(sp.Fill); var ln = sp.ln; '
    'var outV = ln ? paint(ln.Fill) : null; if (fillV === void 0 || outV === void 0) return { ok: false, '
    'reason: "sppr-chain-broken" }; var w = ln && typeof ln.w === "number" ? ln.w : null; '
    'return { ok: true, value: { element: "present", fill: fillV, outline: outV, outlineWidthEmu: w } };'
)
TAIL_NEW = (
    'var h2 = function (n) { return ("0" + n.toString(16).toUpperCase()).slice(-2); }; '
    'var paint = function (uf) { if (uf === null || uf === void 0) return null; var f = uf.fill; '
    'if (f === null || f === void 0) return null; var uc = f.color; if (uc === null || uc === void 0) return '
    '{ form: "other" }; if (uc.Mods && uc.Mods.Mods && uc.Mods.Mods.length > 0) return void 0; var cc = uc.color; '
    'if (!cc || typeof cc !== "object") return void 0; if (typeof cc.id === "number") return { form: "scheme", '
    'id: cc.id }; var rgba = cc.RGBA; if (cc.h === void 0 || !rgba || typeof rgba.R !== "number" || '
    'typeof rgba.G !== "number" || typeof rgba.B !== "number") return void 0; return { form: "rgb", '
    'hex: h2(rgba.R) + h2(rgba.G) + h2(rgba.B) }; }; '
    'var dupInfo = function (list) { var order = []; for (var oi = 0; oi < list.length; oi += 1) { '
    'var oe = list[oi]; var osp = oe && oe.spPr; var of = osp ? paint(osp.Fill) : null; '
    'var oln = osp && osp.ln; var ol = oln ? paint(oln.Fill) : null; order.push({ fill: of, line: ol }); } '
    'return { count: list.length, order: order }; }; '
    'if (holder === null || holder === void 0) return { ok: true, value: { element: "absent", fill: null, '
    'outline: null, outlineWidthEmu: null } }; var sp = holder.spPr; var fillV = sp ? paint(sp.Fill) : null; '
    'var ln = sp && sp.ln; var outV = ln ? paint(ln.Fill) : null; if (fillV === void 0 || outV === void 0) '
    'return { ok: false, reason: "sppr-chain-broken" }; var w = ln && typeof ln.w === "number" ? ln.w : null; '
    'var val = { element: "present", fill: fillV, outline: outV, outlineWidthEmu: w }; '
    'if (onPoint) { val.effective = { brush: fillV, pen: outV }; if (hits.length > 1) val.duplicated = dupInfo(hits); } '
    'return { ok: true, value: val };'
)

CONTRACT_BODY_OLD = (
    'a colour with modifiers or an unreadable chain refuses (sppr-chain-broken); two dPt of one idx refuse; '
    'single-plot charts only for series / dataPoint;'
)
CONTRACT_BODY_NEW = (
    'a colour with modifiers or an unreadable chain refuses (sppr-chain-broken); for dataPoint, more than one '
    'stored entry at the requested idx (GH #346, the position-based public setter\'s own duplicate) no longer '
    'refuses: fill/outline/outlineWidthEmu answer the FIRST document-order entry (the renderer\'s own '
    'first-match-wins rule, LIVE measured and calibrated, GH346-DPT-DUPLICATE-PRECEDENCE-RESEARCH-20261104), '
    'the response also carries effective:{brush,pen} (fill/outline under the renderer\'s own names, present on '
    'every dataPoint read) and, only when more than one entry is stored, duplicated:{count,order:[{fill,line},'
    '...]} (every stored entry\'s own fill/line, read independently off its own spPr, document order); '
    'single-plot charts only for series / dataPoint;'
)

REASONS_OLD = 'dPt-unexpected-shape, datapoint-ambiguous, sppr-chain-broken (GH #346, /20)" },'
REASONS_NEW = 'dPt-unexpected-shape, sppr-chain-broken (GH #346, /20, /27)" },'

BLOCK_ENTRY = 'Api.GetEuroOfficeBridge = function () {'

# ---- megvalositas 2: chart.series.dataPoint.style.write/v1 -----------------------------------------------

DPT_WRITE_REASONS = (
    "bad-args, invalid-fill, invalid-outline, invalid-outlineWidthEmu, sheet-not-found, charts-threw, "
    "chart-not-found, chart-name-ambiguous, chart-type-not-single, series-index-out-of-range, "
    "dPt-unexpected-shape, getnumpts-missing, getnumpts-threw, getnumpts-unexpected-shape, "
    "point-out-of-range (GH #346 v2, the engine's own SetDataPointFill boundary guard, both branches), "
    "xlsx-datapoint-duplicate-unresolvable:<prop>, removedpt-missing, removedpt-threw, "
    "cdpt-missing, cdpt-threw, setidx-missing, setidx-threw, adddpt-missing, adddpt-threw, "
    "setsppr-missing, cspPr-missing, cspPr-threw, setsppr-threw, "
    "setfill-missing, setfill-threw, createsolidfillrgb-missing, createsolidfillrgb-threw, cln-missing, "
    "cln-threw, createduplicate-missing, createduplicate-threw, "
    "lnsetfill-missing, lnsetfill-threw, setw-missing, setw-threw, setln-missing, setln-threw"
)

DPT_WRITE_ENTRY = (
    '{ id: "chart.series.dataPoint.style.write/v1", version: 1, kind: "WRITE", format: "XLSX", mutation: "TARGETED", '
    'semanticContract: "one chart data point (sheet, chart name EXACTLY one match -- chart-name-ambiguous if '
    'more, the SAME strict convention chart.drawing.delete.write/v1 /13 and chart.series.marker.write/v1 '
    'already use for an irreversible write -- series index, point idx), at least one of fill / outline '
    '({form: rgb, hex} or null, the SAME shape chart.series.style.read/v1 already reads back) / '
    'outlineWidthEmu (non-negative integer EMU): GH #346 AUTO-REPAIR normalising write -- the engine\'s own '
    'public setter (CChartSpace.SetDataPointFill/SetDataPointOutLine) addresses oSeries.dPt by ARRAY '
    'POSITION, not by idx, and a cross-session write on a sparse dPt list can create a second <c:dPt> at the '
    'SAME idx (docs/GH346-DPT-DUPLICATE-PRECEDENCE-RESEARCH-20261104.md, LIVE calibrated); this capability '
    'collects ALL dPt entries at the requested idx the SAME way the /27 READ branch does, keeps the FIRST '
    '(document order, mirroring the renderer\'s own first-match-wins rule) and removes every other one via '
    'the internal, History-tracked CSeriesBase.prototype.removeDPt(pos) (NOT idx -- positions are computed '
    'BEFORE any removal and removed in DECREASING order so an earlier removal never shifts a later one), '
    'generalised to any N>1 stored entries; point FIRST checked against the series\' own point count '
    '(point-out-of-range if point >= s.getNumPts().length, the SAME boundary the engine\'s own public setter '
    'checks, GH #346 v2) on EVERY call, before hits are even collected; a discarded entry carrying any dPt '
    'property OTHER than spPr (marker, dLbl, explosion, invertIfNegative, bubble3D, pictureOptions -- the '
    'renderer\'s first-match-wins rule is NOT measured for these) refuses NAMED '
    '(xlsx-datapoint-duplicate-unresolvable:<prop>), 0 mutation, UNCONDITIONALLY -- an explicit fill/outline '
    'on the SAME call does not override this; every constructor/method this call will need (removeDPt for a '
    'dedup, CDPt/setIdx/setSpPr/addDPt for a gap-fill, CreateSolidFillRGB/CSpPr/CLn/createDuplicate/setFill/'
    'setLn/setW for the format write) is checked and the Fill/CLn objects built BEFORE removeDPt runs -- a '
    'missing constructor refuses NAMED with ZERO mutation, never a partially-normalized series (GH #346 v2); '
    'if the series has NO entry at point (gap) and explicit formatting is given, a new AscFormat.CDPt is '
    'constructed, setIdx(point), added via s.addDPt (idx-based, NEVER the position-based dPt[point] the '
    'engine\'s own public setter uses -- GH #346\'s own root cause), then setSpPr(new AscFormat.CSpPr); once '
    'exactly one entry survives (dedup, gap-fill, or already-singular), an explicit fill/outline/'
    'outlineWidthEmu writes that entry\'s OWN spPr only (Fill via AscFormat.CreateSolidFillRGB + spPr.setFill; '
    'outline, when a line already exists, clones it via its own CLn.prototype.createDuplicate and overwrites '
    'only the given Fill/w -- every other line property, e.g. dash pattern, rides along untouched; a fresh '
    'AscFormat.CLn is built only when there was no existing line -- never mutating a shared line object in '
    'place, the SAME lesson chart.series.lineStyle.write/v1 /15 already encodes); if the net result is '
    'already exactly one entry AND (no explicit value given OR every given value already equals the stored '
    'one) the call is SATISFIED with ZERO mutation and ZERO History point BEFORE any Fill/CLn/CSpPr/CDPt '
    'object is even constructed (idempotent retry, gap-fill included); the result is {outcome: '
    'NORMALIZED|WRITTEN|SATISFIED, removed: <dedup count, independent of outcome>, before:{count,effective}, '
    'after:{count,effective}}, effective mirroring chart.series.style.read/v1\'s own {brush,pen} shape; '
    'served only through write(); reasons: ' + DPT_WRITE_REASONS + ' (GH #346, /27)" }, '
)

_DPT_WRITE_LOOKUP = (
    'if (!args || typeof args !== "object" || typeof args.sheet !== "string" || typeof args.chart !== "string") return { ok: false, reason: "bad-args" }; '
    'if (!(Number.isInteger(args.series) && args.series >= 0)) return { ok: false, reason: "bad-args" }; '
    'if (!(Number.isInteger(args.point) && args.point >= 0)) return { ok: false, reason: "bad-args" }; '
    'var hasFill = Object.prototype.hasOwnProperty.call(args, "fill"); '
    'var hasOutline = Object.prototype.hasOwnProperty.call(args, "outline"); '
    'var hasOutlineWidth = Object.prototype.hasOwnProperty.call(args, "outlineWidthEmu"); '
    'var hasExplicit = hasFill || hasOutline || hasOutlineWidth; '
    'var validColor = function (v) { return v === null || (v && typeof v === "object" && v.form === "rgb" && typeof v.hex === "string" && /^[0-9A-Fa-f]{6}$/.test(v.hex)); }; '
    'if (hasFill && !validColor(args.fill)) return { ok: false, reason: "invalid-fill" }; '
    'if (hasOutline && !validColor(args.outline)) return { ok: false, reason: "invalid-outline" }; '
    'if (hasOutlineWidth && !(typeof args.outlineWidthEmu === "number" && args.outlineWidthEmu >= 0 && Math.floor(args.outlineWidthEmu) === args.outlineWidthEmu)) return { ok: false, reason: "invalid-outlineWidthEmu" }; '
    'var ws = Api.GetSheet(args.sheet); if (!ws) return { ok: false, reason: "sheet-not-found" }; '
    'var charts; try { charts = ws.GetAllCharts() || []; } catch (e) { return { ok: false, reason: "charts-threw" }; } '
    'var matches = []; for (var i = 0; i < charts.length; i += 1) { var okName = false; try { okName = charts[i].GetName() === args.chart; } catch (e) { okName = false; } if (okName) matches.push(charts[i]); } '
    'if (matches.length === 0) return { ok: false, reason: "chart-not-found" }; '
    'if (matches.length > 1) return { ok: false, reason: "chart-name-ambiguous" }; '
    'var api = matches[0]; var cs = api.Chart; var plot = cs && cs.chart && cs.chart.plotArea; var charts2 = plot && plot.charts; '
    'if (!charts2 || charts2.length !== 1) return { ok: false, reason: "chart-type-not-single" }; '
    'var ch = charts2[0]; var series = ch.series || []; var s = series[args.series]; '
    'if (!s) return { ok: false, reason: "series-index-out-of-range" }; '
    'if (!Array.isArray(s.dPt)) return { ok: false, reason: "dPt-unexpected-shape" }; '
    'if (typeof s.getNumPts !== "function") return { ok: false, reason: "getnumpts-missing" }; '
    'var numPtsArr; try { numPtsArr = s.getNumPts(); } catch (e) { return { ok: false, reason: "getnumpts-threw" }; } '
    'if (!Array.isArray(numPtsArr)) return { ok: false, reason: "getnumpts-unexpected-shape" }; '
    'if (args.point >= numPtsArr.length) return { ok: false, reason: "point-out-of-range" }; '
)

_DPT_WRITE_HELPERS = (
    'var h2 = function (n) { return ("0" + n.toString(16).toUpperCase()).slice(-2); }; '
    'var paint = function (uf) { if (uf === null || uf === void 0) return null; var f = uf.fill; '
    'if (f === null || f === void 0) return null; var uc = f.color; if (uc === null || uc === void 0) return '
    '{ form: "other" }; if (uc.Mods && uc.Mods.Mods && uc.Mods.Mods.length > 0) return void 0; var cc = uc.color; '
    'if (!cc || typeof cc !== "object") return void 0; if (typeof cc.id === "number") return { form: "scheme", '
    'id: cc.id }; var rgba = cc.RGBA; if (cc.h === void 0 || !rgba || typeof rgba.R !== "number" || '
    'typeof rgba.G !== "number" || typeof rgba.B !== "number") return void 0; return { form: "rgb", '
    'hex: h2(rgba.R) + h2(rgba.G) + h2(rgba.B) }; }; '
    'var effectiveOf = function (entry) { if (!entry) return null; var sp = entry.spPr; if (!sp) return { brush: null, pen: null }; '
    'var fv = paint(sp.Fill); var ln = sp.ln; var ov = ln ? paint(ln.Fill) : null; '
    'return { brush: fv === void 0 ? null : fv, pen: ov === void 0 ? null : ov }; }; '
    'var extraPropOf = function (entry) { '
    'if (entry.marker !== null && entry.marker !== void 0) return "marker"; '
    'if (entry.dLbl !== null && entry.dLbl !== void 0) return "dLbl"; '
    'if (entry.explosion !== null && entry.explosion !== void 0) return "explosion"; '
    'if (entry.invertIfNegative !== null && entry.invertIfNegative !== void 0) return "invertIfNegative"; '
    'if (entry.bubble3D !== null && entry.bubble3D !== void 0) return "bubble3D"; '
    'if (entry.pictureOptions !== null && entry.pictureOptions !== void 0) return "pictureOptions"; '
    'return null; }; '
)

DPT_WRITE_IMPL = (
    '"chart.series.dataPoint.style.write/v1": function (args) { '
    + _DPT_WRITE_LOOKUP + _DPT_WRITE_HELPERS +
    'var hits = []; for (var d = 0; d < s.dPt.length; d += 1) { if (s.dPt[d] && s.dPt[d].idx === args.point) hits.push({ pos: d, entry: s.dPt[d] }); } '
    'var before = { count: hits.length, effective: hits.length ? effectiveOf(hits[0].entry) : null }; '
    'if (hits.length > 1) { for (var ci = 1; ci < hits.length; ci += 1) { var bad = extraPropOf(hits[ci].entry); if (bad) return { ok: false, reason: "xlsx-datapoint-duplicate-unresolvable:" + bad }; } } '
    'var isGap = hits.length === 0; var survivorEntry = hits.length ? hits[0].entry : null; '
    'var sp0 = survivorEntry ? survivorEntry.spPr : null; '
    'var existingFill = sp0 ? paint(sp0.Fill) : null; var existingLn = sp0 && sp0.ln; '
    'var existingOutline = existingLn ? paint(existingLn.Fill) : null; '
    'var existingWidth = (existingLn && typeof existingLn.w === "number") ? existingLn.w : null; '
    'var sameHex = function (a, b) { if (a === null && b === null) return true; if (!a || !b) return false; '
    'if (a.form !== "rgb" || b.form !== "rgb") return false; return a.hex.toUpperCase() === b.hex.toUpperCase(); }; '
    'var fillSame = !hasFill || sameHex(args.fill, existingFill); '
    'var outlineSame = !hasOutline || sameHex(args.outline, existingOutline); '
    'var widthSame = !hasOutlineWidth || args.outlineWidthEmu === existingWidth; '
    'var needsFormatWrite = isGap ? hasExplicit : (hasExplicit && !(fillSame && outlineSame && widthSame)); '
    'if (isGap && !hasExplicit) return { ok: true, value: { outcome: "SATISFIED", removed: 0, before: before, after: before } }; '
    'if (!isGap && hits.length <= 1 && !needsFormatWrite) return { ok: true, value: { outcome: "SATISFIED", removed: 0, before: before, after: before } }; '
    '/* ---- PREFLIGHT: every capability this call will need is validated BEFORE any document mutation (removeDPt is the FIRST one) ---- */ '
    'if (hits.length > 1 && typeof s.removeDPt !== "function") return { ok: false, reason: "removedpt-missing", before: before, removed: 0 }; '
    'var needsSpPrBuild = isGap || (!sp0 && needsFormatWrite); var preparedSpPr = null; '
    'if (needsSpPrBuild) { '
    'if (typeof (AscFormat && AscFormat.CSpPr) !== "function") return { ok: false, reason: "cspPr-missing", before: before, removed: 0 }; '
    'try { preparedSpPr = new AscFormat.CSpPr(); } catch (e) { return { ok: false, reason: "cspPr-threw", before: before, removed: 0 }; } '
    'if (!isGap && typeof survivorEntry.setSpPr !== "function") return { ok: false, reason: "setsppr-missing", before: before, removed: 0 }; } '
    'var fillTarget = sp0 || preparedSpPr; var lnTarget = sp0 || preparedSpPr; '
    'if (needsFormatWrite && hasFill && !fillSame && (!fillTarget || typeof fillTarget.setFill !== "function")) '
    'return { ok: false, reason: "setfill-missing", before: before, removed: 0 }; '
    'var needOutlineObj = needsFormatWrite && (hasOutline || hasOutlineWidth) && !(outlineSame && widthSame) && !(hasOutline && args.outline === null && !hasOutlineWidth); '
    'var needOutlineClear = needsFormatWrite && hasOutline && args.outline === null && !hasOutlineWidth && !outlineSame; '
    'if ((needOutlineObj || needOutlineClear) && (!lnTarget || typeof lnTarget.setLn !== "function")) '
    'return { ok: false, reason: "setln-missing", before: before, removed: 0 }; '
    'var gapPoint = null; '
    'if (isGap) { '
    'if (typeof (AscFormat && AscFormat.CDPt) !== "function") return { ok: false, reason: "cdpt-missing", before: before, removed: 0 }; '
    'try { gapPoint = new AscFormat.CDPt(); } catch (e) { return { ok: false, reason: "cdpt-threw", before: before, removed: 0 }; } '
    'if (typeof gapPoint.setIdx !== "function") return { ok: false, reason: "setidx-missing", before: before, removed: 0 }; '
    'if (typeof gapPoint.setSpPr !== "function") return { ok: false, reason: "setsppr-missing", before: before, removed: 0 }; '
    'if (typeof s.addDPt !== "function") return { ok: false, reason: "adddpt-missing", before: before, removed: 0 }; } '
    'var buildFill = function (hex) { '
    'if (typeof AscFormat === "undefined" || typeof AscFormat.CreateSolidFillRGB !== "function") return { ok: false, reason: "createsolidfillrgb-missing" }; '
    'var r = parseInt(hex.substr(0, 2), 16), g = parseInt(hex.substr(2, 2), 16), b = parseInt(hex.substr(4, 2), 16); '
    'var built; try { built = AscFormat.CreateSolidFillRGB(r, g, b); } catch (e) { return { ok: false, reason: "createsolidfillrgb-threw" }; } '
    'return { ok: true, fill: built }; }; '
    'var preparedFill = null; '
    'if (needsFormatWrite && hasFill && !fillSame && args.fill !== null) { '
    'var bf = buildFill(args.fill.hex); if (!bf.ok) return { ok: false, reason: bf.reason, before: before, removed: 0 }; preparedFill = bf.fill; } '
    'var preparedLn = null; '
    'if (needOutlineObj) { '
    'if (existingLn) { '
    'if (typeof existingLn.createDuplicate !== "function") return { ok: false, reason: "createduplicate-missing", before: before, removed: 0 }; '
    'try { preparedLn = existingLn.createDuplicate(); } catch (e) { return { ok: false, reason: "createduplicate-threw", before: before, removed: 0 }; } '
    '} else { '
    'if (typeof (AscFormat && AscFormat.CLn) !== "function") return { ok: false, reason: "cln-missing", before: before, removed: 0 }; '
    'try { preparedLn = new AscFormat.CLn(); } catch (e) { return { ok: false, reason: "cln-threw", before: before, removed: 0 }; } } '
    'if (hasOutline && args.outline !== null) { '
    'if (typeof preparedLn.setFill !== "function") return { ok: false, reason: "lnsetfill-missing", before: before, removed: 0 }; '
    'var bo = buildFill(args.outline.hex); if (!bo.ok) return { ok: false, reason: bo.reason, before: before, removed: 0 }; '
    'try { preparedLn.setFill(bo.fill); } catch (e) { return { ok: false, reason: "lnsetfill-threw", before: before, removed: 0 }; } } '
    'if (hasOutlineWidth) { '
    'if (typeof preparedLn.setW !== "function") return { ok: false, reason: "setw-missing", before: before, removed: 0 }; '
    'try { preparedLn.setW(args.outlineWidthEmu); } catch (e) { return { ok: false, reason: "setw-threw", before: before, removed: 0 }; } } } '
    '/* ---- MUTATE: removeDPt (dedup) first -- every capability above is already confirmed present ---- */ '
    'var removed = 0; '
    'if (hits.length > 1) { '
    'var toRemove = []; for (var ri = 1; ri < hits.length; ri += 1) toRemove.push(hits[ri].pos); '
    'toRemove.sort(function (a, b) { return b - a; }); '
    'for (var ti = 0; ti < toRemove.length; ti += 1) { '
    'try { s.removeDPt(toRemove[ti]); } catch (e) { return { ok: false, reason: "removedpt-threw", before: before, removed: removed }; } '
    'removed += 1; } } '
    'var survivor; '
    'if (isGap) { '
    'try { gapPoint.setIdx(args.point); } catch (e) { return { ok: false, reason: "setidx-threw", before: before, removed: removed }; } '
    'try { s.addDPt(gapPoint); } catch (e) { return { ok: false, reason: "adddpt-threw", before: before, removed: removed }; } '
    'try { gapPoint.setSpPr(preparedSpPr); } catch (e) { return { ok: false, reason: "setsppr-threw", before: before, removed: removed }; } '
    'if (gapPoint.spPr && typeof gapPoint.spPr.setParent === "function") { try { gapPoint.spPr.setParent(gapPoint); } catch (e) { } } '
    'survivor = gapPoint; '
    '} else { survivor = survivorEntry; } '
    'if (!needsFormatWrite) { '
    'var after0 = { count: 1, effective: effectiveOf(survivor) }; '
    'return { ok: true, value: { outcome: removed > 0 ? "NORMALIZED" : "SATISFIED", removed: removed, before: before, after: after0 } }; } '
    'var sp = survivor.spPr; '
    'if (!sp) { '
    'try { survivor.setSpPr(preparedSpPr); } catch (e) { return { ok: false, reason: "setsppr-threw", before: before, removed: removed }; } '
    'sp = survivor.spPr || preparedSpPr; } '
    'if (hasFill && !fillSame) { '
    'if (args.fill === null) { try { sp.setFill(null); } catch (e) { return { ok: false, reason: "setfill-threw", before: before, removed: removed }; } } '
    'else { try { sp.setFill(preparedFill); } catch (e) { return { ok: false, reason: "setfill-threw", before: before, removed: removed }; } } } '
    'if (needOutlineObj) { '
    'try { sp.setLn(preparedLn); } catch (e) { return { ok: false, reason: "setln-threw", before: before, removed: removed }; } '
    '} else if (needOutlineClear) { '
    'try { sp.setLn(null); } catch (e) { return { ok: false, reason: "setln-threw", before: before, removed: removed }; } } '
    'var after = { count: 1, effective: effectiveOf(survivor) }; '
    'return { ok: true, value: { outcome: "WRITTEN", removed: removed, before: before, after: after } }; }, '
)
MARKERS = [
    '"chart.series.style.read/v1": function', '{ id: "chart.series.style.read/v1"',
    '"chart.series.dataPoint.style.write/v1": function', '{ id: "chart.series.dataPoint.style.write/v1"',
]


def sha(b):
    return hashlib.sha256(b).hexdigest()


def build_block():
    v26 = V26_BLOCK.read_bytes()
    if sha(v26) != V26_BLOCK_SHA or len(v26) != V26_BLOCK_SIZE:
        sys.exit(f"REFUSE: /26 block sha/size {sha(v26)}/{len(v26)} != {V26_BLOCK_SHA}/{V26_BLOCK_SIZE}; NO OUTPUT")
    t = v26.decode("utf-8")
    if t.count(BLOCK_ENTRY) != 1:
        sys.exit("REFUSE: the /26 block does not carry the bridge entrypoint exactly once; NO OUTPUT")
    for new in (HITS_NEW, TAIL_NEW, DPT_WRITE_ENTRY, DPT_WRITE_IMPL):
        for word in FORBIDDEN:
            if word in new:
                sys.exit(f"REFUSE: new fragment contains {word!r}; NO OUTPUT")
    CAPS_ARRAY_TAIL_OLD = (
        'ial-unobservable, credential-not-cleared, history-missing, history-off, setter-threw, '
        'no-effect (GH #338-#342, /24)" }, ], };'
    )
    CAPS_ARRAY_TAIL_NEW = (
        'ial-unobservable, credential-not-cleared, history-missing, history-off, setter-threw, '
        'no-effect (GH #338-#342, /24)" }, ' + DPT_WRITE_ENTRY + '], };'
    )
    WRITES_TAIL_OLD = (
        'return { ok: true, value: { applied: [prop], table: t1.name, observed: observed } }; }, }; '
        'var EO_BRIDGE_HYDRATIONS = {'
    )
    WRITES_TAIL_NEW = (
        'return { ok: true, value: { applied: [prop], table: t1.name, observed: observed } }; }, '
        + DPT_WRITE_IMPL + '}; var EO_BRIDGE_HYDRATIONS = {'
    )
    subs = [
        (PATCHVER_OLD, PATCHVER_NEW),
        (HITS_OLD, HITS_NEW),
        (TAIL_OLD, TAIL_NEW),
        (CONTRACT_BODY_OLD, CONTRACT_BODY_NEW),
        (REASONS_OLD, REASONS_NEW),
        (CAPS_ARRAY_TAIL_OLD, CAPS_ARRAY_TAIL_NEW),
        (WRITES_TAIL_OLD, WRITES_TAIL_NEW),
    ]
    for old, new in subs:
        n = t.count(old)
        if n != 1:
            sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:70]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1:
            sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    if t.count("datapoint-ambiguous") != 0:
        sys.exit("REFUSE: datapoint-ambiguous still referenced after substitution; NO OUTPUT")
    if "effective" not in t or "duplicated" not in t:
        sys.exit("REFUSE: effective/duplicated not found after substitution; NO OUTPUT")
    if "point-out-of-range" not in t or "createDuplicate" not in t:
        sys.exit("REFUSE: point-out-of-range/createDuplicate not found after substitution; NO OUTPUT")
    if "dpt-gap-write-unimplemented" in t:
        sys.exit("REFUSE: dpt-gap-write-unimplemented still referenced -- v2 implements the gap, NO OUTPUT")
    return v26, t.encode("utf-8")


def main():
    v26, v27 = build_block()
    out = {"v26_block_sha": sha(v26), "v26_block_size": len(v26), "v27_block_sha": sha(v27), "v27_block_size": len(v27)}
    (HERE / "bridge-block.js").write_bytes(v27)
    (HERE / "old.txt").write_bytes(v26)
    (HERE / "new.txt").write_bytes(v27)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v27-gh346-dpt-effective-read-and-normalize-write",
                             "old": v26.decode("utf-8"), "new": v27.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (the disposable DS "
                              "first, the live DS needs its own owner GO).",
                 original_sha256=sha(v26), patched_sha256=sha(v27), original_size=len(v26), patched_size=len(v27))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
