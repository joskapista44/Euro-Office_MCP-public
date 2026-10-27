# euro-office-bridge-v1 /9 CANDIDATE (X-O03/6b)

`/9` = `/8` + `chart.series.scatterY.read/v1` (READ, mutation NO). The PUBLIC chartApi's `GetValues()` is a formula
string or THROWS on a scatter series (X-O03/1, internal tracker) -- this capability reads the ACTUAL Y-value array
through the internal model instead, reusing the `chart.dataLabels.effective.read/v1` precedent's own entry point
verbatim (`api.Chart.chart.plotArea.charts[0].series[i]`): X-O03/6a (LIVE r6, internal tracker) measured this is the
SAME object (`===`) as `series.private_GetSeries()` (the path X-O03/4's own probe used to first find the Y-value,
LIVE r4 FOUND_READ_PATH, `m.yVal.getValues()` / `m.getNumPts()`). Design: `docs/X-O03-5-SCATTER-Y-READ-CONSUMER-
DESIGN-20261014.md`.

Build: `python3 build-rev9.py` (input `../rev8-candidate/bridge-block.js`, sha `9f285b897f58a751c3164791b2416e085
d627a1bd10419261622349806a96be7` / 49851 bytes -> output `fdc0724d1300360607b6a83059430a3cc3977e629dec0b6f272ade0
266943e0c` / 52792 bytes, after fixing one insertion-point bug: see the commented-out WRONG anchor in build-rev9.py's
own SUBS list -- the first build attempt inserted the new capability AFTER the `EO_BRIDGE_CAPS` object's own closing
`};`, producing a dangling `"id": function` fragment outside any object literal, `SyntaxError: Unexpected token ':'`
on the FIRST attempt caught this before any output was trusted). Deterministic (reproducible from the pinned /8
input; no timestamps, no environment input).

**Coercion fixed after review (a maintainer, card comment): null/undefined/a boolean/an empty-or-whitespace-only string
all hit a numeric 0/1 under plain `Number(x)` and never reach the `isFinite` guard** (`Number(null)===0`,
`Number("")===0`, `Number(" ")===0`, `Number(true)===1`) -- so an empty Y-cell would have silently read as `0`,
not `null`. Fixed: those four shapes are caught and forced to `NaN` (-> `null`) BEFORE `Number()` runs; every
other value still goes through `Number(x)`, non-finite as `null`. RED/GREEN proved by stashing the fix,
confirming the new test fails against the unfixed dispatch, then restoring it and confirming PASS.

**OFFLINE + SESSION-LOCAL only -- not applied to any Document Server.** Unlike `rev8-candidate/build-rev8.py`, this
build has NO `--sdk` mode: the real served `/8` `sdk-all.js` is a the production host/the production host-host artifact this agent (designated-
uid, no an owner-approval gate access) cannot fetch or verify -- producing a `manifest.json` that CLAIMS an SDK sha
without having read that file would be a fabricated measurement, not a build. Only `manifest-mechanism-proof.json`
is written (the isolated `old.txt` -> `new.txt` transformation, self-contained). A served-sdk-backed `manifest.json`
and the apply itself are a SEPARATE, later step, gated to maintainer (an owner-approval gate) with the real served file in hand
-- NOT this card's scope (X-O03/6b explicitly excludes DS install).

**Verification used (no `node --check` on the raw block -- see below):** `test/acceptance/xlsx-xo03-6b-bridge-
scatter-y-v9-candidate.test.cjs` loads the ACTUAL candidate text with `new Function('Api', src + '; return
Api.GetEuroOfficeBridge();')` (the SAME convention `test/acceptance/xlsx-b10w-bridge-write-v7-candidate.test.cjs`
established) against a fake `Api` built to the X-O03/6a LIVE-confirmed shape, and runs the real dispatch function:
manifest shape (patchVersion /9, /8 capabilities byte-identical, exactly one new READ entry), the byte-identity
reconstruction of /8 from /9 (cutting the two inserted segments back out), class separation (`write()`/`hydrate()`
refuse the READ id; `read()` refuses other versions/ids; `/8` lacks it), POS (own values, two-series no-cross-
contamination), non-finite coercion, an empty series, the 300-point ceiling (inclusive boundary), every named
refusal reason, the chart-lookup-by-name convention (first match wins, a throwing earlier `GetName()`
tolerated), and the null/undefined/empty-or-whitespace-string/boolean coercion regression above -- 15/15
PASS. `test/static/m11-5a2-dual-runtime-compat.test.cjs` and `test/static/runtime-identity.
test.cjs` extend their own existing PARITY/ACCEPTED-LIST coverage to include this REAL `/9` block (not a synthetic
patchVersion-string substitution like the older `/6`-based probes use for a hypothetical future version -- those
moved to a synthetic `/10` to keep testing an actually-unknown version).

**Why not raw `node --check documentserver-patches/.../bridge-block.js`:** every candidate block in this family
(verified on the UNMODIFIED `/8` too) fails `node --check` standalone with `SyntaxError: Unexpected end of input` --
the block is a FRAGMENT meant for splicing into the full `sdk-all.js`, not a standalone, independently-parseable
file. `node --check` on an isolated capability/manifest-entry snippet wrapped in a minimal valid context (`var Api
= {}; var EO_BRIDGE_CAPS = { <impl> };`) DOES pass for this round's additions, and is the targeted syntax check
actually used before the functional test above; the acceptance test file is the authoritative verification.

Capability count: /8 19 -> /9 20 (the same style of count bump every prior candidate's own README records).
