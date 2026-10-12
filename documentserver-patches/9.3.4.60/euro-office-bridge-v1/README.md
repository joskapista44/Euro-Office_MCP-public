# EuroOffice DocumentServer 9.3.4.60 Bridge patch (v1) -- current revision: patchVersion /5

## patchVersion /5 -- B15 RUNTIME_HYDRATION (2026-10-09, zero)

Owner GO: the owner, owner channel 2026-10-09 msg 33257 (internal tracker): B15 = MIGRATE_PRE_M10, KIND=RUNTIME_HYDRATION,
"the /5 adds ONLY the capability B15 needs". **Built and offline-verified only -- NOT applied to any Document Server.**
The apply on the production host is a separate, per-use owner gate (an owner-approval gate), executed by maintainer with the /4 discipline
(PRE_DISK_SHA / PRE_SERVED_SHA = the /4 output below, byte-exact backup, apply, post sha); rollback = restore the /4
output (sha below).

One whole-block replace, as /3 -> /4: `old` = the entire /4 `bridge-block.js` (33630 characters, exactly one
occurrence in the served file), `new` = the entire /5 block (35411 characters). /5 = /4 + ONE entry +
`patchVersion` `euro-office-bridge-v1/9.3.4.60/5`; the fifteen /4 READ entries are byte-identical:

- `sheet.model.ensureReady/v1` -- `kind: RUNTIME_HYDRATION`, `mutation: EPHEMERAL_VIEW_ONLY`. **Not a READ**: it is served
  ONLY by a new `hydrate(capabilityId, args)` entry next to `read()`; `read()` refuses it (`capability-not-advertised`)
  and `hydrate()` refuses every READ id, so the READ contract is unchanged. `{ sheet }` -> `{ modelReady: true }` (no
  document data). It runs the editor's own lazy view creator for that sheet's index -- the same call the driver made
  directly (B15, `WorkbookView.getWorksheet`), now behind the bridge. Allowed: an ephemeral, session-only view object.
  Forbidden: document content/structure, history/undo, persistent metadata, dirty/save state. Reasons: `bad-args`,
  `sheet-not-found`, `model-threw`, `view-creator-missing`, `view-creator-threw`, `no-view`.

Known input / output (computed 2026-10-09 from the SERVED /4 file, fetched over HTTPS, = the accepted /4 artifact):

```
input  (/4 output, served today) def79eed14771fc61bf28162dfd5a83868f3d1c549cc2c1079dff26515b0bd5a / 20950564 bytes
output (/5)                       f1462d87f211c6d6cdfe999465c02b7a0f551139ab4738eefe02657b71d053ee / 20952345 bytes
```

Verified with the unmodified `patch_sdk_all.py`: `--check` on the input -> `PATCHABLE`; `--apply --output <scratch>` ->
`PATCHED`, sha = the output above; `--check` on the output -> `ALREADY_PATCHED`, `"sheet.model.ensureReady/v1": function`
`count=1`; a second `--apply` on the output -> `NO WRITE: exact accepted patched state already present`; an unknown
input (the served file with one prepended byte) -> `INCOMPATIBLE`, `NO WRITE`, rc=3; `node --check` on the output passes.
Mechanism proof on the isolated block (`manifest-mechanism-proof.json`, old.txt -> new.txt): PATCHED, byte-equal to
`new.txt`. `chart-public-getters/` `--check` on the /5 output: `NATIVE_OR_UNKNOWN_SUPPORT` (as for /4).
Capability count: /4 15 -> /5 16.

The sections below document /1-/4 and are kept unchanged.

# (patchVersion /4 header, kept)

## patchVersion /4 -- Wave B READ (2026-10-08, zero; design `docs/DS-BRIDGE-WAVE-B-READ-DESIGN-20261008.md`)

Owner order: the owner, owner channel 2026-10-08, point 9 (direct-internal READs first). **Built and offline-verified only --
NOT applied to any Document Server.** The apply on the production host is a separate, per-use owner gate (an owner-approval gate), executed
by maintainer; rollback = re-apply the /3 output (sha below).

One whole-block replace, as /2 -> /3: `old` = the entire /3 `bridge-block.js` (26548 characters, exactly one occurrence
in the served file), `new` = the entire /4 block (33630 characters). /4 = /3 + three READ capabilities (all
`mutation: NO`, no model object created, no history written) + `patchVersion` `euro-office-bridge-v1/9.3.4.60/4`:

- `range.alignmentExt.read/v1` (B05 READ) -- `{ sheet, cell }` -> `{ readingOrder, indent, shrink }`. A never-set
  readingOrder (`undefined`) is the measured default 0 (S31c); any other gap refuses the whole read: `getAlign-missing`,
  `readingOrder-not-number`, `indent-missing`, `shrink-missing`, plus `bad-args` / `sheet-not-found` / `cell-not-found` /
  `range-missing`.
- `sheet.printSettings.read/v1` (B09) -- `{ sheet }` -> `{ fitW, fitH, fitFlag, oddHeader, rowBreaks, useFirstPage,
  firstPage, hCenter, scale, width, height, oddFooter, vCenter, firstHeader, firstFooter, evenHeader, evenFooter,
  diffFirst, diffOddEven }` with the driver's former normalisations; the row breaks mirror `sheet.columnBreaks.read/v1`
  on the row axis (`idMax` 1048575, `spanMax` 16383). Reasons: `bad-args`, `sheet-not-found`, `model-missing` (any of
  the used getters absent), `rowBreaks-absent`, `rowBreaks-unexpected-shape`, `rowBreaks-element-invalid`.
- `pivot.state.read/v1` (B13) -- `{ sheet, pivot }` (the pivot's public `GetName()`) -> `{ dataRef }`. Reasons:
  `bad-args`, `sheet-not-found`, `pivots-threw`, `pivot-not-found`, `pivot-name-ambiguous`, `pivot-model-missing`,
  `dataRef-not-string`. Known limit kept: after a source-sheet rename the reference still names the OLD sheet.

B15 (sheet hydration) is NOT in /4: it waits for the owner/review decision on the "READ may create a view" contract line.

Known input / output (computed 2026-10-08 from the SERVED /3 file, fetched over HTTPS, byte-identical to the /3 output):

```
input  (/3 output, served today) 4b6fa6535db999b6fbc7d17e3555e8e290d9ad68bf1ef295d796f25c596aa7e8 / 20943482 bytes
output (/4)                       def79eed14771fc61bf28162dfd5a83868f3d1c549cc2c1079dff26515b0bd5a / 20950564 bytes
```

Verified with the unmodified `patch_sdk_all.py`: `--apply --output <scratch>` -> WRITE VERIFY ok; `--check` on the
output -> `ALREADY_PATCHED`, the three new markers `count=1`; `node --check` on the output passes. Mechanism proof on the
isolated block (`manifest-mechanism-proof.json`, old.txt -> new.txt): PATCHABLE -> PATCHED, byte-equal to `new.txt`.
`chart-public-getters/` `--check` on the /4 output: `NATIVE_OR_UNKNOWN_SUPPORT`, its ten getters `count=1` (as for /3).

The sections below document /1-/3 and are kept unchanged.

# EuroOffice DocumentServer 9.3.4.60 Bridge patch (v1, patchVersion /2)

Owner GO: the owner, owner channel, 2026-10-06 (internal tracker). Design:
`docs/DS-BRIDGE-V1-CONTRACT-20261006.md`. Separately versioned from `chart-public-getters/`,
which stays unchanged -- this family's INPUT is that family's own OUTPUT (both target the
currently served `sdk-all.js`).

Adds one entry point to the public Builder `Api` object:

```
Api.GetEuroOfficeBridge() -> {
  manifest(): Manifest,                             // a fresh, independent copy every call
  read(capabilityId: string, args: object): Result   // only READ capabilities of the manifest
}
```

`capabilityId` is a semantic, versioned ID (`textStyle.bold.read/v1`) -- never a class, method
or object-path name. There is no `call`/`eval`/`invoke(methodName)` surface. An ID not in the
manifest -> `{ ok: false, reason: 'capability-not-advertised' }`. Every exception inside a
capability is caught and returned as `{ ok: false, reason: 'internal-threw' }`, never thrown
across the boundary.

## Version history note

**patchVersion /1 shipped (textStyle.bold.read/v1, B01) and is what is actually served today.**
An intermediate five-capability /2 and a four-capability /3 were built and offline-verified in
sequence (a maintainer's dispatch #ref/#ref/#ref) but **neither was ever installed anywhere** --
zero caught this (internal tracker) before delivery to main: chaining two never-shipped
intermediate states onto the real served file would mean two patcher runs and two intermediate
SHAs on the the production host/the production host install step, for no benefit. **This README now documents a single
jump**: one manifest, one operation, input = today's real served file (/1's own output), output
= all nine new capabilities (/2, this version). The prior /2-then-/3 staged history is not kept
here as a superseded trail (unlike /1->/2, which WAS real) -- it never had a real SHA pair
outside this patch family's own git history, and keeping it would invite the exact "LANDOLT !=
KISZOLGALT" confusion the fleet's own discipline warns against.

## Capabilities (all READ, all `mutation: NO`, none creating a model object or writing history)

- `textStyle.bold.read/v1` (B01, unchanged from /1) -- `{ sheet, cell }` -> `{ ok: true, value:
  boolean }` or `{ ok: false, reason }`. Internal path: the SAME one the old B01 reader uses
  (`cell.range.getFont(false).getBold()`), reached via the public `Api.GetSheet(sheet)` +
  `ApiWorksheet.GetRange(cell)` chain. **Parity with the old reader is exact, not best-effort**
  (a maintainer's review, internal tracker): a non-boolean `getBold()` result (or a missing `getFont`) is
  `{ ok: false, reason: 'bold-not-boolean' }` / `'font-unavailable'`, never silently coerced to
  `false` -- the old path's `readInternalFontFlag` (`xlsx-driver.cjs:397-405`) draws the SAME
  three-way distinction, and this capability mirrors it rather than collapsing it to a boolean
  default.
- `textStyle.italic.read/v1` (B02) -- identical shape and reasons to B01, `getItalic()` instead
  of `getBold()` (`xlsx-driver.cjs:397-405,416`).
- `range.appearance.read/v1` (B04) -- `{ sheet, cell }` -> `{ ok: true, value: { align: { h, v
  }, border: { t, b, l, r, d: { s, c }, dd, du } } }` or `{ ok: false, reason }`. Fields cited
  exactly from the current driver read (`xlsx-driver.cjs:456-493`), including the SAME
  shared-edge fixup the driver applies for the engine's own `getBorderFull` gap (row 2's top
  edge, column B's left edge -- read from the neighbour's own public `GetRange(...).range.
  getBorderFull()`). **A failed neighbour read propagates as `neighbour-unreadable`, never
  silently treated as "no override"** (a maintainer's fix, internal tracker): the driver wraps the ENTIRE
  border read in one try/catch and turns ANY throw -- including one from the neighbour lookup
  -- into `UNOBSERVABLE:threw` for the whole axis (`xlsx-driver.cjs` ~490, the M7 S4 edge-probe
  comment). The first readable source swallowed a neighbour-read throw or a missing
  `.range`/`getBorderFull` as "no override found," which is a DIFFERENT, looser result than the
  driver's own fail-closed behaviour. Reasons: `getAlign-missing`, `getBorderFull-missing`,
  `neighbour-unreadable`, `border-unexpected-shape`, plus
  `bad-args`/`sheet-not-found`/`cell-not-found`/`range-missing`.
- `sheet.autoFilter.fieldIdentity.read/v1` (B07, narrowed) -- `{ sheet }` -> `{ ok: true, value:
  { range: <A1 address>|null, fields: [<1-based field#>, ...] } }` or `{ ok: false, reason }`.
  Deliberately narrow: `GetAutoFilter`/`GetFilters`/`GetOn`/`GetOperator`/`GetCriteria1`/
  `GetCriteria2` are ALREADY public (`xlsx-driver.cjs:680-701`) -- only the field identity
  (`f.filter.ColId`, no public getter) crosses the Bridge; `fields` is in the SAME order as the
  public `GetFilters()` list, for the caller to zip by position. Reasons: `autofilter-threw`,
  `autofilter-filters-without-range`, `range-address-unavailable`, `field-unreadable`.
- `sheet.columnBreaks.read/v1` (B12) -- `{ sheet }` -> `{ ok: true, value: [{ id, man, min, max
  }, ...] }` (`[]` if proven empty) or `{ ok: false, reason }`. Mirrors `breaksOfModel`
  (`xlsx-driver.cjs:788-801`) for the `colBreaks` axis exactly (`idMax=16383`,
  `spanMax=1048575`). Reasons: `colBreaks-absent`, `colBreaks-unexpected-shape`,
  `colBreaks-element-invalid`.
- `chart.dataLabels.effective.read/v1` (X-O05b, a maintainer's own feasibility measurement, internal tracker/#ref, `xlsx-xo05b-effective-datalabels-feasibility.live.cjs`) -- `{ sheet, chart,
  series, point }` -> `{ ok: true, value: { level: 'POINT'|'SERIES'|'CHART'|'NONE',
  showSerName, showCatName, showVal, showPercent } }` or `{ ok: false, reason }`. Non-creating
  reads only (`charts[0].dLbls`, `series[i].dLbls`, `dLbls.findDLblByIdx(point)`), precedence
  point -> series -> chart, exactly a maintainer's measured model. **`NONE`'s all-`false` flags are NOT
  the "guessed default for an unobservable read" anti-pattern this codebase elsewhere fail-closes
  on** (that anti-pattern is "we could not determine the state, so we assumed false"; here EVERY
  level was positively checked and found absent, which is itself the ECMA-376-defined "nothing
  configured to show labels" state -- a determined CHART_SIDE semantic, not a measurement gap).
  Flagged here explicitly because the shapes look alike. Reasons: `chart-not-found`,
  `chart-type-not-single`, `series-index-out-of-range`, `<level>-field-not-boolean:<name>`.
- `sheet.freezeState.read/v1` (B06) -- `{ sheet }` -> `{ ok: true, value: { state: 'none' } |
  { state: 'frozen', row0, col0 } }` or `{ ok: false, reason }`. Reads the TARGET sheet's own
  `worksheet.getSheetView().pane` (`xlsx-driver.cjs:595-625`), not the active view -- the public
  `ApiFreezePanes.GetLocation()` candidate reads `wb.getWorksheet()` (the currently-active sheet)
  and loses the tri-state distinction (a maintainer's own code reading ruled it out, no LIVE probe
  needed). A non-null pane with `state` other than `null`/`'frozen'`, or without a valid
  `topLeftFrozenCell`, is a split/unknown pane -- never coerced to `'none'`. Reasons:
  `sheet-view-missing`, `pane-split-or-unknown`, `pane-unexpected-shape`, `worksheet-missing`.
- `drawing.anchor.read/v1` (B14) -- `{ sheet, drawing }` -> `{ ok: true, value: { type,
  to: { col, colOff, row, rowOff } } }` or `{ ok: false, reason }`. `drawing` identifies ONE
  drawing on the sheet by its own `GetName()`, resolved via the public `ws.GetAllDrawings()`
  (the driver's own access pattern, `xlsx-driver.cjs:740-752`) -- no new internal dependency for
  drawing *lookup*, only for the anchor fields themselves (`Drawing.drawingBase.Type` +
  `.boundsFromTo.to`, `xlsx-driver.cjs:190-220`). The public `ApiDrawing.GetPosition()` candidate
  gives only the FROM anchor, never `Type` or the computed `to` bounds (a maintainer's code reading).
  **Checks `base.worksheet` BEFORE reading `boundsFromTo` (a maintainer's #ref fix): on a non-active
  sheet the view is lazy (B15, not built by this patch) and an un-hydrated `boundsFromTo` can
  read as a false/zero value instead of throwing -- a fail-open the driver itself guards against
  with the SAME check, `GEOMETRY_HYDRATED_JS` (`xlsx-driver.cjs:354`, wired into `danchor`
  `~xlsx-driver.cjs:751`): `if (!base.worksheet) return 'UNOBSERVABLE:geometry-not-hydrated'`.**
  Reasons: `drawing-not-found`, `no-drawingBase`, `geometry-not-hydrated`, `no-boundsFromTo-to`,
  `drawingBase-threw`, `boundsFromTo-threw`, plus the shared `no-GetAllDrawings`/
  `drawings-threw`/`drawings-not-a-list` from drawing lookup.
- `drawing.text.read/v1` (B16, explicitly non-creating) -- `{ sheet, drawing }` -> `{ ok: true,
  value: <'|'-joined paragraph text, '' if textless> }` or `{ ok: false, reason }`. Reads via
  `Drawing.getDocContent()` (`xlsx-driver.cjs:360-375`) -- `undefined` means the getter itself is
  unavailable, `null` means a genuinely textless shape (`value: ''`), never the same reason. The
  public `ApiShape.GetContent()` candidate was ruled out by a maintainer: on a textless shape it calls
  `createTextBody()` internally, writing one history point inside what must stay a pure read.
  Content is wrapped via `Api.private_CreateApiDocContent(dc)` -- a private internal API method,
  the SAME one the driver itself uses at this exact call site; its presence in the served bundle
  is exactly what this patch's own `PATCHABLE`/`node --check` verification below exercises.
  Reasons: `no-getDocContent`, `getDocContent-threw`, `getDocContent-undefined`,
  `doc-content-wrap-threw`, `content-unexpected-shape`, `paragraph-text-unavailable`.
- `sheet.columnWidth.canonicalize/v1` (B17) -- `{ sheet, values: number[] }` -> `{ ok: true,
  value: number[] }` (same order as `values`) or `{ ok: false, reason }`. Pure, stateless
  arithmetic: each requested character-count width is round-tripped through the engine's own
  `charCountToModelColWidth -> modelColWidthToColWidth -> colWidthToCharCount`
  (`xlsx-driver.cjs:144,580-586`), via one sheet's model (needed only to reach the three
  functions, not to read or write any document state). The public write+read pair only produces
  this value AFTER a write; this capability needs the canonical value BEFORE one (a maintainer's
  reasoning for ruling the public candidate out). Reasons: `bad-args` (non-array/non-number
  `values`), `api-missing` (any of the three functions absent), `canonicalize-threw`.

### Behaviour note: the chart-getter family's own `--check` on this family's OUTPUT

After this patch is applied, running `chart-public-getters/`'s own `--check` against the
RESULT reports `NATIVE_OR_UNKNOWN_SUPPORT` (not `ALREADY_PATCHED`) -- its manifest only knows
its OWN output SHA, and this family's output is a DIFFERENT, later SHA layered on top. The ten
getter counts still read `1` each (`PASS`), which is the part that actually matters post-
install: a future upgrade-revalidation of `chart-public-getters/` should assert on the getter
counts plus this family's own `patched_sha256`, not on `chart-public-getters/`'s own
`ALREADY_PATCHED` state, once this family is layered on top. Verified directly (not asserted):
`python3 ../patch_sdk_all.py <bridge output> --manifest ../chart-public-getters/manifest.json
--check` -> `NATIVE_OR_UNKNOWN_SUPPORT`, all ten getters `count=1`.

## Insertion point

**NOT an anchor-insert after `Api.AddSheet` -- a body-replace of the entire /1
`Api.GetEuroOfficeBridge=function(){...}` block.** An earlier draft of this README (and the
manifest it described) tried the anchor-insert approach, assuming nothing yet followed the
`Api.AddSheet=function(sName){...}` anchor in the real /1 output. That assumption was wrong, and
zero caught it before delivery (internal tracker): the real /1 output (1229545e) is ITSELF the
result of inserting the /1 Bridge block right after that same anchor -- so appending a SECOND
block after the anchor again produces, in the comma-expression-chained statement sequence,
`Api.AddSheet=...,Api.GetEuroOfficeBridge=<NEW /2 body>,Api.GetEuroOfficeBridge=<OLD /1 body>,
...rest`. In a comma-chain, the LATER assignment to the same property wins -- so after
evaluation `Api.GetEuroOfficeBridge` would resolve to the OLD, one-capability /1 body, with the
new one silently shadowed and unreachable, while `--check` still reported `ALREADY_PATCHED`
(the getter-marker counts only check TEXT presence, not which of two definitions actually binds
at runtime). Fixed: the operation now REPLACES the entire, exactly-one-occurrence /1 block
(`old` = the full /1 `bridge-block.js` text, 1524 characters, the SAME text and the SAME
technique `#ref`'s original /2 delivery used) with the full /2 body -- there is exactly one
`Api.GetEuroOfficeBridge` definition in the result, confirmed below, not two.

`Api` here is the Excel builder's own singleton object (`var Api=Object.create(AscBuilder.
Slide.Api||{})`), not a prototype -- its own methods are own properties (`Api.AddSheet=
function(...)`, `Api.GetSheet=function(...)`, etc.), confirmed by inspecting the served bundle
directly (not assumed from the chart-getter family's own `ApiChart.prototype.*` convention,
which is a DIFFERENT, sibling class).

## Known input / output

**patchVersion /2** input (= /1's own `patched_sha256`, i.e. today's actually-served file):

```
1229545ef5fcac775fd24fa7d858c06deb46b9204555d8208dddff9e4a8c344a / 20918458 bytes
```

Patched output (/2):

```
93775964843fcde586822a6e6f5918d4040855fe2845e366a3128d84d38d59af / 20934705 bytes
```

(+16247 bytes net -- the /1 body, 1524 characters, replaced by the full ten-capability body,
17771 characters, including the geometry-hydration gate below.)

**patchVersion /1** (historical, what is actually served today -- unchanged, kept for the
upgrade trail): input `1ea9602263e0408452608a1862b31b72759a143f6df117f8f6bd7d41faf0cc03` /
20916934 bytes (= `chart-public-getters/`'s own `patched_sha256`), output
`1229545ef5fcac775fd24fa7d858c06deb46b9204555d8208dddff9e4a8c344a` / 20918458 bytes (+1524
bytes, the Bridge entry point + the exact-parity fix).

**The operation's `old` is the entire, exactly-one-occurrence /1 `bridge-block.js` text** (1524
characters, verified via `step1fixed.count(old) == 1` against the real rebuilt /1 output, not
assumed), **`new` is the entire /2 `bridge-block.js` text** (17695 characters) -- a whole-body
replace, not an anchor-insert; see "Insertion point" above for why the anchor-insert approach
was wrong. The entry-point name itself (`Api.GetEuroOfficeBridge=function`) is deliberately NOT
one of the `getters` inventory markers, for the same reason as every prior version: it is
already present at `count=1` in the /1 input (from /1 itself), so it alone cannot discriminate
PATCHABLE-from-ALREADY_PATCHED; each manifest's `getters` list uses markers unique to what THAT
patch version itself adds -- the nine new capabilities' own dispatch-key text, each individually
`count=0` before and `count=1` after.

## Patcher

Shared with every 9.3.4.60 family -- `../patch_sdk_all.py --manifest manifest.json`. See that
script's own header for the generalization note (a maintainer's dispatch via internal tracker/#ref: one
canonical patcher, manifest-driven, no second patcher).

```
python3 ../patch_sdk_all.py <sdk-all.js> --manifest manifest.json --check
python3 ../patch_sdk_all.py <sdk-all.js> --manifest manifest.json --apply --output <dest>
```

Verified offline (not yet installed on any live Document Server -- that is a separate,
an owner-approval gate owner-GO-gated operations step, per the contract doc section 7). All SHAs
below were obtained by rebuilding the REAL served-bundle chain -- a maintainer's handoff original
(`sdk-all.served-1ea96022.js`) -> `chart-public-getters/`'s own `patch_sdk_all.py --apply`
(reaching /1's own input) -> the historical /1 manifest's `patch_sdk_all.py --apply` (reaching
today's real served state, 1229545e) -> THIS manifest's own operation on top, all with the
actual tool, not hand-derived:

- `--check` on the real /1 output (today's served file) -> `PATCHABLE` (all nine new getter
  markers read `count=0`, confirming they are genuinely absent).
- `--apply --output`, run twice -> byte-identical output both times, and byte-identical to the
  independently hand-rebuilt chain (`cmp` clean).
- **`Api.GetEuroOfficeBridge` definition-count check on the patched output -> exactly `1`**
  (a maintainer's #ref gate, added because the getter-marker counts alone cannot catch two
  definitions of the SAME property racing in a comma-chain -- see "Insertion point" above).
  `grep -c`/`String.count` on the literal text, not a regex that could miscount across string
  literals containing that same text.
- **The `new` operation text evaluated against a mocked `Api`, confirming `manifest().
  patchVersion === 'euro-office-bridge-v1/9.3.4.60/2'` and `manifest().capabilities.length ===
  10`** (a maintainer's #ref gate) -- this is the behavioural confirmation that the ONE surviving
  definition is actually the NEW one, not a leftover old one that happens to also satisfy the
  text-presence checks.
- `--check` on the output -> `ALREADY_PATCHED`, all nine new getter markers `count=1`.
- `--check` on the real /1 INPUT (a maintainer's handoff original, before even /1 applied) ->
  `INCOMPATIBLE`, all nine markers `count=0`.
- `node --check` on the full ~20.9 MB patched file -> syntactically valid.
- `chart-public-getters/`'s own `--check` on the output -> `NATIVE_OR_UNKNOWN_SUPPORT`, all ten
  of ITS getters still `count=1` (same documented behaviour as /1, one layer further up).
- The Bridge implementation, minified mechanically (comment-strip + whitespace-collapse, no
  hand-editing) from a readable source, run against the SAME mock harness and diffed
  byte-for-byte identical before the minified version was installed as the patch artifact -- the
  two can drift apart only if someone edits one without re-running this check.
- 19 named smoke/edge cases covering all ten capabilities (manifest shape, B01+B02, five B04
  cases including both the shared-edge-fixup success path AND the two `neighbour-unreadable`
  failure modes -- a thrown neighbour `GetRange` and a neighbour whose `.range` is missing --
  B07, B12, X-O05b, B06/B16/B17 smoke, two B14 cases (`base.worksheet` present -> normal success;
  `base.worksheet` absent with an otherwise-present `boundsFromTo` -> `geometry-not-hydrated`,
  never silently returning the un-hydrated bounds), plus `capability-not-advertised`) against a
  mocked `Api.GetSheet`/`GetRange`/`GetAutoFilter`/`GetAllCharts`/`GetAllDrawings` (not the real
  engine -- that needs a LIVE session): every positive/negative/precedence/fail-closed case
  named in each capability's own bullet above returns the expected, named result.

The full 20 MB `sdk-all.js` bundle is deliberately not stored in Git (same convention as
`chart-public-getters/`).

## Not done by this patch

Nothing is installed on any live Document Server. No MCP-side routing change (the driver's
fail-closed gate chain lives in `src/drivers/eurooffice/ds-bridge.cjs`, a maintainer's own unit, not
this patch). No WRITE capability (v1 is READ-only, per the contract).

## /3 addendum (B11-V1, card an internal branch)

Adds `range.validations.read/v1` (READ, XLSX, mutation NO) to this family. Design doc:
`docs/B11-VALIDATION-ENUMERATION-BRIDGE-DESIGN-20261007.md`. Scope of this patch step: DESIGN +
OFFLINE IMPLEMENTATION + OFFLINE TEST + PATCH_BUILD only -- no live the production host apply, no driver
cutover.

- `bridge-block.js` / `manifest.json` updated in place to `/3` (patchVersion
  `euro-office-bridge-v1/9.3.4.60/3`), following the SAME whole-block replacement convention as
  `/1`->`/2`.
- `manifest.json`'s `original_sha256` is the KNOWN `/2` bundle hash (this family's own prior
  `patched_sha256`, cited, not fabricated). Its `patched_sha256` is **intentionally `null`** with a
  `_patched_sha256_NOTE` field: it is the sha256 of the real ~20 MB served `sdk-all.js` AFTER this
  patch, and that file is not part of this git tree -- it cannot be computed from source alone.
  Whoever has the real `/2`-served file must run `patch_sdk_all.py --apply` for real and fill it
  in before this manifest is trusted for an actual apply.
- `manifest-mechanism-proof.json` is a SEPARATE manifest whose `original_sha256`/`patched_sha256`
  are the hashes of `old.txt`/`new.txt` (the isolated `GetEuroOfficeBridge` block alone, not the
  real bundle) -- used to run the REAL, unmodified `patch_sdk_all.py` end to end (`--check` ->
  `PATCHABLE`, `--apply` -> byte-identical to the shipped `bridge-block.js`, plus
  `ALREADY_PATCHED`/`INCOMPATIBLE` controls) and prove the substitution/getter-inventory logic for
  this new operation is sound, without needing the real bundle.
- Offline tests: `test/acceptance/xlsx-b11-validation-enumeration.test.cjs`, 19 cases, against the
  shipped `bridge-block.js` text directly (sandboxed `new Function('Api', ...)`, no mock-harness
  drift risk since it is not a separate hand-maintained copy).

## /3 COMBINED addendum (supersedes the /3 addendum above;)

the owner's decision (via zero): the validation-only `/3` above (`patched_sha256
2719e5110dcc031440a18ebef8dae25bbd0820a793ea43610a3d269afd633c80`) does **not** ship. Instead, a
single `/2`->`/3` operation adds **both** `range.validations.read/v1` (unchanged from above) and a
new `workbook.definedNames.read/v1` (READ, XLSX, mutation NO), built fresh from the same `/2`
`old.txt` -- not layered on top of the abandoned validation-only `/3`.

- `workbook.definedNames.read/v1` reads `wbModel.dependencyFormulas.defNames` (`{wb:{...},
  sheet:{<sheetId>:{...}}}`) directly off each `DefName`'s stored `.ref`/`.name`/`.hidden`/
  `.isXLNM` fields -- it never calls `getRef(true)` or `getAscCDefName()` (the latter mutates the
  internal `type` field as a side effect, per a maintainer's own static analysis), proven via a spy in
  `test/acceptance/xlsx-b11-defined-names-enumeration.test.cjs` (19 cases). Full design:
  `docs/B11-VALIDATION-ENUMERATION-BRIDGE-DESIGN-20261007.md`'s "/3 COMBINED addendum" section.
- `manifest.json`'s `patched_sha256` is again **intentionally `null`** (per a maintainer's own
  instruction to either leave the old value or null it, and state which): nulled here, with a
  `_patched_sha256_NOTE` explaining that zero computes the real value from the canonical `/2`, as
  for the prior `/3`.
- `manifest-mechanism-proof.json` rebuilt fresh for this combined operation's own `old.txt`/
  `new.txt` (two getters: `range.validations.read/v1` and `workbook.definedNames.read/v1`); ran
  clean through the real, unmodified `patch_sdk_all.py` (`PATCHABLE` -> `PATCHED`, byte-identical
  to the shipped `bridge-block.js`).
- `src/drivers/eurooffice/ds-bridge.cjs` gets a `BRIDGE_CAPABILITIES.DEFINED_NAMES_READ` entry
  for the new capability only -- `range.validations.read/v1` stays unregistered there, matching
  the precedent set when it alone was delivered (no MCP-side driver cutover is in scope for
  either capability).
- `test/acceptance/xlsx-b11-validation-enumeration.test.cjs`'s own capability-count regression
  test updated 11->12 to account for the manifest now carrying both capabilities together.

## Review fixes (on top of the /3 COMBINED addendum above)

Two bugs found by a maintainer's negative-control review of p169, fixed before any `patched_sha256` was
computed: `hidden` now accepts `null` as well as `undefined` (the bundle's own table-name writer
stores it as `null` -- the old check fail-closed the ENTIRE names inventory on any workbook
containing a table); a new `table` boolean field (`type === Asc.c_oAscDefNameType.table`, with its
own `defname-type-enum-missing` fail-closed branch if that enum is unavailable); and the model
resolution no longer tries `Api.GetDocument().wbModel` at all (a wasted/risky call on a path that
never returns `wbModel` in the served bundle) -- only `Asc.editor.wbModel`, proven never-created
via a dedicated spy. Offline tests grew 19->26; the LIVE harness's SEED now also attempts (not
trusts) a table-scoped defined name, and its FILE-parity row excludes `table:true` entries. Full
detail: `docs/B11-VALIDATION-ENUMERATION-BRIDGE-DESIGN-20261007.md`'s own addendum.
