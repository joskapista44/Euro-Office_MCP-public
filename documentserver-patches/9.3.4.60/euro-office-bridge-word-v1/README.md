# EuroOffice DocumentServer 9.3.4.60 WORD Bridge patch (patchVersion /1)

Owner GO (PATCH_BUILD only): the owner, internal tracker (#ref point 4 -- "Ha a /2 a word
bundle-t IS modositja... B19/B20 kesobb kulon"), dispatch: zero, tracker comment #ref. **This
is the build-only step.** The word-bundle LIVE APPLY is a SEPARATE identity-gate, owner-decision
step (a maintainer's own dispatch, citing the contract doc's word-bundle-identity-gate requirement,
hash `f293db63...65c8`) -- "a cell-GO NEM extrapolalhato" (the cell-bundle GO does not extend to
this bundle). Nothing in this family touches any live Document Server; that is a later,
separately-gated operation.

A SEPARATE family from `euro-office-bridge-v1/` (which targets `sdkjs/cell/sdk-all.js`) -- this
one targets the WORD bundle, `sdkjs/word/sdk-all.js`, a DIFFERENT served file with its own
identity. Same handshake, same capability-ID contract, same canonical patcher:

```
Api.GetEuroOfficeBridge() -> {
  manifest(): Manifest,
  read(capabilityId: string, args: object): Result
}
```

`capabilityId` is a semantic, versioned ID -- never a class, method or object-path name. No
`call`/`eval`/`invoke(methodName)` surface. An ID not in the manifest -> `{ ok: false, reason:
'capability-not-advertised' }`. Every exception inside a capability is caught and returned as
`{ ok: false, reason: 'internal-threw' }`, never thrown across the boundary.

## Wave 0: word bundle identity

The currently-served `sdkjs/word/sdk-all.js`:

```
f293db632c599a9ace68c765feea751be12a2b500408d7a997e20e6dea6e65c8 / 18544083 bytes
(a maintainer's own measurement, 2026-10-06 21:34:44 -- verified again here via SHA256SUMS in
agents/an internal branch/handoff/maintainer-word-bridge-20261006/)
```

Confirmed UNPATCHED: `Api.GetEuroOfficeBridge` occurs `0` times in the served file.

Confirmed `Api` is a SINGLETON here too, not a prototype -- same pattern as the cell bundle:

```
let Api=Object.create(AscBuilder.Word.Api||{});
```

(cell bundle: `var Api=Object.create(AscBuilder.Slide.Api||{})`.) Its own methods are own
properties (`Api.AddComment=function(...)`, `Api.AI=function(...)`, `Api.attachEvent=
function(...)`, etc.), confirmed by inspecting the served bundle directly, not assumed.

## Capabilities (both READ, both `mutation: NO`, neither creating a model object or writing history)

- `docx.numbering.levelFormat.read/v1` (B19) -- `{ paragraph: <0-based document-element
  index> }` -> `{ ok: true, value: { present: false } }` (the paragraph is not in a list), or
  `{ ok: true, value: { present: true, level, numId, levelFormats } }` (`levelFormats` always a
  full array when `ok: true` -- never `null`), or `{ ok: false, reason }`. `level` is the 0-based
  numbering level (via the public `ApiNumberingLevel.GetLevelIndex()`); `numId` is the internal
  list-continuity identity (`el.Paragraph.Pr.NumPr.NumId`, which no public getter exposes --
  measured live 2026-09-24, list probe 1, cited in `docx-driver.cjs`'s own comment at the read
  site); `levelFormats` is an array of up to nine per-level LIST TYPE codes
  (`inner.GetLvl(k).GetFormat()` for `k` in `0..8`, `null` for a level that does not exist),
  read from the internal `CNum` via `lv.GetNumbering().Num` -- the SAME internal dependency the
  driver's own comment names ("M5 gate 7... no public getter... read -- READ ONLY -- from the
  internal CNum level format"). **`{ ok: true }` means a successful read with the declared
  type, per the Bridge contract (`docs/DS-BRIDGE-V1-CONTRACT-EVIDENCE-20261006.md`) -- every
  read gap is `{ ok: false, reason }`** (fixed per a maintainer's #ref review: an earlier draft
  returned `{ ok: true, value: { ..., levelFormats: null } }` when the internal `CNum` reader
  was unavailable or threw, which is contract-level fail-open -- a caller cannot tell "the
  reader is unavailable" apart from "successfully read, and this document's numbering happens
  to carry no level formats" by inspecting `ok` alone. If the internal level-format reader is
  unavailable or throws AFTER `present: true` is established, the WHOLE capability now fails
  closed -- it never emits a partially-successful result with a null field standing in for the
  missing part. The driver's own existing `null`-means-unread mapping is UNCHANGED on the
  driver side; this capability's contract is what changed.) Reasons: `bad-args`,
  `document-unavailable`, `paragraph-out-of-range`, `paragraph-unreadable`, `not-a-paragraph`,
  `no-GetNumbering`, `getNumbering-threw`, `level-format-unavailable`, `level-format-threw`.
- `docx.table.effectiveGrid.read/v1` (B20) -- `{ table: <0-based index into
  Api.GetDocument().GetAllTables()> }` -> `{ ok: true, value: { gridCalc: number[] } }`
  (`gridCalc` always a full, non-empty array of finite numbers when `ok: true` -- never `null`)
  or `{ ok: false, reason }`. `gridCalc` is the engine's CALCULATED column grid widths -- no
  public getter exposes them (`GetWidth` absent, split-grid probe 2026-09-25, cited in
  `docx-driver.cjs`'s own comment at the read site) -- read from the internal `t.Table.
  TableGridCalc` array, accepted only if it is a non-empty array of finite numbers (the SAME
  validation the driver itself applies before trusting the value). **Same contract fix as B19
  (a maintainer's #ref): an absent, malshaped, empty, or throwing internal array now fails closed
  (`{ ok: false, reason: 'grid-calc-unavailable' }` / `'grid-calc-threw'`), never `{ ok: true,
  value: { gridCalc: null } }`** -- the driver's own existing null-means-unread mapping stays
  the driver's own concern, unaffected by this capability's own `ok`/`reason` contract.
  Reasons: `bad-args`, `document-unavailable`, `tables-threw`, `tables-not-a-list`,
  `table-not-found`, `grid-calc-threw`, `grid-calc-unavailable`.

## Insertion point

A fresh anchor-insert, NOT a body-replace -- there is no existing word-bundle Bridge block to
collide with (confirmed `0` occurrences of `Api.GetEuroOfficeBridge` in the served input), so
this is the SAME situation as the cell bundle's ORIGINAL /1 insert (`chart-public-getters/` ->
`euro-office-bridge-v1/1`), not the situation that caused the cell-bundle's #ref defect (where
an anchor-insert was wrongly used against an input that ALREADY had a Bridge block glued to the
same anchor). Anchor: the existing, exactly-one-occurrence `Api.AddComment=function(element,
text,author,userId){...}` statement -- a unique, singleton-`Api` own-property, confirmed by
direct inspection, the word bundle's equivalent of the cell bundle's `Api.AddSheet`. `new` =
anchor text, followed by `,Api.GetEuroOfficeBridge=function(){...}` (the full two-capability
body, no trailing comma at the very end -- the comma that already followed the anchor in the
real file remains untouched, outside the matched span, leading into the next statement,
`Api.attachEvent=function(...)`).

Confirmed via the two checks the cell-bundle review rounds established as mandatory: exactly
`1` occurrence of `Api.GetEuroOfficeBridge` in the patched output (not two racing in the
comma-chain), and the inserted text, evaluated against a mocked `Api`, yields `manifest().
patchVersion === 'euro-office-bridge-word-v1/9.3.4.60/1'` and `manifest().capabilities.length
=== 2`.

## Known input / output

```
original_sha256: f293db632c599a9ace68c765feea751be12a2b500408d7a997e20e6dea6e65c8 / 18544083 bytes
patched_sha256:  8362e3a9fe07173e7ce9ece91e73d0829bdd7bd6ecf2b13a28330ae0e08b4699 / 18548802 bytes
```

(+4719 bytes -- the full two-capability implementation, including the #ref contract fix.)

## Patcher

Shared with every 9.3.4.60 family -- the SAME canonical `../patch_sdk_all.py --manifest
manifest.json`, no second patcher.

```
python3 ../patch_sdk_all.py <sdk-all.js> --manifest manifest.json --check
python3 ../patch_sdk_all.py <sdk-all.js> --manifest manifest.json --apply --output <dest>
```

Verified offline (nothing touched any live Document Server -- the word-bundle LIVE APPLY is a
separate, owner-decision identity-gate per the contract doc, not yet reached):

- `--check` on the real served word bundle -> `PATCHABLE` (both getter markers `count=0`).
- `--apply --output`, run twice -> byte-identical output both times, and byte-identical to the
  independently hand-rebuilt insert (`cmp` clean).
- **Exactly `1` occurrence of `Api.GetEuroOfficeBridge` in the patched output** (the cell-family
  #ref lesson, re-verified here even though this is a fresh insert, not a replace).
- **The inserted text, evaluated against a mocked `Api`, yields `patchVersion ===
  'euro-office-bridge-word-v1/9.3.4.60/1'` and `capabilities.length === 2`** (the cell-family
  #ref second gate).
- `--check` on the output -> `ALREADY_PATCHED`, both getter markers `count=1`.
- `--check` on an unrelated file (the CELL bundle's own original) -> `INCOMPATIBLE`, both
  markers `count=0`.
- `node --check` on the full ~18.5 MB patched file -> syntactically valid.
- The Bridge implementation, minified mechanically (comment-strip + whitespace-collapse, no
  hand-editing) from a readable source, run against the SAME 19-case mock harness and diffed
  byte-for-byte identical before the minified version was installed as the patch artifact.
- 19 named smoke/edge cases covering both capabilities (manifest shape; B19: not-in-list, a full
  read with mixed existing/absent levels, the internal-reader-unavailable AND the
  internal-reader-threw paths each now returning `{ ok: false, reason }` -- the #ref fix --
  not-a-paragraph, paragraph-out-of-range, bad-args, document-unavailable; B20: a full read, and
  each of absent/empty/non-numeric/throwing internal array now returning `{ ok: false, reason }`
  -- the same fix -- table-not-found, bad-args, document-unavailable; plus
  `capability-not-advertised`) against a mocked `Api.GetDocument()` (not the real engine -- that
  needs a LIVE session): every positive/negative/fail-closed case named above returns the
  expected, named result.

The full ~18.5 MB `sdk-all.js` (word) bundle is deliberately not stored in Git (same convention
as every other 9.3.4.60 family).

## Not done by this patch

Nothing is installed on any live Document Server. No LIVE apply to the word bundle -- that is a
SEPARATE identity-gate, owner-decision step (a maintainer's dispatch, contract doc section on the word
bundle). No MCP-side routing change. No WRITE capability (READ-only, per the contract). B11 and
any other cell-bundle capability are NOT part of this family (separate target file entirely).
