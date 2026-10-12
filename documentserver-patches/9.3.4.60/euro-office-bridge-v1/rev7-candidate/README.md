# euro-office-bridge-v1 patchVersion /7 -- CANDIDATE (M11.5b.1)

Owner: M11.5 continuous mandate (owner channel, internal tracker). Built and offline-verified; DS apply is the
M11.5b.3 step (compatibility-first: the product accepting /7 ships BEFORE the apply). Rollback = the /6 bundle
`af42c232…` (production today).

## What /7 adds

/7 = /6 + exactly ONE capability + `patchVersion` `euro-office-bridge-v1/9.3.4.60/7`. All 17 /6 entries (15 READ,
1 RUNTIME_HYDRATION, `range.alignmentExt.write/v1`) and every /6 implementation are byte-identical (asserted: /7 minus
the two inserted segments and the version bump == /6).

`sheet.printSettings.write/v1` -- `kind: WRITE`, `mutation: TARGETED`, served by the existing `write()` entry. PUBLIC_FIRST:
the 9.3.4 public ApiWorksheet print surface is margins / orientation / gridlines / headings only; none of the eight B10
settings has a public setter (`docs/evidence/m11-lgb/explore.json`).

```
INPUT   { sheet, setting, ...fields }  ONE setting per call:
        fit {fitToWidth, fitToHeight 0..100}  (0+0 = fit off, also scale 100)
        scale {scale 10..400}  (fit off)       paper {width, height}  (one of the editor's 13 sizes, mm)
        firstPage {firstPage 1..32767 | null}  center {axis horizontal|vertical, value boolean}
        headerFooter {slot odd|first|even Header|Footer, value 1..255 chars | null}
        headerFooterFlag {flag differentFirst|differentOddEven, value boolean}
        pageBreak {axis row|column, id (model break id), remove boolean}  -- manual break
OUTPUT  { ok: true, value: { applied: [sheet.printSettings.read/v1 field names] } }  -- a RECEIPT, not a verification
        { ok: false, reason }  |  { ok: false, reason: '<setter>-threw', partial: bool, applied: [...] }
REASONS bad-args unsupported-property unknown-setting invalid-value sheet-not-found model-missing history-missing
        <setter>-missing <setter>-threw
```
- Same engine path: for every B10 intent the capability makes EXACTLY the calls of the product's current
  direct-internal lowering (`print.cjs lower()`), in order, including the header/footer history transaction (PARITY test,
  17 shapes covering all 8 intents).
- Every input validated and every needed setter checked BEFORE the first mutation; transaction always closed.
- Verification authority: a fresh `sheet.printSettings.read/v1`; partial writes are reported, never rolled back; no fallback.

## Known input / output (computed 2026-10-11)

```
input  (/6 served, production)  af42c2327de11bd2537f7532187de33b13811aef67949d4a6b2f7c2ff3916b63 / 20955549 bytes
output (/7 candidate)           3a73042b6d4314925ea83907fec9d1ef48246facb2dcbf4b35e5741fee070b2e / 20962662 bytes
block  /6 4faf6031… (38615 chars) -> /7 74ecfd1539fa3b158f72d240110162834641ae337293d8b256fb6f149454937d (45728 chars)
```
`build-rev7.py` is the source (three exact, count-asserted substitutions on the /6 block); two rebuilds byte-identical.
With the unmodified `../../patch_sdk_all.py`: input `--check` PATCHABLE; `--apply --output` PATCHED (sha above); output
`--check` ALREADY_PATCHED (3 markers count=1); second apply NO WRITE; unknown input (one prepended byte) INCOMPATIBLE rc=3,
no output; `node --check` of the patched bundle passes; mechanism proof (old.txt -> new.txt) byte-equal.
