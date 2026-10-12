# euro-office-bridge-v1 patchVersion /6 -- CANDIDATE (M11.5a.1, OFFLINE ONLY)

Owner GO: the owner, owner channel (M11.5a.1 = offline /6 Bridge WRITE capability). **Built and offline-verified only --
NOT applied to any Document Server** (`DS_PATCH_APPLY=HOLD`), and the product is NOT changed (`PRODUCT_CUTOVER=HOLD`).
The canonical current revision stays `../bridge-block.js` (/5); this directory is the candidate for M11.5a.2.

## What /6 adds

/6 = /5 + exactly ONE capability + `patchVersion` `euro-office-bridge-v1/9.3.4.60/6`. All 16 /5 entries (15 READ +
1 RUNTIME_HYDRATION) are byte-identical, and so is every /5 implementation (asserted: /6 minus the three inserted
segments and the version bump == /5).

`range.alignmentExt.write/v1` -- `kind: WRITE`, `mutation: TARGETED`, the FIRST Bridge WRITE class. Served ONLY by a new
`write(capabilityId, args)` entry; `read()` / `hydrate()` refuse it and `write()` refuses every READ / HYDRATION id and
any other version (`capability-not-advertised`).

```
INPUT   { sheet: string, cell: <A1 of ONE cell>, indent?: integer 0..15, shrink?: boolean }   (at least one property;
        no other key -- requester/credential fields are refused as unsupported-property)
OUTPUT  { ok: true, value: { applied: [...] } }  -- an operation RECEIPT, not a verification
        { ok: false, reason }  |  { ok: false, reason: '<setter>-threw', partial: bool, applied: [...] }
ORDER   validate every input -> resolve the cell -> check every requested setter -> setIndent -> setShrinkToFit
REASONS bad-args unsupported-property no-property invalid-indent invalid-shrink invalid-target sheet-not-found
        cell-not-found range-missing setIndent-missing setShrinkToFit-missing setIndent-threw setShrinkToFit-threw
```
- Minimal mutation: an unrequested property is never touched (no getter, no setter on it).
- Verification authority: a fresh `range.alignmentExt.read/v1` (WRITE -> Bridge WRITE -> Bridge READ -> L1), never the receipt.
- `PARTIAL_WRITE_POSSIBLE=YES`: a throw in the SECOND setter after the first succeeded is reported as `partial:true`,
  `applied:['indent']`; there is no rollback (a fake one would be a false claim). The product must re-observe.
- No planner logic in the patch: SATISFIED / APPLY / FAIL_CLOSED stays a product decision; the zero-write retry is the
  product's pre-write READ, never the bridge's idempotency.
- History / document mutation: the bridge calls the SAME engine setters on the SAME cell object as today's product
  direct-internal path, inside the caller's callCommand. Offline: targeted mutation only (spy). HISTORY behaviour is
  NOT measured offline -- it is the LIVE parity axis of M11.5a.2.

## Known input / output (computed 2026-10-10)

```
input  (/5 output, SERVED today, fetched in-session)  f1462d87f211c6d6cdfe999465c02b7a0f551139ab4738eefe02657b71d053ee / 20952345 bytes
output (/6 candidate)                                  af42c2327de11bd2537f7532187de33b13811aef67949d4a6b2f7c2ff3916b63 / 20955549 bytes
block  /5 dc879526... (35411 chars) -> /6 4faf6031ab62e715880e995ae184759638e95c93419051f2a807dd303ef29325 (38615 chars)
```
`build-rev6.py` is the source (four exact, count-asserted substitutions on the canonical /5 block); two independent
rebuilds are byte-identical. With the unmodified `../../patch_sdk_all.py`: input `--check` -> PATCHABLE; `--apply
--output` -> PATCHED (sha above); output `--check` -> ALREADY_PATCHED (3 markers count=1); second apply -> NO WRITE;
unknown input (one prepended byte) -> INCOMPATIBLE, rc=3, no output; `node --check` passes; mechanism proof
(old.txt -> new.txt) PATCHED, byte-equal. `build-rev6.py` itself refuses a wrong /5 block or a wrong sdk input with NO OUTPUT.

## Compatibility (measured, product untouched)

The product's runtime identity is an EXACT `patchVersion` match (`src/core/runtime-identity.cjs` `classifyHandshake`).
Today's product (frozen `53434ef`) on a /6 DS -> `PATCH_VERSION_MISMATCH` -> every runtime-gated path fails closed.
So `OLD_PRODUCT + DS6` is FAIL_CLOSED, not PASS -- an owner decision for M11.5a.2 (see the checkpoint report).
