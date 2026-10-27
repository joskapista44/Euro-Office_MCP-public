# euro-office-bridge-v1 /10 CANDIDATE (GH #92)

`/10` = `/9` with ONE behavioural change inside the existing `range.appearance.read/v1` (no capability added or
removed): a border side the engine returns as `null` reads as `{ s: 0, c: 0 }` instead of refusing the whole read with
`border-unexpected-shape`. A missing (`undefined`) side still refuses.

Why (LIVE probe `test/live/xlsx-gh92-merge-border-read-probe.live.cjs`, disposable copy, internal tracker): right after
`ApiRange.Merge(false)`, in the SAME session, `getBorderFull()` returns `null` for every merge member's interior sides
and the diagonal (R3: 11/12 members refused). After forceSave + reopen the same sides read `{ s: 0, c: <rgb 0> }`,
byte-equal to the original merged state (R0 = R4). The keret itself is never changed; only the session-local shape was
unreadable. Symptom in the product (GH #92): `read_range style:true` reports align AND border
`UNOBSERVABLE:BRIDGE:border-unexpected-shape` on re-merged cells after unmerge -> format -> merge.

Build: `python3 build-rev10.py` (input `../rev9-candidate/bridge-block.js`, sha `fdc0724d1300360607b6a83059430a3cc3977e
629dec0b6f272ade0266943e0c` / 52792 bytes -> output `16ad3780e9ba99eb029ea8573cabdf8ff1e41cc7341eaba8cb15e2ad65e169ed` /
53015 bytes; three exact substitutions, each asserted once: patchVersion, the `side()` helper, the capability's
semanticContract). Deterministic (two runs, same sha).

**OFFLINE only -- not applied to any Document Server.** As for /9: no `--sdk` mode; the served-sdk manifest and the
apply on the production host are maintainer's gate (an owner-approval gate), with the real served /9 file in hand.

Verification: `test/acceptance/gh92-bridge-merged-null-side-v10-candidate.test.cjs` runs the ACTUAL /9 and /10 text
against a fake `Api` built to the LIVE-measured shapes: RED on /9 (border-unexpected-shape), GREEN on /10 (session ==
reopened), a missing side / non-boolean dd / non-numeric s still refuse, the row-2 / column-B neighbour fixup unchanged
by a null neighbour side, and /10 reconstructs byte-exact to /9 by undoing the three substitutions. Runtime list grows
/5../10 (compatibility-first, `src/core/runtime-identity.cjs` XLSX_V10): on a /5-/9 DS every read is unchanged.
