# euro-office-bridge-word-v1 /3 CANDIDATE (M11.5e.1, B21-READ)

`/3` = `/2` + `docx.section.variantState.read/v1` (READ) in the existing `read()` map -- the section variant state (title
page / even-and-odd), which no public getter exposes (PUBLIC_FIRST: `ApiSection` has `SetTitlePage` only, no getter; the even/odd flag is document-level `SetEvenAndOddHdrFtr`, no getter -- internal tracker; the English design note lands with the cutover). Every
`/2` capability (2 READ + the B18 WRITE) stays byte-identical. Built by maintainer (internal tracker), verified by a maintainer.
Build: `python3 build-rev3.py --sdk <served /2 word sdk-all.js>` (input `acd8fa14…` / 18552617 -> output
`7aeb04e9b602691c64cbbf0cf665c32b064a75565684ae3be930d53b5fa82a09` / 18554740). Deterministic.
Apply with `../../patch_sdk_all.py <sdk-all.js> --manifest manifest.json --check|--apply [--output]`.
OFFLINE only -- not applied to any Document Server.
