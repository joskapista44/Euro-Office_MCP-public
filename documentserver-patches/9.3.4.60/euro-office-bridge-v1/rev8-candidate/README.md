# euro-office-bridge-v1 /8 CANDIDATE (M11.5c.1, B03-WRITE)

`/8` = `/7` + `worksheet.copy/v1` (WRITE, TARGETED) -- wraps the editor's own `asc_copyWorksheet` (the product's B03
routine; the model routine reorders sheets on save, see `docs/M11-5C-B03-SHEET-COPY-BRIDGE-WRITE-DESIGN.md`).
Build: `python3 build-rev8.py --sdk <served /7 cell sdk-all.js>` (input `3a73042b…` / 20962662 -> output
`73c9211dba0385ff4700f3507526812f20219583b28ed115be909ac7ca171ca6` / 20966785). Deterministic (two rebuilds identical).
Apply with `../../patch_sdk_all.py <sdk-all.js> --manifest manifest.json --check|--apply [--output]`.
OFFLINE + SESSION-LOCAL only -- not applied to any Document Server.
