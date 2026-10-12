# euro-office-bridge-word-v1 /2 CANDIDATE (M11.5d.1, B18-WRITE)

`/2` = `/1` + a `write()` entry + `docx.paragraph.numbering.remove/v1` (WRITE, TARGETED) -- wraps the editor's own
`CParagraph.RemoveNumPr` (the product's B18 routine; no public Builder call removes numbering, see
`docs/M11-5D-B18-NUMBERING-REMOVE-BRIDGE-WRITE-DESIGN.md`). The two `/1` READ capabilities stay byte-identical.
Build: `python3 build-rev2.py --sdk <served /1 word sdk-all.js>` (input `8362e3a9…` / 18548802 -> output
`acd8fa14a65341ef2aa8a5c3adcf5c4e0ffb43627300732219b2e417410be924` / 18552617). Deterministic (two rebuilds identical).
Apply with `../../patch_sdk_all.py <sdk-all.js> --manifest manifest.json --check|--apply [--output]`.
OFFLINE + SESSION-LOCAL only -- not applied to any Document Server.
