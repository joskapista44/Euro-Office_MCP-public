#!/usr/bin/env python3
"""
GH #344 / #346 / #354 / #355 (Bridge /20, DESIGN docs/DS-9.3.5-BRIDGE20-PAGE-SETUP-AND-READS-DESIGN.md) -- OFFLINE build of the
euro-office-bridge-v1 patchVersion /20 READ CANDIDATE for DS 9.3.5, in ONE deterministic pass from the /19 CANDIDATE
block (../rev19-candidate/bridge-block.js, sha pinned).

/20 (READ half) = /19 + THREE new READ capabilities, no mutation, no new history type:
  sheet.printSettings.read/v2     -- the /19 v1 value UNCHANGED (computed by the v1 handler itself) + headerMargin, footerMargin (mm as
                                     stored) + pageOrder ('downThenOver' | 'overThenDown'); v1 stays served
  sheet.autoFilter.colorFilter.read/v1 -- per colour-filter column: kind (cellColor | fontColor) + the colour the filter's dxf holds
  chart.series.style.read/v1      -- fill + outline of ONE of the six #346 targets (chart | plotArea | title | legend | series | dataPoint)
The WRITE half (header/footer margin, page order: new history change types 16-18) waits for the design's gate G0 and is NOT here.

The handler source is the readable, multi-line handler-read-v20.txt next to this file; this builder joins its lines with single
spaces (the block is a single-line fragment) and inserts it, plus the manifest entries, with exact string substitutions, each asserted
to occur exactly once -- the same pattern as build-rev19.py. No timestamps, no environment input. It NEVER writes a Document Server
file and has NO --sdk mode: the served sdk-all.js and its apply are a separate, gated operator step (maintainer; the disposable DS first; the
live DS needs its own owner GO).

NOT MEASURED: that the model paths hold on the REAL 9.3.5 engine in a live session (design gates R1-R2, on the disposable DS). The
accessor / field names were checked by SOURCE READING in the factory sdk-all.js (asc_CPageMargins.asc_getHeader/asc_getFooter, pageSetup.pageOrder,
ColorFilter.CellColor/dxf, CSpPr.Fill/ln, CDPt.idx/spPr, CTitle/CLegend/CPlotArea/CChartSpace .spPr), not by running them.

Usage:
  build-rev20.py   -> writes bridge-block.js, old.txt, new.txt, manifest-mechanism-proof.json next to this file
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V19_BLOCK = HERE.parent / "rev19-candidate" / "bridge-block.js"
V19_BLOCK_SHA = "d8d081513c3abe6dd7cae49cdb6ece9c147c558c4d3a7bab8cd463f1cf352ea2"
V19_BLOCK_SIZE = 123683
HANDLER_SRC = HERE / "handler-read-v20.txt"

PRINT_REASONS = ("bad-args, sheet-not-found, model-missing, margin-getters-missing, headerMargin-not-number, footerMargin-not-number, "
                 "pageOrder-unexpected, plus every v1 reason (rowBreaks-*, model-missing)")
FILTER_REASONS = ("bad-args, sheet-not-found, autofilter-threw, filters-unexpected-shape, field-unreadable, colorfilter-unexpected-shape, "
                  "cellcolor-unexpected, colorfilter-dxf-missing, colour-form-unknown, colour-slots-disagree")
STYLE_REASONS = ("bad-args, unsupported-target, sheet-not-found, charts-threw, chart-not-found, chart-model-missing, chart-type-not-single, "
                 "series-index-out-of-range, dPt-unexpected-shape, datapoint-ambiguous, sppr-chain-broken")

ENTRIES = (
    '{ id: "sheet.printSettings.read/v2", version: 2, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one sheet (by name): EVERYTHING sheet.printSettings.read/v1 returns, unchanged (the v1 handler computes it), '
    'plus headerMargin and footerMargin (mm, exactly as the model stores them; null is the model own unset state) and pageOrder '
    '(downThenOver | overThenDown from the stored ST_PageOrder 0 | 1; null if the model holds none; any other value refuses); '
    'a refused read is a named reason; v1 stays served; reasons: ' + PRINT_REASONS + ' (GH #354, #355, /20)" }, '
    '{ id: "sheet.autoFilter.colorFilter.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one sheet (by name): for every column of the sheet AutoFilter that holds a COLOUR filter, {field (1-based), kind: '
    'cellColor | fontColor (CellColor === false is the font-colour filter), colour, slot}; colour is read from the filter dxf '
    '(font colour first for a font-colour filter, then the fill foreground and background colour; two present colours that differ refuse) as '
    '{form: rgb, hex} | {form: theme, theme, tint} | {form: indexed, indexed} (a model colour object that carries a numeric indexed field; '
    'NOT measured in a session) or null when the dxf holds no colour; columns without a colour filter are omitted; the sheet AutoFilter '
    'only (a Table own filter is not read); never enumerates an object, one unreadable column refuses the WHOLE read; reasons: ' + FILTER_REASONS + ' (GH #344, /20)" }, '
    '{ id: "chart.series.style.read/v1", version: 1, kind: "READ", format: "XLSX", mutation: "NO", '
    'semanticContract: "one chart (sheet, chart name) and ONE of the six style targets (chart | plotArea | title | legend | series {series} | '
    'dataPoint {series, point}): the fill and the outline the model holds for it, read from its spPr (chartSpace.spPr, plotArea.spPr, title.spPr, '
    'legend.spPr, series.spPr, the series dPt with idx === point), fill and outline INDEPENDENTLY: {element: absent | present, fill, outline, '
    'outlineWidthEmu}; fill / outline are null (nothing set), {form: rgb, hex}, {form: scheme, id} (the engine scheme colour id, NOT resolved to rgb) '
    'or {form: other} (a fill that is not a solid colour: no fill, gradient, picture, pattern); a colour with modifiers or an unreadable chain refuses '
    '(sppr-chain-broken); two dPt of one idx refuse; single-plot charts only for series / dataPoint; reasons: ' + STYLE_REASONS + ' (GH #346, /20)" }, '
)

MANIFEST_END = '(GH #332-#335, /19)" }, ], };'
READS_END = '}; var EO_BRIDGE_WRITES = {'
MARKERS = ['"sheet.printSettings.read/v2": function', '"sheet.autoFilter.colorFilter.read/v1": function', '"chart.series.style.read/v1": function',
           '{ id: "sheet.printSettings.read/v2"', '{ id: "sheet.autoFilter.colorFilter.read/v1"', '{ id: "chart.series.style.read/v1"']
FORBIDDEN_IN_HANDLER = ("Object.keys", "Object.values", "Object.entries", "getOwnPropertyNames", "for (var k in", "for (var p in", " in args", "//")


def sha(b): return hashlib.sha256(b).hexdigest()


def handler_text():
    lines = [ln.strip() for ln in HANDLER_SRC.read_text(encoding="utf-8").split("\n")]
    text = " ".join(ln for ln in lines if ln) + " "
    for word in FORBIDDEN_IN_HANDLER:
        if word in text:
            sys.exit(f"REFUSE: the handler text contains {word!r} -- named fields only, no line comments; NO OUTPUT")
    return text


def build_block():
    v19 = V19_BLOCK.read_bytes()
    if sha(v19) != V19_BLOCK_SHA or len(v19) != V19_BLOCK_SIZE:
        sys.exit(f"REFUSE: /19 block sha/size {sha(v19)}/{len(v19)} != {V19_BLOCK_SHA}/{V19_BLOCK_SIZE}; NO OUTPUT")
    impl = handler_text()
    subs = [
        ('patchVersion: "euro-office-bridge-v1/9.3.5.0/19"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/20"'),
        (MANIFEST_END, '(GH #332-#335, /19)" }, ' + ENTRIES + '], };'),
        (READS_END, impl + READS_END),
    ]
    t = v19.decode("utf-8")
    for old, new in subs:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v19, t.encode("utf-8"), impl


def main():
    v19, v20, impl = build_block()
    out = {"v19_block_sha": sha(v19), "v19_block_size": len(v19), "v20_block_sha": sha(v20), "v20_block_size": len(v20)}
    (HERE / "bridge-block.js").write_bytes(v20)
    (HERE / "old.txt").write_bytes(v19)
    (HERE / "new.txt").write_bytes(v20)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS,
            "operations": [{"id": "impl-bridge-entrypoint-v20-gh344-gh346-gh354-gh355-reads", "old": v19.decode("utf-8"), "new": v20.decode("utf-8")}]}
    proof = dict(base, _NOTE="MECHANISM-PROOF MANIFEST ONLY: original/patched sha are the isolated old.txt/new.txt "
                              "blocks, not a served bundle. NO --sdk mode exists for this build (see module docstring) "
                              "-- the served-sdk-backed apply is a SEPARATE, gated operator step (maintainer; the disposable "
                              "DS first, the live DS needs its own owner GO).",
                 original_sha256=sha(v19), patched_sha256=sha(v20), original_size=len(v19), patched_size=len(v20))
    (HERE / "manifest-mechanism-proof.json").write_text(json.dumps(proof, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
