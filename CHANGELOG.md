# Changelog

This project follows [Semantic Versioning](https://semver.org/). Versions before 1.0.0 were internal.

## 1.2.0

**Compatibility note:** no Document Server upgrade is required to keep existing capabilities working -- every
`/16`/`/17`-accepted read and write intent from 1.1.0 is unchanged. `read_tables`, `read_hyperlinks` and the
table-aware structural writes below need Document Server cell Bridge `euro-office-bridge-v1/9.3.5.0/18` or later;
on `/16`/`/17` they refuse with `CAPABILITY_MISSING`, never a guessed value.

- new tool `document_create`: see below -- exposed tools: 4 -> 6 (+ `product_status`, `document_create`)
- `xlsx_execute_task`: 62 -> 87 write intents; `docx_execute_task`: 38 -> 46 write intents
- the full list of accepted read operations and write intents: [CAPABILITIES.md](CAPABILITIES.md), checked against
  the tool schemas by a test
- XLSX, writes on a sheet that holds an Excel Table (read through `read_tables`, Document Server cell Bridge `/18`):
  `insert_rows` / `delete_rows` / `insert_columns` / `delete_columns` beside the table (the table entirely before the
  cut) now apply -- they used to be refused on any sheet with a table; one that would cut through or shift the table is
  refused before the write (`xlsx-structure-overlaps-table`). `merge_range` overlapping a table is refused
  (`xlsx-merge-overlaps-table`); `sort_range` of a range only partly over a table is refused (`xlsx-sort-overlaps-table`),
  a sort of exactly the table's range still sorts it. On a runtime without the table read, merge and sort behave as before.
- XLSX: new write intent `create_table` -- an Excel Table over a range whose first row is the header (public
  `FormatAsTable`: automatic name, the engine's default style). Refused before any write, each named: a range that
  overlaps a table or a pivot table, a merge inside the range, a single row, a blank / formula / duplicate header
  (case-insensitive). Verified through `read_tables`, which needs Document Server cell Bridge `/18`.
- DOCX `read_formatting`: a paragraph whose text also occurs elsewhere (or holds a tab) could report the properties
  of its text without its last characters -- on a paragraph holding a hyperlink, and on some multi-run paragraphs --
  so a mixed paragraph could read as uniformly bold. Its properties now come from a range verified to hold exactly
  the paragraph's text; when no such range is found the properties are reported unread, never guessed.
- DOCX: a paragraph with text around a hyperlink (one or more links) can now be formatted (`format_text`) and given a
  footnote or endnote (`add_footnote`, `add_endnote`) -- the link stays, the note reference goes at the paragraph's end.
  These used to end in `DISPATCHED_OUTCOME_UNKNOWN` without a write. `rewrite_paragraph` on a paragraph holding a
  hyperlink is now refused before any write (`FAILED`, `rewrite-target-holds-hyperlink`) instead of the same unknown
  outcome: a rewrite would delete the link.
- new tool `document_create`: a new, empty `docx`, `xlsx` or `pptx`, produced by DocBuilder from a structure-only seed
  and written to a local path or to a path in the requester's Nextcloud files. An existing file is never replaced; a
  Nextcloud target answers the new `fileId`, which the co-editing tools (`docx_execute_task`, `xlsx_execute_task`) take
  for every further edit.
- `xlsx_read` gains `read_tables` (every Excel Table of one sheet: name, range, header/totals rows, column names,
  style flags, autoFilter presence) and `read_hyperlinks` (every hyperlink of one sheet: target address, internal
  location, screen tip, kind) -- both require Euro-Office Document Server Bridge `euro-office-bridge-v1/9.3.5.0/18`;
  on the 1.1.0-accepted `/16` or `/17` Bridge both operations refuse with `CAPABILITY_MISSING`, never a guessed value.
- The XLSX Bridge-backed writes (alignment indent / shrink, print settings, column-width canonicalization, chart
  gap/overlap and series order, chart delete, series marker / line style) are accepted on the `/18` Bridge as well:
  every pre-existing `/17` Bridge handler is byte-identical on `/18`, so switching the Document Server from `/17` to
  `/18` keeps them working instead of refusing them.

## 1.1.0

**Upgrade note:** this release requires Euro-Office Document Server **9.3.5.0** with the cell (XLSX) Bridge
`euro-office-bridge-v1/9.3.5.0/17` and the word (DOCX) Bridge `euro-office-bridge-word-v1/9.3.5.0/5` (fresh install:
`documentserver-patches/9.3.5.0/fresh-install-v1.1.0.json`, see INSTALL.md). Document Server 9.3.4.60 and its Bridge revisions
are no longer accepted; against the cell `/16` Bridge every XLSX write is refused.

Upgrading from 1.0.x: 1.0.x accepts only Document Server 9.3.4.60 and 1.1.0 only 9.3.5.0, so there is no block set
both releases accept -- the compatibility-first order of UPDATE.md does not apply to this one step. Upgrade the
Document Server to 9.3.5.0 with the blocks above and switch the product to 1.1.0 in the same maintenance window;
between the two steps the editor operations are refused (fail-closed), no document is changed.

- new tool `docx_execute_task`: DOCX write tasks through the same executor, confirmed save and read-back verification
  as XLSX -- 38 write intents (paragraphs, lists, tables, text and table formatting, page setup, headers and footers,
  styles, numbering, footnotes and endnotes, links, bookmarks, comments, images)
- `xlsx_execute_task`: 6 -> 62 write intents (structure, sheets, ranges, formats, conditional formats, validation,
  filters, charts, shapes and images, defined names, print and page layout, pivots, array formulas, comments)
- `xlsx_read`: 3 -> 16 read operations; `docx_read`: 5 -> 12 read operations
- the full list of accepted read operations and write intents: [CAPABILITIES.md](CAPABILITIES.md), checked against the
  tool schemas by a test
- protected sheets: a write into a protected sheet is refused before dispatch, and a hidden formula is not returned
  by a read
- a drawing whose size reads 0 x 0 after a reopen is reported `UNOBSERVABLE`, not as a size
- the public tree carries the test helpers; a test whose subject is a private file (the live harnesses, internal
  documents) is skipped there with a named reason

## 1.0.1

- CI runs on Node.js 22 (the runtime the release is accepted on). On Node.js 20 the test script's glob
  (`node --test "test/**/*.test.cjs"`) is not expanded, so the 1.0.0 CI test job failed before running any test;
  the suite itself passes on Node.js 20 when the files are listed explicitly. No runtime or product change.

## 1.0.0

First public release.

- stdio MCP server with `product_status`, `xlsx_read`, `docx_read` and `xlsx_execute_task`
- requester-own access: every editor operation runs with the calling user's own Nextcloud credential through a
  loopback broker; no service account, no permission layer of its own
- reference broker (`reference-broker/`) implementing the broker contract
- Document Server 9.3.4.60 blocks: Bridge for XLSX (`euro-office-bridge-v1/9.3.4.60/8`) and DOCX
  (`euro-office-bridge-word-v1/9.3.4.60/3`), and a chart read layer (`euro-office-chart-read-v1/9.3.4.60/1`); fresh
  install from the unpatched bundles with pinned input and output hashes
- reproducible artifact build (two builds of one commit are byte-identical)
- fail-closed startup and operation: invalid config, missing broker, unknown Document Server runtime or an unserved
  capability are refused with a named reason
