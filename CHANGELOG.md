# Changelog

This project follows [Semantic Versioning](https://semver.org/). Versions before 1.0.0 were internal.

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
