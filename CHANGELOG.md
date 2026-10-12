# Changelog

This project follows [Semantic Versioning](https://semver.org/). Versions before 1.0.0 were internal.

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
