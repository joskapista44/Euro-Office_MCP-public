# Update

An update installs a new release directory next to the old one and switches one symlink. The old release stays in
place for [rollback](ROLLBACK.md).

1. Read the [CHANGELOG](CHANGELOG.md) entry: it names any Document Server block, config or broker change the new
   release requires.
2. **Document Server first, if the release changes a block.** The order is always compatibility-first: the served
   Document Server blocks must be the ones the new release accepts BEFORE the new release serves, and the running
   release must accept them too. A release needs no outage window unless its CHANGELOG entry names one (1.0.x ->
   1.1.0 does: it changes the Document Server version).
3. Build and verify the new artifact ([INSTALL.md](INSTALL.md) §2) and unpack it into
   `/opt/euro-office-mcp/releases/<new ARTIFACT_VERSION>` with its `artifact.json` (§3) -- without touching `current`.
4. Check it in place: `releases/<new>/bin/euro-office-mcp.cjs --check-config` and `--identity`.
5. Switch: `ln -sfn releases/<new> /opt/euro-office-mcp/current`. MCP clients pick the new release up on their next
   server start; a running server keeps its release until it exits.
6. Smoke: `tools/list` and `product_status` (identity = the new `ARTIFACT_VERSION` / `ARTIFACT_SHA256`), then one
   read per format.

Config files are versioned by `schemaVersion`; a release that needs a new schema says so in the CHANGELOG, and an
old-schema config is refused with a named reason, never reinterpreted.
