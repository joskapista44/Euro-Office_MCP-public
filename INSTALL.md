# Install

Paths below are the defaults (`/opt/euro-office-mcp`, `/etc/euro-office-mcp/config.json`); any other path works if
it is named consistently.

## 1. Document Server blocks

The server only works against an Euro-Office Document Server **9.3.4.60** carrying this project's blocks. They are
inserted into the editor bundles by `documentserver-patches/9.3.4.60/install_bridge.py`, driven by
`fresh-install.json`, which pins the exact UNPATCHED input bundle hash, each block hash and the exact output hash.
Any other input (a different build, an already modified bundle) is refused and nothing is written.

Locate the bundles the editor actually loads (historically
`<documentserver>/sdkjs/cell/sdk-all.js` and `<documentserver>/sdkjs/word/sdk-all.js`), keep a copy of each, then:

```sh
cd documentserver-patches/9.3.4.60
python3 install_bridge.py --target cell --input <cell sdk-all.js> --check
python3 install_bridge.py --target cell --input <cell sdk-all.js> --apply --output /tmp/cell-sdk-all.js
python3 install_bridge.py --target word --input <word sdk-all.js> --check
python3 install_bridge.py --target word --input <word sdk-all.js> --apply --output /tmp/word-sdk-all.js
```

`--check` prints `STATUS=PATCHABLE` (or `ALREADY_PATCHED`, exit 0, nothing to do); `--apply` verifies the output hash
before writing. Move each verified output over the served bundle atomically, keeping owner, group and mode.

The server checks the Bridge identity (`documentServerVersion` + `patchVersion`) on every Bridge call, not at
startup: against a Document Server without the expected blocks the affected fields read `UNOBSERVABLE:<reason>` and
Bridge writes are refused -- there is no fallback. `product_status` does not probe the Document Server (it reports
`ds.status = FAIL_CLOSED`, `reason = LIVE_PROBE_NOT_WIRED`); verify the blocks with a read instead (§7).

## 2. Build the artifact

From a committed revision of this repository (a dirty tree is refused):

```sh
node scripts/build-artifact.cjs --rc <commit> --out ./dist
sha256sum -c ./dist/euro-office-mcp-<ARTIFACT_VERSION>.tgz.sha256
```

Two builds of the same commit are byte-identical. The artifact contains the runtime (`bin/`, `src/`, `config/`, the
package files) and its production dependencies (`playwright-core` 1.62.1), installed from the lockfile.

## 3. Install the artifact

```sh
V=<ARTIFACT_VERSION>
sudo mkdir -p /opt/euro-office-mcp/releases
sudo tar -xzf ./dist/euro-office-mcp-$V.tgz -C /opt/euro-office-mcp/releases
sudo mv /opt/euro-office-mcp/releases/euro-office-mcp-$V /opt/euro-office-mcp/releases/$V
sudo cp ./dist/euro-office-mcp-$V.manifest.json /opt/euro-office-mcp/releases/$V/artifact.json
sudo ln -sfn releases/$V /opt/euro-office-mcp/current
```

`artifact.json` carries the tarball's own hash (a tarball cannot contain it); without it the identity reports
`ARTIFACT_SHA256=NOT_RECORDED`.

## 4. Browser

The pinned headless browser (`chromium-headless-shell` revision 1234) goes into the directory the config names as
`playwright.path`:

```sh
sudo env PLAYWRIGHT_BROWSERS_PATH=/opt/euro-office-mcp/playwright \
  node /opt/euro-office-mcp/current/node_modules/playwright-core/cli.js install --only-shell chromium
```

The server never downloads a browser itself; a missing or different revision is refused.

## 5. Broker

The server needs a broker on loopback that turns the requester's bearer token into that user's own Nextcloud
access. Use the [reference broker](reference-broker/README.md) or an implementation of
[the contract](docs/M12-REFERENCE-BROKER-CONTRACT.md). Each requesting user needs:

- a bearer token, stored in a file readable only by the user that runs the MCP server
  (`EURO_OFFICE_MCP_TOKEN_PATH`, default `~/.agent-token`);
- a Nextcloud app password of their own, known only to the broker.

## 6. Configure

Copy `config/euro-office-mcp.example.json` to `/etc/euro-office-mcp/config.json` and set the Nextcloud URL, the
Document Server URL, the broker route and `playwright.path`. The config never holds a secret: secret-shaped keys are
refused at startup. Validate it:

```sh
/opt/euro-office-mcp/current/bin/euro-office-mcp.cjs --check-config      # CONFIG_OK <identity>, else exit 78
/opt/euro-office-mcp/current/bin/euro-office-mcp.cjs --identity          # product / artifact / runtime identity
```

The config path is `--config <path>`, else `$EURO_OFFICE_MCP_CONFIG`, else `/etc/euro-office-mcp/config.json`.

## 7. Connect an MCP client

The server speaks MCP over stdio: no daemon, no port. Register it as a stdio server, for example:

```json
{
  "mcpServers": {
    "euro-office": {
      "command": "node",
      "args": ["/opt/euro-office-mcp/current/bin/euro-office-mcp.cjs"],
      "env": { "EURO_OFFICE_MCP_TOKEN_PATH": "/home/<user>/.euro-office-mcp-token" }
    }
  }
}
```

Smoke test:

- `tools/list` returns exactly `product_status`, `xlsx_read`, `docx_read`, `xlsx_execute_task`;
- `product_status` reports the released identity (`artifact.version`, `artifact.sha256`, `released: true`);
- `xlsx_read` with `operation: "read_range"` and `format: true` on a small range of a workbook you can open returns
  `ok: true` and no `UNOBSERVABLE:` value -- that exercises the broker, the browser and the cell (XLSX) blocks. For
  DOCX, check a `docx_read` result the same way.
