# Euro-Office_MCP

An MCP server (stdio) that reads and edits XLSX and DOCX files stored in Nextcloud through a real Euro-Office
Document Server editor session -- as the requesting user, with that user's own Nextcloud permissions.

- **One execution framework, format-specific adapters.** Editor lifecycle, execution safety, persistence, retry,
  verification and metrics are shared; XLSX and DOCX semantics live in their own adapters.
- **Nextcloud is the only access authority.** The server never holds a Nextcloud password. A loopback broker
  fetches the editor configuration with the requester's own Nextcloud app password; there is no service account
  and no permission layer of its own. A reference broker ships in [`reference-broker/`](reference-broker/README.md).
- **Fail-closed.** An invalid config, a missing broker or browser, an unknown Document Server runtime or an
  unproven capability refuses the operation -- or, for one field of a read, answers `UNOBSERVABLE:<reason>` -- instead
  of guessing.

## Tools

| tool | what it does |
|---|---|
| `product_status` | read-only status: identity, artifact, config schema, lifecycle gate, dependencies; returns no secrets (the Document Server itself is checked on each editor call, not here) |
| `xlsx_read` | read-only XLSX read of a Nextcloud file as the calling requester; accepted read operations only; never writes or saves |
| `docx_read` | the same for DOCX |
| `xlsx_execute_task` | an XLSX write task (`{ id?, steps }`) run by the executor and persisted with a confirmed save; allowlisted intents only |
| `docx_execute_task` | the same for DOCX: a DOCX write task run by the executor and persisted with a confirmed save; allowlisted intents only |

The exact input schema of each tool, including the accepted operations and intents, is returned by `tools/list`.
The accepted read operations and write intents of this release are listed in [CAPABILITIES.md](CAPABILITIES.md).

## Requirements

- Node.js 22 (the release is built, tested and accepted on 22; `package.json` `engines` keeps the floor at 20)
- Nextcloud with the Euro-Office integration app, reachable over HTTPS
- Euro-Office Document Server **9.3.5.0** with this project's Document Server blocks installed
  ([`documentserver-patches/`](documentserver-patches/)): the cell (XLSX) Bridge `euro-office-bridge-v1/9.3.5.0/17` and the
  word (DOCX) Bridge `euro-office-bridge-word-v1/9.3.5.0/5`; against the cell `/16` Bridge every
  XLSX write is refused (the protected-sheet guard reads through `/17`), and any other Document Server build or
  Bridge revision is refused until revalidated
- upgrading from 1.0.x is a Document Server version change: 1.0.x accepts only Document Server 9.3.4.60 and 1.1.0
  only 9.3.5.0, so the Document Server upgrade and the switch to 1.1.0 need one maintenance window (see
  [CHANGELOG.md](CHANGELOG.md) and [UPDATE.md](UPDATE.md))
- `playwright-core` 1.62.1 (bundled) and the `chromium-headless-shell` revision 1234 browser, at the path named in
  the config
- a broker on loopback, e.g. the reference broker, and one app password per requesting Nextcloud user

## Documentation

- [INSTALL.md](INSTALL.md) -- build, install, configure, connect an MCP client
- [UPDATE.md](UPDATE.md) and [ROLLBACK.md](ROLLBACK.md)
- [SECURITY.md](SECURITY.md) -- trust model and how to report a vulnerability
- [CHANGELOG.md](CHANGELOG.md)
- [CAPABILITIES.md](CAPABILITIES.md) -- the accepted read operations and write intents
- [`docs/M12-REFERENCE-BROKER-CONTRACT.md`](docs/M12-REFERENCE-BROKER-CONTRACT.md) -- the broker contract

## Licence

Apache License 2.0 -- see [LICENSE](LICENSE) and [NOTICE](NOTICE).
