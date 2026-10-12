# Security

## Reporting a vulnerability

Report it privately through this repository's GitHub security advisory form ("Security" -> "Report a
vulnerability"). Please do not open a public issue for a vulnerability.

## Trust model

- **Nextcloud is the only access authority.** Every editor operation is performed with the requesting user's own
  Nextcloud credential, fetched by the broker. The server has no service account, no permission matrix and no admin
  override; a user can reach exactly what Nextcloud lets that user reach.
- **No secret in the server's config.** The config only routes to the broker; secret-shaped keys are refused at
  startup. The requester's bearer token is read from a file (`EURO_OFFICE_MCP_TOKEN_PATH`) that must be readable only
  by the user running the server.
- **The broker listens on loopback only** and keeps one app password per requester in its own file, mode `0600`. It
  stores only a sha256 of each bearer token.
- **Document Server blocks are installed only on the exact pinned bundle** and verified by output hash; any other
  build is refused.
- **Fail-closed.** Anything the server cannot prove (config, runtime identity, capability, persistence) refuses the
  operation, or for one field of a read answers `UNOBSERVABLE:<reason>`, instead of guessing.

## Out of scope

TLS termination, account provisioning and multi-tenant administration are the operator's responsibility.
