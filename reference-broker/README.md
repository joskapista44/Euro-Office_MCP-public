# Reference broker (public release)

A minimal, standalone implementation of the `product_ocs_config` broker contract described in
[`M12-REFERENCE-BROKER-CONTRACT.md`](../docs/M12-REFERENCE-BROKER-CONTRACT.md) (owner decision D4 = B).
It ships with the public release. **The private production broker is not published and not copied**;
this implementation is written from the contract document and the product's own client
(`src/server/broker-config-source.cjs`, `src/drivers/eurooffice/config-source.cjs`) only.

Zero runtime dependencies: Node >= 20 built-ins only.

## What it is for

The product MCP never holds a Nextcloud password. It sends its own agent bearer token to this
broker; the broker maps that token to **the token owner's own Nextcloud app password** (never a
shared service account, never a credential named in the request body) and fetches the editor config
from Nextcloud on the requester's behalf. Nextcloud is the only access authority -- this broker adds
no ACL, no permission matrix and no admin override.

## Configuration

Copy [`config.example.json`](config.example.json) and fill it in. Every top-level and nested key is
checked against an exact allowlist; an unknown key anywhere refuses to start (no silent defaults for
a security-relevant field). `listen.host` must be a loopback address (`127.0.0.1`, `::1` or
`localhost`); anything else also refuses to start.

The config holds only a **sha256 hash** of each requester's bearer token --
`sha256('<the token>') | hex`. Each requester's Nextcloud app password lives in its own file, named
by `appPasswordFile`, which the operator must create with mode `0600`, owned by the broker's own
process user. This broker never reads the app password from the config file itself.

## Running it

```sh
node reference-broker/server.cjs --config ./my-broker-config.json
```

An invalid or unreadable config is never silent: a one-line reason goes to stderr and the process
exits `78` (`EX_CONFIG`) without opening a listener. A listening instance also prints its bound
address to stderr.

Embedding it in another process instead of running the CLI:

```js
const { loadConfig, createServer } = require('./reference-broker/server.cjs');
const config = loadConfig(JSON.parse(require('node:fs').readFileSync('./my-broker-config.json', 'utf8')));
createServer({ config }).listen(config.listen.port, config.listen.host);
```

## What it is not

Not TLS termination (put it behind a loopback-only reverse proxy, or let the host's own network
policy keep it off any non-loopback interface), not account provisioning, not multi-tenant
administration, not a production deployment recipe. See the contract document's own "Out of scope"
line.

## Contract summary

See the contract document for the authoritative rules; in short:

- Exactly four request body keys (`op`, `instance`, `fileId`, `requestedMode`); anything else is
  refused before any Nextcloud call.
- The requester is derived **only** from the bearer token (`sha256` match against the configured
  requesters, constant-time compare); no request field ever selects a different credential.
- One Nextcloud request per call, no retry.
- `READ` passes an edit-capable config through unmodified (no artificial "view" narrowing); `WRITE`
  against a non-edit-capable config is refused.
- Every refusal body is `{ "ok": false, "code": "<CODE>" }`; every log line carries only requester
  id, fileId, requestedMode, HTTP status and the refusal code -- never a token, an app password, a
  `Basic` header or a config body.
