# M12 — reference broker contract (D4 = B)

Owner decision D4 = B (owner channel, 2026-10-12): a **minimal, standalone reference broker** ships with the public
release. The private production broker is **not** published and **not** copied. The reference implementation must fulfil
the accepted requester-own credential + authorization contract and must **not** create an alternative authorization path.

Clean-room rule: the implementation is written from THIS document and the product's own client
(`src/server/broker-config-source.cjs`, `src/drivers/eurooffice/config-source.cjs`) only. The private broker's source is
not read, quoted or ported.

## 1. Role

```text
product MCP (requester's own agent token)
  -> POST <broker.url>  { op, instance, fileId, requestedMode }   Authorization: Bearer <requester token>
  -> reference broker: requester = owner of the token (never a body field)
  -> Nextcloud OCS editor config, fetched with THAT requester's OWN Nextcloud app password
  -> 200 { ok: true, status: 200, config }   or a refusal code the product already maps
```

Nextcloud is the only access authority. The broker holds no ACL, no permission matrix, no admin override, no shared
service account, no impersonation and no fallback credential.

## 2. Configuration (no secret inside)

```json
{
  "listen": { "host": "127.0.0.1", "port": 3421 },
  "path": "/api/euro-broker",
  "instances": { "production": { "nextcloudUrl": "https://nextcloud.example.org" } },
  "requesters": [
    { "id": "alice", "tokenSha256": "<sha256 hex of alice's bearer token>", "nextcloudUser": "alice",
      "appPasswordFile": "/etc/euro-office-broker/alice.app-password" }
  ],
  "timeoutMs": 30000
}
```

- `listen.host` must be a loopback address (`127.0.0.1`, `::1`, `localhost`); anything else refuses to start.
- The config holds token **hashes** only; each app password lives in its own file (the operator sets 0600, owned by the broker user).
- An unknown key anywhere in the config refuses to start (no silent defaults for security fields).

## 3. Request

`POST <path>`, `Content-Type: application/json`, body ≤ 4 KiB, exactly these four keys:

| key | rule | refusal |
|---|---|---|
| `op` | `"product_ocs_config"` | 400 `EURO_BROKER_DENY_UNKNOWN_OP` |
| `instance` | a key of `instances` | 400 `EURO_BROKER_DENY_INSTANCE` |
| `fileId` | string `^[0-9]{1,20}$` | 400 `PRODUCT_OCS_CONFIG_BAD_BODY` |
| `requestedMode` | `"READ"` or `"WRITE"` | 400 `PRODUCT_OCS_CONFIG_BAD_BODY` |
| any other key (`user`, `uid`, `requester`, `credential`, …) | refused | 400 `PRODUCT_OCS_CONFIG_BAD_BODY` |

Non-JSON, non-object, oversize body: 400 `PRODUCT_OCS_CONFIG_BAD_BODY`. Other method or path: 404.

## 4. Requester and credential

| situation | answer |
|---|---|
| no / malformed `Authorization: Bearer` | 401 `AUTHZ_DENY_NO_AUTH` |
| `sha256(token)` matches no requester (constant-time compare) | 401 `AUTHZ_DENY_NO_AUTH` |
| requester's app-password file missing / unreadable / empty | 403 `PRODUCT_OCS_CONFIG_DENY_NO_REQUESTER_SECRET` |
| Nextcloud rejects the requester's credential (401) | 403 `PRODUCT_OCS_CONFIG_REQUESTER_CREDENTIAL_UNAVAILABLE` |

The credential used is ALWAYS the one mapped to the token owner. There is no code path that selects another requester's
credential.

## 5. Nextcloud call and verdict

`GET <nextcloudUrl>/ocs/v2.php/apps/eurooffice/api/v1/config/<fileId>?format=json`, Basic `nextcloudUser:appPassword`,
headers `OCS-APIRequest: true`, `Accept: application/json`; the config is `ocs.data` (or the body when there is no
`ocs` envelope). One request, no retry.

| Nextcloud answer | broker answer |
|---|---|
| 403 / 404, or an OCS `meta.statuscode` 403 / 404 / 997 / 998 | 403 `PRODUCT_OCS_CONFIG_DENY_AUTHZ` (missing == not allowed; no probing) |
| network error, timeout, 5xx, non-JSON | 502 `PRODUCT_OCS_CONFIG_UPSTREAM_FAILED` |
| config, `requestedMode = WRITE`, config not edit-capable | 403 `PRODUCT_OCS_CONFIG_DENY_CONFIG_BELOW_REQUEST` |
| config, `requestedMode = READ` | 200, passed through even when edit-capable (no VIEW narrowing) |
| config, `requestedMode = WRITE`, edit-capable | 200 |

Edit-capable = `config.document.permissions.edit === true` and `config.editorConfig.mode !== "view"`.
The config is returned **unmodified**. The broker does not check `document.referenceData.fileKey`; the product does.

## 6. Hygiene

- Never logs, returns or echoes a token, an app password, a Basic header or a config body. A log line carries only:
  requester id, fileId, requestedMode, HTTP status, refusal code.
- Every refusal body is `{ "ok": false, "code": "<CODE>" }` with a code from sections 3-5 (all of them are in the
  product's `BROKER_CODE_MAP`, so the product maps every answer to a controlled class).
- Zero runtime dependencies (Node ≥ 20 built-ins only).

## 7. Acceptance (offline, a fake Nextcloud on loopback)

The product's real client (`createBrokerConfigSource`) against the reference broker against a fake Nextcloud:

1. own READ → config (edit-capable passes through); own WRITE edit-capable → config; WRITE view-only → `ACCESS_DENIED`.
2. Nextcloud 404 and 403 → `ACCESS_DENIED`, the same answer for both.
3. two requesters: alice's token only ever reaches Nextcloud with alice's credential (the fake records the Basic user).
4. identity-bearing body keys (`user`, `uid`, `requester`) → 400; the product's own `validateTarget` refusal still holds.
5. unknown / missing token → `CONTROLLED_ACCESS_ERROR`; missing app-password file → `CONTROLLED_ACCESS_ERROR`.
6. Nextcloud down / 500 / timeout → `BROKER_UNAVAILABLE`.
7. non-loopback `listen.host` → refuses to start; unknown config key → refuses to start.
8. secret scan: no token, app password or Basic header in any response body or log line.
9. sabotage controls (each must turn a test RED, then restore): accept a `user` body key; select the credential by a
   body field; narrow READ to view; retry with another credential.

Out of scope: TLS termination, account provisioning, multi-tenant administration, production deployment.
