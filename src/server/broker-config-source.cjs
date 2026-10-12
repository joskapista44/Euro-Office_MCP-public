'use strict';

/**
 * M11.2 -- requester-own credential routing (owner GO, owner channel 2026-10-09; plan
 * docs/M11-CAPABILITY-SERVING-PRODUCT-MCP.md, M11.2).
 *
 * The product runtime's editor-config source. Path:
 *   authenticated agent (this process's UID and its own agent token)
 *     -> maintainer broker op `product_ocs_config` {fileId, requestedMode}
 *     -> the broker derives the requester from the token and fetches the Nextcloud OCS editor config with the
 *        requester's OWN Nextcloud credential
 *     -> the config object -> the driver (config-source seam).
 *
 * What this module does NOT do (owner decision 2026-10-09, Nextcloud/DS/storage are the access authority):
 *   - no ACL / permission matrix / share policy / SAC decision: it never judges whether the requester may access the
 *     file, and it accepts an edit-capable config for a READ operation (READ safety is the operation's semantics:
 *     no dispatch, no save);
 *   - no Nextcloud user, password, app password or vault secret: none is read, sent, received or stored here; the only
 *     secret this module touches is the agent's OWN bearer token, read at call time and never logged or returned;
 *   - no fallback: one broker request per call, no retry with another mode, credential, user or direct Basic auth.
 *
 * Every broker answer maps onto exactly one controlled outcome; anything unrecognised fails closed.
 * Broker contract mirrored (maintainer repo, src/web/product-ocs-config-broker.ts + routes/connectors.ts):
 *   200 {ok, status, config} | 400 *_BAD_BODY | 401 | 403 EURO_BROKER_DENY_PRINCIPAL / PRODUCT_OCS_CONFIG_DENY_* /
 *   AUTHZ_DENY_NO_NC_IDENTITY / AUTHZ_DENY_USE | 500/502/503.
 */

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { OPERATIONS, ConfigSourceError } = require('../drivers/eurooffice/config-source.cjs');

const FILE_ID = /^[0-9]{1,20}$/;
const LOOPBACK_URL = /^https?:\/\/(127\.0\.0\.1|localhost)(:\d+)?\//;

// tool-argument keys that would carry a requester identity: refused before the broker is asked (input hygiene -- the
// requester is ONLY the token owner). `identity` / `owner` are deliberately absent: adapters use them for cell/table identity.
const IDENTITY_KEYS = /^(requester|principal|uid|user|username|userid|user_id|nextclouduser|nc_user|ncuser|as_user|asuser|impersonate|agent|agent_id|actor|password|app_password|apppassword|credential)$/i;

// The product error classes (owner list, owner channel 2026-10-09, M11.2 point 12).
const ERRORS = Object.freeze({
  BROKER_REQUEST_INVALID: 'BROKER_REQUEST_INVALID',
  CONTROLLED_ACCESS_ERROR: 'CONTROLLED_ACCESS_ERROR',
  ACCESS_DENIED: 'ACCESS_DENIED',
  EXTERNAL_POLICY_DENIED: 'EXTERNAL_POLICY_DENIED',
  BROKER_UNAVAILABLE: 'BROKER_UNAVAILABLE',
  BROKER_RESPONSE_INVALID: 'BROKER_RESPONSE_INVALID',
});

// broker refusal code -> product error class. The broker code is kept as the (secret-free) detail.
const BROKER_CODE_MAP = Object.freeze({
  // the requester itself: not allowed on the broker, no mapped NC identity, no own credential, credential use refused
  EURO_BROKER_DENY_PRINCIPAL: ERRORS.CONTROLLED_ACCESS_ERROR,
  PRODUCT_OCS_CONFIG_DENY_PRINCIPAL: ERRORS.CONTROLLED_ACCESS_ERROR,
  AUTHZ_DENY_NO_NC_IDENTITY: ERRORS.CONTROLLED_ACCESS_ERROR,
  AUTHZ_DENY_NOT_DEVICE_PRINCIPAL: ERRORS.CONTROLLED_ACCESS_ERROR,
  AUTHZ_DENY_NO_AUTH: ERRORS.CONTROLLED_ACCESS_ERROR,
  PRODUCT_OCS_CONFIG_DENY_NO_REQUESTER_SECRET: ERRORS.CONTROLLED_ACCESS_ERROR,
  AUTHZ_DENY_USE: ERRORS.CONTROLLED_ACCESS_ERROR,
  PRODUCT_OCS_CONFIG_REQUESTER_CREDENTIAL_UNAVAILABLE: ERRORS.CONTROLLED_ACCESS_ERROR,
  // the target: the broker does not tell "missing" from "not yours" (G2) -- one inaccessible-target class, and the
  // product must not probe further to tell them apart
  PRODUCT_OCS_CONFIG_DENY_AUTHZ: ERRORS.ACCESS_DENIED,
  PRODUCT_OCS_CONFIG_DENY_LEVEL: ERRORS.ACCESS_DENIED,
  // Nextcloud granted less than the operation needs (a WRITE answered with a view config): the target is read-only
  PRODUCT_OCS_CONFIG_DENY_CONFIG_BELOW_REQUEST: ERRORS.ACCESS_DENIED,
  // G1 (external interface defect, owner an internal branch): READ + an edit-capable Nextcloud config is refused by the
  // broker's superseded policy. Controlled, and NEVER worked around (no WRITE retry, no Basic auth, no other credential).
  PRODUCT_OCS_CONFIG_DENY_CONFIG_EXCEEDS_VERDICT: ERRORS.EXTERNAL_POLICY_DENIED,
  // request shape
  PRODUCT_OCS_CONFIG_BAD_BODY: ERRORS.BROKER_REQUEST_INVALID,
  EURO_BROKER_DENY_UNKNOWN_OP: ERRORS.BROKER_REQUEST_INVALID,
  EURO_BROKER_DENY_INSTANCE: ERRORS.BROKER_REQUEST_INVALID,
  // availability
  PRODUCT_ZUUL2_UNAVAILABLE: ERRORS.BROKER_UNAVAILABLE,
  PRODUCT_OCS_CONFIG_UPSTREAM_FAILED: ERRORS.BROKER_UNAVAILABLE,
  PRODUCT_OCS_CONFIG_AUDIT_FAILED: ERRORS.BROKER_UNAVAILABLE,
});

const refusal = (code, detail) => ({ ok: false, code, detail: detail || null });

/** The fileId a config is bound to (`document.referenceData.fileKey`), or null when the config does not say. */
function boundFileId(config) {
  const ref = config && config.document && config.document.referenceData;
  return ref && (typeof ref.fileKey === 'string' || typeof ref.fileKey === 'number') ? String(ref.fileKey) : null;
}

/** Shape check only -- what the driver needs to start an editor; NOT an authorization judgement. */
function editorConfigShapeError(config, fileId) {
  if (typeof config !== 'object' || config === null || Array.isArray(config)) return 'the config is not an object';
  if (!config.document || typeof config.document.key !== 'string' || !config.document.key) return 'the config has no document.key';
  if (typeof config.documentServerUrl !== 'string' || !/^https?:\/\//.test(config.documentServerUrl)) return 'the config has no documentServerUrl';
  if (boundFileId(config) !== fileId) return 'the config is not bound to the requested fileId';
  return null;
}

/**
 * The target of one tool call -> { ok:true, fileId, operation } or a refusal. Identity-bearing keys anywhere in the
 * arguments are refused (BROKER_REQUEST_INVALID): the requester is only ever the authenticated token owner.
 */
function validateTarget(args, operation) {
  if (!OPERATIONS.includes(operation)) return refusal(ERRORS.BROKER_REQUEST_INVALID, `unknown operation: ${String(operation)}`);
  if (typeof args !== 'object' || args === null || Array.isArray(args)) return refusal(ERRORS.BROKER_REQUEST_INVALID, 'tool arguments must be an object');
  const identity = [];
  const scan = (node, at, depth) => {
    if (depth > 8 || typeof node !== 'object' || node === null) return;
    for (const [k, v] of Object.entries(node)) {
      if (IDENTITY_KEYS.test(k)) identity.push(`${at}${k}`);
      scan(v, `${at}${k}.`, depth + 1);
    }
  };
  scan(args, '', 0);
  if (identity.length) return refusal(ERRORS.BROKER_REQUEST_INVALID, `identity-bearing argument(s) refused: ${identity.join(', ')} -- the requester is the authenticated token owner`);
  const fileId = typeof args.fileId === 'number' ? String(args.fileId) : args.fileId;
  if (typeof fileId !== 'string' || !FILE_ID.test(fileId)) return refusal(ERRORS.BROKER_REQUEST_INVALID, 'fileId must be a numeric string');
  return { ok: true, fileId, operation };
}

/**
 * The broker-backed editor-config source of the product runtime.
 *   brokerUrl  the product config's broker.url (loopback, the broker's /api/euro-broker route)
 *   instance   the broker instance name (default 'production')
 *   env        EURO_OFFICE_MCP_TOKEN_PATH, else ~/.agent-token (the consumer UID's OWN agent token)
 *   fetchImpl / readToken: test seams.
 * -> { kind, requestEditorConfig(fileId, operation) -> {ok:true, config} | refusal (never throws),
 *      getEditorConfig({fileId, operation}) -> config | throws ConfigSourceError (the driver seam) }
 */
function createBrokerConfigSource({ brokerUrl, instance = 'production', env = process.env, fetchImpl = (...a) => globalThis.fetch(...a), readToken = null, timeoutMs = 30000 } = {}) {
  const tokenPath = env.EURO_OFFICE_MCP_TOKEN_PATH || path.join(os.homedir(), '.agent-token');
  const token = readToken || (() => fs.readFileSync(tokenPath, 'utf8').trim());

  async function requestEditorConfig(fileId, operation) {
    if (typeof brokerUrl !== 'string' || !LOOPBACK_URL.test(brokerUrl)) return refusal(ERRORS.BROKER_REQUEST_INVALID, 'no loopback broker configured');
    if (!OPERATIONS.includes(operation)) return refusal(ERRORS.BROKER_REQUEST_INVALID, `unknown operation: ${String(operation)}`);
    if (typeof fileId !== 'string' || !FILE_ID.test(fileId)) return refusal(ERRORS.BROKER_REQUEST_INVALID, 'fileId must be a numeric string');
    let bearer;
    try { bearer = token(); } catch { return refusal(ERRORS.CONTROLLED_ACCESS_ERROR, 'the agent identity (own token) is unavailable'); }
    if (!bearer) return refusal(ERRORS.CONTROLLED_ACCESS_ERROR, 'the agent identity (own token) is unavailable');

    let status;
    let body;
    try {
      const r = await fetchImpl(brokerUrl, {
        method: 'POST',
        headers: { Authorization: `Bearer ${bearer}`, 'Content-Type': 'application/json' },
        // the ONLY fields: no user, uid, requester or credential -- the broker derives the requester from the token
        body: JSON.stringify({ op: 'product_ocs_config', instance, fileId, requestedMode: operation }),
        signal: AbortSignal.timeout(timeoutMs),
      });
      status = r.status;
      body = await r.json().catch(() => null);
    } catch (err) {
      const timedOut = err && (err.name === 'TimeoutError' || err.name === 'AbortError');
      return refusal(ERRORS.BROKER_UNAVAILABLE, timedOut ? 'the broker did not answer in time' : 'the broker did not answer');
    }

    if (status === 200) {
      if (!body || body.ok !== true) return refusal(ERRORS.BROKER_RESPONSE_INVALID, 'malformed broker answer');
      const shape = editorConfigShapeError(body.config, fileId);
      if (shape) return refusal(ERRORS.BROKER_RESPONSE_INVALID, shape);
      return { ok: true, config: body.config };
    }
    const code = body && typeof body.code === 'string' ? body.code : null;
    if (code && BROKER_CODE_MAP[code]) return refusal(BROKER_CODE_MAP[code], `broker ${code}`);
    if (status === 401) return refusal(ERRORS.CONTROLLED_ACCESS_ERROR, 'the agent token was not accepted');
    if (status === 400) return refusal(ERRORS.BROKER_REQUEST_INVALID, `broker HTTP 400${code ? ` ${code}` : ''}`);
    if (status >= 500) return refusal(ERRORS.BROKER_UNAVAILABLE, `broker HTTP ${status}${code ? ` ${code}` : ''}`);
    return refusal(ERRORS.BROKER_RESPONSE_INVALID, `unrecognised broker answer: HTTP ${status}${code ? ` ${code}` : ''}`);
  }

  return {
    kind: 'BROKER',
    requestEditorConfig,
    async getEditorConfig({ fileId, operation }) {
      const r = await requestEditorConfig(String(fileId), operation);
      if (!r.ok) throw new ConfigSourceError(r.code, r.detail);
      return r.config;
    },
  };
}

module.exports = { ERRORS, BROKER_CODE_MAP, IDENTITY_KEYS, boundFileId, editorConfigShapeError, validateTarget, createBrokerConfigSource };
