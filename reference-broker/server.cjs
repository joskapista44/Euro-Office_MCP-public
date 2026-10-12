'use strict';

/**
 * M12 reference broker (owner decision D4 = B, owner channel 2026-10-12).
 *
 * A minimal, standalone implementation of the `product_ocs_config` broker contract for the public
 * release. The private production broker is NOT published and NOT copied; this file is written from
 * the spec (M12-REFERENCE-BROKER-CONTRACT.md) and the product's own client
 * (`src/server/broker-config-source.cjs`, `src/drivers/eurooffice/config-source.cjs`) only.
 *
 * Zero runtime dependencies: Node >= 20 built-ins only (`node:http`, `node:crypto`, `node:fs`).
 *
 * Nextcloud is the only access authority. This broker holds no ACL, no permission matrix, no admin
 * override, no shared service account, no impersonation and no fallback credential: it derives the
 * requester from the caller's bearer token ONLY, and fetches the Nextcloud OCS editor config with
 * THAT requester's own Nextcloud app password. One upstream request per call, no retry.
 */

const http = require('node:http');
const crypto = require('node:crypto');
const fs = require('node:fs');

const FILE_ID = /^[0-9]{1,20}$/;
const LOOPBACK_HOST = new Set(['127.0.0.1', '::1', 'localhost']);
const REQUEST_KEYS = new Set(['op', 'instance', 'fileId', 'requestedMode']);
const DEFAULT_MAX_BODY_BYTES = 4096;

function sha256hex(s) {
  return crypto.createHash('sha256').update(s, 'utf8').digest('hex');
}

/** Constant-time comparison of two hex digests (equal-length requirement enforced first, outside crypto). */
function hexDigestsEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  const ba = Buffer.from(a.toLowerCase(), 'hex');
  const bb = Buffer.from(b.toLowerCase(), 'hex');
  if (ba.length === 0 || ba.length !== bb.length) return false;
  return crypto.timingSafeEqual(ba, bb);
}

/** A strict allowlist check: every key of `obj` must be in `keys`, else throws naming the offending path. */
function requireExactKeys(obj, keys, where) {
  for (const k of Object.keys(obj)) if (!keys.includes(k)) throw new Error(`unknown config key: ${where}${k}`);
}

/**
 * Strict-shape config loader. ANY unknown key, at any level, refuses (throws) -- no silent defaults
 * for security fields (spec §2).
 */
function loadConfig(raw) {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) throw new Error('config must be a JSON object');
  requireExactKeys(raw, ['listen', 'path', 'instances', 'requesters', 'timeoutMs'], '');

  const { listen, path: routePath, instances, requesters, timeoutMs = 30000 } = raw;

  if (typeof routePath !== 'string' || !routePath.startsWith('/')) throw new Error('path must be a string starting with "/"');
  if (typeof timeoutMs !== 'number' || !(timeoutMs > 0)) throw new Error('timeoutMs must be a positive number');

  if (typeof listen !== 'object' || listen === null || Array.isArray(listen)) throw new Error('listen must be an object');
  requireExactKeys(listen, ['host', 'port'], 'listen.');
  if (typeof listen.host !== 'string' || !LOOPBACK_HOST.has(listen.host)) throw new Error('listen.host must be a loopback address (127.0.0.1, ::1 or localhost)');
  if (typeof listen.port !== 'number' || !Number.isInteger(listen.port) || listen.port < 0) throw new Error('listen.port must be a non-negative integer');

  if (typeof instances !== 'object' || instances === null || Array.isArray(instances)) throw new Error('instances must be an object');
  if (Object.keys(instances).length === 0) throw new Error('instances must name at least one instance');
  for (const [name, inst] of Object.entries(instances)) {
    if (typeof inst !== 'object' || inst === null || Array.isArray(inst)) throw new Error(`instances.${name} must be an object`);
    requireExactKeys(inst, ['nextcloudUrl'], `instances.${name}.`);
    if (typeof inst.nextcloudUrl !== 'string' || !/^https?:\/\//.test(inst.nextcloudUrl)) throw new Error(`instances.${name}.nextcloudUrl must be an http(s) URL`);
  }

  if (!Array.isArray(requesters)) throw new Error('requesters must be an array');
  const byTokenHash = new Map();
  for (const r of requesters) {
    if (typeof r !== 'object' || r === null || Array.isArray(r)) throw new Error('each requesters[] entry must be an object');
    requireExactKeys(r, ['id', 'tokenSha256', 'nextcloudUser', 'appPasswordFile'], 'requesters[].');
    if (typeof r.id !== 'string' || !r.id) throw new Error('requesters[].id must be a non-empty string');
    if (typeof r.tokenSha256 !== 'string' || !/^[0-9a-fA-F]{64}$/.test(r.tokenSha256)) throw new Error(`requesters[${r.id}].tokenSha256 must be a 64-hex-char sha256`);
    if (typeof r.nextcloudUser !== 'string' || !r.nextcloudUser) throw new Error(`requesters[${r.id}].nextcloudUser must be a non-empty string`);
    if (typeof r.appPasswordFile !== 'string' || !r.appPasswordFile) throw new Error(`requesters[${r.id}].appPasswordFile must be a non-empty string`);
    const key = r.tokenSha256.toLowerCase();
    if (byTokenHash.has(key)) throw new Error(`duplicate tokenSha256 for requesters ${byTokenHash.get(key).id} and ${r.id}`);
    byTokenHash.set(key, r);
  }

  return { listen, path: routePath, instances, requesters, timeoutMs };
}

function send(res, status, body) {
  const text = JSON.stringify(body);
  // the 200 answer carries the editor config, which in turn carries a Document Server token/JWT: never cached.
  res.writeHead(status, { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(text), 'Cache-Control': 'no-store' });
  res.end(text);
}

/** Reads the request body, rejecting once it exceeds `maxBytes` (destroys the socket; no partial processing). */
function readBody(req, maxBytes) {
  return new Promise((resolve, reject) => {
    let total = 0;
    const chunks = [];
    let oversize = false;
    req.on('data', (chunk) => {
      total += chunk.length;
      if (total > maxBytes) {
        // Do NOT destroy the socket here: the caller still needs to write a proper 400 response on
        // this same connection. Stop accumulating and let the remaining bytes drain unread instead.
        if (!oversize) { oversize = true; reject(new Error('body too large')); }
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => { if (!oversize) resolve(Buffer.concat(chunks)); });
    req.on('error', (err) => { if (!oversize) reject(err); });
  });
}

/** `config.document.permissions.edit === true` AND `config.editorConfig.mode !== "view"` (spec §5). */
function isEditCapable(config) {
  const editFlag = !!(config && config.document && config.document.permissions && config.document.permissions.edit === true);
  const notView = !(config && config.editorConfig && config.editorConfig.mode === 'view');
  return editFlag && notView;
}

/** The OCS `ocs.data` envelope, or the raw body when there is no `ocs` envelope (spec §5). */
function unwrapOcs(body) {
  return body && typeof body === 'object' && body.ocs && typeof body.ocs === 'object' ? body.ocs.data : body;
}

/** The OCS `ocs.meta.statuscode`, or null when absent. */
function ocsMetaStatuscode(body) {
  const meta = body && typeof body === 'object' && body.ocs && typeof body.ocs === 'object' ? body.ocs.meta : null;
  return meta && typeof meta.statuscode === 'number' ? meta.statuscode : null;
}

/**
 * createServer({ config, fetchImpl, readAppPassword, log })
 *   config          the object returned by loadConfig()
 *   fetchImpl       test seam for the upstream Nextcloud call; default: global fetch
 *   readAppPassword test seam for reading an app-password file; default: fs.readFileSync(...).trim()
 *   log             test seam for the hygiene-constrained log line; default: console.log(JSON.stringify(...))
 * -> a plain node:http server (not yet listening).
 */
function createServer({ config, fetchImpl = (...a) => globalThis.fetch(...a), readAppPassword = (file) => fs.readFileSync(file, 'utf8').trim(), log = (fields) => { try { console.log(JSON.stringify(fields)); } catch { /* never throw from logging */ } } }) {
  const requestersByHash = new Map(config.requesters.map((r) => [r.tokenSha256.toLowerCase(), r]));

  async function handle(req, res) {
    const fields = { requester: null, fileId: null, requestedMode: null, status: null, code: null };
    const finish = (status, body) => { fields.status = status; fields.code = body && body.code ? body.code : null; log(fields); send(res, status, body); };

    if (req.method !== 'POST' || req.url !== config.path) { res.writeHead(404, { 'Cache-Control': 'no-store' }).end(); return; }

    let bodyBuf;
    try { bodyBuf = await readBody(req, DEFAULT_MAX_BODY_BYTES); }
    catch { return finish(400, { ok: false, code: 'PRODUCT_OCS_CONFIG_BAD_BODY' }); }

    let body;
    try { body = JSON.parse(bodyBuf.toString('utf8')); } catch { return finish(400, { ok: false, code: 'PRODUCT_OCS_CONFIG_BAD_BODY' }); }
    if (typeof body !== 'object' || body === null || Array.isArray(body)) return finish(400, { ok: false, code: 'PRODUCT_OCS_CONFIG_BAD_BODY' });
    for (const k of Object.keys(body)) if (!REQUEST_KEYS.has(k)) return finish(400, { ok: false, code: 'PRODUCT_OCS_CONFIG_BAD_BODY' });

    if (body.op !== 'product_ocs_config') return finish(400, { ok: false, code: 'EURO_BROKER_DENY_UNKNOWN_OP' });
    if (typeof body.instance !== 'string' || !Object.prototype.hasOwnProperty.call(config.instances, body.instance)) return finish(400, { ok: false, code: 'EURO_BROKER_DENY_INSTANCE' });
    if (typeof body.fileId !== 'string' || !FILE_ID.test(body.fileId)) return finish(400, { ok: false, code: 'PRODUCT_OCS_CONFIG_BAD_BODY' });
    if (body.requestedMode !== 'READ' && body.requestedMode !== 'WRITE') return finish(400, { ok: false, code: 'PRODUCT_OCS_CONFIG_BAD_BODY' });
    fields.fileId = body.fileId;
    fields.requestedMode = body.requestedMode;

    const authz = req.headers.authorization;
    const m = typeof authz === 'string' ? /^Bearer (.+)$/.exec(authz) : null;
    if (!m) return finish(401, { ok: false, code: 'AUTHZ_DENY_NO_AUTH' });
    const tokenHash = sha256hex(m[1]);
    let requester = null;
    // every candidate is compared (no early exit on hash length / shape) to keep the lookup itself constant-time-ish
    for (const [hash, r] of requestersByHash) if (hexDigestsEqual(tokenHash, hash)) requester = r;
    if (!requester) return finish(401, { ok: false, code: 'AUTHZ_DENY_NO_AUTH' });
    fields.requester = requester.id;

    let appPassword;
    try { appPassword = readAppPassword(requester.appPasswordFile); } catch { appPassword = ''; }
    if (!appPassword) return finish(403, { ok: false, code: 'PRODUCT_OCS_CONFIG_DENY_NO_REQUESTER_SECRET' });

    const instanceCfg = config.instances[body.instance];
    const url = `${instanceCfg.nextcloudUrl.replace(/\/$/, '')}/ocs/v2.php/apps/eurooffice/api/v1/config/${body.fileId}?format=json`;
    const basic = `Basic ${Buffer.from(`${requester.nextcloudUser}:${appPassword}`).toString('base64')}`;

    let ncStatus;
    let ncBody;
    try {
      const r = await fetchImpl(url, { method: 'GET', headers: { Authorization: basic, 'OCS-APIRequest': 'true', Accept: 'application/json' }, signal: AbortSignal.timeout(config.timeoutMs) });
      ncStatus = r.status;
      const text = await r.text();
      try { ncBody = JSON.parse(text); } catch { ncBody = null; }
    } catch {
      return finish(502, { ok: false, code: 'PRODUCT_OCS_CONFIG_UPSTREAM_FAILED' });
    }

    if (ncStatus === 401) return finish(403, { ok: false, code: 'PRODUCT_OCS_CONFIG_REQUESTER_CREDENTIAL_UNAVAILABLE' });
    const metaCode = ocsMetaStatuscode(ncBody);
    if (ncStatus === 403 || ncStatus === 404 || metaCode === 403 || metaCode === 404 || metaCode === 997 || metaCode === 998) {
      return finish(403, { ok: false, code: 'PRODUCT_OCS_CONFIG_DENY_AUTHZ' });
    }
    if (ncStatus >= 500 || ncBody === null) return finish(502, { ok: false, code: 'PRODUCT_OCS_CONFIG_UPSTREAM_FAILED' });
    // fail-closed hygiene (review, point 1): any Nextcloud HTTP status outside the ones named above
    // (e.g. 302, 400, 410) is NOT a recognised success -- never fall through to the config path on an unmapped status.
    if (ncStatus !== 200) return finish(502, { ok: false, code: 'PRODUCT_OCS_CONFIG_UPSTREAM_FAILED' });
    // fail-closed hygiene (point 2): an OCS envelope whose meta.statuscode is neither a known-good value (100/200) nor
    // one of the named deny codes (403/404/997/998) is unrecognised -- never pass it through as a success.
    if (metaCode !== null && metaCode !== 100 && metaCode !== 200) return finish(502, { ok: false, code: 'PRODUCT_OCS_CONFIG_UPSTREAM_FAILED' });

    const effectiveConfig = unwrapOcs(ncBody);
    if (body.requestedMode === 'WRITE' && !isEditCapable(effectiveConfig)) {
      return finish(403, { ok: false, code: 'PRODUCT_OCS_CONFIG_DENY_CONFIG_BELOW_REQUEST' });
    }
    return finish(200, { ok: true, status: 200, config: effectiveConfig });
  }

  return http.createServer((req, res) => {
    handle(req, res).catch(() => { try { send(res, 502, { ok: false, code: 'PRODUCT_OCS_CONFIG_UPSTREAM_FAILED' }); } catch { /* socket already gone */ } });
  });
}

module.exports = { loadConfig, createServer, isEditCapable, unwrapOcs, ocsMetaStatuscode, hexDigestsEqual, sha256hex };

/**
 * CLI entry point: `node reference-broker/server.cjs --config <path-to-config.json>`.
 * An invalid or unreadable config is never silent: a one-line reason goes to stderr and the
 * process exits 78 (EX_CONFIG, sysexits.h) without starting a listener.
 */
if (require.main === module) {
  const args = process.argv.slice(2);
  const i = args.indexOf('--config');
  const configPath = i !== -1 ? args[i + 1] : null;
  if (!configPath) {
    console.error('usage: node reference-broker/server.cjs --config <path-to-config.json>');
    process.exitCode = 78;
  } else {
    let config = null;
    try {
      config = loadConfig(JSON.parse(fs.readFileSync(configPath, 'utf8')));
    } catch (err) {
      console.error(`reference broker: invalid config (${configPath}): ${err.message}`);
      process.exitCode = 78;
    }
    if (config) {
      createServer({ config }).listen(config.listen.port, config.listen.host, function () {
        console.error(`reference broker listening on ${config.listen.host}:${this.address().port}${config.path}`);
      });
    }
  }
}
