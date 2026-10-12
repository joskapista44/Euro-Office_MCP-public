'use strict';

/**
 * The product MCP configuration (owner decision, owner channel 2026-10-08, point 8).
 *
 * Production location: /etc/euro-office-mcp/config.json (override: --config or EURO_OFFICE_MCP_CONFIG). The file
 * holds endpoints, timeouts, runtime policy and the broker ROUTE -- never a secret value and, since M11.2 (schemaVersion 2,
 * owner decision 2026-10-09), no credential reference either: the product holds no Nextcloud identity of its own; the
 * broker uses the REQUESTER's own credential. A config that carries anything secret-shaped is refused at startup
 * (fail-closed), so a secret cannot slip into it by accident.
 * Schema: config/euro-office-mcp.schema.json; example: config/euro-office-mcp.example.json.
 */

const fs = require('node:fs');
const crypto = require('node:crypto');

const CONFIG_SCHEMA_VERSION = 2;
const DEFAULT_CONFIG_PATH = '/etc/euro-office-mcp/config.json';

// key names that may never appear anywhere in the config (case-insensitive substring match)
const SECRET_KEY_PATTERN = /pass(word)?|secret|token|apikey|api_key|authorization|cookie|private_?key|credential_?value/i;
const TOP_KEYS = new Set(['schemaVersion', 'nextcloud', 'documentServer', 'broker', 'timeouts', 'playwright', 'log', 'policy']);
const LOOPBACK_URL = /^https?:\/\/(127\.0\.0\.1|localhost)(:\d+)?\//;

const isObj = v => v !== null && typeof v === 'object' && !Array.isArray(v);
const isHttpsUrl = v => typeof v === 'string' && /^https:\/\/[^\s/]+(\/[^\s]*)?$/.test(v);
const isPosInt = v => Number.isInteger(v) && v > 0;

function secretShapedKeys(value, at = '') {
  const hits = [];
  if (Array.isArray(value)) value.forEach((v, i) => hits.push(...secretShapedKeys(v, `${at}[${i}]`)));
  else if (isObj(value)) {
    for (const [k, v] of Object.entries(value)) {
      if (SECRET_KEY_PATTERN.test(k)) hits.push(`${at}.${k}`);
      hits.push(...secretShapedKeys(v, `${at}.${k}`));
    }
  }
  return hits;
}

/** { ok:true, config } | { ok:false, errors:[...] } -- pure, no I/O. */
function validateConfig(raw) {
  const errors = [];
  if (!isObj(raw)) return { ok: false, errors: ['config is not a JSON object'] };
  for (const k of Object.keys(raw)) {
    if (k === 'credential') errors.push('credential is superseded (M11.2, schemaVersion 2): the product holds no credential reference -- configure broker.url');
    else if (!TOP_KEYS.has(k)) errors.push(`unknown key: ${k}`);
  }
  if (raw.schemaVersion !== CONFIG_SCHEMA_VERSION) errors.push(`schemaVersion must be ${CONFIG_SCHEMA_VERSION}${raw.schemaVersion === 1 ? ' (schemaVersion 1 carried the superseded product-credential reference; migrate credential.brokerUrl -> broker.url)' : ''}`);
  const secrets = secretShapedKeys(raw);
  if (secrets.length) errors.push(`secret-shaped key(s) are not allowed in the config: ${secrets.join(', ')}`);

  if (!isObj(raw.nextcloud) || !isHttpsUrl(raw.nextcloud.url)) errors.push('nextcloud.url must be an https URL');

  if (raw.documentServer !== undefined) {
    const ds = raw.documentServer;
    if (!isObj(ds)) errors.push('documentServer must be an object');
    else {
      if (ds.expectedUrl !== undefined && !isHttpsUrl(ds.expectedUrl)) errors.push('documentServer.expectedUrl must be an https URL');
      if (ds.cacheKeyIsIdentity !== undefined) errors.push('documentServer.cacheKeyIsIdentity is not configurable: the URL cache key is never an identity');
    }
  }

  // M11.2: the broker ROUTE (non-secret). Absent = valid config, but every editor operation is refused (fail-closed):
  // the product has no other way to an editor config.
  if (raw.broker !== undefined) {
    const b = raw.broker;
    if (!isObj(b)) errors.push('broker must be an object');
    else {
      if (!(typeof b.url === 'string' && LOOPBACK_URL.test(b.url))) errors.push('broker.url must be a loopback http(s) URL (the broker route, e.g. http://127.0.0.1:3420/api/euro-broker)');
      if (b.instance !== undefined && !(typeof b.instance === 'string' && /^[a-z0-9_-]{1,32}$/.test(b.instance))) errors.push('broker.instance must match ^[a-z0-9_-]{1,32}$');
      if (b.timeoutMs !== undefined && !isPosInt(b.timeoutMs)) errors.push('broker.timeoutMs must be a positive integer');
      for (const k of Object.keys(b)) if (!['url', 'instance', 'timeoutMs'].includes(k)) errors.push(`unknown key: broker.${k}`);
    }
  }

  if (raw.timeouts !== undefined) {
    if (!isObj(raw.timeouts)) errors.push('timeouts must be an object');
    else for (const k of ['openMs', 'commandMs']) if (raw.timeouts[k] !== undefined && !isPosInt(raw.timeouts[k])) errors.push(`timeouts.${k} must be a positive integer`);
  }
  if (raw.playwright !== undefined && !(isObj(raw.playwright) && (raw.playwright.path === undefined || typeof raw.playwright.path === 'string'))) errors.push('playwright.path must be a string');
  if (raw.log !== undefined && !(isObj(raw.log) && (raw.log.level === undefined || ['error', 'warn', 'info', 'debug'].includes(raw.log.level)))) errors.push('log.level must be error|warn|info|debug');
  if (raw.policy !== undefined) {
    if (!isObj(raw.policy)) errors.push('policy must be an object');
    else for (const k of Object.keys(raw.policy)) errors.push(`unknown key: policy.${k} (no exposure override exists: only the reviewed release list EXPOSED_CAPABILITY_TOOLS releases a tool)`);
  }

  if (errors.length) return { ok: false, errors };
  return { ok: true, config: raw };
}

/** sha256 of the canonical (key-sorted) config: the CONFIG_IDENTITY. Safe to publish (the config holds no secret). */
function configIdentity(config) {
  const canon = v => (Array.isArray(v) ? v.map(canon) : isObj(v) ? Object.fromEntries(Object.keys(v).sort().map(k => [k, canon(v[k])])) : v);
  return crypto.createHash('sha256').update(JSON.stringify(canon(config))).digest('hex');
}

/** Load + validate a config file. Never throws. */
function loadConfig(file) {
  let text;
  try { text = fs.readFileSync(file, 'utf8'); } catch (e) { return { ok: false, errors: [`config not readable: ${file} (${e.code || 'ERR'})`] }; }
  let raw;
  try { raw = JSON.parse(text); } catch { return { ok: false, errors: [`config is not valid JSON: ${file}`] }; }
  const v = validateConfig(raw);
  return v.ok ? { ok: true, config: v.config, path: file, identity: configIdentity(v.config) } : { ok: false, errors: v.errors, path: file };
}

module.exports = { CONFIG_SCHEMA_VERSION, DEFAULT_CONFIG_PATH, validateConfig, loadConfig, configIdentity, secretShapedKeys };
