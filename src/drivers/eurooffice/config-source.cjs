'use strict';

/**
 * M11.2 (owner GO, owner channel 2026-10-09) -- the editor-config source seam of the Euro-Office drivers.
 *
 * A driver opens a document with the editor config of ONE fileId. Where that config comes from is not the driver's
 * business: it only calls `configSource.getEditorConfig({ fileId, operation })` and gets either the config object or a
 * thrown `ConfigSourceError` (a controlled refusal: the driver's open() fails, the editor is never started). The
 * driver knows nothing about broker tokens, the vault, Nextcloud passwords or requester mappings.
 *
 * Two sources exist:
 *   - PRODUCT runtime: the broker source (src/server/broker-config-source.cjs) -- the requester's OWN Nextcloud
 *     credential is used by the broker, never by this process;
 *   - TEST / HARNESS / LEGACY only: `createBasicAuthConfigSource` below (direct OCS call with a Basic-auth user/pass).
 *     The product runtime never constructs it (guarded by test/acceptance/m11-2-credential-routing.test.cjs).
 */

const OPERATIONS = Object.freeze(['READ', 'WRITE']);

/** A controlled config-source refusal: `code` is the product error class, `detail` a secret-free reason. */
class ConfigSourceError extends Error {
  constructor(code, detail) {
    super(`editor config refused: ${code}${detail ? ` (${detail})` : ''}`);
    this.name = 'ConfigSourceError';
    this.code = code;
    this.detail = detail || null;
  }
}

const isConfigSource = s => s !== null && typeof s === 'object' && typeof s.getEditorConfig === 'function';

/**
 * TEST / HARNESS / LEGACY ONLY -- the pre-M11.2 direct path: Basic auth with a user/pass against the Nextcloud OCS
 * config endpoint. `operation` is ignored (the OCS endpoint has no operation parameter).
 */
function createBasicAuthConfigSource({ url, user, pass, fetchImpl = (...a) => globalThis.fetch(...a) }) {
  const base = String(url).replace(/\/$/, '');
  const auth = 'Basic ' + Buffer.from(`${user}:${pass}`).toString('base64');
  return {
    kind: 'BASIC_AUTH_LEGACY',
    async getEditorConfig({ fileId }) {
      const r = await fetchImpl(`${base}/ocs/v2.php/apps/eurooffice/api/v1/config/${encodeURIComponent(fileId)}?format=json`,
        { headers: { Authorization: auth, 'OCS-APIRequest': 'true', Accept: 'application/json' } });
      if (!r.ok) throw new Error(`config HTTP ${r.status}`);
      const j = await r.json();
      return (j && j.ocs && j.ocs.data) || j;
    },
  };
}

/**
 * The driver's config source: an explicit `configSource` wins; without one, the legacy Basic-auth source is built from
 * `url` / `user` / `pass` exactly as before M11.2 (the existing LIVE harnesses construct drivers that way).
 */
function resolveConfigSource({ configSource = null, url, user, pass } = {}) {
  if (configSource !== null && configSource !== undefined) {
    if (!isConfigSource(configSource)) throw new TypeError('configSource must provide getEditorConfig({ fileId, operation })');
    return configSource;
  }
  return createBasicAuthConfigSource({ url, user, pass });
}

module.exports = { OPERATIONS, ConfigSourceError, isConfigSource, createBasicAuthConfigSource, resolveConfigSource };
