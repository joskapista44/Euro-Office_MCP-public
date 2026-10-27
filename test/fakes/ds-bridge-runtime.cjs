'use strict';

/**
 * The REAL DS Bridge v1 block, as the canonical patcher installs it (documentserver-patches/9.3.5.0/
 * euro-office-bridge-v1/bridge-block.js, a standalone script -- no installer-diff manifest wrapper, unlike the
 * retired 9.3.4.60 family), evaluated against a fake `Api` -- so the offline tests exercise the exact bridge code
 * the Document Server serves, not an imitation. `installBridge(api)` adds `api.GetEuroOfficeBridge` and returns
 * the runtime id its manifest advertises (core/runtime-identity.cjs: documentServerVersion|patchVersion -- the
 * stable identity).
 *
 * DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): was documentserver-patches/9.3.4.60/
 * euro-office-bridge-v1/manifest.json's `operations[0].new` text -- retired with the rest of the 9.3.4.60
 * family. The 9.3.5.0 bridge-block.js self-reports `documentServerVersion: "9.3.5"`, `patchVersion:
 * "euro-office-bridge-v1/9.3.5.0/16"` (MEASURED, grep on the file), matching RUNTIMES.xlsx exactly.
 */
const fs = require('node:fs');
const path = require('node:path');

const BLOCK = fs.readFileSync(path.join(__dirname, '../../documentserver-patches/9.3.5.0/euro-office-bridge-v1/bridge-block.js'), 'utf8').replace(/,\s*$/, '');

function installBridge(api) {
  // eslint-disable-next-line no-new-func
  new Function('Api', `${BLOCK};`)(api);
  const m = api.GetEuroOfficeBridge().manifest();
  return `${m.documentServerVersion}|${m.patchVersion}`;
}

/** The same served bridge, but its manifest reports another patch revision -- an UNACCEPTED runtime (owner decision (B)). */
const OTHER_PATCH = 'euro-office-bridge-v1/9.3.5.0/4';
function withPatch(api, patchVersion = OTHER_PATCH) {
  const real = api.GetEuroOfficeBridge;
  return { ...api, GetEuroOfficeBridge: () => { const b = real(); const m = b.manifest(); return { manifest: () => ({ ...m, patchVersion }), read: b.read, hydrate: b.hydrate }; } };
}

module.exports = { installBridge, withPatch, OTHER_PATCH, BRIDGE_RUNTIME_ID: (() => { const a = {}; return installBridge(a); })() };
