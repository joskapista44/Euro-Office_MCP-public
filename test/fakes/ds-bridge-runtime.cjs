'use strict';

/**
 * The REAL DS Bridge v1 block, as the canonical patcher installs it (documentserver-patches/9.3.4.60/euro-office-bridge-v1,
 * the operation's `new` text), evaluated against a fake `Api` -- so the offline tests exercise the exact bridge code the
 * Document Server serves, not an imitation. `installBridge(api)` adds `api.GetEuroOfficeBridge` and returns the runtime
 * id its manifest advertises (core/runtime-identity.cjs: documentServerVersion|patchVersion -- the stable identity).
 */
const path = require('node:path');

const MANIFEST = require(path.join(__dirname, '../../documentserver-patches/9.3.4.60/euro-office-bridge-v1/manifest.json'));
const BLOCK = MANIFEST.operations[0].new.replace(/,\s*$/, '');

function installBridge(api) {
  // eslint-disable-next-line no-new-func
  new Function('Api', `${BLOCK};`)(api);
  const m = api.GetEuroOfficeBridge().manifest();
  return `${m.documentServerVersion}|${m.patchVersion}`;
}

/** The same served bridge, but its manifest reports another patch revision -- an UNACCEPTED runtime (owner decision (B)). */
const OTHER_PATCH = 'euro-office-bridge-v1/9.3.4.60/4';
function withPatch(api, patchVersion = OTHER_PATCH) {
  const real = api.GetEuroOfficeBridge;
  return { ...api, GetEuroOfficeBridge: () => { const b = real(); const m = b.manifest(); return { manifest: () => ({ ...m, patchVersion }), read: b.read, hydrate: b.hydrate }; } };
}

module.exports = { installBridge, withPatch, OTHER_PATCH, BRIDGE_RUNTIME_ID: (() => { const a = {}; return installBridge(a); })() };
