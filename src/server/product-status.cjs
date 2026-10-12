'use strict';

/**
 * product_status -- the read-only status tool and the source of `--identity` (owner decision, owner channel
 * 2026-10-08, points 6-7).
 *
 * Returns identity, config schema/identity, the DS handshake classification, the tool-exposure gate state and a
 * dependency summary. Never a secret, a token, a credential value or a secret-bearing path: the broker route appears
 * only as "configured / missing", the config only as its hash. Since M11.2 the product holds no credential reference:
 * the credential model is the requester's OWN Nextcloud credential, used by the broker only.
 *
 * The DS handshake needs an open editor session of a requester's file, so the probe is injected: production uses
 * `defaultDsProbe`, which REFUSES (fail-closed) -- no broker route, or live probing not wired yet (a READ-tool stage);
 * tests inject a handshake result. Whatever the probe returns is
 * classified by the one canonical classifier (core/runtime-identity.cjs), so a mismatch can only fail closed.
 */

const { ACCEPTED_RUNTIMES, XLSX_RUNTIME_IDS, classifyHandshake } = require('../core/runtime-identity.cjs');
const { CONFIG_SCHEMA_VERSION } = require('./config.cjs');
const { releaseIdentity } = require('./identity.cjs');
const { exposureGate, exposedTools, TOOL_ROWS } = require('./tools.cjs');
const { createBrowserRuntime } = require('./browser-runtime.cjs');

/** The production probe: fail-closed -- no broker route, or the live probe is not wired (it needs a requester's file). */
async function defaultDsProbe(config) {
  if (!config || !config.broker || !config.broker.url) return { refused: 'BROKER_ROUTE_MISSING' };
  return { refused: 'LIVE_PROBE_NOT_WIRED' };
}

/** Classify a probe answer into a DS section: never PASS without a classified manifest. */
function dsSection(probe) {
  if (!probe || probe.refused) {
    return { status: 'FAIL_CLOSED', reason: probe ? probe.refused : 'NO_PROBE', expectedRuntime: XLSX_RUNTIME_IDS.join(',') };
  }
  const c = classifyHandshake(probe.handshake, ACCEPTED_RUNTIMES.xlsx);
  if (!c.ok) return { status: 'FAIL_CLOSED', reason: c.reason, expectedRuntime: XLSX_RUNTIME_IDS.join(',') };
  return {
    status: 'PASS', runtimeId: c.runtimeId, expectedRuntime: XLSX_RUNTIME_IDS.join(','),
    documentServerVersion: c.manifest.documentServerVersion, patchVersion: c.manifest.patchVersion,
    bridgeVersion: c.manifest.bridgeVersion,
    capabilities: c.manifest.capabilities.map(x => `${x.id}@${x.version}`),
  };
}

// M11.3b.2: the pinned browser runtime check (playwright-core 1.62.1 + headless shell 1234 under playwright.path) --
// the same check the READ tools run before any broker call; no browser is started here
function playwrightResolvable(config) {
  const r = createBrowserRuntime({ browsersPath: config && config.playwright ? config.playwright.path : undefined }).check();
  return r.ok ? 'AVAILABLE' : r.code;
}

/** Build the status. `opts`: { config, configIdentity, dsProbe, root }. */
async function productStatus({ config, configIdentity, dsProbe = defaultDsProbe, root } = {}) {
  const id = releaseIdentity(root);
  const gate = exposureGate();
  const ds = dsSection(await Promise.resolve().then(() => dsProbe(config)).catch(() => ({ refused: 'PROBE_THREW' })));
  const exposed = exposedTools();
  return {
    product: id.PRODUCT,
    sourceCommit: id.SOURCE_COMMIT,
    artifact: { version: id.ARTIFACT_VERSION, sha256: id.ARTIFACT_SHA256, released: id.RELEASED },
    config: { schemaVersion: CONFIG_SCHEMA_VERSION, identity: configIdentity || null,
      credentialModel: 'REQUESTER_OWN_VIA_BROKER', brokerRoute: config && config.broker && config.broker.url ? 'CONFIGURED' : 'MISSING' },
    ds,
    exposureGate: gate,
    tools: { exposed, withheld: TOOL_ROWS.filter(r => !exposed.includes(r.TOOL)).map(r => r.TOOL) },
    dependencies: { node: process.version, playwright: playwrightResolvable(config) },
    health: ds.status === 'PASS' ? 'READY' : `DEGRADED: ds ${ds.reason}; exposure gate ${gate.state}`,
  };
}

/** The --identity lines (KEY=VALUE), from the same status. */
function identityLines(status) {
  const ds = status.ds;
  return [
    `PRODUCT=${status.product}`,
    `SOURCE_COMMIT=${status.sourceCommit}`,
    `ARTIFACT_VERSION=${status.artifact.version}`,
    `ARTIFACT_SHA256=${status.artifact.sha256}`,
    `CONFIG_SCHEMA_VERSION=${status.config.schemaVersion}`,
    `CONFIG_IDENTITY=${status.config.identity || 'NO_CONFIG'}`,
    `DS_DOCUMENT_SERVER_VERSION=${ds.status === 'PASS' ? ds.documentServerVersion : `UNAVAILABLE(${ds.reason})`}`,
    `DS_PATCH_VERSION=${ds.status === 'PASS' ? ds.patchVersion : `UNAVAILABLE(${ds.reason})`}`,
    `DS_BRIDGE_CAPABILITIES=${ds.status === 'PASS' ? ds.capabilities.join(',') : `UNAVAILABLE(${ds.reason})`}`,
    `EXPOSURE_GATE=${status.exposureGate.state}`,
  ];
}

module.exports = { productStatus, identityLines, dsSection, defaultDsProbe };
