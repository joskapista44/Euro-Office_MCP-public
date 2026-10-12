'use strict';

/**
 * The product's runtime identity (owner decision, owner channel 2026-10-08, point 6).
 *
 * Nothing here is computed or guessed at run time:
 *   release.json   written INTO the artifact tree by scripts/build-artifact.cjs (PRODUCT, SOURCE_COMMIT,
 *                  ARTIFACT_VERSION, LOCKFILE_SHA256)
 *   artifact.json  written NEXT TO the unpacked tree by the installer (the tarball's own ARTIFACT_SHA256 -- a tarball
 *                  cannot contain its own hash)
 * A source checkout has neither: it reports UNRELEASED, never a git HEAD it happens to sit on.
 */

const fs = require('node:fs');
const path = require('node:path');

const PRODUCT = 'Euro-Office_MCP';
const ROOT = path.resolve(__dirname, '..', '..');

function readJson(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return null; }
}

/** { PRODUCT, SOURCE_COMMIT, ARTIFACT_VERSION, ARTIFACT_SHA256, RELEASED } -- `root` is injectable for tests. */
function releaseIdentity(root = ROOT) {
  const rel = readJson(path.join(root, 'release.json'));
  const art = readJson(path.join(root, 'artifact.json'));
  const released = !!(rel && rel.PRODUCT === PRODUCT && /^[0-9a-f]{40}$/.test(String(rel.SOURCE_COMMIT || '')));
  return {
    PRODUCT,
    RELEASED: released,
    SOURCE_COMMIT: released ? rel.SOURCE_COMMIT : 'UNRELEASED',
    ARTIFACT_VERSION: released && typeof rel.ARTIFACT_VERSION === 'string' ? rel.ARTIFACT_VERSION : 'UNRELEASED',
    ARTIFACT_SHA256: released && art && /^[0-9a-f]{64}$/.test(String(art.ARTIFACT_SHA256 || '')) ? art.ARTIFACT_SHA256 : 'NOT_RECORDED',
  };
}

module.exports = { PRODUCT, ROOT, releaseIdentity };
