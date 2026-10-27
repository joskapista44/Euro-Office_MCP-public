'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { needsPrivate } = require('../helpers/private-source.cjs');

/**
 * GH #76 MODEL 1 (owner decision 2026-10-18): the stored-content release gate is MANDATORY for every new xlsx WRITE
 * capability, newly supported operation shape, or change to the save / verification semantics. Mechanically: every
 * test/live/release-*.live.cjs that is NOT on the frozen pre-gate list must import ./stored-content-gate.cjs, and the
 * frozen list never grows.
 */
const LIVE = path.join(__dirname, '..', 'live');
const LEGACY = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'fixtures', 'release-harness-legacy-before-stored-gate.json'), 'utf8')).legacy;
const releaseHarnesses = () => fs.readdirSync(LIVE).filter(f => /^release-.*\.live\.cjs$/.test(f));
const usesGate = src => /require\(['"]\.\/stored-content-gate\.cjs['"]\)/.test(src);

test('every release harness added after the gate became mandatory imports the stored-content gate', { skip: needsPrivate('test/live') }, () => {
  const missing = releaseHarnesses().filter(f => !LEGACY.includes(f) && !usesGate(fs.readFileSync(path.join(LIVE, f), 'utf8')));
  assert.deepEqual(missing, [], `new release harness(es) without ./stored-content-gate.cjs: ${missing.join(', ')}`);
});

test('the pre-gate list is frozen: 77 names, every one still a real file, and the gate harness itself is not on it', { skip: needsPrivate('test/live') }, () => {
  assert.equal(LEGACY.length, 77);
  const present = new Set(releaseHarnesses());
  assert.deepEqual(LEGACY.filter(f => !present.has(f)), [], 'a legacy entry whose file is gone must be removed, never replaced');
  assert.ok(!LEGACY.includes('release-gh76-stored-content-gate-product.live.cjs'));
  assert.ok(usesGate(fs.readFileSync(path.join(LIVE, 'release-gh76-stored-content-gate-product.live.cjs'), 'utf8')));
});

test('the detector is not blind: a new harness without the import is caught, one with it passes', () => {
  assert.equal(usesGate("const gate = require('./stored-content-gate.cjs');"), true);
  assert.equal(usesGate("const w1 = require('./xlsx-w1-write-range.live.cjs');"), false);
});
