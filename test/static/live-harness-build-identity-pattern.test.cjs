'use strict';

/**
 * The runtime-identity migration (owner decision (B), internal tracker, fca99fc -- "migrate the
 * LIVE harnesses to the manifest runtime identity") removed `buildOf` from
 * `src/drivers/eurooffice/xlsx-driver.cjs`. Six X-O08 probes (16fd90b9, a maintainer's own LIVE run:
 * "LIVE r1 ERROR -- buildOf is not a function") were made on a pre-migration base and landed on
 * this branch WITHOUT migration -- `xlsx-x-o08-1a/-1b/-2/-3/-4/-5`, every one still importing
 * `buildOf` from the removed export, none of them runnable today. They are migrated in the SAME
 * patch as this test (only the build-identity import + assignment lines, per a maintainer's request --
 * no other line touched).
 *
 * This test's own job is NOT re-measuring that: it is a going-forward guard, so the same pattern
 * (a LIVE harness importing the removed `buildOf` symbol) cannot land again unnoticed. It walks
 * every `test/live/**​/*.live.cjs` file and asserts NONE contain the literal string -- the exact
 * same horgony the migration itself removes, not a narrower or looser one.
 */

const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');
const { needsPrivate } = require('../helpers/private-source.cjs');

const LIVE_DIR = path.join(__dirname, '..', 'live');

function liveFiles() {
  return fs.readdirSync(LIVE_DIR, { recursive: true })
    .filter(p => p.endsWith('.live.cjs'))
    .map(p => path.join(LIVE_DIR, p));
}

test('live-harness-build-identity-pattern: a POSITIVE control -- the probe itself can find a known string, so a zero result is not scanner-blindness', { skip: needsPrivate('test/live') }, () => {
  const files = liveFiles();
  assert.ok(files.length > 50, `expected the LIVE harness tree to contain well over 50 files, found ${files.length}`);
  const knownPresent = files.filter(f => fs.readFileSync(f, 'utf8').includes('xlsx-wave0-probe.live.cjs'));
  assert.ok(knownPresent.length > 10, 'xlsx-wave0-probe.live.cjs is required by most LIVE harnesses -- if this is 0, the file list itself is wrong, not the horgony');
});

test('live-harness-build-identity-pattern: a NEGATIVE control -- a freshly invented string is not found, so the scanner does not say yes to everything', { skip: needsPrivate('test/live') }, () => {
  const files = liveFiles();
  const hit = files.filter(f => fs.readFileSync(f, 'utf8').includes('ZSIRAF7788_NEVER_WRITTEN_HERE'));
  assert.deepEqual(hit, []);
});

test('live-harness-build-identity-pattern: NO live harness imports the removed `buildOf` export -- the migrated six (1a/1b/2/3/4/5) plus every harness after fca99fc', { skip: needsPrivate('test/live') }, () => {
  const files = liveFiles();
  const offenders = files.filter(f => fs.readFileSync(f, 'utf8').includes('buildOf')).map(f => path.relative(LIVE_DIR, f));
  assert.deepEqual(offenders, [], `the following LIVE harness(es) still reference the removed buildOf export: ${offenders.join(', ')}`);
});
