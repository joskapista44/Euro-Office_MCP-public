'use strict';

/**
 * Release 3 trim helper: the tests of an intent that is routed + coded but NOT in the accepted DOCX WRITE allowlist are SKIPPED with a named reason
 * and run again by themselves the moment the intent is added back to the allowlist (no test edit on re-release). Only the allowlist decides.
 */
const nodeTest = require('node:test');
const { DOCX_WRITE_ALLOWLIST } = require('../../src/server/write-tools.cjs');

/** false when the intent IS released (the test runs), else the named skip reason. */
const skipUnlessReleased = intent => (DOCX_WRITE_ALLOWLIST.includes(intent) ? false : `${intent} is withheld by the release 3 trim -- runs again once it is back in DOCX_WRITE_ALLOWLIST`);

/** a drop-in `test` whose cases are skipped while `intent` is withheld. */
const testFor = intent => (name, fn) => nodeTest(name, { skip: skipUnlessReleased(intent) }, fn);

module.exports = { skipUnlessReleased, testFor };
