'use strict';

/**
 * v1.1.0 release condition: the README and the CHANGELOG both say that upgrading from 1.0.x is a Document Server
 * version change (1.0.x accepts only 9.3.4.60, 1.1.0 only 9.3.5.0) and needs one maintenance window. The docs live
 * at release/public/root/ in the canonical tree and at the root of the derived public tree; the test runs in both.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const ROOT = path.join(__dirname, '..', '..');
const doc = name => [`release/public/root/${name}`, name].map(rel => path.join(ROOT, rel)).find(p => fs.existsSync(p));

/** The note names both Document Server versions, the 1.0.x upgrade and the maintenance window. */
const states = text => /1\.0\.x/.test(text) && /9\.3\.4\.60/.test(text) && /9\.3\.5\.0/.test(text) && /maintenance window/.test(text);

for (const name of ['README.md', 'CHANGELOG.md']) {
  test(`${name} states that 1.0.x -> 1.1.0 is a Document Server version change with one maintenance window`, () => {
    assert.ok(doc(name), `${name} not found`);
    assert.ok(states(fs.readFileSync(doc(name), 'utf8')));
  });
}

test('RED CONTROL: the same check fails on a copy without the maintenance window or without the old DS version', () => {
  const real = fs.readFileSync(doc('README.md'), 'utf8');
  assert.equal(states(real.split('maintenance window').join('restart')), false);
  assert.equal(states(real.split('9.3.4.60').join('9.3.x')), false);
});
