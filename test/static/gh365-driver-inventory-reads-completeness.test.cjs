'use strict';

/**
 * GH #365 option C U2 review (SZO SZERINT): "ez MASODSZOR tortenik -- #363
 * compiledStyleReads, most veryHiddenReads" -- a new `*Reads`-suffixed key is added to
 * xlsx-driver.cjs's readQuery() return statement, but the matching pass-through into
 * `src/adapters/xlsx/inventory.cjs`'s normalizeInventory() is forgotten. The family projection
 * then always sees `undefined`/`null` for that axis on the real product path, even though the
 * driver read it correctly -- a "forgot to widen a sibling SEAM" bug, one layer down from
 * test/static/ds935-bridge29-runtime-list-completeness.test.cjs's "forgot to widen a sibling
 * LIST" (same source-walk technique, same deliberate-exclusion convention, different pairing).
 *
 * The walk is SOURCE-BASED (not a fixed name list): every `*Reads`-suffixed identifier named in
 * xlsx-driver.cjs's readQuery() return statement must ALSO be a key normalizeInventory() forwards,
 * or be explicitly, nameably excluded with a comment carrying DELIBERATE_EXCLUSION_MARK right next
 * to the key in readQuery()'s return line's own file (the same "deliberate exclusion" convention
 * the XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS precedent already uses) -- a FUTURE key added to the return
 * statement and never widened into inventory.cjs fails this test too.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const DRIVER_FILE = path.join(__dirname, '..', '..', 'src', 'drivers', 'eurooffice', 'xlsx-driver.cjs');
const INVENTORY_FILE = path.join(__dirname, '..', '..', 'src', 'adapters', 'xlsx', 'inventory.cjs');
const DELIBERATE_EXCLUSION_MARK = 'DELIBERATELY NOT FORWARDED TO INVENTORY';

/** the readQuery() return statement line itself -- a SINGLE line, source-based (not line-numbered by hand). */
function findReadQueryReturnLine(driverSrc) {
  const lines = driverSrc.split('\n');
  const idx = lines.findIndex(l => l.includes('return { ok: true, sheets: sheets.map'));
  if (idx === -1) return null;
  return { line: lines[idx], lineNo: idx + 1, context: lines.slice(Math.max(0, idx - 3), idx + 1).join('\n') };
}

/** every distinct `*Reads`-suffixed identifier named anywhere in that return line (covers both the unconditional and the `${cond ? 'xReads, ' : ''}` template-fragment forms). */
function readsKeysOf(returnLine) {
  const found = returnLine.match(/\b[A-Za-z_]\w*Reads\b/g) || [];
  return [...new Set(found)];
}

/** does inventory.cjs's own returned object literal carry this key (as `key:` -- the exact pattern every existing *Reads pass-through line already uses)? */
function inventoryForwards(invSrc, key) {
  return new RegExp(`\\b${key}:`).test(invSrc);
}

test('STATIC GUARD: every *Reads key in xlsx-driver.cjs\'s readQuery() return statement is ALSO forwarded by inventory.cjs\'s normalizeInventory(), or documents the deliberate exclusion right next to it', () => {
  const driverSrc = fs.readFileSync(DRIVER_FILE, 'utf8');
  const invSrc = fs.readFileSync(INVENTORY_FILE, 'utf8');
  const ret = findReadQueryReturnLine(driverSrc);
  assert.ok(ret, 'readQuery()\'s return statement must be found (source shape changed?)');
  const keys = readsKeysOf(ret.line);
  assert.ok(keys.length >= 15, `found ${keys.length} *Reads keys in the return line (expected at least the 16 already known)`);
  const unaccountedFor = keys.filter(k => !inventoryForwards(invSrc, k) && !ret.context.includes(DELIBERATE_EXCLUSION_MARK));
  assert.deepEqual(unaccountedFor, [], 'every *Reads key readQuery() returns must be forwarded by normalizeInventory(), or the return line\'s own nearby comment must name the deliberate exclusion');
});

test('RED CONTROL: a *Reads key present in readQuery()\'s return line but ABSENT from inventory.cjs (simulated, no real file touched) is flagged', () => {
  const fixtureReturnLine = "  return { ok: true, sheets: sheets.map(s => s.row), cells, newlyAddedReads, cfReads };";
  const fixtureInventorySrc = "cfReads: payload.cfReads ? payload.cfReads : null,"; // newlyAddedReads intentionally missing
  const keys = readsKeysOf(fixtureReturnLine);
  assert.deepEqual(keys, ['newlyAddedReads', 'cfReads']);
  const unaccountedFor = keys.filter(k => !inventoryForwards(fixtureInventorySrc, k) && !fixtureReturnLine.includes(DELIBERATE_EXCLUSION_MARK));
  assert.deepEqual(unaccountedFor, ['newlyAddedReads'], 'a key the driver returns but inventory.cjs never forwards must be caught -- this is the exact GH #365 U2 review bug shape, reproduced');
});

test('NEGATIVE CONTROL: the completeness check DOES discriminate -- a real, already-forwarded key (veryHiddenReads, this round\'s own fix) stripped from a COPY of inventory.cjs\'s source reproduces the flag', () => {
  const driverSrc = fs.readFileSync(DRIVER_FILE, 'utf8');
  const invSrc = fs.readFileSync(INVENTORY_FILE, 'utf8');
  const ret = findReadQueryReturnLine(driverSrc);
  const keys = readsKeysOf(ret.line);
  assert.ok(keys.includes('veryHiddenReads'), 'sanity: the real return line carries veryHiddenReads today');
  assert.equal(inventoryForwards(invSrc, 'veryHiddenReads'), true, 'sanity: the real inventory.cjs forwards it today (this round\'s fix is in place)');
  // simulate the PRE-FIX state by stripping the one line the fix added, from an in-memory copy only
  const strippedInvSrc = invSrc.split('\n').filter(l => !l.includes('veryHiddenReads: payload.veryHiddenReads')).join('\n');
  assert.equal(inventoryForwards(strippedInvSrc, 'veryHiddenReads'), false, 'sanity: the in-memory strip actually removed the forwarding line');
  const unaccountedFor = keys.filter(k => !inventoryForwards(strippedInvSrc, k) && !ret.context.includes(DELIBERATE_EXCLUSION_MARK));
  assert.ok(unaccountedFor.includes('veryHiddenReads'), 'stripping a real, already-forwarded key must reproduce the flag -- the check discriminates, not just passes');
});

test('veryHiddenReads itself: present in both files today, no deliberate-exclusion marker needed', () => {
  const driverSrc = fs.readFileSync(DRIVER_FILE, 'utf8');
  const invSrc = fs.readFileSync(INVENTORY_FILE, 'utf8');
  assert.ok(driverSrc.includes('veryHiddenReads'), 'driver carries it');
  assert.equal(inventoryForwards(invSrc, 'veryHiddenReads'), true, 'inventory.cjs forwards it (GH #365 U2 review fix)');
});
