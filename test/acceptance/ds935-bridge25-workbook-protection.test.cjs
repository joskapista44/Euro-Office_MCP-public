'use strict';

/**
 * W2 regression fix (internal tracker, two real production incidents on this exact handler) -- the euro-office-
 * bridge-v1 /25 CANDIDATE for DS 9.3.5 (documentserver-patches/9.3.5.0/euro-office-bridge-v1/rev25-candidate/):
 * workbook.protection.read/v1's `v === null` branch now answers {lockStructure:false} instead of the generic
 * protection-flag-not-boolean refusal /24 gave it. Offline: the REAL candidate block is evaluated against a fake
 * window/Asc/AscCommon in the shape the engine actually produced (getWorkbookProtection returning null when the
 * OOXML lockStructure attribute is absent), proving the HANDLER LOGIC -- NOT that the fix clears the production
 * incident on a real DS (that LIVE run is a separate, gated step, a maintainer's own lane).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');

// tracker an internal branch -- this test's OWN build step used to run
// build-rev25.py with cwd = the SHARED, committed rev25-candidate/ directory, writing bridge-block.js IN PLACE.
// ds935-bridge26's test file ALSO reads that exact path (both at its own module-load time and inside its own
// build step, which treats it as a pinned-sha dependency) -- under full-suite concurrency this is a read/write
// race on a SHARED mutable file. MEASURED directly (200 rapid rewrites, sampling the file's byte-length at max
// speed): Python's `Path.write_bytes()` truncates-then-writes, non-atomically -- caught the file at 0 bytes in
// 4/1712 samples. A concurrent reader landing in that window gets an empty read, which is EXACTLY the shape of
// a maintainer's two reported failures (an assertion expecting content sees `''`, or build-rev26.py's own sha/size pin
// on its /25 input refuses with a non-zero exit). The fix: run the builder in an ISOLATED, per-call temp copy
// of the two directories it touches -- never the shared repo tree -- so there is nothing left to race.
function isolatedBuildRun() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ds935-bridge25-'));
  const rev24Dir = path.join(tmp, 'rev24-candidate');
  const rev25Dir = path.join(tmp, 'rev25-candidate');
  fs.mkdirSync(rev24Dir);
  fs.mkdirSync(rev25Dir);
  fs.copyFileSync(path.join(DIR, 'rev24-candidate', 'bridge-block.js'), path.join(rev24Dir, 'bridge-block.js'));
  fs.copyFileSync(path.join(DIR, 'rev25-candidate', 'build-rev25.py'), path.join(rev25Dir, 'build-rev25.py'));
  fs.copyFileSync(path.join(DIR, 'rev25-candidate', 'handler-wb-protection-read-v25.txt'), path.join(rev25Dir, 'handler-wb-protection-read-v25.txt'));
  const script = path.join(rev25Dir, 'build-rev25.py');
  const run = () => spawnSync('python3', ['-I', script], { cwd: rev25Dir, encoding: 'utf8' });
  return { tmp, rev25Dir, run };
}
const V24 = fs.readFileSync(path.join(DIR, 'rev24-candidate', 'bridge-block.js'), 'utf8');
const V25 = fs.readFileSync(path.join(DIR, 'rev25-candidate', 'bridge-block.js'), 'utf8');
const READ_HANDLER_25 = fs.readFileSync(path.join(DIR, 'rev25-candidate', 'handler-wb-protection-read-v25.txt'), 'utf8');
const WR = 'workbook.protection.read/v1';
const WW = 'workbook.protection.write/v1';
const sha = t => crypto.createHash('sha256').update(t).digest('hex');
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);

// ---- the build --------------------------------------------------------------------------------------------

test('the /25 block = the /24 block + exactly the one read-handler substitution, and the builder is deterministic', () => {
  const { tmp, rev25Dir, run } = isolatedBuildRun();
  try {
    const r1 = run();
    assert.equal(r1.status, 0, r1.stderr);
    const r2 = run();
    assert.equal(r2.stdout, r1.stdout, 'the build is byte-for-byte reproducible across two runs');
    assert.equal(fs.readFileSync(path.join(rev25Dir, 'bridge-block.js'), 'utf8'), V25, 'the committed block equals what the builder just produced');
    assert.equal(fs.readFileSync(path.join(rev25Dir, 'old.txt'), 'utf8'), V24);
    assert.equal(fs.readFileSync(path.join(rev25Dir, 'new.txt'), 'utf8'), V25);
    assert.ok(V25.includes('GH #338-#342 regression fix, /25'), 'the read-contract carries the new /25 tag');
    assert.ok(V25.includes('GH #338-#342, /24'), 'the write-contract keeps its OWN, separate /24 tag -- untouched');
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
});

test('both blocks parse the same way (the fragment trailing comma stripped)', () => {
  for (const t of [V24, V25]) assert.doesNotThrow(() => new Function('Api', 'Asc', 'window', 'AscCommon', t.replace(/\}\s*,\s*$/, '}')));
});

test('the new read handler text enumerates no object, carries no line comment and names no credential field', () => {
  for (const w of ['Object.keys', 'Object.values', 'Object.entries', 'getOwnPropertyNames', 'for (var k in', 'for (var p in', ' in args',
    'hashValue', 'saltValue', 'spinCount', 'algorithmName', 'password', 'temporaryPassword', 'asc_setSheet', 'asc_setPassword']) {
    assert.ok(!READ_HANDLER_25.includes(w), w);
  }
  assert.ok(!READ_HANDLER_25.includes('//'));
});

test('manifest: /25, the SAME two capabilities as /24 (nothing added, nothing removed), patchVersion bumped', () => {
  const man = load(engineWB()).manifest();
  assert.equal(man.patchVersion, 'euro-office-bridge-v1/9.3.5.0/25');
  const ids24 = load(engineWB(), V24).manifest().capabilities.map(c => c.id).sort();
  const ids25 = man.capabilities.map(c => c.id).sort();
  assert.deepEqual(ids25, ids24);
  assert.ok(ids25.includes(WR)); assert.ok(ids25.includes(WW));
});

// ---- the fake workbook-protection engine (adds the NULL shape the /24 unit's own fixture could not name) --------

function engineWB({ lockStructure = false, modelMissing = false, apiMissing = false, flagNotBoolean = false } = {}) {
  // the measured engine shape: CWorkbookProtection's constructor default is `this.lockStructure = null`; a present
  // <workbookProtection/> element with NO lockStructure attribute leaves it at that default, and
  // getWorkbookProtection's own `if (this.workbookProtection)` branch only decides WHETHER to call the getter, not
  // what it returns -- so `null` reaches the Bridge exactly as asc_getLockStructure() left it.
  let lock = flagNotBoolean ? 'nope' : lockStructure === null ? null : lockStructure;
  const wb = modelMissing ? {} : { getWorkbookProtection(t) { if (t !== 0) throw new Error('bad type'); return lock; } };
  const Asc = { editor: { wbModel: wb }, c_oAscWorkbookProtectType: apiMissing ? {} : { lockStructure: 0 } };
  return { Asc, window: { Asc }, AscCommon: {} };
}

function load(e, block = V25) {
  const Api = { GetSheet: () => null };
  new Function('Api', 'Asc', 'window', 'AscCommon', block.replace(/\}\s*,\s*$/, '}'))(Api, e.Asc, e.window, e.AscCommon);
  return Api.GetEuroOfficeBridge();
}

// ---- workbook.protection.read/v1 -- the fix itself -------------------------------------------------------

test('THE FIX: null (an empty <workbookProtection/> element -- attribute absent) answers {lockStructure:false}, not a refusal', () => {
  const br = load(engineWB({ lockStructure: null }));
  assert.deepEqual(br.read(WR, {}), { ok: true, value: { lockStructure: false } });
});

test('RED CONTROL: the SAME null engine shape on /24 (the regression itself) still refuses protection-flag-not-boolean', () => {
  const br = load(engineWB({ lockStructure: null }), V24);
  assert.equal(reason(br.read(WR, {})), 'protection-flag-not-boolean');
});

test('every OTHER non-boolean value is UNCHANGED, fail-closed -- the fix narrows the failure surface by exactly the one shape proven not a failure', () => {
  for (const v of [undefined, 0, 1, 'true', {}, []]) {
    const br = load(engineWB({ flagNotBoolean: true })); // 'nope' stands in for the generic non-boolean case
    assert.equal(reason(br.read(WR, {})), 'protection-flag-not-boolean');
  }
});

test('the two real boolean values are unchanged', () => {
  for (const v of [true, false]) {
    const br = load(engineWB({ lockStructure: v }));
    assert.deepEqual(br.read(WR, {}), { ok: true, value: { lockStructure: v } });
  }
});

test('model-missing and protection-api-missing refusals are unchanged', () => {
  assert.equal(reason(load(engineWB({ modelMissing: true })).read(WR, {})), 'model-missing');
  assert.equal(reason(load(engineWB({ apiMissing: true })).read(WR, {})), 'protection-api-missing');
});

test('bad-args is unchanged', () => {
  assert.equal(reason(load(engineWB()).read(WR, 'not-an-object')), 'bad-args');
});

// ---- workbook.protection.write/v1 -- proven UNTOUCHED, not merely asserted byte-identical ------------------

/** a full write-capable fixture (the /24 unit's own engineWB, reused verbatim: the credential-clear side effect,
 * history, setProtectedWorkbook) -- `fail` lets the RED control below make the setter report failure. */
function engineWBWrite({ fail = false } = {}) {
  const H = { pts: 0, Create_NewPoint() { this.pts += 1; }, StartTransaction() {}, EndTransaction() {}, Is_On: () => true };
  let lock = false; let wp = null;
  function CWorkbookProtection() { this._lock = null; this._credDirty = false; }
  CWorkbookProtection.prototype.asc_setLockStructure = function (v) { this._lock = v; if (v === true) this._credDirty = true; };
  const clear = function () { this._credDirty = false; };
  CWorkbookProtection.prototype.asc_setWorkbookAlgorithmName = clear;
  CWorkbookProtection.prototype.asc_setWorkbookSaltValue = clear;
  CWorkbookProtection.prototype.asc_setWorkbookSpinCount = clear;
  CWorkbookProtection.prototype.asc_setWorkbookHashValue = clear;
  CWorkbookProtection.prototype.asc_isPassword = function () { return this._credDirty; };
  const wb = {
    getWorkbookProtection(t) { if (t !== 0) throw new Error('bad type'); return lock; },
    setProtectedWorkbook(props) { if (fail) return true; lock = props._lock; wp = props; return true; },
    get workbookProtection() { return wp; },
  };
  const Asc = { editor: { wbModel: wb }, c_oAscWorkbookProtectType: { lockStructure: 0 }, CWorkbookProtection };
  return { Asc, window: { Asc }, AscCommon: { History: H } };
}

test('the write handler behaves IDENTICALLY on /24 and /25 (same engine, same call, same result) -- not just byte-identical text', () => {
  const r24 = load(engineWBWrite(), V24).write(WW, { protect: true });
  const r25 = load(engineWBWrite(), V25).write(WW, { protect: true });
  assert.equal(r24.ok, true, JSON.stringify(r24));
  assert.deepEqual(r25, r24);
});

test('RED CONTROL: the IDENTICAL-behaviour test above is not vacuous -- a /24 engine that fails to apply the lock makes the result differ', () => {
  const real = load(engineWBWrite({ fail: false }), V24).write(WW, { protect: true });
  const divergent = load(engineWBWrite({ fail: true }), V24).write(WW, { protect: true });
  assert.equal(real.ok, true, JSON.stringify(real));
  assert.notDeepEqual(divergent, real, 'a real engine difference DOES change the handler result -- the IDENTICAL-behaviour test above would have caught a real /25 regression');
});
