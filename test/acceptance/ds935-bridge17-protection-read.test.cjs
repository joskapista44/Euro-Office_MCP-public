'use strict';

/**
 * GH #326 / #327 -- the euro-office-bridge-v1 /17 candidate for DS 9.3.5 (documentserver-patches/9.3.5.0/
 * euro-office-bridge-v1/rev17-candidate/): ONE new READ capability, sheet.protection.read/v1. Offline: the REAL
 * candidate block is evaluated against a fake Api whose worksheet model has the shape LIVE-measured on DS 9.3.5
 * (sheetProtection with the 16 boolean flags AND the credential fields; getRange3()._foreachNoEmpty over cells
 * whose xfs is null or {locked, hidden}). The owner condition "never returns a password / hash / credential" is
 * asserted on the serialized output with sentinel values planted in every credential field.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V16 = fs.readFileSync(path.join(DIR, 'bridge-block.js'), 'utf8');
const V17 = fs.readFileSync(path.join(DIR, 'rev17-candidate', 'bridge-block.js'), 'utf8');
const CAP = 'sheet.protection.read/v1';
const SECRETS = { algorithmName: 'SENTINEL-ALG', hashValue: 'SENTINEL-HASH', saltValue: 'SENTINEL-SALT', spinCount: 'SENTINEL-SPIN', password: 'SENTINEL-PW', temporaryPassword: 'SENTINEL-TMP' };
const FLAGS = ['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks',
  'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'];
const sha = t => crypto.createHash('sha256').update(t).digest('hex');

const a1 = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { r: Number(m[2]) - 1, c: c - 1 }; };
const bbox = range => { const [p, q = p] = range.split(':'); const x = a1(p); const y = a1(q); return { r1: x.r, c1: x.c, r2: y.r, c2: y.c }; };

/** a fake worksheet in the LIVE-measured shape; cells: { A1: xfs-or-null } (only the non-empty cells) */
function sheet({ sp = null, cells = {}, walkThrows = false } = {}) {
  const list = Object.entries(cells).map(([ref, xfs]) => { const p = a1(ref); return { nRow: p.r, nCol: p.c, xfs }; });
  const worksheet = {
    sheetProtection: sp,
    getRange3: (r1, c1, r2, c2) => ({ _foreachNoEmpty: fn => { if (walkThrows) throw new Error('boom'); for (const cell of list) if (cell.nRow >= r1 && cell.nRow <= r2 && cell.nCol >= c1 && cell.nCol <= c2) fn(cell); } }),
  };
  return { worksheet, GetRange: r => { if (!/^[A-Z]+\d+(:[A-Z]+\d+)?$/.test(r)) throw new Error('bad range'); return { range: { bbox: bbox(r) } }; } };
}
const protectedSp = (over = {}) => Object.assign({ sheet: true, _ws: { name: 'cyclic' } }, SECRETS, Object.fromEntries(FLAGS.map(f => [f, f === 'insertRows' || f === 'formatCells'])), over);

function bridgeWith(sheets) {
  const Api = { GetSheet: n => (Object.prototype.hasOwnProperty.call(sheets, n) ? sheets[n] : null) };
  // the block is a fragment that ends with "}," (it is spliced into sdk-all.js before an anchor); the same
  // trailing-comma strip makes the shipped /16 block parse, see the syntax test below
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(Api);
  return Api.GetEuroOfficeBridge();
}

test('the /17 block = the shipped /16 block + exactly the builder\'s three substitutions', () => {
  const r = spawnSync('python3', ['-I', '-c', 'import importlib.util,json,sys;s=importlib.util.spec_from_file_location("b",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);print(json.dumps({"entry":m.PROTECTION_ENTRY,"impl":m.PROTECTION_IMPL,"sha":m.V16_BLOCK_SHA}))',
    path.join(DIR, 'rev17-candidate', 'build-rev17.py')], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stderr);
  const b = JSON.parse(r.stdout);
  assert.equal(sha(V16), b.sha, 'the builder pins the shipped /16 block');
  assert.equal(V17.split(b.entry).length - 1, 1);
  assert.equal(V17.split(b.impl).length - 1, 1);
  const back = V17.replace(b.entry, '').replace(b.impl, '').replace('patchVersion: "euro-office-bridge-v1/9.3.5.0/17"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/16"');
  assert.equal(back, V16, 'nothing else moved');
  assert.equal(fs.readFileSync(path.join(DIR, 'rev17-candidate', 'old.txt'), 'utf8'), V16);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev17-candidate', 'new.txt'), 'utf8'), V17);
});

test('both blocks parse the same way (the fragment trailing comma stripped)', () => {
  for (const t of [V16, V17]) assert.doesNotThrow(() => new Function('Api', t.replace(/\}\s*,\s*$/, '}')));
});

test('the manifest advertises /17 and the capability as a non-mutating READ', () => {
  const m = JSON.parse(bridgeWith({}).manifest ? JSON.stringify(bridgeWith({}).manifest()) : '{}');
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.5.0/17');
  const c = m.capabilities.filter(x => x.id === CAP);
  assert.equal(c.length, 1);
  assert.equal(c[0].kind, 'READ');
  assert.equal(c[0].mutation, 'NO');
});

test('unprotected sheet (sheetProtection null) -> protected false, forbids null', () => {
  const r = bridgeWith({ U: sheet() }).read(CAP, { sheet: 'U' });
  assert.deepEqual(r, { ok: true, value: { protected: false, forbids: null, cells: null } });
});

test('a protection object with sheet=false reads as unprotected (no flags reported)', () => {
  const r = bridgeWith({ U: sheet({ sp: protectedSp({ sheet: false }) }) }).read(CAP, { sheet: 'U' });
  assert.deepEqual(r.value, { protected: false, forbids: null, cells: null });
});

test('protected sheet -> the 15 named flags, and NO credential value anywhere in the output', () => {
  const r = bridgeWith({ P: sheet({ sp: protectedSp() }) }).read(CAP, { sheet: 'P' });
  assert.equal(r.ok, true);
  assert.equal(r.value.protected, true);
  assert.deepEqual(Object.keys(r.value.forbids).sort(), [...FLAGS].sort());
  assert.equal(r.value.forbids.insertRows, true);
  assert.equal(r.value.forbids.sort, false);
  const out = JSON.stringify(r);
  for (const v of Object.values(SECRETS)) assert.ok(!out.includes(v), `leaked ${v}`);
  for (const k of Object.keys(SECRETS)) assert.ok(!out.includes(k), `named ${k}`);
  assert.ok(!out.includes('cyclic'), 'the _ws back-reference is not followed');
});

test('a non-boolean flag is a named refusal, never a guessed value', () => {
  const r = bridgeWith({ P: sheet({ sp: protectedSp({ insertRows: 1 }) }) }).read(CAP, { sheet: 'P' });
  assert.deepEqual(r, { ok: false, reason: 'protection-flag-not-boolean:insertRows' });
  const r2 = bridgeWith({ P: sheet({ sp: protectedSp({ sheet: 'yes' }) }) }).read(CAP, { sheet: 'P' });
  assert.deepEqual(r2, { ok: false, reason: 'protection-flag-not-boolean:sheet' });
});

test('range walk: unlocked and hidden non-empty cells; an xfs-less cell is the locked default', () => {
  const s = sheet({ sp: protectedSp(), cells: { A1: null, B2: { locked: false, hidden: null }, C1: { locked: null, hidden: true }, D9: { locked: false, hidden: true } } });
  const r = bridgeWith({ P: s }).read(CAP, { sheet: 'P', range: 'A1:C3' });
  assert.equal(r.ok, true);
  assert.deepEqual(r.value.cells, { range: 'A1:C3', nonEmpty: 3, unlocked: ['B2'], hidden: ['C1'] });
});

test('the range walk also works on an unprotected sheet (the guard decides what it means)', () => {
  const r = bridgeWith({ U: sheet({ cells: { B2: { locked: false, hidden: null } } }) }).read(CAP, { sheet: 'U', range: 'A1:B2' });
  assert.deepEqual(r.value, { protected: false, forbids: null, cells: { range: 'A1:B2', nonEmpty: 1, unlocked: ['B2'], hidden: [] } });
});

test('the walk is capped: more than 10000 non-empty cells is a refusal, never a truncated list', () => {
  const cells = {}; for (let i = 1; i <= 10001; i += 1) cells['A' + i] = null;
  const r = bridgeWith({ P: sheet({ sp: protectedSp(), cells }) }).read(CAP, { sheet: 'P', range: 'A1:A10001' });
  assert.deepEqual(r, { ok: false, reason: 'cells-over-cap:10000' });
});

test('named refusals: bad args, missing sheet, invalid range, a throwing walk, a non-boolean cell flag', () => {
  const b = bridgeWith({ P: sheet({ sp: protectedSp(), cells: { A1: { locked: 'x', hidden: null } } }), T: sheet({ walkThrows: true }) });
  assert.deepEqual(b.read(CAP, {}), { ok: false, reason: 'bad-args' });
  assert.deepEqual(b.read(CAP, { sheet: 'P', range: 5 }), { ok: false, reason: 'bad-args' });
  assert.deepEqual(b.read(CAP, { sheet: 'nope' }), { ok: false, reason: 'sheet-not-found' });
  assert.deepEqual(b.read(CAP, { sheet: 'P', range: 'not a range' }), { ok: false, reason: 'range-invalid' });
  assert.deepEqual(b.read(CAP, { sheet: 'T', range: 'A1' }), { ok: false, reason: 'cells-walk-threw' });
  assert.deepEqual(b.read(CAP, { sheet: 'P', range: 'A1' }), { ok: false, reason: 'cell-flag-not-boolean' });
});

test('the handler calls no setter: the fake model is unchanged after every read', () => {
  const sp = protectedSp();
  const before = JSON.stringify(sp, (k, v) => (k === '_ws' ? undefined : v));
  bridgeWith({ P: sheet({ sp, cells: { B2: { locked: false, hidden: null } } }) }).read(CAP, { sheet: 'P', range: 'A1:C3' });
  assert.equal(JSON.stringify(sp, (k, v) => (k === '_ws' ? undefined : v)), before);
});

test('plan-rev17.py derives a plan the UNCHANGED install_bridge.py accepts; an unknown input is refused, nothing written', () => {
  const os = require('node:os');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-b17-plan-'));
  try {
    const root = path.join(__dirname, '..', '..', 'documentserver-patches');
    fs.cpSync(path.join(root, '9.3.5.0'), path.join(tmp, '9.3.5.0'), { recursive: true });
    fs.copyFileSync(path.join(root, '9.3.4.60', 'install_bridge.py'), path.join(tmp, 'install_bridge.py'));
    const data = Buffer.from('head;AscBuilder.Cell.init=function(x){};mid;Api.GetSheets=function(){};tail');
    const planPath = path.join(tmp, '9.3.5.0', 'fresh-install.json');
    const plan = JSON.parse(fs.readFileSync(planPath, 'utf8'));
    plan.cell.input_sha256 = sha(data);
    fs.writeFileSync(planPath, JSON.stringify(plan));
    fs.writeFileSync(path.join(tmp, 'in.js'), data);
    const helper = path.join(tmp, '9.3.5.0', 'euro-office-bridge-v1', 'rev17-candidate', 'plan-rev17.py');
    const out = path.join(tmp, 'p17.json');
    const d = spawnSync('python3', ['-I', helper, '--input', path.join(tmp, 'in.js'), '--out', out], { encoding: 'utf8' });
    assert.equal(d.status, 0, d.stderr);
    const p17 = JSON.parse(fs.readFileSync(out, 'utf8'));
    assert.equal(p17.cell.patchVersion, 'euro-office-bridge-v1/9.3.5.0/17');
    assert.equal(p17.cell.blocks.find(b => b.file.startsWith('euro-office-bridge-v1/')).sha256, sha(V17));
    fs.copyFileSync(out, path.join(tmp, '9.3.5.0', 'p17.json'));
    const c = spawnSync('python3', ['-I', path.join(tmp, 'install_bridge.py'), '--target', 'cell', '--plan', path.join(tmp, '9.3.5.0', 'p17.json'), '--input', path.join(tmp, 'in.js'), '--check'], { encoding: 'utf8' });
    assert.equal(c.status, 0, c.stderr);
    assert.match(c.stdout, /STATUS=PATCHABLE PATCH_VERSION=euro-office-bridge-v1\/9\.3\.5\.0\/17/);
    fs.writeFileSync(path.join(tmp, 'bad.js'), Buffer.concat([data, Buffer.from('x')]));
    const bad = path.join(tmp, 'bad.json');
    const n = spawnSync('python3', ['-I', helper, '--input', path.join(tmp, 'bad.js'), '--out', bad], { encoding: 'utf8' });
    assert.equal(n.status, 2);
    assert.match(n.stderr, /UNKNOWN_INPUT/);
    assert.equal(fs.existsSync(bad), false);
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
});
