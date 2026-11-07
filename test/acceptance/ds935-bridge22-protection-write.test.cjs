'use strict';

/**
 * GH #338-#342 -- the euro-office-bridge-v1 /22 WRITE candidate for DS 9.3.5 (documentserver-patches/9.3.5.0/euro-office-bridge-v1/
 * rev22-candidate/): sheet.protection.write/v1 + sheet.cellProtection.write/v1. Offline: the REAL candidate block is evaluated against a
 * fake window / Api / History in the SHAPES read from the served 9.3.5 cell sdk (Worksheet.setProtectedSheet(props, addToHistory) ->
 * CSheetProtection.set -> checkProperty -> History.Add per changed field, the object dropped when isDefault(); setDefaultInterface;
 * asc_isPassword; Range.setLocked / setHiddenFormulas; getCell3().getXfs()). This proves the HANDLER LOGIC -- NOT that a protection
 * change reaches the server-side save, survives co-editing and a fresh reopen on the real engine (design gates G1-G3, NOT MEASURED).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V21 = fs.readFileSync(path.join(DIR, 'rev21-candidate', 'bridge-block.js'), 'utf8');
const V22 = fs.readFileSync(path.join(DIR, 'rev22-candidate', 'bridge-block.js'), 'utf8');
const PW = 'sheet.protection.write/v1';
const CW = 'sheet.cellProtection.write/v1';
const sha = t => crypto.createHash('sha256').update(t).digest('hex');
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);

const FLAGS = ['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks',
  'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'];
const DEF = Object.fromEntries(FLAGS.map(f => [f, !(f === 'selectLockedCells' || f === 'selectUnlockedCells')]));
const CRED = ['algorithmName', 'hashValue', 'saltValue', 'spinCount', 'password', 'temporaryPassword'];
const cap = s => s.charAt(0).toUpperCase() + s.slice(1);

// ---- the fake engine ----------------------------------------------------------------------------------

function engine({ protectedWith = null, credential = false, setterThrows = false, setterIgnores = false, historyOn = true, cells = {} } = {}) {
  const H = {
    calls: [],
    Create_NewPoint() { this.calls.push('point'); }, StartTransaction() { this.calls.push('start'); }, EndTransaction() { this.calls.push('end'); },
    Is_On: () => historyOn,
    Add(type, from, to) { this.calls.push(['add', type, from, to]); },
  };
  const touched = [];
  function CSheetProtection() { }
  CSheetProtection.prototype.setDefaultInterface = function () {
    for (const c of CRED) this[c] = null;
    this.sheet = false; for (const f of FLAGS) this[f] = DEF[f];
  };
  CSheetProtection.prototype.setSheet = function (v) { this.sheet = v; };
  for (const f of FLAGS) CSheetProtection.prototype[`set${cap(f)}`] = function (v) { this[f] = v; };
  CSheetProtection.prototype.asc_isPassword = function () { return this.algorithmName != null || this.password != null; };
  CSheetProtection.prototype.isDefault = function () {
    return CRED.slice(0, 4).every(c => this[c] === null) && this.sheet === false && FLAGS.every(f => this[f] === DEF[f]);
  };
  /** a credential field read by anyone but the engine's own asc_isPassword is recorded (the handler must never touch one) */
  const watch = sp => new Proxy(sp, { get(t, k, r) { if (CRED.includes(k) && !watch.inEngine) touched.push(k); return Reflect.get(t, k, r); } });
  watch.inEngine = false;
  const wrapIsPassword = sp => { const orig = sp.asc_isPassword; sp.asc_isPassword = function () { watch.inEngine = true; try { return orig.call(this); } finally { watch.inEngine = false; } }; return sp; };
  const m = {
    sheetProtection: null,
    getId: () => 'sheet-7',
    setProtectedSheet(props, add) {
      if (setterThrows === true || (typeof setterThrows === 'function' && setterThrows())) throw new Error('x');
      if (setterIgnores) return true;
      watch.inEngine = true;
      try {
        let sp = this.sheetProtection;
        if (!sp) { sp = new CSheetProtection(); sp.setDefaultInterface(); }
        for (const k of [...CRED.slice(0, 5), 'sheet', ...FLAGS]) {
          if (sp[k] !== props[k]) { if (add) H.Add(k, sp[k], props[k]); sp[k] = props[k]; }
        }
        this.sheetProtection = sp.isDefault() ? null : watch(wrapIsPassword(sp));
      } finally { watch.inEngine = false; }
      return true;
    },
  };
  if (protectedWith) {
    const sp = new CSheetProtection(); sp.setDefaultInterface(); sp.sheet = true; Object.assign(sp, protectedWith);
    if (credential) { sp.algorithmName = 4; sp.hashValue = 'SECRET-HASH'; sp.saltValue = 'SECRET-SALT'; sp.spinCount = 100000; }
    m.sheetProtection = watch(wrapIsPassword(sp));
  }
  // cell xf model: cells['B2'] = { locked, hidden } (absent = null xfs = the Excel default: locked, not hidden)
  const col = c => String.fromCharCode(65 + c);
  const xf = { ...cells };
  const cellSetter = (r1, c1, r2, c2, key) => v => {
    if (setterThrows === 'cell') throw new Error('x');
    if (setterIgnores) return;
    for (let r = r1; r <= r2; r += 1) for (let c = c1; c <= c2; c += 1) { const a = `${col(c)}${r + 1}`; xf[a] = { ...(xf[a] || {}), [key]: v }; H.calls.push(['cell', key, a, v]); }
  };
  m.getRange3 = (r1, c1, r2, c2) => ({ setLocked: cellSetter(r1, c1, r2, c2, 'locked'), setHiddenFormulas: cellSetter(r1, c1, r2, c2, 'hidden') });
  m.getCell3 = (r, c) => ({ getXfs: () => xf[`${col(c)}${r + 1}`] || null });
  const parseA1 = s => {
    const mm = /^([A-Z])(\d+)(?::([A-Z])(\d+))?$/.exec(s);
    if (!mm) throw new Error('bad');
    const c1 = mm[1].charCodeAt(0) - 65; const r1 = Number(mm[2]) - 1;
    return { r1, c1, r2: mm[3] ? Number(mm[4]) - 1 : r1, c2: mm[3] ? mm[3].charCodeAt(0) - 65 : c1 };
  };
  const sheetObj = { worksheet: m, GetRange: a => ({ range: { bbox: parseA1(a) } }) };
  const views = [];
  const view = { model: m, updateAfterChangeSheetProtection() { views.push('refresh'); } };
  const Asc = { CSheetProtection, editor: { wb: { getWorksheet: () => view } } };
  return { H, m, xf, touched, views, sheetObj, window: { Asc }, AscCommon: { History: H } };
}

function load(e, block = V22) {
  const sheets = { S: e.sheetObj };
  const Api = { GetSheet: n => (Object.prototype.hasOwnProperty.call(sheets, n) ? sheets[n] : null) };
  new Function('Api', 'Asc', 'window', 'AscCommon', block.replace(/\}\s*,\s*$/, '}'))(Api, {}, e.window, e.AscCommon);
  return Api.GetEuroOfficeBridge();
}
const stored = e => (e.m.sheetProtection ? { on: e.m.sheetProtection.sheet, ...Object.fromEntries(FLAGS.map(f => [f, e.m.sheetProtection[f]])) } : null);

// ---- the build ----------------------------------------------------------------------------------------

test('the /22 block = the /21 block with exactly the builder substitutions; the builder is deterministic', () => {
  const r = spawnSync('python3', ['-I', '-c', 'import importlib.util,json,sys;s=importlib.util.spec_from_file_location("b",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);v21,v22=m.build_block();print(json.dumps({"sha21":m.V21_BLOCK_SHA,"v22":v22.decode("utf-8"),"impl":m.one_line(m.HANDLER_SRC),"entries":m.ENTRIES}))',
    path.join(DIR, 'rev22-candidate', 'build-rev22.py')], { encoding: 'utf8', maxBuffer: 1 << 24 });
  assert.equal(r.status, 0, r.stderr);
  const b = JSON.parse(r.stdout);
  assert.equal(sha(V21), b.sha21, 'the builder pins the /21 block');
  assert.equal(b.v22, V22, 'the committed block is exactly what the builder produces');
  const back = V22.replace(b.impl, '').replace(b.entries, '')
    .replace('patchVersion: "euro-office-bridge-v1/9.3.5.0/22"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/21"');
  assert.equal(back, V21, 'every /21 capability (manifest + handlers + the layout installer) is byte-for-byte unchanged');
  for (const word of CRED) assert.equal(b.impl.includes(word), false, `the handler text names no credential field: ${word}`);
});

test('RED control: the /21 block advertises neither write', () => {
  const br = load(engine(), V21);
  assert.equal(reason(br.write(PW, { sheet: 'S', protect: true })), 'capability-not-advertised');
  assert.equal(reason(br.write(CW, { sheet: 'S', range: 'B2', locked: false })), 'capability-not-advertised');
});

test('manifest: /22, both writes listed as WRITE / TARGETED, every /21 capability still listed', () => {
  const man = load(engine()).manifest();
  assert.equal(man.patchVersion, 'euro-office-bridge-v1/9.3.5.0/22');
  const ids = man.capabilities.map(c => c.id);
  for (const id of [PW, CW]) {
    const c = man.capabilities.find(x => x.id === id);
    assert.equal(c.kind, 'WRITE'); assert.equal(c.mutation, 'TARGETED');
  }
  const ids21 = load(engine(), V21).manifest().capabilities.map(c => c.id);
  for (const id of ids21) assert.ok(ids.includes(id), id);
  assert.equal(ids.length, ids21.length + 2);
});

// ---- sheet.protection.write/v1 -----------------------------------------------------------------------

test('protect passwordless: one point + transaction, sheet + the two named flags recorded, no credential field set, view refreshed', () => {
  const e = engine();
  const r = load(e).write(PW, { sheet: 'S', protect: true, forbids: { formatCells: false, insertRows: false } });
  assert.deepEqual(r, { ok: true, value: { applied: ['protect'] } });
  assert.deepEqual(stored(e), { on: true, ...DEF, formatCells: false, insertRows: false });
  assert.deepEqual(e.H.calls, ['point', 'start', ['add', 'sheet', false, true], ['add', 'formatCells', true, false], ['add', 'insertRows', true, false], 'end']);
  assert.deepEqual(e.touched, [], 'the handler read no credential field');
  for (const c of CRED) assert.equal(e.m.sheetProtection[c], null, `${c} stays null (passwordless)`);
  assert.deepEqual(e.views, ['refresh']);
});

test('protect with no forbids = the Excel default flags; the retry with the state already held is SATISFIED with NO history', () => {
  const e = engine();
  const br = load(e);
  assert.equal(reason(br.write(PW, { sheet: 'S', protect: true })), 'OK:{"applied":["protect"]}');
  assert.deepEqual(stored(e), { on: true, ...DEF });
  e.H.calls.length = 0;
  assert.deepEqual(br.write(PW, { sheet: 'S', protect: true }), { ok: true, value: { applied: [], satisfied: true } });
  assert.deepEqual(e.H.calls, []);
});

test('#342: a protected sheet re-protected with other flags changes exactly those flags', () => {
  const e = engine({ protectedWith: { formatCells: false } });
  const r = load(e).write(PW, { sheet: 'S', protect: true, forbids: { formatCells: true, selectLockedCells: true, sort: false } });
  assert.equal(reason(r), 'OK:{"applied":["protect"]}');
  assert.deepEqual(stored(e), { on: true, ...DEF, sort: false, selectLockedCells: true });
  assert.deepEqual(e.H.calls.filter(Array.isArray).map(c => c[1]), ['formatCells', 'selectLockedCells', 'sort']);
});

test('unprotect: the engine drops the object (default props); unprotecting an unprotected sheet is SATISFIED', () => {
  const e = engine({ protectedWith: { formatCells: false, insertRows: false } });
  const br = load(e);
  assert.equal(reason(br.write(PW, { sheet: 'S', protect: false })), 'OK:{"applied":["unprotect"]}');
  assert.equal(e.m.sheetProtection, null);
  e.H.calls.length = 0;
  assert.deepEqual(br.write(PW, { sheet: 'S', protect: false }), { ok: true, value: { applied: [], satisfied: true } });
  assert.deepEqual(e.H.calls, []);
  assert.deepEqual(e.touched, []);
});

test('a credential-protected sheet: protect AND unprotect refused, nothing changed, no credential field read by the handler', () => {
  for (const args of [{ protect: false }, { protect: true, forbids: { sort: false } }, { protect: true }]) {
    const e = engine({ protectedWith: { formatCells: false }, credential: true });
    const before = JSON.stringify(stored(e));
    assert.equal(reason(load(e).write(PW, { sheet: 'S', ...args })), 'credential-protected', JSON.stringify(args));
    assert.equal(JSON.stringify(stored(e)), before);
    assert.equal(e.m.sheetProtection.hashValue === 'SECRET-HASH', true, 'the credential is untouched');
    assert.deepEqual(e.H.calls, []);
    assert.deepEqual(e.touched.filter(k => k !== 'hashValue'), [], 'only this test itself read hashValue');
  }
});

test('the receipt never carries a credential: sentinel values on the engine object never reach the serialized result', () => {
  const e = engine({ protectedWith: {}, credential: true });
  const out = JSON.stringify(load(e).write(PW, { sheet: 'S', protect: false }));
  for (const s of ['SECRET-HASH', 'SECRET-SALT', '100000']) assert.equal(out.includes(s), false, s);
});

test('protection write refusals before any mutation, named', () => {
  const e = engine();
  const br = load(e);
  const cases = [
    [null, 'bad-args'], [{ protect: true }, 'bad-args'], [{ sheet: 'S' }, 'bad-args'], [{ sheet: 'S', protect: 'yes' }, 'bad-args'],
    [{ sheet: 'S', protect: true, forbids: null }, 'bad-args'], [{ sheet: 'S', protect: true, forbids: 'x' }, 'bad-args'],
    [{ sheet: 'S', protect: false, forbids: {} }, 'bad-args'],
    [{ sheet: 'S', protect: true, forbids: { sort: 'no' } }, 'flag-not-boolean:sort'],
    [{ sheet: 'S', protect: true, forbids: { sorting: true } }, 'forbids-not-canonical'],
    [{ sheet: 'S', protect: true, forbids: { sort: true, formatCells: false } }, 'forbids-not-canonical'],
    [{ sheet: 'S', protect: true, forbids: { password: 'x' } }, 'forbids-not-canonical'],
    [{ sheet: 'Nope', protect: true }, 'sheet-not-found'],
  ];
  for (const [args, want] of cases) assert.equal(reason(br.write(PW, args)), want, JSON.stringify(args));
  assert.deepEqual(e.H.calls, []);
  assert.equal(e.m.sheetProtection, null);
});

test('protection write engine-side refusals: history off / no-effect / a throwing setter restores the old state', () => {
  let e = engine({ historyOn: false });
  assert.equal(reason(load(e).write(PW, { sheet: 'S', protect: true })), 'history-off');
  assert.equal(e.m.sheetProtection, null);
  e = engine({ setterIgnores: true });
  assert.equal(reason(load(e).write(PW, { sheet: 'S', protect: true })), 'no-effect', 'an unchanged read-back is never success');
  let first = true;
  e = engine({ protectedWith: { formatCells: false }, setterThrows: () => { const t = first; first = false; return t; } });
  assert.equal(reason(load(e).write(PW, { sheet: 'S', protect: false })), 'setter-threw');
  assert.deepEqual(stored(e), { on: true, ...DEF, formatCells: false }, 'the old protection is still in place');
  assert.equal(e.H.calls[0], 'point'); assert.equal(e.H.calls.at(-1), 'end', 'the transaction is always closed');
});

// ---- sheet.cellProtection.write/v1 -------------------------------------------------------------------

test('cell protection: unlock B2:B3 + hide C1 on an unprotected sheet, one point + transaction each, read back per cell', () => {
  const e = engine();
  const br = load(e);
  assert.deepEqual(br.write(CW, { sheet: 'S', range: 'B2:B3', locked: false }), { ok: true, value: { applied: ['locked'] } });
  assert.deepEqual(br.write(CW, { sheet: 'S', range: 'C1', hidden: true }), { ok: true, value: { applied: ['hidden'] } });
  assert.deepEqual(e.xf, { B2: { locked: false }, B3: { locked: false }, C1: { hidden: true } });
  assert.deepEqual(e.H.calls, ['point', 'start', ['cell', 'locked', 'B2', false], ['cell', 'locked', 'B3', false], 'end',
    'point', 'start', ['cell', 'hidden', 'C1', true], 'end']);
});

test('cell protection: both flags in one call; the state already held (incl. the empty-cell default) is SATISFIED with NO history', () => {
  const e = engine({ cells: { A1: { locked: false, hidden: true } } });
  const br = load(e);
  assert.deepEqual(br.write(CW, { sheet: 'S', range: 'A1', locked: false, hidden: true }), { ok: true, value: { applied: [], satisfied: true } });
  assert.deepEqual(br.write(CW, { sheet: 'S', range: 'D4', locked: true, hidden: false }), { ok: true, value: { applied: [], satisfied: true } }, 'an empty cell is the Excel default');
  assert.deepEqual(e.H.calls, []);
  assert.equal(reason(br.write(CW, { sheet: 'S', range: 'A1', locked: true, hidden: false })), 'OK:{"applied":["locked","hidden"]}');
  assert.deepEqual(e.xf.A1, { locked: true, hidden: false });
});

test('cell protection refusals before any mutation: a PROTECTED sheet, whole row / column, over cap, bad args', () => {
  let e = engine({ protectedWith: { formatCells: false } });
  assert.equal(reason(load(e).write(CW, { sheet: 'S', range: 'C1', hidden: false })), 'sheet-protected', '#327 redaction cannot be turned off by a write');
  assert.deepEqual(e.H.calls, []);
  e = engine();
  const br = load(e);
  const cases = [
    [null, 'bad-args'], [{ sheet: 'S', range: 'A1' }, 'bad-args'], [{ sheet: 'S', locked: true }, 'bad-args'], [{ sheet: 'S', range: '', locked: true }, 'bad-args'],
    [{ sheet: 'S', range: 'A1', locked: 'no' }, 'invalid-value'], [{ sheet: 'S', range: 'A1', hidden: 1 }, 'invalid-value'],
    [{ sheet: 'S', range: 'B:B', locked: false }, 'range-shape-not-supported'], [{ sheet: 'S', range: '3:4', locked: false }, 'range-shape-not-supported'],
    [{ sheet: 'S', range: 'A', locked: false }, 'range-shape-not-supported'],
    [{ sheet: 'S', range: 'zz!', locked: false }, 'range-invalid'],
    [{ sheet: 'Nope', range: 'A1', locked: false }, 'sheet-not-found'],
  ];
  for (const [args, want] of cases) assert.equal(reason(br.write(CW, args)), want, JSON.stringify(args));
  e.sheetObj.GetRange = () => ({ range: { bbox: { r1: 0, c1: 0, r2: 9999, c2: 1 } } });
  assert.equal(reason(br.write(CW, { sheet: 'S', range: 'A1:B10000', locked: false })), 'range-over-cap:10000');
  assert.deepEqual(e.H.calls, []);
});

test('cell protection engine-side: history off / no-effect / a throwing setter, the transaction always closed', () => {
  let e = engine({ historyOn: false });
  assert.equal(reason(load(e).write(CW, { sheet: 'S', range: 'B2', locked: false })), 'history-off');
  e = engine({ setterIgnores: true });
  assert.equal(reason(load(e).write(CW, { sheet: 'S', range: 'B2', locked: false })), 'no-effect');
  e = engine({ setterThrows: 'cell' });
  assert.equal(reason(load(e).write(CW, { sheet: 'S', range: 'B2', locked: false })), 'setter-threw');
  assert.deepEqual(e.H.calls, ['point', 'start', 'end']);
});

// ---- the plan ----------------------------------------------------------------------------------------

test('plan-rev22.py: refuses an unknown input and writes nothing', () => {
  const tmp = fs.mkdtempSync(path.join(require('node:os').tmpdir(), 'p22-'));
  const input = path.join(tmp, 'sdk-all.js'); fs.writeFileSync(input, 'not the unpatched sdk');
  const out = path.join(tmp, 'plan.json');
  const r = spawnSync('python3', ['-I', path.join(DIR, 'rev22-candidate', 'plan-rev22.py'), '--input', input, '--out', out], { encoding: 'utf8' });
  assert.equal(r.status, 2, r.stderr);
  assert.match(r.stderr, /UNKNOWN_INPUT/);
  assert.equal(fs.existsSync(out), false);
  fs.rmSync(tmp, { recursive: true, force: true });
});
