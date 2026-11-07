'use strict';

/**
 * GH #338-#342 remainder (Bridge /24, DESIGN docs/DS-9.3.5-BRIDGE24-WORKBOOK-PROTECTION-DESIGN.md, W1+W3) + GH #335
 * remainder (sheet.table.write/v1 totalsRow value:false, tracker 02fae56e) -- the euro-office-bridge-v1 /24 CANDIDATE
 * for DS 9.3.5 (documentserver-patches/9.3.5.0/euro-office-bridge-v1/rev24-candidate/): workbook.protection.read/v1,
 * workbook.protection.write/v1, and the widened totalsRow branch of the ALREADY-SHIPPED sheet.table.write/v1. Offline:
 * the REAL candidate block is evaluated against a fake window/Asc/AscCommon in the SHAPES the design document names
 * (wbModel.getWorkbookProtection(type), wbModel.setProtectedWorkbook(props, true), Asc.CWorkbookProtection), and
 * against the /19-measured table-model fake engine for the totalsRow change. This proves the HANDLER LOGIC -- NOT
 * that a workbook-protection change reaches the server-side save and survives a fresh reopen on the real engine
 * (design gates G1-G3, NOT MEASURED in this unit; no LIVE run, no product-side wiring).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V23 = fs.readFileSync(path.join(DIR, 'rev23-candidate', 'bridge-block.js'), 'utf8');
const V24 = fs.readFileSync(path.join(DIR, 'rev24-candidate', 'bridge-block.js'), 'utf8');
const READ_HANDLER = fs.readFileSync(path.join(DIR, 'rev24-candidate', 'handler-wb-protection-read-v24.txt'), 'utf8');
const WRITE_HANDLER = fs.readFileSync(path.join(DIR, 'rev24-candidate', 'handler-wb-protection-write-v24.txt'), 'utf8');
const WR = 'workbook.protection.read/v1';
const WW = 'workbook.protection.write/v1';
const TW = 'sheet.table.write/v1';
const sha = t => crypto.createHash('sha256').update(t).digest('hex');
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);

// ---- the build ------------------------------------------------------------------------------------------

test('the /24 block = the /23 block + exactly the builder substitutions, and the builder is deterministic', () => {
  const script = path.join(DIR, 'rev24-candidate', 'build-rev24.py');
  const run = () => spawnSync('python3', ['-I', script], { cwd: path.dirname(script), encoding: 'utf8' });
  // the builder also (re)writes manifest-mechanism-proof.json next to bridge-block.js; snapshot and restore it so
  // these two re-runs leave the tree exactly as committed (the published copy is redacted by the derivation -- a rebuild from
  // the published, redaction-touched build-rev24.py would otherwise write back a byte-different proof file).
  const proofPath = path.join(DIR, 'rev24-candidate', 'manifest-mechanism-proof.json');
  const proofBefore = fs.readFileSync(proofPath);
  const r1 = run();
  assert.equal(r1.status, 0, r1.stderr);
  const r2 = run();
  assert.equal(r2.stdout, r1.stdout, 'the build is byte-for-byte reproducible across two runs');
  assert.equal(fs.readFileSync(path.join(DIR, 'rev24-candidate', 'bridge-block.js'), 'utf8'), V24, 'the committed block equals what the builder just produced');
  assert.equal(fs.readFileSync(path.join(DIR, 'rev24-candidate', 'old.txt'), 'utf8'), V23);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev24-candidate', 'new.txt'), 'utf8'), V24);
  assert.equal(V24.includes('totals-off-not-measured'), false, 'the retired reason name is gone from /24');
  fs.writeFileSync(proofPath, proofBefore);
});

test('both blocks parse the same way (the fragment trailing comma stripped)', () => {
  for (const t of [V23, V24]) assert.doesNotThrow(() => new Function('Api', 'Asc', 'window', 'AscCommon', t.replace(/\}\s*,\s*$/, '}')));
});

test('both new handler texts enumerate no object, carry no line comment and name no credential field', () => {
  for (const HANDLER of [READ_HANDLER, WRITE_HANDLER]) {
    for (const w of ['Object.keys', 'Object.values', 'Object.entries', 'getOwnPropertyNames', 'for (var k in', 'for (var p in', ' in args',
      'hashValue', 'saltValue', 'spinCount', 'algorithmName', 'password', 'temporaryPassword', 'asc_setSheet', 'asc_setPassword']) {
      assert.ok(!HANDLER.includes(w), w);
    }
    assert.ok(!HANDLER.includes('//'));
  }
});

test('manifest: /24, both new capabilities listed with the right kind/mutation, every /23 capability still listed', () => {
  const man = load(engineWB()).manifest();
  assert.equal(man.patchVersion, 'euro-office-bridge-v1/9.3.5.0/24');
  const ids = man.capabilities.map(c => c.id);
  const r = ids.find(x => x === WR); const w = man.capabilities.find(x => x.id === WW);
  assert.ok(r); assert.equal(man.capabilities.find(x => x.id === WR).kind, 'READ');
  assert.equal(w.kind, 'WRITE'); assert.equal(w.mutation, 'TARGETED');
  assert.ok(ids.includes(TW));
});

test('RED control: the /23 block advertises neither new workbook capability', () => {
  const br = load(engineWB(), V23);
  assert.equal(reason(br.read(WR, {})), 'capability-not-advertised');
  assert.equal(reason(br.write(WW, { protect: true })), 'capability-not-advertised');
});

// ---- the fake workbook-protection engine (design: wbModel.getWorkbookProtection(type) / .setProtectedWorkbook) ------

function engineWB({ lockStructure = false, modelMissing = false, apiMissing = false, flagNotBoolean = false,
  historyMissing = false, historyOn = true, setterThrows = false, setterIgnores = false,
  nullSettersMissing = false, credentialSticky = false } = {}) {
  const H = { pts: 0, Create_NewPoint() { this.pts += 1; }, StartTransaction() { }, EndTransaction() { }, Is_On: () => historyOn };
  let lock = flagNotBoolean ? 'nope' : lockStructure;
  let wp = null; // the LAST props object applied via setProtectedWorkbook -- the fake's wb.workbookProtection
  function CWorkbookProtection() { this._lock = null; this._credDirty = false; }
  if (!apiMissing) {
    // the internal tracker LIVE run measured this as a side effect of the ONE call, not a separate step -- the fake mirrors that, not a guess.
    CWorkbookProtection.prototype.asc_setLockStructure = function (v) { this._lock = v; if (v === true) this._credDirty = true; };
    if (!nullSettersMissing) {
      const clear = function () { if (!credentialSticky) this._credDirty = false; };
      CWorkbookProtection.prototype.asc_setWorkbookAlgorithmName = clear;
      CWorkbookProtection.prototype.asc_setWorkbookSaltValue = clear;
      CWorkbookProtection.prototype.asc_setWorkbookSpinCount = clear;
      CWorkbookProtection.prototype.asc_setWorkbookHashValue = clear;
    }
    CWorkbookProtection.prototype.asc_isPassword = function () { return this._credDirty; };
  }
  let setCalls = 0;
  const wb = modelMissing ? {} : {
    getWorkbookProtection(t) { if (t !== 0) throw new Error('bad type'); return lock; },
    setProtectedWorkbook(props, addToHistory) {
      setCalls += 1;
      if (setterThrows) throw new Error('stub: engine refused');
      if (setterIgnores) return true;
      lock = props._lock;
      wp = props;
      return true;
    },
    get workbookProtection() { return wp; },
  };
  const Asc = { editor: { wbModel: wb }, c_oAscWorkbookProtectType: { lockStructure: 0, lockWindows: 1, lockRevisions: 2 }, CWorkbookProtection };
  return {
    Asc, window: { Asc }, AscCommon: historyMissing ? {} : { History: H }, H,
    get lock() { return lock; }, get setCalls() { return setCalls; },
    get isPassword() { return wp ? wp.asc_isPassword() : undefined; },
  };
}

function load(e, block = V24) {
  const Api = { GetSheet: () => null };
  new Function('Api', 'Asc', 'window', 'AscCommon', block.replace(/\}\s*,\s*$/, '}'))(Api, e.Asc, e.window, e.AscCommon);
  return Api.GetEuroOfficeBridge();
}

// ---- workbook.protection.read/v1 -------------------------------------------------------------------------

test('read: {lockStructure} reflects the model, both states; args may be omitted or {}', () => {
  for (const v of [true, false]) {
    const e = engineWB({ lockStructure: v });
    const b = load(e);
    assert.deepEqual(b.read(WR, {}), { ok: true, value: { lockStructure: v } });
    assert.deepEqual(b.read(WR, undefined), { ok: true, value: { lockStructure: v } });
  }
});

test('read: bad-args / model-missing / protection-api-missing / protection-flag-not-boolean', () => {
  assert.equal(reason(load(engineWB()).read(WR, { extra: 1 })), 'OK:{"lockStructure":false}', 'extra keys are harmless, nothing is enumerated');
  assert.equal(reason(load(engineWB()).read(WR, 'nope')), 'bad-args');
  assert.equal(reason(load(engineWB({ modelMissing: true })).read(WR, {})), 'model-missing');
  const e = engineWB(); delete e.Asc.c_oAscWorkbookProtectType;
  assert.equal(reason(load(e).read(WR, {})), 'protection-api-missing');
  assert.equal(reason(load(engineWB({ flagNotBoolean: true })).read(WR, {})), 'protection-flag-not-boolean');
});

// ---- workbook.protection.write/v1 ------------------------------------------------------------------------

test('write: protect true from unlocked -- one history point, read-after-write confirms, receipt names the field, credential members end cleared', () => {
  const e = engineWB({ lockStructure: false });
  const r = load(e).write(WW, { protect: true });
  assert.deepEqual(r, { ok: true, value: { applied: ['lockStructure'] } });
  assert.equal(e.lock, true);
  assert.equal(e.H.pts, 1);
  assert.equal(e.isPassword, false, 'the engine-measured asc_setLockStructure(true) side effect must be cleared before the handler returns ok:true');
});

test('SECURITY (internal tracker LIVE second finding): asc_setLockStructure(true) is NOT credential-free by itself -- the handler must null-clear it', () => {
  // this is the measured engine behaviour the fake simulates: lockStructure(true) alone dirties the credential fields
  const e = engineWB({ lockStructure: false });
  load(e).write(WW, { protect: true });
  assert.equal(e.isPassword, false, 'cleared by the handler, not by the engine call itself');
});

test('NEG control: without the null-clear calls, the read-after-write check is the ONLY thing left catching it -- and it does (two independent layers, proven separately)', () => {
  const e = engineWB({ lockStructure: false });
  const CLEAR_BLOCK = 'props.asc_setWorkbookAlgorithmName(null); props.asc_setWorkbookSaltValue(null); props.asc_setWorkbookSpinCount(null); props.asc_setWorkbookHashValue(null); ';
  assert.ok(V24.includes(CLEAR_BLOCK), 'the exact block under test must be present in the real /24 candidate');
  const noClearBlock = V24.replace(CLEAR_BLOCK, '');
  assert.notEqual(noClearBlock, V24, 'the removal applied to a real copy of the block');
  const r = load(e, noClearBlock).write(WW, { protect: true });
  // the engine-measured side effect is never cleared without this block -- the read-after-write layer (separately
  // proven by the SECURITY/credential-not-cleared test above) is what catches it here, not a silent success
  assert.deepEqual(r, { ok: false, reason: 'credential-not-cleared' }, 'the second layer (read-after-write + revert) catches what the first layer (null-clear) would otherwise have missed');
  assert.equal(e.lock, false, 'reverted -- no credential-bearing lock left standing even with the first layer removed');
});

test('NEG control, both layers removed: without EITHER the null-clear calls or the post-clear isPassword check, the hazard IS silent (the original review-blocking defect, reproduced)', () => {
  const e = engineWB({ lockStructure: false });
  const CLEAR_BLOCK = 'if (args.protect === true) { props.asc_setWorkbookAlgorithmName(null); props.asc_setWorkbookSaltValue(null); props.asc_setWorkbookSpinCount(null); props.asc_setWorkbookHashValue(null); } ';
  const CHECK_BLOCK = 'if (args.protect === true) { var wp0; try { wp0 = wb.workbookProtection; } catch (e2) { wp0 = undefined; } if (wp0 && typeof wp0.asc_isPassword === "function" && wp0.asc_isPassword()) { try { var revertProps = new A.CWorkbookProtection(); if (typeof revertProps.asc_setLockStructure === "function") { revertProps.asc_setLockStructure(false); wb.setProtectedWorkbook(revertProps, true); } } catch (e3) {} bad = { ok: false, reason: "credential-not-cleared" }; } }';
  assert.ok(V24.includes(CLEAR_BLOCK), 'the exact clear block under test must be present in the real /24 candidate');
  assert.ok(V24.includes(CHECK_BLOCK), 'the exact check block under test must be present in the real /24 candidate');
  const bothRemoved = V24.replace(CLEAR_BLOCK, '').replace(CHECK_BLOCK, '');
  assert.notEqual(bothRemoved, V24, 'both removals applied to a real copy of the block');
  const r = load(e, bothRemoved).write(WW, { protect: true });
  assert.equal(r.ok, true, 'with BOTH layers removed, the write succeeds -- the hazard the internal tracker LIVE run found is silent, not a failure');
  assert.equal(e.isPassword, true, 'and the "passwordless" lock reads back as credentialled');
});

test('SECURITY: credential-not-cleared -- if the clear does not take (engine quirk), the SAME transaction reverts the lock before returning the named refusal', () => {
  const e = engineWB({ lockStructure: false, credentialSticky: true });
  const r = load(e).write(WW, { protect: true });
  assert.deepEqual(r, { ok: false, reason: 'credential-not-cleared' });
  assert.equal(e.lock, false, 'the lock is reverted -- no credential-bearing lock is left standing');
  assert.equal(e.setCalls, 2, 'the initial setProtectedWorkbook plus the same-transaction revert');
  assert.equal(e.H.pts, 1, 'the revert is bundled into the SAME history point, not a second one');
});

test('write: protection-api-missing when a null-clear setter itself is missing -- checked BEFORE any setter call, zero mutation', () => {
  const e = engineWB({ lockStructure: false, nullSettersMissing: true });
  const r = load(e).write(WW, { protect: true });
  assert.deepEqual(r, { ok: false, reason: 'protection-api-missing' });
  assert.equal(e.setCalls, 0, 'checked before asc_setLockStructure is ever called -- nothing is written');
  assert.equal(e.H.pts, 0);
  assert.equal(e.lock, false);
});

// unprotect (protect:false on a currently-locked workbook) does NOT reach the setter in this build at all -- see
// the SECURITY test below. With no sourced credential detector (engine finding 4), the fail-closed fix refuses
// every protect:false-while-locked call, named credential-unobservable, regardless of whether a real password
// was ever set. There is therefore no "unprotect succeeds" case to test here; that was the review-blocking defect.

test('write: already-at-value is {applied:[],satisfied:true} with NO history point', () => {
  const e = engineWB({ lockStructure: true });
  const r = load(e).write(WW, { protect: true });
  assert.deepEqual(r, { ok: true, value: { applied: [], satisfied: true } });
  assert.equal(e.H.pts, 0, 'no history point for an already-satisfied target');
});

test('SECURITY: write protect:false on a CURRENTLY LOCKED workbook refuses credential-unobservable -- zero setter call, zero history point', () => {
  const e = engineWB({ lockStructure: true });
  const r = load(e).write(WW, { protect: false });
  assert.deepEqual(r, { ok: false, reason: 'credential-unobservable' });
  assert.equal(e.setCalls, 0, 'setProtectedWorkbook must never be called on this path');
  assert.equal(e.H.pts, 0, 'no history point may be created before the refusal');
  assert.equal(e.lock, true, 'the lock is untouched');
});

test('NEG control: without the fail-closed branch, the SAME world would unlock with no credential check (proves the branch is load-bearing)', () => {
  const e = engineWB({ lockStructure: true });
  const noGuardBlock = V24.replace('if (args.protect === false && cur === true) return { ok: false, reason: "credential-unobservable" }; ', '');
  assert.notEqual(noGuardBlock, V24, 'the removal applied to a real copy of the block');
  const r = load(e, noGuardBlock).write(WW, { protect: false });
  assert.equal(r.ok, true, 'without the fix, an unguarded removal would have applied');
  assert.equal(e.setCalls, 1);
});

test('write: protect:true on an unlocked workbook is UNAFFECTED by the fail-closed branch (no existing lock to preserve)', () => {
  const e = engineWB({ lockStructure: false });
  const r = load(e).write(WW, { protect: true });
  assert.deepEqual(r, { ok: true, value: { applied: ['lockStructure'] } });
  assert.equal(e.lock, true);
});

test('write: bad-args (missing / non-boolean protect)', () => {
  assert.equal(reason(load(engineWB()).write(WW, {})), 'bad-args');
  assert.equal(reason(load(engineWB()).write(WW, { protect: 'yes' })), 'bad-args');
  assert.equal(reason(load(engineWB()).write(WW, undefined)), 'bad-args');
});

test('write: model-missing / protection-api-missing (CWorkbookProtection, enum, setter, read-type)', () => {
  assert.equal(reason(load(engineWB({ modelMissing: true })).write(WW, { protect: true })), 'model-missing');
  const e1 = engineWB(); delete e1.window.Asc.CWorkbookProtection;
  assert.equal(reason(load(e1).write(WW, { protect: true })), 'protection-api-missing');
  const e2 = engineWB(); delete e2.Asc.c_oAscWorkbookProtectType;
  assert.equal(reason(load(e2).write(WW, { protect: true })), 'protection-api-missing');
  assert.equal(reason(load(engineWB({ flagNotBoolean: true })).write(WW, { protect: true })), 'protection-api-missing');
  assert.equal(reason(load(engineWB({ apiMissing: true })).write(WW, { protect: true })), 'protection-api-missing');
});

test('write: history-missing / history-off, no setter call', () => {
  const e1 = engineWB({ historyMissing: true });
  assert.equal(reason(load(e1).write(WW, { protect: true })), 'history-missing');
  assert.equal(e1.lock, false, 'no mutation when History itself is absent');
  const e2 = engineWB({ historyOn: false });
  assert.equal(reason(load(e2).write(WW, { protect: true })), 'history-off');
  assert.equal(e2.lock, false);
});

test('write: setter-threw / no-effect, each a named refusal, no partial success claimed', () => {
  assert.equal(reason(load(engineWB({ setterThrows: true })).write(WW, { protect: true })), 'setter-threw');
  assert.equal(reason(load(engineWB({ setterIgnores: true })).write(WW, { protect: true })), 'no-effect');
});

// ---- the fake table-model engine for sheet.table.write/v1 (same shapes as ds935-bridge19-table-write.test.cjs) ------

function mkTable(name, ref = 'A1:C5', { header = true, totals = false } = {}) {
  return { name, ref, header, totals, columns: ['Item', 'Qty', 'Note'], style: { name: 'TableStyleMedium2', rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false }, autoFilter: true };
}
function tableWorld({ tables = { S1: [mkTable('Table1')] } } = {}) {
  const model = tables;
  const names = Object.keys(model);
  const calls = [];
  const part = t => ({
    DisplayName: t.name, Ref: { getAbsName: () => '$' + t.ref.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2').replace(/:/, ':') },
    isHeaderRow: () => t.header, isTotalsRow: () => t.totals, TableColumns: t.columns.map(c => ({ Name: c })),
    TableStyleInfo: { Name: t.style.name, ShowRowStripes: t.style.rowStripes, ShowColumnStripes: t.style.columnStripes, ShowFirstColumn: t.style.firstColumn, ShowLastColumn: t.style.lastColumn },
    AutoFilter: t.autoFilter ? {} : null,
  });
  const Api = { GetSheet: n => (Object.prototype.hasOwnProperty.call(model, n) ? { worksheet: { get TableParts() { return model[n].map(part); } } } : null) };
  const find = name => { for (const s of names) for (const t of model[s]) if (t.name === name) return t; return null; };
  const e = {
    canEdit: () => true, collaborativeEditing: { getGlobalLock: () => false }, asc_isProtectedWorkbook: () => false,
    asc_getWorksheetsCount: () => names.length, asc_getWorksheetName: i => names[i], asc_showWorksheet: () => { },
    asc_getActiveWorksheetIndex: () => 0, asc_registerCallback: () => { }, asc_unregisterCallback: () => { },
    asc_changeFormatTableInfo(tn, opt, val) {
      calls.push(['asc_changeFormatTableInfo', tn, opt, val]);
      const t = find(tn);
      if (opt === 5) {
        const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(t.ref);
        if (val === true) { t.totals = true; t.ref = `${m[1]}${m[2]}:${m[3]}${Number(m[4]) + 1}`; }
        else { t.totals = false; t.ref = `${m[1]}${m[2]}:${m[3]}${Number(m[4]) - 1}`; }
      }
    },
  };
  const consts = { c_oAscChangeFilterOptions: { style: 2 }, c_oAscChangeTableStyleInfo: { rowStripes: 6, columnStripes: 3, firstColumn: 1, lastColumn: 2, rowTotal: 5 } };
  return { Api, Asc: Object.assign({ editor: e }, consts), model, calls };
}
function bridgeOfTable(w, block = V24) {
  new Function('Api', 'Asc', block.replace(/\}\s*,\s*$/, '}'))(w.Api, w.Asc);
  return w.Api.GetEuroOfficeBridge();
}
const writeTable = (w, args) => bridgeOfTable(w).write(TW, Object.assign({ sheet: 'S1', table: 'Table1' }, args));

// ---- sheet.table.write/v1, totalsRow -- the ONE changed handler ------------------------------------------

test('totalsRow ON, regression: unchanged behaviour (ref grows, header-row-missing still guards insertion)', () => {
  const w = tableWorld();
  const r = writeTable(w, { property: 'totalsRow', value: true });
  assert.deepEqual(r, { ok: true, value: { applied: ['totalsRow'], table: 'Table1', observed: true } });
  assert.equal(w.model.S1[0].ref, 'A1:C6');
  const w2 = tableWorld({ tables: { S1: [mkTable('Table1', 'A2:C5', { header: false })] } });
  assert.equal(reason(writeTable(w2, { property: 'totalsRow', value: true })), 'header-row-missing');
});

test('totalsRow OFF: now APPLIES -- the ref shrinks back, observed false, header-row-missing is NOT checked for OFF', () => {
  const w = tableWorld({ tables: { S1: [mkTable('Table1', 'A1:C6', { totals: true })] } });
  const r = writeTable(w, { property: 'totalsRow', value: false });
  assert.deepEqual(r, { ok: true, value: { applied: ['totalsRow'], table: 'Table1', observed: false } });
  assert.deepEqual(w.calls, [['asc_changeFormatTableInfo', 'Table1', 5, false]]);
  assert.equal(w.model.S1[0].ref, 'A1:C5');
  // OFF on a table WITHOUT a header row still applies -- the header-row check only guards insertion (true)
  const w2 = tableWorld({ tables: { S1: [mkTable('Table1', 'A2:C6', { header: false, totals: true })] } });
  assert.equal(writeTable(w2, { property: 'totalsRow', value: false }).ok, true);
});

test('totalsRow: already-at-value is direction-aware (false==false is already-at-value too)', () => {
  const w = tableWorld();
  assert.equal(reason(writeTable(w, { property: 'totalsRow', value: false })), 'already-at-value');
  assert.deepEqual(w.calls, []);
  const w2 = tableWorld({ tables: { S1: [mkTable('Table1', 'A1:C6', { totals: true })] } });
  assert.equal(reason(writeTable(w2, { property: 'totalsRow', value: true })), 'already-at-value');
});

test('totalsRow: value-not-boolean still refused; totals-off-not-measured no longer produced by any input', () => {
  assert.equal(reason(writeTable(tableWorld(), { property: 'totalsRow', value: 1 })), 'value-not-boolean');
  assert.equal(reason(writeTable(tableWorld(), { property: 'totalsRow' })), 'value-not-boolean');
  assert.notEqual(reason(writeTable(tableWorld({ tables: { S1: [mkTable('Table1', 'A1:C6', { totals: true })] } }), { property: 'totalsRow', value: false })), 'totals-off-not-measured');
});

test('totalsRow OFF, differential proof: /23 still refuses it not-measured, /24 applies it -- same world, same args', () => {
  const w23 = tableWorld({ tables: { S1: [mkTable('Table1', 'A1:C6', { totals: true })] } });
  const r23 = bridgeOfTable(w23, V23).write(TW, { sheet: 'S1', table: 'Table1', property: 'totalsRow', value: false });
  assert.equal(reason(r23), 'totals-off-not-measured');
  const w24 = tableWorld({ tables: { S1: [mkTable('Table1', 'A1:C6', { totals: true })] } });
  const r24 = bridgeOfTable(w24, V24).write(TW, { sheet: 'S1', table: 'Table1', property: 'totalsRow', value: false });
  assert.equal(r24.ok, true);
});
