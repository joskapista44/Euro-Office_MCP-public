'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const sheets = require('../../src/adapters/xlsx/families/sheets.cjs');
const { b03BridgeWritePrelude, b03DispatchResult } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

/**
 * M11.5c.4 (owner: M11.5 continuous mandate, internal tracker) -- the B03 cutover end to end, offline: the PRODUCT's
 * copy_sheet command (the driver's injected B03 client + the adapter's lowered body) runs against the REAL /8 bridge block
 * over a fake editor (Asc.editor) that records every call. One worksheet.copy/v1 write per command; the gate refuses a
 * runtime without the capability before any call; a refusal after the copy call is uncertain (the copy may exist).
 */
const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const block = rev => fs.readFileSync(path.join(DIR, rev, 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V7 = block('rev7-candidate'); const V8 = block('rev8-candidate');

function editor({ sheets: list = ['L', 'H'], error = false, copy = 'append' } = {}) {
  const calls = []; let onError = null;
  const e = {
    asc_getWorksheetsCount: () => list.length, asc_getWorksheetName: i => list[i], canEdit: () => true,
    collaborativeEditing: { getGlobalLock: () => 0 }, asc_isProtectedWorkbook: () => false,
    asc_registerCallback: (n, f) => { onError = f; }, asc_unregisterCallback: () => {},
    asc_copyWorksheet: (w, names, idx) => { calls.push([w, names, idx]); if (error && onError) onError(-1, 1); if (copy === 'append') list.push(names[0]); },
  };
  return { e, calls, list };
}
function runCopy(src, ed, s = { sheet: 'L', name: 'L_c' }) {
  const api = { GetSheet: () => null };
  new Function('Api', 'AscCommon', 'Asc', `${src};`)(api, {}, { editor: ed.e });
  const cmd = sheets.lower({ intent: 'copy_sheet', selector: s });
  return { cmd, result: new Function('Api', b03BridgeWritePrelude() + cmd.body)(api) };
}

test('the lowered copy_sheet is ONE Bridge WRITE command (no internal kind), marked for the driver second line', () => {
  const cmd = sheets.lower({ intent: 'copy_sheet', selector: { sheet: 'L', name: 'L_c' } });
  assert.equal(cmd.kind, 'callCommand');
  assert.deepEqual([cmd.meta.family, cmd.meta.action, cmd.meta.bridgeCopyWrite], ['sheets', 'copy_sheet', true]);
  assert.match(cmd.body, /__eoB03Copy\("L", "L_c"\)/);
  assert.doesNotMatch(cmd.body, /asc_/, 'no editor internal in the product body');
});

test('/8: one worksheet.copy/v1 write -> the product call (asc_copyWorksheet(-1, [name], [index])), body ok, dispatch ok', () => {
  const ed = editor({ sheets: ['A', 'L', 'H'] });
  const { result } = runCopy(V8, ed);
  assert.deepEqual(result, { ok: true, index: 3 });
  assert.deepEqual(ed.calls, [[-1, ['L_c'], [1]]]);
  assert.deepEqual(b03DispatchResult({ ok: true, outcome: 'command-completed', result }), { ok: true, dispatched: true, outcome: 'command-completed', error: null });
});

test('/7 (no B03 capability): the gate refuses BEFORE any editor call; mutation NONE; not ok', () => {
  const ed = editor();
  const { result } = runCopy(V7, ed);
  assert.deepEqual(result, { ok: false, stage: 'bridge-write-gate', reason: 'PATCH_VERSION_MISMATCH', mutation: 'NONE' });
  assert.deepEqual(ed.calls, []);
  assert.equal(b03DispatchResult({ ok: true, result }).ok, false);
});

test('a refusal BEFORE the copy call (name taken) -> mutation NONE, not uncertain; one AFTER (copy-error, copy-not-observed) -> uncertain, mutation UNKNOWN', () => {
  const taken = runCopy(V8, editor(), { sheet: 'L', name: 'h' }).result;
  assert.deepEqual([taken.ok, taken.stage, taken.reason, taken.uncertain, taken.mutation], [false, 'bridge-write', 'BRIDGE:name-taken', false, 'NONE']);
  const err = editor({ error: true }); const e = runCopy(V8, err).result;
  assert.deepEqual([e.ok, e.reason, e.uncertain, e.mutation], [false, 'BRIDGE:copy-error', true, 'UNKNOWN']);
  assert.deepEqual(err.list, ['L', 'H', 'L_c'], 'the copy exists although the receipt is a refusal: uncertain, never retried here');
  assert.equal(err.calls.length, 1);
  const none = runCopy(V8, editor({ copy: 'none' })).result;
  assert.deepEqual([none.reason, none.uncertain, none.mutation], ['BRIDGE:copy-not-observed', true, 'UNKNOWN']);
  for (const r of [taken, e, none]) { const d = b03DispatchResult({ ok: true, result: r }); assert.deepEqual([d.ok, d.dispatched, d.outcome], [false, true, 'b03-bridge-write-failed']); }
});

test('fail-closed classification: ONLY the handler\'s own pre-call reasons are mutation NONE; an unknown reason is uncertain', () => {
  const api = { GetSheet: () => null };
  const m = new Function('Api', 'AscCommon', 'Asc', `${V8}; return Api.GetEuroOfficeBridge();`)({ GetSheet: () => null }, {}, {}).manifest();
  const run = reason => { const cmd = sheets.lower({ intent: 'copy_sheet', selector: { sheet: 'L', name: 'L_c' } }); return new Function('Api', b03BridgeWritePrelude() + cmd.body)(Object.assign(api, { GetEuroOfficeBridge: () => ({ manifest: () => m, write: () => ({ ok: false, reason }) }) })); };
  for (const reason of ['bad-args', 'unsupported-property', 'invalid-name', 'editor-missing', 'edit-not-allowed', 'workbook-locked', 'workbook-protected', 'name-taken', 'source-not-found', 'source-ambiguous', 'asc_copyWorksheet-missing', 'canEdit-missing']) {
    const r = run(reason); assert.deepEqual([r.uncertain, r.mutation], [false, 'NONE'], reason);
  }
  for (const reason of ['copy-threw', 'copy-error', 'copy-not-observed', 'copy-unexpected', 'copy-partially-done-in-a-future-revision', 'internal-threw', 'something-missing']) {
    const r = run(reason); assert.deepEqual([r.uncertain, r.mutation], [true, 'UNKNOWN'], reason);
  }
});

test('a malformed receipt (wrong name / extra key) is UNEXPECTED_SHAPE, uncertain -- never APPLIED on the receipt', () => {
  const api = { GetSheet: () => null };
  const fake = r => ({ GetEuroOfficeBridge: () => ({ manifest: () => new Function('Api', 'AscCommon', 'Asc', `${V8}; return Api.GetEuroOfficeBridge();`)({ GetSheet: () => null }, {}, {}).manifest(), write: () => r }) });
  for (const r of [{ ok: true, value: { applied: ['copy'], name: 'X', index: 2 } }, { ok: true, value: { applied: ['copy'], name: 'L_c', index: 2, extra: 1 } }, { ok: true }, null]) {
    const cmd = sheets.lower({ intent: 'copy_sheet', selector: { sheet: 'L', name: 'L_c' } });
    const out = new Function('Api', b03BridgeWritePrelude() + cmd.body)(Object.assign(api, fake(r)));
    assert.deepEqual([out.ok, out.reason, out.uncertain], [false, 'UNEXPECTED_SHAPE', true], JSON.stringify(r));
  }
});
