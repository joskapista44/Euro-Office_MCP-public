'use strict';

/**
 * GH #332-#335 -- the euro-office-bridge-v1 /19 candidate for DS 9.3.5 (documentserver-patches/9.3.5.0/euro-office-bridge-v1/
 * rev19-candidate/): ONE new WRITE capability, sheet.table.write/v1 (style | flag | totalsRow | name | ref, one property per
 * call). Offline: the REAL candidate block is evaluated against a fake Api (the /18-measured worksheet model: TableParts) and a
 * fake EDITOR (Asc.editor) whose table actions take the table NAME and mutate that model the way DS 9.3.5 was measured to
 * (docs/evidence/gh332-gh335-engine-access-*.json). This proves the HANDLER LOGIC -- NOT that an editor table action called from
 * inside a Bridge write() lands in history and persists on the real engine; that is the design's FIRST LIVE GATE (NOT MEASURED).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V18 = fs.readFileSync(path.join(DIR, 'rev18-candidate', 'bridge-block.js'), 'utf8');
const V19 = fs.readFileSync(path.join(DIR, 'rev19-candidate', 'bridge-block.js'), 'utf8');
const HANDLER = fs.readFileSync(path.join(DIR, 'rev19-candidate', 'handler-table-write.txt'), 'utf8');
const CAP = 'sheet.table.write/v1';
const sha = t => crypto.createHash('sha256').update(t).digest('hex');
const CONSTS = { c_oAscChangeTableStyleInfo: { columnFirst: 1, columnLast: 2, columnBanded: 3, rowHeader: 4, rowTotal: 5, rowBanded: 6, filterButton: 7, advancedSettings: 8 }, c_oAscChangeFilterOptions: { filter: 1, style: 2 } };

const colName = n => { let s = ''; for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s; return s; };
const colNum = s => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

/** a model table, in the /18-measured shape */
function mkTable(name, ref = 'A1:C5', { header = true, totals = false, style = 'TableStyleMedium2', stripes = true } = {}) {
  return { name, ref, header, totals, columns: ['Item', 'Qty', 'Note'], style: style === null ? null : { name: style, rowStripes: stripes, columnStripes: false, firstColumn: false, lastColumn: false }, autoFilter: true };
}

/** the world: sheets -> model tables; the Api reads them in the TableParts shape, the editor mutates them */
function world({ tables = { S1: [mkTable('Table1')] }, fault = null, canEdit = true, lock = false, protectedWb = false, omit = [], consts = CONSTS, active = 0, errorOnCall = false } = {}) {
  const model = tables;
  const names = Object.keys(model);
  const log = { calls: [], shows: [], registered: 0, unregistered: 0 };
  const part = t => ({
    DisplayName: t.name,
    Ref: { getAbsName: () => '$' + t.ref.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2').replace(/:/, ':') },
    isHeaderRow: () => t.header,
    isTotalsRow: () => t.totals,
    TableColumns: t.columns.map(c => ({ Name: c })),
    TableStyleInfo: t.style ? { Name: t.style.name, ShowRowStripes: t.style.rowStripes, ShowColumnStripes: t.style.columnStripes, ShowFirstColumn: t.style.firstColumn, ShowLastColumn: t.style.lastColumn, _sentinel: 'LEAK-STYLE' } : null,
    AutoFilter: t.autoFilter ? {} : null,
    _sentinel: 'LEAK-TABLE',
  });
  const Api = { GetSheet: n => (Object.prototype.hasOwnProperty.call(model, n) ? { worksheet: { get TableParts() { return model[n].map(part); } } } : null) };
  const find = name => { for (const s of names) for (const t of model[s]) if (t.name === name) return t; return null; };
  let activeIdx = active;
  const e = {
    canEdit: () => canEdit,
    collaborativeEditing: { getGlobalLock: () => lock },
    asc_isProtectedWorkbook: () => protectedWb,
    asc_getWorksheetsCount: () => names.length,
    asc_getWorksheetName: i => names[i],
    asc_showWorksheet: i => { log.shows.push(i); activeIdx = i; },
    asc_getActiveWorksheetIndex: () => activeIdx,
    asc_registerCallback: (id, fn) => { log.registered += 1; e._cb = fn; },
    asc_unregisterCallback: () => { log.unregistered += 1; },
    asc_changeAutoFilter(tn, opt, val) {
      log.calls.push(['asc_changeAutoFilter', tn, opt, val]);
      if (errorOnCall && e._cb) e._cb('stub-error', 'critical');
      if (fault === 'throw') throw new Error('stub: engine refused');
      if (fault === 'ignore') return undefined;
      const t = find(tn);
      if (opt === 2 && t.style) t.style.name = val;
      if (opt === 2 && !t.style) t.style = { name: val, rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false };
      if (fault === 'collateral') t.ref = 'A1:C6';
      if (fault === 'other') { const o = names.flatMap(s => model[s]).find(x => x !== t); if (o) o.style.name = 'TableStyleDark1'; }
      return undefined;
    },
    asc_changeFormatTableInfo(tn, opt, val) {
      log.calls.push(['asc_changeFormatTableInfo', tn, opt, val]);
      if (fault === 'throw') throw new Error('stub: engine refused');
      if (fault === 'ignore') return undefined;
      const t = find(tn);
      if (opt === 6) t.style.rowStripes = val;
      if (opt === 3) t.style.columnStripes = val;
      if (opt === 1) t.style.firstColumn = val;
      if (opt === 2) t.style.lastColumn = val;
      if (opt === 5 && val === true) {
        t.totals = true;
        const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(t.ref);
        t.ref = `${m[1]}${m[2]}:${m[3]}${Number(m[4]) + 1}`;
      }
      return undefined;
    },
    asc_changeDisplayNameTable(tn, nn) {
      log.calls.push(['asc_changeDisplayNameTable', tn, nn]);
      if (fault === 'throw') throw new Error('stub: engine refused');
      if (fault === 'ignore') return undefined;
      find(tn).name = nn;
      return undefined;
    },
    asc_changeTableRange(tn, r) {
      log.calls.push(['asc_changeTableRange', tn, r]);
      if (fault === 'throw') throw new Error('stub: engine refused');
      if (fault === 'ignore') return undefined;
      find(tn).ref = r;
      if (fault === 'columns-lost') find(tn).columns = ['Item'];
      return undefined;
    },
  };
  for (const k of omit) delete e[k];
  return { Api, Asc: Object.assign({ editor: e }, consts), model, log, editor: e };
}

function bridgeOf(w, block = V19) {
  new Function('Api', 'Asc', block.replace(/\}\s*,\s*$/, '}'))(w.Api, w.Asc);
  return w.Api.GetEuroOfficeBridge();
}
const writeOf = (w, args) => bridgeOf(w).write(CAP, Object.assign({ sheet: 'S1', table: 'Table1' }, args));
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);

// ---- the build --------------------------------------------------------------------------------------

test('the /19 block = the /18 block + exactly the builder\'s substitutions, and the builder is deterministic', () => {
  const r = spawnSync('python3', ['-I', '-c', 'import importlib.util,json,sys;s=importlib.util.spec_from_file_location("b",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);print(json.dumps({"entry":m.TABLE_WRITE_ENTRY,"impl":m.handler_text(),"sha":m.V18_BLOCK_SHA}))',
    path.join(DIR, 'rev19-candidate', 'build-rev19.py')], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stderr);
  const b = JSON.parse(r.stdout);
  assert.equal(sha(V18), b.sha, 'the builder pins the /18 candidate block');
  for (const frag of [b.entry, b.impl]) assert.equal(V19.split(frag).length - 1, 1, `fragment occurs once: ${frag.slice(0, 40)}`);
  const back = V19.split(b.entry).join('').split(b.impl).join('').replace('patchVersion: "euro-office-bridge-v1/9.3.5.0/19"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/18"');
  assert.equal(back, V18, 'nothing else moved');
  assert.equal(fs.readFileSync(path.join(DIR, 'rev19-candidate', 'old.txt'), 'utf8'), V18);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev19-candidate', 'new.txt'), 'utf8'), V19);
});

test('both blocks parse the same way (the fragment trailing comma stripped)', () => {
  for (const t of [V18, V19]) assert.doesNotThrow(() => new Function('Api', 'Asc', t.replace(/\}\s*,\s*$/, '}')));
});

test('the handler text enumerates no object and carries no line comment (build-time check, and re-asserted here)', () => {
  for (const w of ['Object.keys', 'Object.values', 'Object.entries', 'getOwnPropertyNames', 'for (var k in', 'for (var p in', ' in args']) assert.ok(!HANDLER.includes(w), w);
  assert.ok(!HANDLER.includes('//'));
});

test('the manifest advertises /19 and the capability as a TARGETED WRITE; the capability count grew by exactly one', () => {
  const m = bridgeOf(world()).manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.5.0/19');
  const c = m.capabilities.filter(x => x.id === CAP);
  assert.equal(c.length, 1);
  assert.equal(c[0].kind, 'WRITE');
  assert.equal(c[0].mutation, 'TARGETED');
  const m18 = new Function('Api', 'Asc', V18.replace(/\}\s*,\s*$/, '}') + '; return Api.GetEuroOfficeBridge().manifest();');
  const before = (() => { const w = world(); new Function('Api', 'Asc', V18.replace(/\}\s*,\s*$/, '}'))(w.Api, w.Asc); return w.Api.GetEuroOfficeBridge().manifest(); })();
  assert.equal(m.capabilities.length, before.capabilities.length + 1);
  assert.equal(typeof m18, 'function');
});

// ---- the five properties -------------------------------------------------------------------------------

test('style: the editor call takes (name, filter-style option 2, style) and the success is read back', () => {
  const w = world();
  const r = writeOf(w, { property: 'style', name: 'TableStyleMedium9' });
  assert.deepEqual(r, { ok: true, value: { applied: ['style'], table: 'Table1', observed: 'TableStyleMedium9' } });
  assert.deepEqual(w.log.calls, [['asc_changeAutoFilter', 'Table1', 2, 'TableStyleMedium9']]);
  assert.equal(w.model.S1[0].style.name, 'TableStyleMedium9');
});

test('flag: each of the four flags maps to its option and the value is read back', () => {
  for (const [flag, opt, key] of [['rowStripes', 6, 'rowStripes'], ['columnStripes', 3, 'columnStripes'], ['firstColumn', 1, 'firstColumn'], ['lastColumn', 2, 'lastColumn']]) {
    const w = world({ tables: { S1: [mkTable('Table1', 'A1:C5', { stripes: key === 'rowStripes' })] } });
    const value = key !== 'rowStripes';
    const r = writeOf(w, { property: 'flag', flag, value });
    assert.equal(r.ok, true, `${flag}: ${reason(r)}`);
    assert.deepEqual(w.log.calls, [['asc_changeFormatTableInfo', 'Table1', opt, value]], flag);
    assert.equal(w.model.S1[0].style[key], value);
  }
});

test('totalsRow ON: the engine inserts the row (ref grows by one), observed true; OFF is refused as not measured', () => {
  const w = world();
  const r = writeOf(w, { property: 'totalsRow', value: true });
  assert.deepEqual(r, { ok: true, value: { applied: ['totalsRow'], table: 'Table1', observed: true } });
  assert.deepEqual(w.log.calls, [['asc_changeFormatTableInfo', 'Table1', 5, true]]);
  assert.equal(w.model.S1[0].ref, 'A1:C6');
  const w2 = world({ tables: { S1: [mkTable('Table1', 'A1:C6', { totals: true })] } });
  assert.equal(reason(writeOf(w2, { property: 'totalsRow', value: false })), 'totals-off-not-measured');
  assert.deepEqual(w2.log.calls, []);
});

test('name: renamed, read back; the old name is gone', () => {
  const w = world();
  const r = writeOf(w, { property: 'name', name: 'Sales_2026' });
  assert.deepEqual(r, { ok: true, value: { applied: ['name'], table: 'Sales_2026', observed: 'Sales_2026' } });
  assert.deepEqual(w.log.calls, [['asc_changeDisplayNameTable', 'Table1', 'Sales_2026']]);
});

test('ref: a resize with the same top-left cell; $ and lower case are normalised before the call', () => {
  const w = world();
  const r = writeOf(w, { property: 'ref', ref: '$a$1:$c$7' });
  assert.deepEqual(r, { ok: true, value: { applied: ['ref'], table: 'Table1', observed: 'A1:C7' } });
  assert.deepEqual(w.log.calls, [['asc_changeTableRange', 'Table1', 'A1:C7']]);
});

// ---- the refusals BEFORE any engine call ----------------------------------------------------------------------

const PRE = [
  ['no args', undefined, 'bad-args'],
  ['empty sheet', { sheet: '', table: 'Table1', property: 'style', name: 'TableStyleMedium9' }, 'bad-args'],
  ['table not a string', { sheet: 'S1', table: 5, property: 'style', name: 'TableStyleMedium9' }, 'bad-args'],
  ['unknown property', { property: 'headerRow', value: false }, 'unsupported-property'],
  ['a foreign field for the property', { property: 'style', name: 'TableStyleMedium9', value: true }, 'unsupported-property'],
  ['style: missing name', { property: 'style' }, 'bad-args'],
  ['style: free string', { property: 'style', name: 'Pretty' }, 'style-unknown'],
  ['style: out of range', { property: 'style', name: 'TableStyleMedium29' }, 'style-unknown'],
  ['style: Light22', { property: 'style', name: 'TableStyleLight22' }, 'style-unknown'],
  ['flag: unknown flag', { property: 'flag', flag: 'headerRow', value: false }, 'flag-unknown'],
  ['flag: value coerced from a string', { property: 'flag', flag: 'rowStripes', value: 'false' }, 'value-not-boolean'],
  ['flag: value 0', { property: 'flag', flag: 'rowStripes', value: 0 }, 'value-not-boolean'],
  ['totals: value not boolean', { property: 'totalsRow', value: 1 }, 'value-not-boolean'],
  ['name: has a space', { property: 'name', name: 'My Table' }, 'name-invalid'],
  ['name: a cell reference', { property: 'name', name: 'AB12' }, 'name-invalid'],
  ['name: R1C1-like', { property: 'name', name: 'R' }, 'name-invalid'],
  ['name: starts with a digit', { property: 'name', name: '1abc' }, 'name-invalid'],
  ['name: empty', { property: 'name', name: '' }, 'name-invalid'],
  ['ref: unparsable', { property: 'ref', ref: 'A1' }, 'ref-unparsable'],
  ['ref: whole column', { property: 'ref', ref: 'A:C' }, 'ref-unparsable'],
  ['ref: not a string', { property: 'ref', ref: 7 }, 'bad-args'],
];
for (const [label, args, want] of PRE) {
  test(`refused before any engine call -- ${label} -> ${want}`, () => {
    const w = world();
    const b = bridgeOf(w);
    const a = args === undefined ? undefined : (args.sheet === undefined ? Object.assign({ sheet: 'S1', table: 'Table1' }, args) : args);
    assert.equal(reason(b.write(CAP, a)), want);
    assert.deepEqual(w.log.calls, []);
  });
}

// ---- the refusals from the model --------------------------------------------------------------------------------

test('table-not-found / sheet-not-found / table-ambiguous', () => {
  assert.equal(reason(writeOf(world(), { property: 'style', name: 'TableStyleMedium9', table: 'Nope' })), 'table-not-found');
  assert.equal(reason(writeOf(world(), { property: 'style', name: 'TableStyleMedium9', sheet: 'S9' })), 'sheet-not-found');
  const w = world({ tables: { S1: [mkTable('Table1'), mkTable('Table1', 'F1:G5')] } });
  assert.equal(reason(writeOf(w, { property: 'style', name: 'TableStyleMedium9' })), 'table-ambiguous');
  assert.deepEqual(w.log.calls, []);
});

test('a value already at the target is already-at-value with NO engine call (the product treats it as SATISFIED)', () => {
  for (const [args, tb] of [
    [{ property: 'style', name: 'TableStyleMedium2' }, mkTable('Table1')],
    [{ property: 'flag', flag: 'rowStripes', value: true }, mkTable('Table1')],
    [{ property: 'totalsRow', value: true }, mkTable('Table1', 'A1:C6', { totals: true })],
    [{ property: 'name', name: 'Table1' }, mkTable('Table1')],
    [{ property: 'ref', ref: 'A1:C5' }, mkTable('Table1')],
  ]) {
    const w = world({ tables: { S1: [tb] } });
    assert.equal(reason(writeOf(w, args)), 'already-at-value', args.property);
    assert.deepEqual(w.log.calls, [], args.property);
  }
});

test('flag on a table with no style info; totals without a header row', () => {
  assert.equal(reason(writeOf(world({ tables: { S1: [mkTable('Table1', 'A1:C5', { style: null })] } }), { property: 'flag', flag: 'rowStripes', value: false })), 'table-style-info-missing');
  assert.equal(reason(writeOf(world({ tables: { S1: [mkTable('Table1', 'A2:C5', { header: false })] } }), { property: 'totalsRow', value: true })), 'header-row-missing');
});

test('name-taken: case-insensitive, across EVERY sheet; the table\'s own case change is allowed', () => {
  const w = world({ tables: { S1: [mkTable('Table1')], S2: [mkTable('Sales', 'A1:B3')] } });
  assert.equal(reason(writeOf(w, { property: 'name', name: 'SALES' })), 'name-taken');
  assert.deepEqual(w.log.calls, []);
  const w2 = world();
  assert.equal(writeOf(w2, { property: 'name', name: 'TABLE1x' }).ok, true);
  const w3 = world();
  assert.equal(writeOf(w3, { property: 'name', name: 'table1' }).ok, true, 'a case-only rename of itself is not a collision');
});

test('ref: origin moves / too small / overlaps another table', () => {
  assert.equal(reason(writeOf(world(), { property: 'ref', ref: 'B1:C7' })), 'ref-moves-origin');
  assert.equal(reason(writeOf(world(), { property: 'ref', ref: 'A1:C1' })), 'ref-too-small');
  assert.equal(reason(writeOf(world(), { property: 'ref', ref: 'A1:A0' })), 'ref-unparsable');
  const w = world({ tables: { S1: [mkTable('Table1'), mkTable('Table2', 'A7:C9')] } });
  assert.equal(reason(writeOf(w, { property: 'ref', ref: 'A1:C8' })), 'ref-overlaps');
  assert.deepEqual(w.log.calls, []);
  assert.equal(writeOf(w, { property: 'ref', ref: 'A1:C6' }).ok, true, 'adjacent, not overlapping, is fine');
});

// ---- the editor preflight ------------------------------------------------------------------------------------

test('editor / API / constants / permission preflight, each a named refusal and no call', () => {
  const style = { property: 'style', name: 'TableStyleMedium9' };
  const w0 = world(); w0.Asc.editor = undefined;
  assert.equal(reason(writeOf(w0, style)), 'editor-missing');
  assert.equal(reason(writeOf(world({ omit: ['asc_changeAutoFilter'] }), style)), 'asc_changeAutoFilter-missing');
  assert.equal(reason(writeOf(world({ omit: ['asc_changeFormatTableInfo'] }), { property: 'flag', flag: 'rowStripes', value: false })), 'asc_changeFormatTableInfo-missing');
  assert.equal(reason(writeOf(world({ omit: ['asc_showWorksheet'] }), style)), 'asc_showWorksheet-missing');
  assert.equal(reason(writeOf(world({ consts: { c_oAscChangeTableStyleInfo: CONSTS.c_oAscChangeTableStyleInfo } }), style)), 'c_oAscChangeFilterOptions-missing');
  assert.equal(reason(writeOf(world({ consts: { c_oAscChangeFilterOptions: CONSTS.c_oAscChangeFilterOptions } }), { property: 'totalsRow', value: true })), 'c_oAscChangeTableStyleInfo-missing');
  assert.equal(reason(writeOf(world({ canEdit: false }), style)), 'edit-not-allowed');
  assert.equal(reason(writeOf(world({ lock: true }), style)), 'workbook-locked');
  assert.equal(reason(writeOf(world({ protectedWb: true }), style)), 'workbook-protected');
});

// ---- the verification after the call ---------------------------------------------------------------------------

test('an engine that ignores the call -> no-effect (never a reported success)', () => {
  for (const args of [{ property: 'style', name: 'TableStyleMedium9' }, { property: 'flag', flag: 'rowStripes', value: false }, { property: 'totalsRow', value: true }, { property: 'name', name: 'Sales' }, { property: 'ref', ref: 'A1:C7' }]) {
    const w = world({ fault: 'ignore' });
    assert.equal(reason(writeOf(w, args)), 'no-effect', args.property);
    assert.equal(w.log.calls.length, 1);
  }
});

test('an engine that throws -> <call>-threw; an editor error callback -> <call>-error', () => {
  assert.equal(reason(writeOf(world({ fault: 'throw' }), { property: 'style', name: 'TableStyleMedium9' })), 'asc_changeAutoFilter-threw');
  assert.equal(reason(writeOf(world({ fault: 'throw' }), { property: 'name', name: 'Sales' })), 'asc_changeDisplayNameTable-threw');
  const w = world({ errorOnCall: true });
  const r = writeOf(w, { property: 'style', name: 'TableStyleMedium9' });
  assert.equal(r.reason, 'asc_changeAutoFilter-error');
  assert.deepEqual(r.errors, [{ id: 'stub-error', level: 'critical' }]);
  assert.equal(w.log.registered, w.log.unregistered, 'the callback is always unregistered');
});

test('a call that changes ANOTHER table is named, a call that changes ANOTHER field of the target is named', () => {
  const w = world({ tables: { S1: [mkTable('Table1'), mkTable('Table2', 'F1:G5')] }, fault: 'other' });
  assert.equal(reason(writeOf(w, { property: 'style', name: 'TableStyleMedium9' })), 'other-table-changed');
  const w2 = world({ fault: 'collateral' });
  assert.equal(reason(writeOf(w2, { property: 'style', name: 'TableStyleMedium9' })), 'target-collateral-change');
  const w3 = world({ fault: 'columns-lost' });
  const r3 = writeOf(w3, { property: 'ref', ref: 'A1:C7' });
  assert.equal(r3.ok, true, 'a resize may legitimately change the column list (the product refuses a width change)');
});

test('the active sheet is shown for the call and restored afterwards; the calls are made once each', () => {
  const w = world({ tables: { S0: [mkTable('Other', 'A1:B3')], S1: [mkTable('Table1')] }, active: 0 });
  const r = writeOf(w, { property: 'style', name: 'TableStyleMedium9' });
  assert.equal(r.ok, true);
  assert.deepEqual(w.log.shows, [1, 0]);
  assert.equal(w.log.calls.length, 1);
});

// ---- the owner conditions ---------------------------------------------------------------------------------------

test('no unnamed field ever reaches a result: sentinels planted on the table and style objects never appear', () => {
  const outs = [
    writeOf(world(), { property: 'style', name: 'TableStyleMedium9' }),
    writeOf(world({ fault: 'collateral' }), { property: 'style', name: 'TableStyleMedium9' }),
    writeOf(world({ fault: 'ignore' }), { property: 'name', name: 'Sales' }),
  ];
  for (const o of outs) assert.ok(!JSON.stringify(o).includes('LEAK'), JSON.stringify(o));
});

test('an exception inside the handler becomes internal-threw through write(), not a thrown error', () => {
  const w = world();
  w.Asc.editor.canEdit = () => { throw new Error('boom'); };
  const r = writeOf(w, { property: 'style', name: 'TableStyleMedium9' });
  assert.deepEqual(r, { ok: false, reason: 'internal-threw' });
});

test('the capability is served only through write(); read() and an unknown id are not advertised', () => {
  const b = bridgeOf(world());
  assert.equal(b.read(CAP, { sheet: 'S1' }).reason, 'capability-not-advertised');
  assert.equal(b.write('sheet.table.nope/v1', {}).reason, 'capability-not-advertised');
});

test('helpers are self-consistent', () => {
  assert.equal(colName(colNum('AB')), 'AB');
});
