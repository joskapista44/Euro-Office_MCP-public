'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { INTENTS, STATUS } = require('../../src/adapters/docx/intents.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { PROPERTY_QUERY } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * M4 closure gates (owner W5 points 40-42), computed from the code, not from documentation:
 *
 *   DOCX_REGISTERED_INTENTS=28  DOCX_MIGRATED_FULL=28  MISSING=[]  PARTIAL=[]  EXTRA=[]
 *   DOCX_LEGACY_EXECUTION_PATH=NONE
 *   FRAME_EVALUATE_MUTATION=NONE  DIRECT_DRIVER_MUTATION_OUTSIDE_COMMON_DISPATCH=NONE
 *   (ADAPTER_OWNS_PERSISTENCE / ADAPTER_OPENS_SESSION / ADAPTER_CLOSES_SESSION = NO: the structural
 *   form is test/static/docx-adapter-boundary.test.cjs; this file does not repeat it.)
 *
 * Each measuring rule is also run on a synthetic violation (NEGATIVE CONTROL).
 */

const ROOT = path.join(__dirname, '..', '..');
const snapshot = JSON.parse(fs.readFileSync(path.join(ROOT, 'test', 'fixtures', 'docx-source-registry.json'), 'utf8'));
const strip = s => s.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/(^|[^:])\/\/.*$/gm, '$1');
// string literals blanked: generated editor commands and notes are DATA, not calls
const codeOnly = code => code.replace(/`(?:\\[\s\S]|[^`\\])*`/g, '``').replace(/'(?:\\.|[^'\\\n])*'/g, "''").replace(/"(?:\\.|[^"\\\n])*"/g, '""');

function tree(dir, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) tree(full, out);
    else if (e.name.endsWith('.cjs')) out.push(full);
  }
  return out;
}
const SRC = tree(path.join(ROOT, 'src')).map(full => ({ file: path.relative(ROOT, full).replace(/\\/g, '/'), code: strip(fs.readFileSync(full, 'utf8')) }));

/** FULL = migrated (a status other than NOT_MIGRATED), not partial, no LIVE pending. */
function accounting(intents, registry) {
  const reg = registry.map(i => i.intent);
  const mine = intents.map(r => r.intent);
  const full = intents.filter(r => r.status !== STATUS.NOT_MIGRATED && r.partial !== true && r.pendingLive !== true).map(r => r.intent);
  return {
    registered: reg.length,
    migratedFull: full.length,
    missing: reg.filter(i => !mine.includes(i)),
    extra: mine.filter(i => !reg.includes(i)),
    partial: intents.filter(r => r.partial === true || r.pendingLive === true || r.status === STATUS.NOT_MIGRATED).map(r => r.intent),
  };
}

test('DOCX_REGISTERED_INTENTS=28, DOCX_MIGRATED_FULL=28, MISSING=[], PARTIAL=[], EXTRA=[] (from the registry snapshot + intents.cjs)', () => {
  const a = accounting(INTENTS, snapshot.intents);
  assert.deepEqual(a, { registered: 28, migratedFull: 28, missing: [], extra: [], partial: [] });
});

test('NEGATIVE CONTROL: the accounting sees a pending, a missing and an extra intent', () => {
  const pending = INTENTS.map((r, i) => (i === 0 ? { ...r, pendingLive: true } : r));
  assert.equal(accounting(pending, snapshot.intents).migratedFull, accounting(INTENTS, snapshot.intents).migratedFull - 1);
  assert.deepEqual(accounting(INTENTS.slice(1), snapshot.intents).missing, [INTENTS[0].intent]);
  assert.deepEqual(accounting([...INTENTS, { intent: 'invented_intent', status: 'M4_WAVE5' }], snapshot.intents).extra, ['invented_intent']);
});

// ---- DOCX_LEGACY_EXECUTION_PATH=NONE --------------------------------------------------------------

// The SOURCE repository's execution modules (joskapista44/euro-mcp): none may be reachable from src/.
const LEGACY = /require\(\s*['"][^'"]*(docx-agent-|docx-session-|coedit-|lib-operations|\/lib\.cjs|euro-mcp)[^'"]*['"]\s*\)/;
const outsideRequire = ({ file, code }) => [...code.matchAll(/require\(\s*['"]([^'"]+)['"]\s*\)/g)].map(m => m[1])
  .filter(r => !r.startsWith('node:') && !(r.startsWith('.') && path.resolve(ROOT, path.dirname(file), r).startsWith(path.join(ROOT, 'src'))));

test('DOCX_LEGACY_EXECUTION_PATH=NONE: src/ requires no source-repo module and nothing outside src/ (node builtins aside)', () => {
  assert.ok(SRC.length >= 40, 'positive control: the gate sees src/');
  assert.deepEqual(SRC.filter(m => LEGACY.test(m.code)).map(m => m.file), []);
  assert.deepEqual(SRC.map(m => [m.file, outsideRequire(m)]).filter(([, r]) => r.length), []);
});

test('DOCX_LEGACY_EXECUTION_PATH=NONE: every registered intent plans through the M4 adapter into its declared family module', () => {
  const adapterSrc = fs.readFileSync(path.join(ROOT, 'src', 'adapters', 'docx', 'adapter.cjs'), 'utf8');
  for (const row of INTENTS) {
    const planned = planOperations({ steps: [{ intent: row.intent, ...STEP[row.intent] }] });
    assert.equal(planned.ok, true, `${row.intent}: ${planned.detail}`);
    assert.ok(fs.existsSync(path.join(ROOT, row.targetModule)), `${row.intent}: ${row.targetModule}`);
    const mod = `./families/${path.basename(row.targetModule)}`;
    assert.ok(adapterSrc.includes(`require('${mod}')`), `${row.intent}: the adapter wires ${mod}`);
  }
});

test('NEGATIVE CONTROL: the legacy rule fires on a source-repo require and on an outside require', () => {
  assert.equal(LEGACY.test("const x = require('../../zero-euro-col-wt/docx-agent-executor.cjs');"), true);
  assert.equal(LEGACY.test("const s = require('./docx-session-tables.cjs');"), true);
  assert.deepEqual(outsideRequire({ file: 'src/core/x.cjs', code: "require('../../scripts/y.cjs')" }), ['../../scripts/y.cjs']);
});

// ---- FRAME_EVALUATE_MUTATION=NONE, DIRECT_DRIVER_MUTATION_OUTSIDE_COMMON_DISPATCH=NONE --------------

const DRIVER = SRC.find(m => m.file === 'src/drivers/eurooffice/docx-driver.cjs').code;

test('FRAME_EVALUATE_MUTATION=NONE: the driver evaluates in exactly three places -- the read/lifecycle helper, the callCommand runner, requestClose', () => {
  const sites = [...DRIVER.matchAll(/\.evaluate\s*\(/g)].length;
  assert.equal(sites, 3);
  assert.match(DRIVER, /const inEditor = fn => state\.frame\.evaluate\(fn\);/);
  assert.match(DRIVER, /async function runInEditor\(body\) \{\s*return state\.frame\.evaluate\(/);
  assert.match(DRIVER, /lifecycle = await state\.page\.evaluate\(async \(\) => \{\s*const d = window\.__docEditor;/);
  // callCommand -- the ONLY editor mutation route -- appears only inside the runner
  assert.equal([...DRIVER.matchAll(/callCommand\s*\(/g)].length, 1);
  // every inEditor body calls only readers, forceSave (the M1 barrier) and the readiness projection
  const bodies = [...DRIVER.matchAll(/inEditor\(\(\) => \{([\s\S]*?)\}\)/g)].map(m => m[1]);
  assert.ok(bodies.length >= 3, 'positive control: the gate sees the inEditor bodies');
  for (const b of bodies) for (const [, name] of b.matchAll(/\b(?:e|d)\.(\w+)\s*\(/g)) assert.match(name, /^(pluginMethod_GetFileHTML|forceSave|asc_isDocumentCanSave)$/, `inEditor calls ${name}`);
});

test('DIRECT_DRIVER_MUTATION_OUTSIDE_COMMON_DISPATCH=NONE: driver.dispatch is reached only through the core session handle; the executor hands the adapter a guarded capability', () => {
  const callers = SRC.filter(m => /\.dispatch\s*\(/.test(codeOnly(m.code))).map(m => m.file);
  assert.deepEqual(callers, ['src/core/session/handle.cjs']);
  const adapterCalls = SRC.filter(m => m.file.startsWith('src/adapters/') && /\bdispatch\s*\(/.test(codeOnly(m.code))).map(m => m.file);
  assert.deepEqual(adapterCalls, ['src/adapters/docx/adapter.cjs', 'src/adapters/xlsx/adapter.cjs'], 'each adapter only calls the capability it is handed');
});

test('the driver READ never creates a header/footer part: every GetHeader/GetFooter in the property query passes create=false', () => {
  const calls = [...PROPERTY_QUERY.matchAll(/\.(GetHeader|GetFooter)\(([^)]*)\)/g)];
  assert.ok(calls.length >= 2, 'positive control: the gate sees the section read');
  for (const [, name, args] of calls) assert.match(args, /,\s*false\s*$/, `${name}(${args})`);
});

test('NEGATIVE CONTROL: the evaluate / read rules fire on violations', () => {
  const extra = `${DRIVER}\nstate.frame.evaluate(() => window.editor.asc_AddText('x'));`;
  assert.equal([...extra.matchAll(/\.evaluate\s*\(/g)].length, 4);
  assert.deepEqual([{ file: 'src/core/x.cjs', code: 'driver.dispatch(h, c);' }].filter(m => /\.dispatch\s*\(/.test(codeOnly(m.code))).map(m => m.file), ['src/core/x.cjs']);
  assert.equal(/\.dispatch\s*\(/.test(codeOnly("const note = 'before x.dispatch(';")), false);
  assert.doesNotMatch('sec.GetHeader(v, true)', /,\s*false\s*\)$/);
});

// one valid step per intent (the adapter's plan check, not a live run)
const STEP = {
  rewrite_paragraph: { selector: { text: 'x' }, text: 'y' },
  format_text: { selector: { text: 'x' }, format: { bold: true } },
  insert_paragraph_before: { selector: { text: 'x' }, text: 'y' },
  insert_paragraph_after: { selector: { text: 'x' }, text: 'y' },
  delete_paragraph: { selector: { text: 'x' } },
  move_paragraph_before: { selector: { text: 'x' }, destinationSelector: { text: 'y' } },
  move_paragraph_after: { selector: { text: 'x' }, destinationSelector: { text: 'y' } },
  apply_list: { selector: { text: 'x' }, type: 'bullet', level: 0 },
  set_list_level: { selector: { text: 'x' }, level: 1 },
  indent_list_item: { selector: { text: 'x' } },
  set_list_type: { selector: { text: 'x' }, type: 'numbered' },
  outdent_list_item: { selector: { text: 'x' } },
  remove_list: { selector: { text: 'x' } },
  add_bookmark: { selector: { text: 'x' }, name: 'b_1' },
  add_link: { selector: { text: 'x' }, url: 'https://example.invalid/x' },
  add_footnote: { selector: { text: 'x' }, text: 'n' },
  add_endnote: { selector: { text: 'x' }, text: 'n' },
  replace_table_cell: { selector: { text: 'x' }, text: 'y' },
  insert_table_row_before: { selector: { text: 'x' }, terminalRowCount: 4 },
  insert_table_row_after: { selector: { text: 'x' }, terminalRowCount: 4 },
  delete_table_row: { selector: { tableCell: { text: 'x', tableAnchor: 'a' } }, terminalRowCount: 2 },
  insert_table_column_before: { selector: { text: 'x' }, terminalColumnCount: 4 },
  insert_table_column_after: { selector: { text: 'x' }, terminalColumnCount: 4 },
  delete_table_column: { selector: { tableCell: { text: 'x', tableAnchor: 'a' } }, terminalColumnCount: 2 },
  merge_table_cells: { selector: { tableCell: { text: 'x', tableAnchor: 'a' } }, endSelector: { text: 'y' }, terminalMergedText: 'x y', terminalMergedSpan: { rows: 1, cols: 2 } },
  split_table_cell: { selector: { tableCell: { text: 'x', tableAnchor: 'a' } }, splitRows: 1, splitCols: 2 },
  edit_header: { section: { index: 0, of: 1 }, text: 'h' },
  edit_footer: { section: { index: 0, of: 1 }, text: 'f' },
};
