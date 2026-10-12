'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * M6 XLSX adapter boundary (the M4 DOCX gate, same rules, on src/adapters/xlsx).
 *
 *   XLSX_ADAPTER_SESSION_OPEN=NONE   XLSX_ADAPTER_SESSION_CLOSE=NONE   XLSX_ADAPTER_PERSISTENCE=NONE
 *   XLSX_MUTATION_PATH=COMMON_DISPATCH
 *
 * The strongest form is structural: the adapter imports NOTHING outside src/adapters/xlsx (plus
 * node builtins), so it has no route to the session, the persistence barrier, the driver or a
 * browser. Its only way to mutate is the `dispatch` capability handed to executePrimitive, and its
 * only way to read is the observation it is handed. Each rule is also run on a synthetic violation.
 */

const DIR = path.join(__dirname, '..', '..', 'src', 'adapters', 'xlsx');

function stripComments(source) {
  return source.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/(^|[^:])\/\/.*$/gm, '$1');
}

function files(dir = DIR, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) files(full, out);
    else if (e.name.endsWith('.cjs')) out.push(full);
  }
  return out;
}

const MODULES = files().map(full => ({
  file: path.relative(DIR, full).replace(/\\/g, '/'),
  full,
  code: stripComments(fs.readFileSync(full, 'utf8')),
}));

// The generated editor command is DATA (a string the driver executes), not adapter code. Its API
// names (Delete, AddText, ...) are therefore excluded from the token rules by blanking string
// literals first; the rules then see only what the adapter itself calls.
const codeOnly = code => code.replace(/`(?:\\[\s\S]|[^`\\])*`/g, '``').replace(/'(?:\\.|[^'\\\n])*'/g, "''").replace(/"(?:\\.|[^"\\\n])*"/g, '""');

const RULES = {
  'imports-outside-the-adapter': ({ full, code }) => [...code.matchAll(/require\(\s*['"]([^'"]+)['"]\s*\)/g)].map(m => m[1]).some(r => {
    if (r.startsWith('node:')) return false;
    if (!r.startsWith('.')) return true; // a package (e.g. playwright) is a transport route
    return !path.resolve(path.dirname(full), r).startsWith(DIR);
  }),
  'session-open-or-close': ({ code }) => /\b(openSession|closeSession|withSession|awaitReady|requestClose)\s*\(|\.(open|close)\s*\(/.test(codeOnly(code)),
  'persistence-or-save': ({ code }) => /\b(persist|forceSave|canSave|forceSaveBarrier)\b/.test(codeOnly(code)),
  'transport-backdoor': ({ code }) => /\b(evaluate|callCommand|readLive|readPackage|fetch|frames?|page)\s*\(|\.frame\b|\.page\b/.test(codeOnly(code)),
  'receipt-as-authority': ({ code }) => /receiptEvidence/.test(codeOnly(code)),
};

test('positive control: the gate sees the XLSX adapter tree', () => {
  const names = MODULES.map(m => m.file).sort();
  for (const f of ['adapter.cjs', 'inventory.cjs', 'a1.cjs', 'families/cells.cjs']) assert.ok(names.includes(f), f);
});

for (const [name, rule] of Object.entries(RULES)) {
  test(`XLSX adapter: ${name} -- none`, () => {
    assert.deepEqual(MODULES.filter(rule).map(m => m.file), []);
  });
}

test('XLSX_MUTATION_PATH=COMMON_DISPATCH: exactly one dispatch call, the capability handed to executePrimitive', () => {
  const calls = MODULES.flatMap(m => [...codeOnly(m.code).matchAll(/\bdispatch\s*\(/g)].map(() => m.file));
  assert.deepEqual(calls, ['adapter.cjs']);
  const adapter = MODULES.find(m => m.file === 'adapter.cjs').code;
  assert.match(adapter, /executePrimitive\s*\([^)]*\{\s*dispatch,\s*target\s*\}\s*\)[\s\S]{0,400}return dispatch\(/);
});

test('NEGATIVE CONTROL: every rule fires on its own violation', () => {
  const fake = code => ({ file: 'families/x.cjs', full: path.join(DIR, 'families', 'x.cjs'), code });
  assert.equal(RULES['imports-outside-the-adapter'](fake("const s = require('../../../core/session/handle.cjs');")), true);
  assert.equal(RULES['imports-outside-the-adapter'](fake("const pw = require('playwright');")), true);
  assert.equal(RULES['session-open-or-close'](fake('await closeSession(h);')), true);
  assert.equal(RULES['persistence-or-save'](fake('await driver.forceSave();')), true);
  assert.equal(RULES['transport-backdoor'](fake('await handle.frame.evaluate(() => 1);')), true);
  assert.equal(RULES['receipt-as-authority'](fake('if (prior.receiptEvidence) skip();')), true);
  // and the string-literal blanking does not hide a real call next to a string
  assert.equal(RULES['persistence-or-save'](fake("const x = 'forceSave'; forceSave();")), true);
});
