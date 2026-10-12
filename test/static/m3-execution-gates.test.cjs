'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { validateAdapter } = require('../../src/core/adapter/contract.cjs');
const { createDocumentAdapter } = require('../fakes/fake-document.cjs');

/**
 * M3 static gates — owner brief point 26.
 *
 * Each rule is a pure predicate over (file, code), run first on the REAL tree and then on a
 * synthetic source that violates it. A rule that cannot go red on its own violation is not a gate.
 * Comments are stripped before matching, because the prose in these modules states the very
 * prohibitions the gates look for (`detector-matches-itself`).
 */

const EXEC = path.join(__dirname, '..', '..', 'src', 'core', 'execution');

function stripComments(source) {
  return source.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/(^|[^:])\/\/.*$/gm, '$1');
}

const FILES = fs.readdirSync(EXEC).filter(f => f.endsWith('.cjs'))
  .map(f => ({ file: f, code: stripComments(fs.readFileSync(path.join(EXEC, f), 'utf8')) }));

const requiresOf = code => [...code.matchAll(/require\(\s*['"]([^'"]+)['"]\s*\)/g)].map(m => m[1]);

const RULES = {
  'executor-opens-or-closes-a-session': ({ code }) =>
    requiresOf(code).some(r => /session\/(open|lifecycle|internal)\.cjs$/.test(r)) ||
    /\b(openSession|closeSession|awaitReady|withSession)\s*\(/.test(code),

  'executor-persists': ({ code }) =>
    requiresOf(code).some(r => /persistence\//.test(r)) ||
    /\b(persist|forceSave|canSave)\s*\(/.test(code),

  'strategy-reaches-transport': ({ file, code }) =>
    file.startsWith('strategy-') &&
    (requiresOf(code).some(r => r !== './plan.cjs') ||
      /\b(dispatch|executePrimitive|executePrimitiveBatch|readLive|observeFresh|verifyPrimitive)\b/.test(code)),

  'receipt-read-as-authority': ({ code }) =>
    // Building a receipt is allowed; READING one (a property access on receiptEvidence) is not.
    /receiptEvidence\s*(\.|\[)/.test(code) || /\.receiptEvidence\b/.test(code),

  'fixed-correctness-delay': ({ code }) => /\b(setTimeout|setInterval|sleep|delay|setImmediate)\s*\(/.test(code),

  // The only source of groups is the strategy's own grouping, and there is exactly one place a
  // group is dispatched -- so there is no second, faster path that could skip a barrier.
  'batch-fast-path': ({ file, code }) =>
    file === 'executor.cjs' &&
    ((code.match(/\.groups\(/g) || []).length !== 1 || (code.match(/await dispatchQueue\(/g) || []).length !== 1),

  'batch-strategy-not-dependency-cut': ({ file, code }) =>
    file === 'strategy-batch.cjs' && !/groupIntoBatches\(/.test(code),

  // R1: grouping must consult the explicit independence proof; unknown independence is not batched.
  'batch-grouping-ignores-independence': ({ file, code }) => {
    if (file !== 'plan.cjs') return false;
    const fn = code.slice(code.indexOf('function groupIntoBatches'), code.indexOf('function singletonGroups'));
    return !/batchIndependence\(/.test(fn) || !/PROVEN_BATCH_SAFE/.test(fn);
  },

  // The batch hook receives only the guarded dispatch (one physical dispatch per round).
  'batch-hook-unguarded': ({ file, code }) =>
    file === 'executor.cjs' && !/executePrimitiveBatch\([\s\S]{0,200}\{\s*dispatch:\s*guardedDispatch\s*\}/.test(code),
};

test('positive control: the gates see the execution tree', () => {
  const names = FILES.map(f => f.file).sort();
  for (const expected of ['executor.cjs', 'plan.cjs', 'strategy-batch.cjs', 'strategy-sequential.cjs', 'strategy-transaction.cjs', 'uncertainty.cjs', 'result.cjs']) {
    assert.ok(names.includes(expected), `missing ${expected}`);
  }
  const executor = FILES.find(f => f.file === 'executor.cjs');
  assert.ok(requiresOf(executor.code).length >= 8, 'the require extractor must find the executor edges');
});

for (const [name, rule] of Object.entries(RULES)) {
  test(`M3 gate: ${name} -- clean on the real tree`, () => {
    const hits = FILES.filter(rule).map(f => f.file);
    assert.deepEqual(hits, [], `${name}: ${hits.join(', ')}`);
  });
}

test('NEGATIVE CONTROL: every M3 gate fires on its own violation', () => {
  const synthetic = {
    'executor-opens-or-closes-a-session': { file: 'executor.cjs', code: "const { closeSession } = require('../session/open.cjs'); closeSession(h);" },
    'executor-persists': { file: 'executor.cjs', code: "const { persist } = require('../persistence/barrier.cjs'); await persist(h);" },
    'strategy-reaches-transport': { file: 'strategy-batch.cjs', code: "const { dispatch } = require('../session/handle.cjs'); dispatch(h, c);" },
    'receipt-read-as-authority': { file: 'executor.cjs', code: 'if (prior.receiptEvidence.operations.length) skip();' },
    'fixed-correctness-delay': { file: 'executor.cjs', code: 'await new Promise(r => setTimeout(r, 500));' },
    'batch-fast-path': { file: 'executor.cjs', code: 'const g = strat.groups(ops); await dispatchQueue(handle, a, q); await dispatchQueue(handle, a, all);' },
    'batch-strategy-not-dependency-cut': { file: 'strategy-batch.cjs', code: 'module.exports = { groups: ops => ({ groups: [ops] }) };' },
    'batch-grouping-ignores-independence': { file: 'plan.cjs', code: 'function groupIntoBatches(ops) { return { groups: [ops] }; } function singletonGroups() {}' },
    'batch-hook-unguarded': { file: 'executor.cjs', code: 'await adapter.executePrimitiveBatch(handle, items, { dispatch: dispatchFn });' },
  };
  for (const [name, fixture] of Object.entries(synthetic)) {
    assert.equal(RULES[name](fixture), true, `${name} did not fire on its own violation`);
  }
});

test('adapter does not own lifecycle: the M3 document adapter passes the contract, a lifecycle-owning one fails', () => {
  assert.deepEqual(validateAdapter(createDocumentAdapter()), { ok: true, problems: [] });
  const owning = { ...createDocumentAdapter(), close: async () => {}, forceSave: async () => {} };
  const v = validateAdapter(owning);
  assert.equal(v.ok, false);
  assert.ok(v.problems.some(p => /close/.test(p)) && v.problems.some(p => /forceSave/.test(p)));
});

test('common executor is format-agnostic: one executor, no per-format executor module', () => {
  const executors = FILES.filter(f => /executor/i.test(f.file)).map(f => f.file);
  assert.deepEqual(executors, ['executor.cjs']);
  // NEGATIVE CONTROL: the same filter would catch a per-format executor next to it.
  assert.equal(['executor.cjs', 'docx-executor.cjs'].filter(f => /executor/i.test(f)).length, 2);
});
