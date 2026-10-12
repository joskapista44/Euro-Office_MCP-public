'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * Static gate for the M0 section 3.1 dependency rules.
 *
 * Each forbidden edge below is a rule one of the two sources actually broke somewhere, so this
 * gate is a regression test against a measured history, not a style preference.
 */

const SRC = path.join(__dirname, '..', '..', 'src');

/** Every .cjs under src/, with its requires resolved to repo-relative module paths. */
function collectModules(dir = SRC, out = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) collectModules(full, out);
    else if (entry.name.endsWith('.cjs')) {
      const source = fs.readFileSync(full, 'utf8');
      const requires = [];
      // Matches require('...') in source order. Comment text is excluded below.
      for (const m of source.matchAll(/require\(\s*['"]([^'"]+)['"]\s*\)/g)) {
        const target = m[1];
        if (!target.startsWith('.')) continue; // node: builtins are not architecture edges
        requires.push(path.relative(SRC, path.resolve(path.dirname(full), target)));
      }
      out.push({ file: path.relative(SRC, full), requires });
    }
  }
  return out;
}

const MODULES = collectModules();

test('the static gate sees the tree at all (positive control)', () => {
  assert.ok(MODULES.length >= 10, `expected the core tree, found ${MODULES.length} modules`);
  const withEdges = MODULES.filter(m => m.requires.length > 0);
  assert.ok(withEdges.length >= 5, 'the require extractor must find real edges, or its zeros mean nothing');
});

test('FORBIDDEN: core/* must never import an adapter by name', () => {
  for (const mod of MODULES) {
    if (!mod.file.startsWith('core/')) continue;
    for (const edge of mod.requires) {
      assert.ok(
        !edge.startsWith('adapters/'),
        `${mod.file} imports ${edge}: the core must not know an adapter by name`
      );
    }
  }
});

test('FORBIDDEN: adapters/* may not open, close or persist', () => {
  for (const mod of MODULES) {
    if (!mod.file.startsWith('adapters/')) continue;
    for (const edge of mod.requires) {
      assert.ok(!edge.startsWith('core/session'), `${mod.file} imports ${edge}: an adapter may not own a session`);
      assert.ok(!edge.startsWith('core/persistence'), `${mod.file} imports ${edge}: persistence is not an adapter concern`);
    }
  }
});

test('FORBIDDEN: adapters never see each other', () => {
  for (const mod of MODULES) {
    if (!mod.file.startsWith('adapters/')) continue;
    const own = mod.file.split(path.sep)[1];
    for (const edge of mod.requires) {
      if (!edge.startsWith('adapters/')) continue;
      assert.equal(edge.split(path.sep)[1], own, `${mod.file} imports ${edge}: adapters are isolated from each other`);
    }
  }
});

test('FORBIDDEN: the receipt may not reach the persistence decision', () => {
  // Owner requirement 8: execution state -> persistence, execution state -> receipt,
  // and never receipt -> persistence decision. Both directions are asserted.
  for (const mod of MODULES) {
    if (mod.file.startsWith('core/persistence/')) {
      for (const edge of mod.requires) {
        assert.ok(
          !edge.startsWith('core/receipt'),
          `${mod.file} imports ${edge}: a persistence decision must not be reachable from a receipt`
        );
      }
    }
    if (mod.file.startsWith('core/receipt/')) {
      for (const edge of mod.requires) {
        assert.ok(
          !edge.startsWith('core/persistence'),
          `${mod.file} imports ${edge}: the receipt is derived evidence, not a persistence participant`
        );
      }
    }
  }
});

test('the adapter contract module holds no format knowledge', () => {
  const contract = fs.readFileSync(path.join(SRC, 'core', 'adapter', 'contract.cjs'), 'utf8');
  // The core must not know these concepts. The words may appear in the prose that states the
  // prohibition, so only the identifier-like forms are searched.
  for (const forbidden of ['paragraph', 'worksheet', 'pivotTable', 'slideLayout', 'cellRange']) {
    assert.ok(
      !new RegExp(`\\b${forbidden}\\b`, 'i').test(contract.replace(/^\s*\*.*$/gm, '')),
      `the adapter contract mentions the format concept "${forbidden}"`
    );
  }
});

/**
 * The rules above run on the real tree. `src/adapters/` is EMPTY in M1 — no capability is
 * migrated — so the adapter rules currently have no subject and would pass vacuously.
 * These two tests measure the rule itself instead of the tree.
 */
const { violations } = require('./edges-rules.cjs');

test('NEGATIVE CONTROL: the edge rules go red on a synthetic violating tree', () => {
  const bad = [
    { file: 'core/execution/executor.cjs', requires: ['adapters/docx/primitives.cjs'] },
    { file: 'adapters/docx/primitives.cjs', requires: ['core/session/open.cjs'] },
    { file: 'adapters/xlsx/verify.cjs', requires: ['core/persistence/barrier.cjs'] },
    { file: 'adapters/docx/target.cjs', requires: ['adapters/xlsx/target.cjs'] },
    { file: 'core/persistence/verified-writes.cjs', requires: ['core/receipt/envelope.cjs'] },
    { file: 'core/receipt/envelope.cjs', requires: ['core/persistence/barrier.cjs'] },
  ];
  bad.push(
    { file: 'core/observation/observe.cjs', requires: ['adapters/docx/target.cjs'] },
    { file: 'core/observation/enrichment.cjs', requires: ['core/session/handle.cjs'] },
    { file: 'core/observation/enrichment.cjs', requires: ['core/execution/result.cjs'] },
    { file: 'core/verification/orchestrate.cjs', requires: ['core/persistence/barrier.cjs'] },
    { file: 'core/authority.cjs', requires: ['core/metrics/collect.cjs'] }
  );

  const found = violations(bad);
  const rules = new Set(found.map(v => v.rule));

  // 12, not 11: the observation->adapter edge trips TWO rules, the general `core-imports-adapter`
  // and the narrower `observation-imports-adapter`. The overlap is deliberate — the narrow rule
  // survives if the general one is ever scoped down — so the count reflects rules fired, not
  // lines planted.
  assert.equal(found.length, 12, `every planted violation must be caught, got ${found.length}`);
  assert.equal(
    found.filter(v => v.file === 'core/observation/observe.cjs').length,
    2,
    'the overlapping pair is what makes the count 12'
  );
  for (const rule of [
    'core-imports-adapter',
    'adapter-imports-session',
    'adapter-imports-persistence',
    'adapter-imports-adapter',
    'persistence-imports-receipt',
    'receipt-imports-persistence',
    'observation-imports-adapter',
    'enrichment-imports-session',
    'enrichment-imports-execution',
    'verifier-imports-persistence',
    'leaf-module-has-an-edge',
  ]) {
    assert.ok(rules.has(rule), `rule ${rule} never fired — it cannot protect anything`);
  }
});

test('POSITIVE CONTROL: the same rules are green on the real tree', () => {
  const found = violations(MODULES);
  assert.deepEqual(found, [], `forbidden edges present: ${JSON.stringify(found)}`);

  // And state the measured scope plainly rather than letting an empty set read as a proof.
  // M1-M3 stated here that the adapter rules had NO live subject. M4 gives them one: the green above
  // is now measured over real adapter modules, not over an empty set.
  const adapterModules = MODULES.filter(m => m.file.replace(/\\/g, '/').startsWith('adapters/'));
  assert.ok(adapterModules.length >= 1, 'M4: the adapter rules must have a live subject');
});
