'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * Static gate: the core knows no format concept — owner requirement 4, PROJECT section 5.
 *
 * The prose in the core STATES the prohibition ("nothing here names a paragraph, a list, ...").
 * A naive grep therefore matches the rule's own wording and reports the sentence that forbids the
 * thing as an instance of it — our named `detector-matches-itself` shape. The gate strips comments
 * first, so it measures CODE.
 */

const SRC = path.join(__dirname, '..', '..', 'src');

const FORBIDDEN_CONCEPTS = Object.freeze([
  'paragraph', 'heading', 'bookmark', 'footnote',
  'worksheet', 'spreadsheet', 'cell', 'pivot', 'chart', 'formula',
  'slide', 'slideLayout',
]);

function stripComments(source) {
  return source.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/(^|[^:])\/\/.*$/gm, '$1');
}

// The rule is about the CORE. M1-M3 had no adapter, so scanning all of src/ and scanning src/core
// were the same set; M4 adds src/adapters/docx, which is SUPPOSED to name format concepts.
const CORE = path.join(SRC, 'core');

function coreFiles(dir = CORE, out = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) coreFiles(full, out);
    else if (entry.name.endsWith('.cjs')) out.push(full);
  }
  return out;
}

/**
 * The pattern is NOT `\bconcept\b` case-insensitively. That form was the first thing written
 * here, and the negative control below caught it going blind: `\bparagraph\b` does not match
 * `paragraphIndex`, which is exactly the identifier shape a leak would take. It reported a clean
 * core while being unable to see the real risk.
 *
 * Instead: a word boundary in front, the concept in either capitalisation, and a negative
 * lookahead for a lowercase letter after it. `paragraphIndex` and `PivotTable` match;
 * `charter` does not match `chart`.
 */
function conceptPattern(concept) {
  // Two traps were measured writing this gate, and both produced a CLEAN-looking core:
  //
  //   `\\bconcept\\b`        misses `paragraphIndex` — the identifier shape a leak actually takes
  //   `/concept(?![a-z])/i`  misses it too: the `i` flag makes `[a-z]` match `I` as well, so the
  //                          lookahead excludes exactly the camelCase boundary it should allow
  //
  // So: three explicit casings, no `i` flag. The lookahead stays case-sensitive and does the
  // discriminating — `chart` matches `chartData` and `buildChartAxis` but not `charter`;
  // `cell` matches `cellRef` and `cell_ref` but not `cellar`, `excellent` or `cancellation`.
  const lower = concept.toLowerCase();
  const pascal = lower[0].toUpperCase() + lower.slice(1);
  const upper = lower.toUpperCase();
  return new RegExp(`(?:${lower}|${pascal}|${upper})(?![a-z])`);
}

function scan(files) {
  const hits = [];
  for (const full of files) {
    const rel = path.relative(SRC, full).replace(/\\/g, '/');
    const body = stripComments(fs.readFileSync(full, 'utf8'));
    for (const concept of FORBIDDEN_CONCEPTS) {
      if (conceptPattern(concept).test(body)) hits.push({ file: rel, concept });
    }
  }
  return hits;
}

test('the comment stripper works, and the raw text would have matched (positive control)', () => {
  const contract = fs.readFileSync(path.join(SRC, 'core/adapter/contract.cjs'), 'utf8');

  // The prose that states the prohibition DOES contain the words...
  assert.match(contract, /\bparagraph\b/i, 'the prohibition sentence must exist in the source');
  // ...and must disappear once comments are stripped.
  assert.doesNotMatch(stripComments(contract), conceptPattern('paragraph'), 'the stripper failed to remove the prose');
});

test('NEGATIVE CONTROL: the scan goes red on code that really names a format concept', () => {
  const tmp = path.join(require('node:os').tmpdir(), `m1-format-probe-${process.pid}.cjs`);
  fs.writeFileSync(tmp, 'const paragraphIndex = 3;\nmodule.exports = { paragraphIndex };\n');
  try {
    const hits = scan([tmp]);
    assert.ok(hits.length > 0, 'the scan must be able to find a real format identifier');
    assert.equal(hits[0].concept, 'paragraph');
  } finally {
    fs.unlinkSync(tmp);
  }
});

test('no format concept appears in core CODE', () => {
  const hits = scan(coreFiles());
  assert.deepEqual(hits, [], `the core names format concepts in code: ${JSON.stringify(hits)}`);
});

test('the scan covers the whole core tree, not a subset', () => {
  const files = coreFiles();
  assert.ok(files.length >= 12, `expected the full core tree, scanned ${files.length} files`);
});

test('scope: the adapter tree is outside the rule, and it DOES name format concepts (positive control)', () => {
  const adapterDir = path.join(SRC, 'adapters');
  if (!fs.existsSync(adapterDir)) return;
  const adapterFiles = coreFiles(adapterDir);
  assert.ok(scan(adapterFiles).length > 0, 'the adapter names format concepts; the scan must see them there');
  assert.ok(coreFiles().every(f => f.startsWith(CORE)), 'the core scan contains no adapter file');
});

test('the pattern catches identifier shapes but not unrelated words', () => {
  // The exact discrimination the first version of this gate lacked.
  const matches = (concept, text) => conceptPattern(concept).test(text);

  assert.equal(matches('paragraph', 'const paragraphIndex = 3;'), true, 'camelCase identifier');
  assert.equal(matches('pivot', 'function buildPivotTable() {}'), true, 'PascalCase inside an identifier');
  assert.equal(matches('cell', 'cell_ref = 1'), true, 'snake_case identifier');
  assert.equal(matches('chart', 'const CHART_TYPE = 1;'), true, 'SCREAMING_CASE constant');
  assert.equal(matches('chart', 'buildChartAxis()'), true, 'concept inside a longer identifier');

  assert.equal(matches('chart', 'const charter = 1;'), false, 'a longer unrelated word is not a hit');
  assert.equal(matches('cell', 'const cellar = 1;'), false, 'nor is "cellar"');
  assert.equal(matches('cell', 'excellent cancellation'), false, 'nor words that merely contain it');
});
