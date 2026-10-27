'use strict';

/**
 * GH #141 MASTER reconciliation (docs-only, 2026-10-18): every evidence file the reconciliation
 * doc cites must actually exist in this repo, at the path cited. This never re-derives the Wave
 * table -- it only proves the evidence links are not dangling.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { needsPrivate } = require('../helpers/private-source.cjs');

const REPO_ROOT = path.join(__dirname, '..', '..');

// Every path here must appear, byte-for-byte, in the "Evidence-file index" section of
// docs/GH141-MASTER-RECONCILIATION-20261018.md.
const EVIDENCE_FILES = Object.freeze([
  'docs/CAPABILITY-COMPLETION-STATUS-20261012.md',
  'docs/GH100-SESSION-REUSE-DESIGN.md',
  'docs/GH100-U3-FRESHNESS-PLAN.md',
  'docs/GH100-U4C-READ-REUSE-DESIGN.md',
  'docs/GH133-READ-SHEET-STATE-RELEASE.md',
  'docs/GH135-STRUCTURAL-READBACK-CONTRACT.md',
  'docs/GH136-CF-READ-COMPLETENESS.md',
  'docs/GH137-U1-READ-RUNS-RELEASE.md',
  'docs/GH137-U2-FORMAT-RUN-RELEASE.md',
  'docs/GH138-U1-DEFINED-NAME-WRITE-RELEASE.md',
  'docs/GH138-U2-PRINT-TITLES-WRITE-RELEASE.md',
  'docs/GH138-U3-PAGE-LAYOUT-WRITE-RELEASE.md',
  'docs/GH138-U4-PRINT-FAMILY-RELEASE.md',
  'docs/GH139-CHART-DRAWING-SYMMETRY-MATRIX.md',
  'docs/GH139-U2-MOVE-CHART-RELEASE.md',
  'docs/GH139-U3-INSERT-DRAWING-RELEASE.md',
  'docs/GH140-PIVOT-CONTRACT.md',
  'docs/GH140-U4-REFRESH-PIVOT-WRITE-RELEASE.md',
  'docs/L16-FORMAT-REMOVE-REFUSALS-PLAN.md',
  'docs/L16-3-C5-UNMERGE-COMPOSITE-ANALYSIS.md',
  'docs/evidence/gh141-final-readwrite-acceptance-live-final-20261018.json',
]);

test('every evidence file cited by the GH #141 master reconciliation doc exists in the repo', { skip: needsPrivate('docs/GH141-MASTER-RECONCILIATION-20261018.md') }, () => {
  for (const rel of EVIDENCE_FILES) {
    const abs = path.join(REPO_ROOT, rel);
    assert.ok(fs.existsSync(abs), `missing evidence file: ${rel}`);
  }
});

test('the reconciliation doc itself exists and its evidence-file index is byte-identical to the list this test checks', { skip: needsPrivate('docs/GH141-MASTER-RECONCILIATION-20261018.md') }, () => {
  const docPath = path.join(REPO_ROOT, 'docs', 'GH141-MASTER-RECONCILIATION-20261018.md');
  assert.ok(fs.existsSync(docPath), 'docs/GH141-MASTER-RECONCILIATION-20261018.md should exist');
  const body = fs.readFileSync(docPath, 'utf8');
  const m = body.match(/## Evidence-file index[\s\S]*?```\n([\s\S]*?)```/);
  assert.ok(m, 'doc should have an "Evidence-file index" fenced block');
  const listed = m[1].trim().split('\n').map(s => s.trim()).filter(Boolean);
  assert.deepEqual(listed, EVIDENCE_FILES.slice(), 'the doc\'s own evidence-file list must match what this test checks -- otherwise the test could pass while the doc cites something unverified');
});

test('RED CONTROL: a non-existent evidence path is correctly reported as missing -- proves the check actually discriminates', () => {
  const abs = path.join(REPO_ROOT, 'docs', 'GH141-DEFINITELY-NOT-A-REAL-FILE-20261018.md');
  assert.ok(!fs.existsSync(abs));
});
