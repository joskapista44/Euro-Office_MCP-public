'use strict';

/**
 * DS v9.3.5 #11 (kártya: an internal branch): a
 * docs/DS-9.3.5-LIVE-REGRESSION-SET.md táblázatának EGYETLEN célja -- minden test/live/*.live.cjs
 * fájl PONTOSAN EGYSZER szerepeljen a dokumentum saját táblájában. A fájl-halmazt a LEMEZRŐL olvassa
 * (fs.readdirSync), nem egy korábban mentett listából -- egy ÚJ harness így nem eshet ki csendben.
 *
 * RED CONTROL: ha a dokumentum táblájából kitörölnek egy sort (vagy duplikálnak egyet), ez a teszt
 * elbukik -- lásd az utolsó `test()`-et, ami egy SAJÁT MÁSOLATON szintetikusan reprodukálja ezt.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { needsPrivate } = require('../helpers/private-source.cjs');

const LIVE_DIR = path.join(__dirname, '..', 'live');
const DOC_PATH = path.join(__dirname, '..', '..', 'docs', 'DS-9.3.5-LIVE-REGRESSION-SET.md');

function liveHarnessFiles() {
  return fs.readdirSync(LIVE_DIR).filter(f => f.endsWith('.live.cjs')).sort();
}

// Row shape: "| `<harness>` | <class> | ..." -- the harness name is the first backtick-quoted token.
function harnessesInDoc(text) {
  const names = [];
  const re = /^\|\s*`([a-zA-Z0-9._-]+\.live\.cjs)`\s*\|\s*(A|B|C)\s*\|/gm;
  let m;
  while ((m = re.exec(text))) names.push({ harness: m[1], cls: m[2] });
  return names;
}

test('every test/live/*.live.cjs file (read fresh from disk) appears exactly once in the DS-9.3.5-LIVE-REGRESSION-SET.md table', { skip: needsPrivate('test/live', 'docs/DS-9.3.5-LIVE-REGRESSION-SET.md') }, () => {
  const onDisk = liveHarnessFiles();
  assert.ok(onDisk.length > 300, `sanity floor: expected a large harness set, got ${onDisk.length}`);

  const doc = fs.readFileSync(DOC_PATH, 'utf8');
  const rows = harnessesInDoc(doc);
  assert.ok(rows.length > 300, `the table itself must have rows -- parsed ${rows.length}`);

  const counts = new Map();
  for (const r of rows) counts.set(r.harness, (counts.get(r.harness) || 0) + 1);

  const missing = onDisk.filter(f => !counts.has(f));
  const duplicated = [...counts.entries()].filter(([, n]) => n > 1).map(([f]) => f);
  const extra = [...counts.keys()].filter(f => !onDisk.includes(f));

  assert.deepEqual(missing, [], `harness(es) on disk but missing from the table: ${missing.join(', ')}`);
  assert.deepEqual(duplicated, [], `harness(es) appearing more than once in the table: ${duplicated.join(', ')}`);
  assert.deepEqual(extra, [], `table row(s) naming a harness that does not exist on disk: ${extra.join(', ')}`);
  assert.equal(rows.length, onDisk.length, 'row count must equal the on-disk file count (both checks above already prove this, kept as a direct sanity assert)');
});

test('each row classifies as exactly one of A/B/C, and the summary line sums to the on-disk count', { skip: needsPrivate('test/live', 'docs/DS-9.3.5-LIVE-REGRESSION-SET.md') }, () => {
  const onDisk = liveHarnessFiles();
  const doc = fs.readFileSync(DOC_PATH, 'utf8');
  const rows = harnessesInDoc(doc);
  for (const r of rows) assert.ok(['A', 'B', 'C'].includes(r.cls), `${r.harness} has an invalid class: ${r.cls}`);

  const m = doc.match(/A \(DS-runtime-kötött\): \*\*(\d+)\*\*\s*\|\s*B \(DS-független\): \*\*(\d+)\*\*\s*\|\s*C \(NEM-MÉRT\): \*\*(\d+)\*\*/);
  assert.ok(m, 'the summary line must be present and parseable');
  const [a, b, c] = m.slice(1).map(Number);
  assert.equal(a + b + c, onDisk.length, 'A+B+C in the summary line must equal the on-disk harness count');

  const byClass = { A: 0, B: 0, C: 0 };
  for (const r of rows) byClass[r.cls] += 1;
  assert.deepEqual(byClass, { A: a, B: b, C: c }, 'the summary counts must match the actual row classes, not just the on-disk total');
});

test('RED CONTROL: a table with one row deleted (synthetic copy, not the real file) fails the coverage check', { skip: needsPrivate('test/live') }, () => {
  const onDisk = liveHarnessFiles();
  // Build a synthetic doc text missing the first on-disk harness entirely.
  const victim = onDisk[0];
  const fakeDoc = onDisk
    .filter(f => f !== victim)
    .map(f => `| \`${f}\` | A | \`${f}:1\` | LIVE | - | test |`)
    .join('\n');
  const rows = harnessesInDoc(fakeDoc);
  const counts = new Set(rows.map(r => r.harness));
  const missing = onDisk.filter(f => !counts.has(f));
  assert.deepEqual(missing, [victim], 'deleting one row from the table must surface exactly that one harness as missing');
});

test('RED CONTROL: a table with one row duplicated (synthetic copy) fails the coverage check', { skip: needsPrivate('test/live') }, () => {
  const onDisk = liveHarnessFiles();
  const dup = onDisk[0];
  const fakeDoc = [dup, ...onDisk].map(f => `| \`${f}\` | A | \`${f}:1\` | LIVE | - | test |`).join('\n');
  const rows = harnessesInDoc(fakeDoc);
  const counts = new Map();
  for (const r of rows) counts.set(r.harness, (counts.get(r.harness) || 0) + 1);
  const duplicated = [...counts.entries()].filter(([, n]) => n > 1).map(([f]) => f);
  assert.deepEqual(duplicated, [dup], 'duplicating one row in the table must surface exactly that one harness as duplicated');
});
