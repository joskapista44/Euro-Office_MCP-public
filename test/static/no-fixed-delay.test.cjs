'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * Static gate: no fixed delay as a correctness primitive — PROJECT invariant 6, owner req. 17.
 *
 * Condition/predicate based waiting IS allowed. The distinction the gate enforces is WHERE the
 * timer may live: exactly one named function, whose only job is to yield between polls of an
 * observed predicate. Anywhere else in the core, a timer is a delay standing in for correctness.
 */

const SRC = path.join(__dirname, '..', '..', 'src');

/**
 * The single deliberate exception, named here so the gate reads it rather than a reviewer
 * remembering it. If the timer moves or a second one appears, this gate goes red.
 */
const ALLOWED = Object.freeze([
  { file: 'core/session/open.cjs', fn: 'defaultWait', reason: 'yields between polls of an observed readiness/canSave predicate' },
  // M4 transport: an UPPER BOUND on waiting for the editor's callCommand callback. The happy path
  // exits on the callback event; the timer fires only for an editor that never answers, and then the
  // dispatch is reported as not completed (the core classifies it uncertain) -- never as success.
  { file: 'drivers/eurooffice/docx-driver.cjs', fn: 'runInEditor', reason: 'bounds the wait for the callCommand completion callback; exits on the event' },
  // M6: the XLSX transport's copy of the same bounded callback wait (xlsx-driver.cjs).
  { file: 'drivers/eurooffice/xlsx-driver.cjs', fn: 'runInEditor', reason: 'bounds the wait for the callCommand completion callback; exits on the event' },
  // (M11.5c.4 removed the M6 W3 copyInFrame poll: copy_sheet is now ONE synchronous worksheet.copy/v1 Bridge WRITE inside
  // a callCommand, bounded by runInEditor's own timer above.)
  // M11.3b.2 bounded shutdown (owner decision, owner channel 2026-10-10): an UPPER BOUND on waiting for the active
  // session's own close(). The happy path exits on the close event; the timer fires only for a session that never
  // closes, and the result is reported TIMEOUT -- never a correctness wait.
  { file: 'server/session-gate.cjs', fn: 'shutdown', reason: 'bounds the wait for the active driver close(); exits on the event' },
]);

const TIMER_PATTERN = /\b(setTimeout|setInterval|sleep|delay|setImmediate)\s*\(/g;

function collectFiles(dir = SRC, out = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) collectFiles(full, out);
    else if (entry.name.endsWith('.cjs')) out.push(full);
  }
  return out;
}

/** Strips comments so the prose that STATES the prohibition is not counted as a violation. */
function code(source) {
  return source.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
}

function findTimers() {
  const hits = [];
  for (const full of collectFiles()) {
    const rel = path.relative(SRC, full).replace(/\\/g, '/');
    const body = code(fs.readFileSync(full, 'utf8'));
    for (const m of body.matchAll(TIMER_PATTERN)) {
      const line = body.slice(0, m.index).split('\n').length;
      hits.push({ file: rel, token: m[1], line });
    }
  }
  return hits;
}

test('the timer scan is not blind (positive control)', () => {
  // The allowed exception must actually be found; if the scanner cannot see the one timer we
  // know exists, its zeros elsewhere would mean nothing.
  const hits = findTimers();
  const allowedFile = ALLOWED[0].file;
  assert.ok(
    hits.some(h => h.file === allowedFile),
    `the scanner did not find the known timer in ${allowedFile} — it is not measuring`
  );
});

test('no fixed delay in the core outside the one named exception', () => {
  const allowedFiles = new Set(ALLOWED.map(a => a.file));
  const offenders = findTimers().filter(h => !allowedFiles.has(h.file));
  assert.deepEqual(
    offenders,
    [],
    `fixed delay used outside the named exception: ${JSON.stringify(offenders)}`
  );
});

test('each named exception is exactly one timer, inside its named function', () => {
  for (const allowed of ALLOWED) {
    const inFile = findTimers().filter(h => h.file === allowed.file);
    const named = ALLOWED.filter(x => x.file === allowed.file).length; // one timer per named function
    assert.equal(inFile.length, named, `expected exactly ${named} timer(s) in ${allowed.file}, found ${inFile.length}`);

    const body = fs.readFileSync(path.join(SRC, allowed.file), 'utf8');
    const fnStart = body.indexOf(`function ${allowed.fn}(`);
    assert.ok(fnStart > -1, `the named exception function ${allowed.fn} no longer exists`);
    const next = body.indexOf('\n  async function ', fnStart + 1);
    const fnBody = body.slice(fnStart, next > -1 ? next : body.indexOf('\n}', fnStart));
    assert.match(fnBody, /setTimeout/, `the timer must live inside ${allowed.fn}`);
  }
});

test('every wait in the core is bounded by an observed predicate, not a bare duration', () => {
  // Both loops that wait do so inside a `for (;;)` that exits on a probe result. A wait whose
  // only exit is the timer would show up as an await on the helper with no surrounding probe.
  const open = fs.readFileSync(path.join(SRC, 'core/session/open.cjs'), 'utf8');
  const barrier = fs.readFileSync(path.join(SRC, 'core/persistence/barrier.cjs'), 'utf8');

  assert.match(open, /last = await adapter\.readinessProbe\(handle\)/, 'readiness waits on a probe');
  assert.match(barrier, /probe\.ok === true && probe\.canSave === false/, 'the barrier waits on canSave');
  for (const [name, body] of [['open', open], ['barrier', barrier]]) {
    assert.match(body, /for \(;;\)/, `${name} must poll a condition rather than sleep a duration`);
  }
});
