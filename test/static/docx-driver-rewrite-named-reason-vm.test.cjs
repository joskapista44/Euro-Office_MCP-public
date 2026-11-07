'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { commandBody, REWRITE_NO_WRITE_REASONS } = require('../../src/adapters/docx/families/text.cjs');

/**
 * GH #359 (1) 1/2: `commandBody` used to abort every guard with a bare `return;` -- a LIVE probe
 * (MULTIRUN, a production Document Server) proved this makes a no-write abort (the range-text-
 * mismatch guard, text.cjs:248) indistinguishable from a real success, because `e.callCommand`'s
 * callback receives the function's return VALUE: both an early `return;` and falling off the end
 * of the try block resolve to `undefined`/`null`.
 *
 * These tests run the REAL generated source through `vm` (per the `docx-driver-linespacing-read-
 * vm.test.cjs` precedent) against a minimal, mutable fake paragraph/range -- not a re-proof of the
 * live engine's own run-boundary quirk (that is LIVE-only, already measured), but a proof that
 * `commandBody`'s JS now reports a NAMED reason for each guard, and still performs the SAME write
 * it always did when none of them fire.
 */

/** A paragraph backed by a single mutable string, with GetRange() returning a range bound to a
 * fixed [start, end) that re-reads `state.text` on every GetText() call (a live reference, not a
 * frozen slice) -- the same semantics commandBody's own exactRange() helper assumes when it reads
 * a just-inserted range back. `rangeTextOverride(start, end, natural, callNum)` can corrupt one
 * SPECIFIC GetText() call on one specific [start,end) range object (callNum counts per-object, so
 * exactRange's own internal probe-and-match can still succeed before a LATER read on the SAME
 * object is made to diverge -- the only way to reach a post-write guard without an early one
 * absorbing it first). */
function makeParagraph(text, { hyperlink = false, classType = 'paragraph', rangeTextOverride = null, addTextFails = false, setTextPrFails = false, deleteFails = false } = {}) {
  const state = { text };
  const para = {
    GetClassType: () => classType,
    GetText: () => state.text,
    GetElementsCount: () => (hyperlink ? 1 : 0),
    GetElement: () => (hyperlink ? { GetClassType: () => 'hyperlink' } : undefined),
    GetRange: (start, end) => {
      let getTextCalls = 0;
      return {
        GetText: () => {
          getTextCalls += 1;
          const natural = state.text.slice(start, end);
          return rangeTextOverride ? rangeTextOverride(start, end, natural, getTextCalls) : natural;
        },
        GetTextPr: () => ({ marker: 'old-run-props' }),
        AddText: (t, pos) => {
          if (addTextFails || pos !== 'before') return false;
          state.text = state.text.slice(0, start) + t + state.text.slice(start);
          return true;
        },
        SetTextPr: () => (setTextPrFails ? null : true),
        Delete: () => {
          if (deleteFails) return false;
          state.text = state.text.slice(0, start) + state.text.slice(end);
          return true;
        },
      };
    },
  };
  return { para, state };
}

function runCommandBody(targetText, replacement, paragraph) {
  const sandbox = { Api: { GetDocument: () => ({ GetElementsCount: () => 1, GetElement: () => paragraph }) } };
  vm.createContext(sandbox);
  return vm.runInContext(`(function(){ ${commandBody(targetText, replacement)} })()`, sandbox);
}

// --- pre-write guards: every one of these is in REWRITE_NO_WRITE_REASONS -------------------------

test('range-text-mismatch (text.cjs:248, the MULTIRUN culprit): a paragraph that full-text-matches, whose RANGE read disagrees, returns the named reason and never writes', () => {
  const { para, state } = makeParagraph('Hello World', { rangeTextOverride: () => 'Hello Worlx' });
  const r = runCommandBody('Hello World', 'Hi', para);
  assert.equal(r, 'rewrite-live-range-text-mismatch');
  assert.equal(state.text, 'Hello World', 'no write happened');
  assert.ok(REWRITE_NO_WRITE_REASONS.includes(r));
});

// (text.cjs:251, hyperlink-refused, is not exercised here: `linkedSpanOf`'s hyperlink range-position
// arithmetic, invoked by the EARLIER unitRangeOf lookup at :248, would need its own faithful fake to
// even reach this guard -- out of scope for this unit, which only adds a name to an already-correct,
// unchanged `return`. It stays in REWRITE_NO_WRITE_REASONS on inspection, same class as :246/:252/:255.)

// --- the success paths: 'ok', and the live text actually changed ---------------------------------

test("GREEN: a clean, unique paragraph is rewritten end to end -- the old text is gone, the new text is there once, and the body reports 'ok'", () => {
  const { para, state } = makeParagraph('Hello World');
  const r = runCommandBody('Hello World', 'Hi', para);
  assert.equal(r, 'ok');
  assert.equal(state.text, 'Hi');
  assert.equal(REWRITE_NO_WRITE_REASONS.includes(r), false, "'ok' must never be treated as a no-write reason");
});

test("GREEN: an empty replacement takes the delete-only branch (text.cjs:264) and still reports 'ok' -- it IS a write", () => {
  const { para, state } = makeParagraph('Hello World');
  const r = runCommandBody('Hello World', '', para);
  assert.equal(r, 'ok');
  assert.equal(state.text, '');
});

// --- AddText itself failing, and the post-AddText rollback guards: a write WAS (or may have been)
//     attempted, so none of these belong in the no-write set ------------------------------------

test('addtext-failed (text.cjs:287, review #ref point 2): AddText IS the mutating call -- a `!== true` return does not prove nothing was inserted, so this is NOT a no-write reason', () => {
  const { para, state } = makeParagraph('Hello World', { addTextFails: true });
  const r = runCommandBody('Hello World', 'Hi', para);
  assert.equal(r, 'rewrite-live-addtext-failed');
  assert.equal(state.text, 'Hello World', "this fake's AddText leaves the text untouched on failure, but the live engine's own change log is not something this unit can observe");
  assert.equal(REWRITE_NO_WRITE_REASONS.includes(r), false, 'AddText returning non-true is not proof of no write -- same class as the rollback branches below');
});

test('rollback-inserted-text-mismatch (text.cjs:273): AddText already ran before this check fires -- the reason is reported, but it is NOT a no-write reason', () => {
  // AddText runs normally (the replacement is really spliced in). exactRange's own internal
  // probe-and-match (the FIRST GetText() call on the [0,2) range it creates for 'Hi') must still
  // see the natural text, or it would never pick this range as `inserted` in the first place; only
  // the SECOND GetText() call on that SAME object -- commandBody's own explicit check at :273 --
  // is made to diverge, by `callNum`, which counts per range object (not globally).
  const { para, state } = makeParagraph('Hello World', {
    rangeTextOverride: (start, end, natural, callNum) => (start === 0 && end === 2 && callNum === 2 ? `${natural}X` : natural),
  });
  const r = runCommandBody('Hello World', 'Hi', para);
  assert.equal(r, 'rewrite-live-rollback-inserted-text-mismatch');
  assert.equal(REWRITE_NO_WRITE_REASONS.includes(r), false, 'a write was attempted (AddText ran) before this guard -- must stay mutated=true');
  // rollback() re-reads with a FRESH range object (its own, un-corrupted first GetText() call), so
  // it finds and deletes the insert -- the visible document ends up back at the original text, but
  // the reason the driver sees is still "a write was attempted", not "no write" (an in-editor
  // rollback is best-effort and is not proof the Document Server's own change log stayed empty).
  assert.equal(state.text, 'Hello World');
});

test('rollback-settextpr-failed (text.cjs:271): AddText already ran -- not a no-write reason', () => {
  const { para } = makeParagraph('Hello World', { setTextPrFails: true });
  const r = runCommandBody('Hello World', 'Hi', para);
  assert.equal(r, 'rewrite-live-rollback-settextpr-failed');
  assert.equal(REWRITE_NO_WRITE_REASONS.includes(r), false);
});

// --- the reason catalogue itself ------------------------------------------------------------------

test('REWRITE_NO_WRITE_REASONS holds exactly the guards that fire before commandBody\'s first mutating engine call -- AddText itself excluded (review #ref point 2)', () => {
  assert.deepEqual([...REWRITE_NO_WRITE_REASONS].sort(), [
    'rewrite-live-hyperlink-refused',
    'rewrite-live-no-document-or-empty-target',
    'rewrite-live-not-paragraph',
    'rewrite-live-paragraph-not-found',
    'rewrite-live-paragraph-text-not-unique',
    'rewrite-live-range-text-mismatch',
  ].sort());
});
