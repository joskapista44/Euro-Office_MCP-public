'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery, richOutcome } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * gntw3 (nested-table document, 2026-10-19): read_tables / readLive on a document carrying a nested table
 * (3 top-level tables, one of them holding a nested table) came back LIVE as `{tables:null}` with
 * `incomplete: ["rich: UNOBSERVABLE:property-query: no paragraphs array"]` -- the whole rich payload gone,
 * not just the table list (maintainer card #ref / RAW_READ evidence gntw3-docx-nested-table-write-product-rawdump.out).
 *
 * Part (a), the root cause: the PROPERTY_QUERY script itself is innocent. Run offline in a vm against a
 * document shaped exactly like the stored fixture (3 top-level tables, one nested one level deep) it never
 * throws, leaks no raw API object into the returned tree (every value funnels through readCell/readTableRows/
 * tblFormatOf, which only ever extract plain primitives), and the whole result is JSON.stringify-able. This
 * is the SAME conclusion maintainer's docx-nested-tables-in-read-tables-vm.test.cjs already reaches for
 * correctness of VALUES; this file adds the serializability/no-leaked-object angle the "API object in
 * nestedTables" hypothesis in the card named. Since the script is clean, and it always ends with an explicit
 * `return {...}` (never an implicit/undefined return), the loss has to happen OUTSIDE this file: at the
 * native callCommand bridge, whose callback fired with no value at all (runInEditor's own
 * `r === undefined ? null : r` line, docx-driver.cjs ~1024). That is NOT reproducible offline (closed
 * DS/ONLYOFFICE component) -- this file does not claim to.
 *
 * Part (b), the fix: richOutcome() used to treat "callCommand completed but the callback got nothing"
 * (result: null) IDENTICALLY to "the script computed and returned a value whose own .paragraphs field is
 * missing" -- both produced 'UNOBSERVABLE:property-query: no paragraphs array', a message that reads as a
 * SHAPE problem on a result that was in fact never received. richOutcome now names the escape on its own,
 * ahead of the generic branches.
 */

const run = (doc, src = propertyQuery()) => {
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${src} })()`, sandbox)));
};
const paragraph = text => ({ GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, GetElementsCount: () => 0, GetElement: () => null, GetTextPr: () => null, GetRange: () => ({ GetTextPr: () => null }), GetStyle: () => null, GetParaPr: () => null });
const cellOf = (text, ...nested) => {
  const els = [paragraph(text), ...nested];
  return { GetContent: () => ({ GetElementsCount: () => els.length, GetElement: i => els[i], GetText: () => text }) };
};
const table = (tag, nestedAt = {}) => {
  const cells = [[0, 0], [0, 1], [1, 0], [1, 1]].map(([r, c]) => cellOf(`${tag} r${r}c${c}`, ...(nestedAt[`${r}${c}`] ? [].concat(nestedAt[`${r}${c}`]) : [])));
  return { GetClassType: () => 'table', GetRowsCount: () => 2, GetRow: r => ({ GetCellsCount: () => 2, GetCell: c => cells[r * 2 + c] }), GetTableTitle: () => null };
};
const docOf = (body, all = body) => ({
  GetElementsCount: () => body.length, GetElement: i => body[i], GetAllTables: () => all, Search: () => [],
  GetAllBookmarksNames: () => [], GetBookmarkRange: () => null, GetFootnotesFirstParagraphs: () => [], GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null,
});

// finds any function, circular reference, or raw API-shaped object (has GetClassType) anywhere in a plain tree
function firstNonPlain(v, path, seen) {
  if (v === null || typeof v !== 'object') return typeof v === 'function' ? `${path} (function)` : null;
  if (seen.has(v)) return `${path} (CIRCULAR)`;
  seen.add(v);
  if (typeof v.GetClassType === 'function') return `${path} (RAW API OBJECT leaked -- has GetClassType)`;
  for (const k of Object.keys(v)) { const bad = firstNonPlain(v[k], `${path}.${k}`, seen); if (bad) return bad; }
  return null;
}

test('vm: a document shaped like the gntw3 fixture (3 top-level tables, one nested one level deep) -- the query never throws, returns a plain tree, and is fully JSON-serializable', () => {
  const nested = table('NIN');
  const t1 = table('T1', { '00': nested });
  const t2 = table('T2');
  const t3 = table('T3');
  const doc = docOf([t1, t2, t3], [t1, nested, t2, t3]);

  let threw = null;
  let r;
  try { r = run(doc); } catch (e) { threw = e; }

  assert.equal(threw, null, 'the query must not throw on a nested-table document');
  assert.equal(r.error, undefined, 'the outer catch must not have fired');
  assert.equal(Array.isArray(r.paragraphs), true, 'paragraphs stays an array even with no top-level paragraphs');
  assert.equal(r.document.tables.length, 3, 'all three top-level tables are read');
  // the page hands the nested tables over as a JSON string on the holding cell (gnpq2: the object form does not cross the callCommand boundary); readLive inflates it
  assert.equal(typeof r.document.tables[0].rows[0][0].nestedTablesJson, 'string', 'the nested table travels as a string on its holding cell');
  assert.equal(JSON.parse(r.document.tables[0].rows[0][0].nestedTablesJson).length, 1, 'the nested table is read on its holding cell');

  const leaked = firstNonPlain(r, 'r', new Set());
  assert.equal(leaked, null, `the returned tree must hold no function/circular ref/raw API object; found: ${leaked}`);

  // the whole structure must survive a JSON round trip -- the same shape the real callCommand boundary
  // has to carry back to the host
  let jsonOk = true;
  try { JSON.parse(JSON.stringify(r)); } catch { jsonOk = false; }
  assert.equal(jsonOk, true, 'the result must be JSON.stringify/parse-safe');
});

test('richOutcome: a callCommand that completed but delivered NO result value is named as an escape, not as a shape problem', () => {
  // this is exactly what runInEditor hands richOutcome when its callback fires with `r === undefined`
  // (docx-driver.cjs: `result: r === undefined ? null : r`) -- the PROPERTY_QUERY script itself never
  // legitimately returns null (proved above and in the previous test), so this shape is unambiguous: an
  // escape, not an empty/malformed read.
  const o = richOutcome({ ok: true, outcome: 'command-completed', result: null });
  assert.equal(o.paragraphs, null);
  assert.equal(o.document, null);
  assert.equal(o.richUnobservable, 'UNOBSERVABLE:property-query: command-completed with no result value (escaped before delivery, not computed as empty)');
  // the OLD, misleading message (reads as "the script ran and its .paragraphs field had the wrong shape")
  // must NOT be the one reported for this case anymore
  assert.notEqual(o.richUnobservable, 'UNOBSERVABLE:property-query: no paragraphs array');
});

test('richOutcome NEG: a result the script genuinely computed (paragraphs present but not an array, e.g. the key missing) keeps the OLD, shape-specific message -- the new branch is scoped to result===null only', () => {
  const o = richOutcome({ ok: true, outcome: 'command-completed', result: { document: {} } });
  assert.equal(o.richUnobservable, 'UNOBSERVABLE:property-query: no paragraphs array');
});

test('richOutcome NEG: a genuinely empty-but-valid read (paragraphs: [], document: {}) reports no richUnobservable at all', () => {
  const o = richOutcome({ ok: true, outcome: 'command-completed', result: { paragraphs: [], document: {} } });
  assert.equal(o.richUnobservable, undefined);
  assert.deepEqual(o.paragraphs, []);
  assert.deepEqual(o.document, {});
});

test('richOutcome NEG: a call that never reached command-completed (timeout / call-threw) keeps its own, earlier-branch message, not the new one', () => {
  assert.equal(richOutcome({ ok: false, outcome: 'command-callback-timeout' }).richUnobservable, 'UNOBSERVABLE:property-query-call: command-callback-timeout');
  assert.equal(richOutcome({ ok: false, outcome: 'call-threw', error: 'boom' }).richUnobservable, 'UNOBSERVABLE:property-query-call: boom');
});
