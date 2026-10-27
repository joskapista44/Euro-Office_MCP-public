'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

/**
 * GH #226 3 -- the REAL propertyQuery() string (not a retyped copy) run in a vm against a fake Api whose notes are linked paragraph lists
 * (the OnlyOffice GetNext() chain). Pins: the additive footnoteChains/endnoteChains, the stop condition (GetNext() === null), and the NAMED
 * incomplete outcomes (no-getnext | error | overrun-suspected | bound). The first-paragraph fields stay exactly as before.
 */

/** A note = a linked list of paragraph fakes; `tail` is what the LAST paragraph's GetNext() yields (null = the real end of the note). */
function linkNote(texts, tail = null, opts = {}) {
  const paras = texts.map(t => ({ GetText: () => t }));
  paras.forEach((p, i) => {
    if (opts.noGetNext) { delete p.GetNext; return; }
    p.GetNext = () => { if (opts.throwAt === i) throw new Error('boom'); return i + 1 < paras.length ? paras[i + 1] : tail; };
  });
  return paras;
}

function run(foot, end) {
  const doc = {
    GetFootnotesFirstParagraphs: () => foot.map(n => n[0]),
    GetEndNotesFirstParagraphs: () => end.map(n => n[0]),
    GetCurrentFootEndnote: () => null,
    GetAllBookmarksNames: () => [],
    GetElementsCount: () => 0,
  };
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${propertyQuery()} })()`, sandbox)));
}
const docOf = r => r.document || r.docInfo || r;

test('vm: every note\'s full paragraph list through GetNext() === null is read, complete, stop=null; the first-paragraph arrays are unchanged', () => {
  const d = docOf(run(
    [linkNote(['F-A1']), linkNote(['F-B1', 'F-B2']), linkNote(['F-C1', 'F-C2', 'F-C3'])],
    [linkNote(['E-A1', 'E-A2', 'E-A3']), linkNote(['E-B1']), linkNote(['E-C1', 'E-C2'])],
  ));
  assert.deepEqual(d.footnotes, ['F-A1', 'F-B1', 'F-C1']);
  assert.deepEqual(d.endnotes, ['E-A1', 'E-B1', 'E-C1']);
  assert.deepEqual(d.footnoteChains.map(c => c.paragraphs), [['F-A1'], ['F-B1', 'F-B2'], ['F-C1', 'F-C2', 'F-C3']]);
  assert.deepEqual(d.endnoteChains.map(c => c.paragraphs), [['E-A1', 'E-A2', 'E-A3'], ['E-B1'], ['E-C1', 'E-C2']]);
  for (const c of [...d.footnoteChains, ...d.endnoteChains]) { assert.equal(c.complete, true); assert.equal(c.stop, 'null'); }
});

test('vm: a chain that runs into the NEXT note\'s first paragraph is NAMED overrun-suspected (incomplete), never cut by text', () => {
  const b = linkNote(['F-B1']);
  const a = linkNote(['F-A1', 'F-A2'], b[0]); // the real boundary leaks: A's last GetNext() yields B's first paragraph
  const d = docOf(run([a, b], []));
  assert.equal(d.footnoteChains[0].complete, false);
  assert.equal(d.footnoteChains[0].stop, 'overrun-suspected');
  assert.deepEqual(d.footnoteChains[0].paragraphs, ['F-A1', 'F-A2']);
  assert.equal(d.footnoteChains[1].complete, true);
});

test('vm: no GetNext / a throwing GetNext / an endless chain are NAMED incomplete (no-getnext | error | bound)', () => {
  const endless = { GetText: () => 'x' }; endless.GetNext = () => endless;
  const d = docOf(run([linkNote(['N-1', 'N-2'], null, { noGetNext: true }), linkNote(['T-1', 'T-2'], null, { throwAt: 0 }), [endless]], []));
  assert.deepEqual(d.footnoteChains.map(c => [c.complete, c.stop]), [[false, 'no-getnext'], [false, 'error'], [false, 'bound']]);
  assert.deepEqual(d.footnoteChains[0].paragraphs, ['N-1']);
});

test('inventory: footnoteChains/endnoteChains pass through frozen; absent = null; a malformed record is a NAMED incomplete note', () => {
  const inv = normalizeInventory({ html: '<p>x</p>', paragraphs: [], document: { footnotes: ['a'], endnotes: ['b'], footnoteChains: [{ paragraphs: ['a', 'a2'], complete: true, stop: 'null' }, 7], endnoteChains: undefined } });
  const d = inv.document;
  assert.deepEqual(d.footnoteChains[0], { paragraphs: ['a', 'a2'], complete: true, stop: 'null' });
  assert.deepEqual(d.footnoteChains[1], { paragraphs: [], complete: false, stop: 'malformed' });
  assert.equal(d.endnoteChains, null);
  assert.ok(Object.isFrozen(d.footnoteChains) && Object.isFrozen(d.footnoteChains[0].paragraphs));
});

const { FORMATS, OPERATIONS, STATUS, inputSchema } = require('../../src/server/read-tools.cjs');
const project = doc => FORMATS.docx.project({ document: doc }, { operation: 'read_notes' });
const chain = (paragraphs, complete = true, stop = 'null') => ({ paragraphs, complete, stop });

test('projection: full text per note (paragraphs joined by a newline); complete only when EVERY note of both kinds is complete', () => {
  const r = project({ footnoteChains: [chain(['a']), chain(['b1', 'b2'])], endnoteChains: [chain(['e1', 'e2', 'e3'])] });
  assert.deepEqual(r.footnotes.map(n => n.text), ['a', 'b1\nb2']);
  assert.equal(r.endnotes[0].text, 'e1\ne2\ne3');
  assert.equal(r.complete, true);
});

test('projection: an untrusted chain is NAMED (unobservable) and makes the whole read incomplete; an unread kind stays null', () => {
  const r = project({ footnoteChains: [chain(['a']), chain(['b1'], false, 'overrun-suspected')], endnoteChains: [chain(['e'])] });
  assert.equal(r.footnotes[1].unobservable, 'UNOBSERVABLE:note-chain-overrun-suspected');
  assert.equal(r.footnotes[0].unobservable, undefined);
  assert.equal(r.complete, false);
  const u = project({ footnoteChains: [chain(['a'])], endnoteChains: null });
  assert.equal(u.endnotes, null);
  assert.equal(u.complete, false);
  assert.deepEqual(project(null), { footnotes: null, endnotes: null, complete: false });
});

test('status gate: read_notes is ACCEPTED (the LIVE gh226chain-live-1792394096767-10bc02 proved the stop condition) and the schema offers it', () => {
  assert.equal(OPERATIONS.docx.read_notes.status, STATUS.ACCEPTED);
  assert.equal(inputSchema('docx').properties.operation.enum.includes('read_notes'), true);
});

test('LIVE harness classifier: classifyProductNotes judges the product read_notes result against the fixture model', { skip: needsPrivate('test/live/docx-notes-chain-boundary.live.cjs') }, () => {
  const { classifyProductNotes } = require('../live/docx-notes-chain-boundary.live.cjs');
  const model = { foot: [['a'], ['b1', 'b2']], end: [['e']] };
  const n = (paragraphs, complete = true) => ({ paragraphs, text: paragraphs.join('\n'), complete, stop: 'null' });
  const good = classifyProductNotes({ footnotes: [n(['a']), n(['b1', 'b2'])], endnotes: [n(['e'])], complete: true }, model);
  assert.deepEqual([...good.foot, ...good.end].map(x => x.finding), ['FULL', 'FULL', 'FULL']);
  assert.equal(good.complete, true);
  const bad = classifyProductNotes({ footnotes: [n(['a']), n(['b1'])], endnotes: [], complete: false }, model);
  assert.deepEqual([...bad.foot, ...bad.end].map(x => x.finding), ['FULL', 'MISMATCH', 'NOT_READ']);
  assert.equal(bad.complete, false);
});
