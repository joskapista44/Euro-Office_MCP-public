'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const rich = require('../../src/adapters/docx/families/rich.cjs');
const notes = require('../../src/adapters/docx/families/notes.cjs');

/**
 * The noteRefs field of the rich read names an unreadable child instead of silently UNDERCOUNTING (the per-element guard's follow-up). The REAL propertyQuery() string run in a vm against a fake Api
 * in which one child of a paragraph cannot be fetched / its style cannot be read; then the three consumers that must treat the named string as UNREAD, never as a value:
 * the inventory (keeps it named, not null), notes.cjs (fail-closed 'note-references-unreadable'), rich.cjs (a NAMED count on both sides is NOT 'kept'; unit keys do not depend on the message text).
 */

const run = doc => {
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${propertyQuery()} })()`, sandbox)));
};
const noteRun = style => ({ GetClassType: () => 'run', GetRange: () => ({ GetTextPr: () => ({}) }), GetText: () => '', GetTextPr: () => ({ GetStyle: () => ({ GetName: () => style }) }) });
const para = (text, children) => ({ GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, GetElementsCount: () => children.length, GetElement: j => { const c = children[j]; if (c instanceof Error) throw c; return c; } });
const docOf = paragraphs => ({ GetElementsCount: () => paragraphs.length, GetElement: i => paragraphs[i], Search: () => [], GetAllBookmarksNames: () => [], GetFootnotesFirstParagraphs: () => [], GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null });
const entry = (r, text) => r.paragraphs.find(p => p.text === text);

test('vm: a paragraph whose children are all readable counts its note references; a child that cannot be fetched makes noteRefs the NAMED string (hyperlinks too), never a silent undercount', () => {
  const r = run(docOf([
    para('P1 two refs', [noteRun('footnote reference'), noteRun('endnote reference'), noteRun('footnote reference')]),
    para('P2 broken child', [noteRun('footnote reference'), new Error('child fetch')]),
    para('P3 one ref', [noteRun('endnote reference')]),
  ]));
  assert.equal(r.error, undefined);
  assert.deepEqual(entry(r, 'P1 two refs').noteRefs, { footnote: 2, endnote: 1 });
  assert.equal(entry(r, 'P2 broken child').noteRefs, 'UNOBSERVABLE:child 1: child fetch');
  assert.equal(entry(r, 'P2 broken child').hyperlinks, 'UNOBSERVABLE:child 1: child fetch');
  assert.deepEqual(entry(r, 'P3 one ref').noteRefs, { footnote: 0, endnote: 1 });
});

test('vm: a run whose style chain throws (the note-reference classification) is ALSO a named child failure -- the count would be wrong, so it is not reported', () => {
  const badRun = { ...noteRun('footnote reference'), GetTextPr: () => { throw new Error('no style chain'); } };
  const r = run(docOf([para('P1 bad style', [noteRun('footnote reference'), badRun])]));
  assert.equal(entry(r, 'P1 bad style').noteRefs, 'UNOBSERVABLE:child 1: no style chain');
});

const NAMED = 'UNOBSERVABLE:child 0: boom';
const inv = (noteRefs, extra = {}) => normalizeInventory({ html: '<p>Target</p>', paragraphs: [{ text: 'Target', props: { bold: false }, runs: [], nonRunChildren: 0, runsCoverText: true, hyperlinks: [], noteRefs, ...extra }], document: { bookmarks: [], footnotes: [], endnotes: [], inMainBody: true } });
const nodeOf = i => i.blocks.find(b => b.text === 'Target');

test('inventory: a NAMED noteRefs stays named; an absent one is null; an object stays an object -- never conflated, and a string is NEVER spread into an object', () => {
  assert.equal(nodeOf(inv(NAMED)).noteRefs, NAMED);
  assert.equal(nodeOf(inv(undefined)).noteRefs, null);
  assert.deepEqual(nodeOf(inv({ footnote: 1, endnote: 0 })).noteRefs, { footnote: 1, endnote: 0 });
});

test('notes.cjs: a named noteRefs is fail-closed note-references-unreadable (the count is not a value)', () => {
  const s = notes.noteState(inv(NAMED), { intent: 'add_footnote', selector: { text: 'Target' }, text: 'a note' });
  assert.equal(s.ok, false); assert.equal(s.reason, 'note-references-unreadable');
});

test('rich.cjs verifyPreservation: the SAME named noteRefs on both sides is NOT kept (it used to compare equal as strings); readable equal counts still are', () => {
  const op = { intent: 'add_link', selector: { text: 'Target' } };
  const named = rich.verifyPreservation(op, { primary: inv(NAMED) }, { primary: inv(NAMED) });
  assert.equal(named.ok, false); assert.equal(named.evidence.changed.length, 0, 'only the unreadable counts block it');
  const kept = rich.verifyPreservation(op, { primary: inv({ footnote: 1, endnote: 0 }) }, { primary: inv({ footnote: 1, endnote: 0 }) });
  assert.equal(kept.ok, true);
  const moved = rich.verifyPreservation(op, { primary: inv({ footnote: 1, endnote: 0 }) }, { primary: inv({ footnote: 2, endnote: 0 }) });
  assert.equal(moved.ok, false);
});
