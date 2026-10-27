'use strict';

/**
 * GH #226 (DOCX P2 READ: notes completeness). history: the audit (docs/GH226-NOTES-READ-DECISION.md) found `read_notes` first-paragraph only;
 * GH #226 2 found three public full-text paths; GH #226 3 (docs/GH226-NOTES-FULL-TEXT-RESEARCH.md s.7) implemented the GetNext() chain read and
 * the LIVE (DS 9.3.5, run gh226chain-live-1792394096767-10bc02) proved GetNext() === null at every note end, every note FULL, fresh and reopen,
 * no write -> read_notes PARTIAL -> ACCEPTED. These tests pin the status, the evidence and the product behaviour through the real read tool.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { OPERATIONS, STATUS, createReadTool } = require('../../src/server/read-tools.cjs');
const { fakeBroker, docxEditor, FILE, PROPERTY_RESULT } = require('../fakes/read-editors.cjs');

const chain = (paragraphs, complete = true, stop = 'null') => ({ paragraphs, complete, stop });

async function readNotes(documentPatch) {
  const saved = { ...PROPERTY_RESULT.document };
  Object.assign(PROPERTY_RESULT.document, documentPatch);
  try {
    const b = fakeBroker('docx');
    const ed = docxEditor();
    const t = createReadTool('docx', { configSource: b.source, loadPlaywright: ed.loadPlaywright });
    return { r: await t.handler({ fileId: FILE, operation: 'read_notes' }), b };
  } finally {
    for (const k of Object.keys(PROPERTY_RESULT.document)) delete PROPERTY_RESULT.document[k];
    Object.assign(PROPERTY_RESULT.document, saved);
  }
}

test('GH #226 3: read_notes is ACCEPTED (args {}), the evidence names the GetNext() chain, the stop condition and the LIVE run', () => {
  const o = OPERATIONS.docx.read_notes;
  assert.equal(o.status, STATUS.ACCEPTED);
  assert.deepEqual(o.args, {});
  assert.match(o.evidence, /GetNext\(\) === null/);
  assert.match(o.evidence, /gh226chain-live-1792394096767-10bc02/);
  assert.doesNotMatch(o.evidence, /MEASURED TRUNCATED/);
  assert.doesNotMatch(o.evidence, /NOT MEASURED/);
});

test('GH #226 3: read_notes returns the FULL text of every note, complete when every chain ended by null', async () => {
  const { r, b } = await readNotes({
    footnotes: ['f1', 'g1'], endnotes: ['e1'],
    footnoteChains: [chain(['f1']), chain(['g1', 'g2'])], endnoteChains: [chain(['e1', 'e2', 'e3'])],
  });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(b.calls.length > 0, true, 'an ACCEPTED op reaches the broker');
  assert.deepEqual(r.result.footnotes.map(n => n.text), ['f1', 'g1\ng2']);
  assert.deepEqual(r.result.endnotes[0].paragraphs, ['e1', 'e2', 'e3']);
  assert.equal(r.result.complete, true);
});

test('GH #226 3: an untrusted chain is NAMED (unobservable), never a guessed full text, and the read is not complete', async () => {
  const { r } = await readNotes({
    footnoteChains: [chain(['f1']), chain(['g1'], false, 'overrun-suspected')], endnoteChains: [chain(['e1'])],
  });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.footnotes[1].unobservable, 'UNOBSERVABLE:note-chain-overrun-suspected');
  assert.equal(r.result.footnotes[0].unobservable, undefined);
  assert.equal(r.result.complete, false);
});

test('GH #226 3: a document whose chains were never read answers null (unread), not an empty list', async () => {
  const { r } = await readNotes({});
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result, { footnotes: null, endnotes: null, complete: false });
});
