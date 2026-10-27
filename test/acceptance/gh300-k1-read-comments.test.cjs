'use strict';
/**
 * GH #300 K1: read_comments -- the XLSX READ of the cell comments. The driver's commentSheets read (a bounded per-cell
 * GetComment scan reconciled against ws.GetComments(), run here against a fake Api), the projection
 * (families/comments-read.cjs), the request it makes and the catalogue entry.
 */
const test = require('node:test');
const assert = require('node:assert');
const { OPERATIONS, STATUS, validateReadRequest, xlsxRequest } = require('../../src/server/read-tools.cjs');
const { readQuery, MAX_COMMENT_SCAN_CELLS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { projectCommentsRead } = require('../../src/adapters/xlsx/families/comments-read.cjs');

/** A fake ApiComment from a plain row; a getter named in `throws` throws. */
function comment(c, throws = []) {
  const g = (k, v) => () => { if (throws.includes(k)) throw new Error(`${k} boom`); return v; };
  return { GetText: g('GetText', c.text), GetAuthorName: g('GetAuthorName', c.author), GetUserId: g('GetUserId', c.userId ?? ''), GetTimeUTC: g('GetTimeUTC', c.timeUtc ?? 0),
    IsSolved: g('IsSolved', !!c.solved), GetRepliesCount: () => (c.replies || []).length, GetReply: i => comment(c.replies[i]), GetId: () => { throw new Error('the id is not read'); } };
}
/** A fake sheet: `at` maps A1 -> comment row; `extra` = comments GetComments() counts that no cell carries; `used` = GetUsedRange address. */
function sheet(name, at, { used = 'A1:C5', extra = 0, throws = [] } = {}) {
  const list = Object.values(at).map(c => comment(c, throws));
  return { GetName: () => name, GetIndex: () => 0, GetUsedRange: () => (used ? { GetAddress: () => used } : null),
    GetComments: () => [...list, ...Array.from({ length: extra }, () => ({}))],
    GetRange: a => ({ GetComment: () => (at[a] ? comment(at[a], throws) : null), GetAddress: () => a }) };
}
const run = (sheets, asked) => {
  const Api = { GetSheets: () => sheets, GetSheet: n => sheets.find(s => s.GetName() === n) || null, WorksheetFunction: { COUNT: () => 0 } };
  return new Function('Api', readQuery([], { commentSheets: asked, usedRanges: false }))(Api).commentReads;
};

test('GH #300 K1: read_comments is an ACCEPTED xlsx read operation taking { sheet }, asking the driver for that sheet\'s comments only', () => {
  assert.strictEqual(OPERATIONS.xlsx.read_comments.status, STATUS.ACCEPTED);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_comments', sheet: 'Notes' }).ok, true);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_comments' }).ok, false);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_comments', sheet: 'Notes', range: 'A1' }).ok, false);
  assert.deepEqual(xlsxRequest('read_comments', { sheet: 'Notes' }), { windows: [], commentSheets: ['Notes'] });
});

test('the driver: every anchor in row-major order -- also a comment-only cell outside the used range (K1 LIVE probe 3: used A1:C5, notes on E6 / J30)', () => {
  const at = {
    J30: { text: 'far', author: 'Far Author' },
    A1: { text: 'one', author: 'Excel Author' },
    E6: { text: 'edge', author: 'Edge Author' },
    B2: { text: 'threaded', author: 'Thread Owner', userId: 'owner@example.test', timeUtc: 1790848800000, replies: [{ text: 'reply', author: 'Thread Replier', userId: 'replier@example.test', timeUtc: 1790854200000 }] },
    D4: { text: 'solved', author: 'Thread Owner', userId: 'owner@example.test', timeUtc: 1790932500000, solved: true },
  };
  const r = run([sheet('Notes', at)], ['Notes']).Notes;
  assert.deepEqual(r.comments.map(c => c.cell), ['A1', 'B2', 'D4', 'E6', 'J30']);
  assert.deepEqual(r.comments[1], { cell: 'B2', text: 'threaded', author: 'Thread Owner', userId: 'owner@example.test', timeUtc: 1790848800000, solved: false,
    replies: [{ text: 'reply', author: 'Thread Replier', userId: 'replier@example.test', timeUtc: 1790854200000 }] });
  assert.strictEqual(r.comments[2].solved, true);
});

test('the driver: no comment -> an empty list without a scan; an asked sheet that does not exist is NAMED', () => {
  const ws = sheet('Plain', {});
  ws.GetRange = () => { throw new Error('a sheet with no comment must not be scanned'); };
  const r = run([ws], ['Plain', 'Nope']);
  assert.deepEqual(r.Plain, { comments: [] });
  assert.strictEqual(r.Nope, 'UNOBSERVABLE:sheet-missing');
});

test('the driver: an anchor the bounded scan does not reach names the WHOLE sheet -- never a partial list', () => {
  const r = run([sheet('Notes', { A1: { text: 'one', author: 'a' } }, { extra: 1 })], ['Notes']).Notes;
  assert.strictEqual(r, 'UNOBSERVABLE:comment-anchor-outside-scan:1/2:A1:Z200');
});

test('the driver: the scan rectangle grows with the used range and stays inside MAX_COMMENT_SCAN_CELLS', () => {
  // used A1:AH500 -> 44 columns (34 + 10), rows min(600, floor(20000 / 44) = 454) < 500 used rows -> the used range itself is not coverable
  const big = run([sheet('Big', { A1: { text: 'x', author: 'a' } }, { used: 'A1:AH500', extra: 1 })], ['Big']).Big;
  assert.strictEqual(big, `IMPLEMENTATION_LIMIT:comment-scan-over-${MAX_COMMENT_SCAN_CELLS}-cells`);
  // used A1:B300 -> 26 columns, rows max(200, 400) = 400 -> a comment on Z400 is found, the scan names A1:Z400 when short
  const at = { Z400: { text: 'corner', author: 'a' } };
  assert.deepEqual(run([sheet('Tall', at, { used: 'A1:B300' })], ['Tall']).Tall.comments.map(c => c.cell), ['Z400']);
  assert.strictEqual(run([sheet('Tall', at, { used: 'A1:B300', extra: 1 })], ['Tall']).Tall, 'UNOBSERVABLE:comment-anchor-outside-scan:1/2:A1:Z400');
});

test('the driver: a getter that throws is that field\'s NAMED reason, the row and the other fields stay', () => {
  const r = run([sheet('Notes', { A1: { text: 'one', author: 'a' } }, { throws: ['GetAuthorName'] })], ['Notes']).Notes;
  assert.deepEqual(r.comments[0], { cell: 'A1', text: 'one', author: 'UNOBSERVABLE:GetAuthorName boom', userId: '', timeUtc: 0, solved: false, replies: [] });
});

const inv = (reads, sheets = ['Notes', 'Plain']) => ({ sheets: sheets.map((name, index) => ({ name, index })), commentReads: reads });

test('the projection: resolved / replies; a stored-empty user id and a stored-absent time are null, never \'\' or the 1970 epoch', () => {
  const reads = { Notes: { comments: [
    { cell: 'A1', text: 'legacy note one', author: 'Excel Author', userId: '', timeUtc: 0, solved: false, replies: [] },
    { cell: 'B2', text: 'threaded one', author: 'Thread Owner', userId: 'owner@example.test', timeUtc: 1790848800000, solved: true,
      replies: [{ text: 'threaded reply', author: 'Thread Replier', userId: 'replier@example.test', timeUtc: 1790854200000 }] },
  ] } };
  assert.deepEqual(projectCommentsRead(inv(reads), { sheet: 'Notes' }), { sheet: 'Notes', comments: [
    { cell: 'A1', text: 'legacy note one', author: 'Excel Author', userId: null, timeUtc: null, resolved: false, replies: [] },
    { cell: 'B2', text: 'threaded one', author: 'Thread Owner', userId: 'owner@example.test', timeUtc: 1790848800000, resolved: true,
      replies: [{ text: 'threaded reply', author: 'Thread Replier', userId: 'replier@example.test', timeUtc: 1790854200000 }] },
  ] });
});

test('the projection: no comment -> []; a named sheet read -> comments null + the reason; not read -> NAMED; no such sheet -> SHEET_NOT_FOUND', () => {
  assert.deepEqual(projectCommentsRead(inv({ Plain: { comments: [] } }), { sheet: 'Plain' }), { sheet: 'Plain', comments: [] });
  assert.deepEqual(projectCommentsRead(inv({ Notes: 'UNOBSERVABLE:comment-anchor-outside-scan:1/2:A1:Z200' }), { sheet: 'Notes' }),
    { sheet: 'Notes', comments: null, reason: 'UNOBSERVABLE:comment-anchor-outside-scan:1/2:A1:Z200' });
  assert.deepEqual(projectCommentsRead(inv(null), { sheet: 'Notes' }), { sheet: 'Notes', comments: null, reason: 'UNOBSERVABLE:not-read' });
  assert.deepEqual(projectCommentsRead(inv({ Notes: { rows: [] } }), { sheet: 'Notes' }), { sheet: 'Notes', comments: null, reason: 'UNOBSERVABLE:malformed-read' });
  assert.deepEqual(projectCommentsRead(inv({}), { sheet: 'Nope' }), { sheet: 'Nope', error: 'SHEET_NOT_FOUND' });
});
