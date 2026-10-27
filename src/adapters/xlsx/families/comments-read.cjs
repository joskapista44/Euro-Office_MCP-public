'use strict';

/**
 * GH #300 K1: the PROJECTION behind the `read_comments` read tool -- a pure field-select over `inv.commentReads[sheet]`, the
 * driver's own in-session read (xlsx-driver.cjs readQuery commentSheets: the PUBLIC ApiComment getters, the anchor found by a
 * per-cell GetComment scan, reconciled against ws.GetComments()). One row per comment, in row-major cell order:
 *   { cell, text, author, userId, timeUtc, resolved, replies: [{ text, author, userId, timeUtc }] }
 * `cell` is the key: GetId() is NOT stable across a reopen (K1 LIVE probe 1), so no id is part of the row. A stored-empty user id
 * ('' -- a legacy note, or a comment made without one) and a stored-absent time (GetTimeUTC() 0 -- K1 probe 1: every API-made
 * comment) are null, never '' / the 1970 epoch. The legacy notes and the threaded comments read alike (K1 probes 2-3); the
 * "Author:" prefix of a legacy note is already stripped by the editor. A field the driver could not read stays its NAMED
 * 'UNOBSERVABLE:<reason>'.
 *   the sheet has no comment                -> { sheet, comments: [] }
 *   the read could not be done or reconciled -> { sheet, comments: null, reason: 'UNOBSERVABLE:...' | 'IMPLEMENTATION_LIMIT:...' }
 *                                              (never a partial list: an anchor the bounded scan did not reach names the sheet)
 */

const orNull = (v, absent) => (v === absent ? null : v);
const replyOf = r => ({ text: r.text, author: r.author, userId: orNull(r.userId, ''), timeUtc: orNull(r.timeUtc, 0) });

function projectCommentsRead(inv, a) {
  if (!inv.sheets.some(s => s.name === a.sheet)) return { sheet: a.sheet, error: 'SHEET_NOT_FOUND' };
  const read = inv.commentReads ? inv.commentReads[a.sheet] : undefined;
  if (read === undefined) return { sheet: a.sheet, comments: null, reason: 'UNOBSERVABLE:not-read' };
  if (typeof read === 'string') return { sheet: a.sheet, comments: null, reason: read };
  if (!read || !Array.isArray(read.comments)) return { sheet: a.sheet, comments: null, reason: 'UNOBSERVABLE:malformed-read' };
  return { sheet: a.sheet, comments: read.comments.map(c => ({ cell: c.cell, text: c.text, author: c.author, userId: orNull(c.userId, ''), timeUtc: orNull(c.timeUtc, 0),
    resolved: c.solved, replies: Array.isArray(c.replies) ? c.replies.map(replyOf) : c.replies })) };
}

module.exports = { projectCommentsRead };
