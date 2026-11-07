'use strict';

/**
 * GH #305 -- `document_create`: a NEW docx | xlsx | pptx, produced by DocBuilder and written to a target the caller names.
 * Owner decisions (owner channel, 2026-10-29/30): "A: DocBuilder creates, then co-editing edits"; the MCP is target
 * independent and runs no authorization round of its own -- the user writes wherever the target lets them; this holds for
 * every format.
 *
 *   1. the format's minimal seed (document-create-seeds.cjs) -> DocBuilder through the broker's `ds_job` op
 *      (OpenFile / SaveFile, returnDoc) -> the returned bytes, checked to be a package of the asked format;
 *   2. the bytes -> ONE target adapter:
 *        local      an absolute path on this host; an existing file is never replaced (exclusive create)
 *        nextcloud  a path in the requester's own files, through the broker's create-only `nc_webdav` PUT (the broker
 *                   sends If-None-Match: *, so an existing file answers 412); the answer carries the new fileId, which
 *                   the co-editing tools take next.
 *
 * What it does NOT do: overwrite, ask any authority whether the write is allowed (the target answers that), or accept a
 * requester identity (the requester is the token owner, as on every broker route). Never throws: every outcome is one
 * { ok, stage, code, ... } answer.
 */

const crypto = require('node:crypto');
const fsDefault = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { SEEDS } = require('./document-create-seeds.cjs');
const { IDENTITY_KEYS } = require('./broker-config-source.cjs');

const FORMATS = Object.freeze(['docx', 'xlsx', 'pptx']);
const TARGET_KINDS = Object.freeze(['local', 'nextcloud']);
const MAIN_PART = Object.freeze({ docx: 'word/document.xml', xlsx: 'xl/workbook.xml', pptx: 'ppt/presentation.xml' });
const LOOPBACK_URL = /^https?:\/\/(127\.0\.0\.1|localhost)(:\d+)?\//;

const refuse = (stage, code, reason, extra = {}) => ({ ok: false, stage, code, reason, ...extra });

/** The part names of a zip, from its central directory; null when the bytes are not a readable zip. */
function zipPartNames(buf) {
  const end = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  if (buf.length < 22 || buf.readUInt32LE(0) !== 0x04034b50 || end < 0 || end + 22 > buf.length) return null;
  const names = [];
  let p = buf.readUInt32LE(end + 16);
  for (let i = buf.readUInt16LE(end + 10); i > 0; i -= 1) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) return null;
    const n = buf.readUInt16LE(p + 28);
    names.push(buf.subarray(p + 46, p + 46 + n).toString('utf8'));
    p += 46 + n + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32);
  }
  return names;
}

const EXISTS = 'the target file already exists; it is never replaced';
const bad = reason => refuse('VALIDATE', 'BAD_REQUEST', reason);
const isObject = v => typeof v === 'object' && v !== null && !Array.isArray(v);
const others = (obj, allowed) => Object.keys(obj).filter(k => !allowed.includes(k));

/** the shape of `target.path` for its kind and format -> a refusal reason, or null */
function pathProblem(kind, p, format) {
  if (typeof p !== 'string' || !p || p.includes('\0') || p.includes('\\')) return 'target.path must be a non-empty path';
  if (path.extname(p).toLowerCase() !== `.${format}`) return `target.path must end in .${format}`;
  if (kind === 'local') return path.isAbsolute(p) ? null : 'a local target.path must be absolute';
  const relative = !p.startsWith('/') && p.split('/').every(seg => seg !== '' && seg !== '.' && seg !== '..');
  return relative ? null : 'a nextcloud target.path is relative to the requester\'s files root, without empty, "." or ".." segments';
}

/** the `target` object -> a refusal reason, or null */
function targetProblem(target, format) {
  if (!isObject(target)) return 'target must be an object';
  const identity = Object.keys(target).filter(k => IDENTITY_KEYS.test(k));
  if (identity.length) return `identity-bearing argument(s) refused: target.${identity.join(', target.')} -- the requester is the authenticated token owner`;
  const unknown = others(target, ['kind', 'path']);
  if (unknown.length) return `unknown target field(s): ${unknown.join(', ')}`;
  if (!TARGET_KINDS.includes(target.kind)) return `target.kind must be one of ${TARGET_KINDS.join(' | ')}`;
  return pathProblem(target.kind, target.path, format);
}

function validateRequest(args) {
  if (!isObject(args)) return bad('arguments must be an object');
  const extra = others(args, ['format', 'target']);
  if (extra.length) return bad(`unknown argument(s): ${extra.join(', ')}`);
  if (!FORMATS.includes(args.format)) return bad(`format must be one of ${FORMATS.join(' | ')}`);
  const problem = targetProblem(args.target, args.format);
  if (problem) return bad(problem);
  return { ok: true, format: args.format, target: { kind: args.target.kind, path: args.target.path } };
}

/** One broker call -> { status, body } | { denied } | { unavailable }. `ctx`: { brokerUrl, instance, timeoutMs, fetchImpl, token }. */
async function brokerCall(ctx, payload) {
  let bearer = '';
  try { bearer = ctx.token(); } catch { /* no own token: refused below */ }
  if (!bearer) return { denied: 'the agent identity (own token) is unavailable' };
  try {
    const r = await ctx.fetchImpl(ctx.brokerUrl, {
      method: 'POST',
      headers: { Authorization: `Bearer ${bearer}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ instance: ctx.instance, ...payload }),
      signal: AbortSignal.timeout(ctx.timeoutMs),
    });
    return { status: r.status, body: (await r.json().catch(() => null)) || {} };
  } catch (err) {
    const late = err && (err.name === 'TimeoutError' || err.name === 'AbortError');
    return { unavailable: late ? 'the broker did not answer in time' : 'the broker did not answer' };
  }
}

function brokerRefusal(stage, r) {
  if (r.denied) return refuse(stage, 'CONTROLLED_ACCESS_ERROR', r.denied);
  if (r.unavailable) return refuse(stage, 'BROKER_UNAVAILABLE', r.unavailable);
  const code = typeof r.body.code === 'string' ? ` ${r.body.code}` : '';
  let cls = 'BROKER_REQUEST_INVALID';
  if (r.status === 401 || r.status === 403) cls = 'CONTROLLED_ACCESS_ERROR';
  else if (r.status >= 500) cls = 'BROKER_UNAVAILABLE';
  return refuse(stage, cls, `broker HTTP ${r.status}${code}`);
}

/** DocBuilder (broker ds_job) on the format's seed -> { ok:true, bytes } | refusal */
async function build(ctx, format) {
  const script = `builder.OpenFile("__DOC_URL__");\nbuilder.SaveFile("${format}", "created.${format}");\nbuilder.CloseFile();\n`;
  const traceId = `document-create-${format}-${crypto.randomBytes(4).toString('hex')}`;
  const r = await brokerCall(ctx, { op: 'ds_job', script, documentBase64: SEEDS[format]().toString('base64'), returnDoc: true, traceId });
  if (r.status !== 200) return brokerRefusal('BUILD', r);
  // the job's ok/outcome is NOT the verdict (docx answers ok:false "unchanged" for a good save): the returned bytes are
  const b64 = r.body.documentBase64;
  if (typeof b64 !== 'string' || !b64) return refuse('BUILD', 'DOCBUILDER_NO_OUTPUT', `DocBuilder returned no document (outcome ${String(r.body.outcome)})`);
  const bytes = Buffer.from(b64, 'base64');
  const names = zipPartNames(bytes) || [];
  if (!names.includes('[Content_Types].xml') || !names.includes(MAIN_PART[format])) return refuse('BUILD', 'DOCBUILDER_OUTPUT_INVALID', `the returned bytes are not a ${format} package`);
  return { ok: true, bytes };
}

const LOCAL_ERRORS = Object.freeze({
  EEXIST: ['TARGET_EXISTS', EXISTS],
  ENOENT: ['TARGET_PARENT_MISSING', 'the target folder does not exist'],
  EACCES: ['ACCESS_DENIED', 'the target folder is not writable'],
  EPERM: ['ACCESS_DENIED', 'the target folder is not writable'],
  EROFS: ['ACCESS_DENIED', 'the target folder is not writable'],
});

/** exclusive create: an existing file -- even one that appeared since the pre-check -- is never replaced */
function writeLocal(fs, p, bytes) {
  try {
    fs.writeFileSync(p, bytes, { flag: 'wx' });
    return { ok: true, fields: {} };
  } catch (e) {
    const code = (e && e.code) || 'ERR';
    const [c, reason] = LOCAL_ERRORS[code] || ['TARGET_WRITE_FAILED', `local write failed (${code})`];
    return refuse('WRITE', c, reason);
  }
}

// the Nextcloud status the broker relays for a refused create-only PUT
const NC_ERRORS = Object.freeze({
  412: ['TARGET_EXISTS', EXISTS],
  409: ['TARGET_PARENT_MISSING', 'the target folder does not exist'],
  401: ['ACCESS_DENIED', 'Nextcloud refused the write'],
  403: ['ACCESS_DENIED', 'Nextcloud refused the write'],
});

/** the broker's create-only PUT, then the new fileId */
async function writeNextcloud(ctx, p, bytes) {
  const put = await brokerCall(ctx, { op: 'nc_webdav', root: 'files', method: 'PUT', path: p, dataBase64: bytes.toString('base64') });
  if (put.status !== 200) return brokerRefusal('WRITE', put);
  const s = put.body.status;
  if (s !== 201) {
    const [c, reason] = NC_ERRORS[s] || ['TARGET_WRITE_FAILED', `Nextcloud answered ${String(s)}`];
    return refuse('WRITE', c, reason);
  }
  const id = await brokerCall(ctx, { op: 'nc_webdav', root: 'files', method: 'PROPFIND', propfindKind: 'fileid', path: p });
  const raw = id.status === 200 ? id.body.fileid : undefined;
  // created, but without a fileId the caller cannot go on to co-editing: say so, do not hide it
  if (raw === undefined || raw === null) return { ok: true, fields: { fileId: null, fileIdUnavailable: 'created, but the fileId lookup did not answer' } };
  return { ok: true, fields: { fileId: String(raw) } };
}

const INPUT_SCHEMA = Object.freeze({
  type: 'object',
  properties: {
    format: { type: 'string', enum: [...FORMATS] },
    target: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: [...TARGET_KINDS] },
        path: { type: 'string', description: 'local: an absolute path; nextcloud: a path relative to your files root. Must end in .<format>.' },
      },
      required: ['kind', 'path'],
      additionalProperties: false,
    },
  },
  required: ['format', 'target'],
  additionalProperties: false,
});

/** everything decidable before DocBuilder runs -> the validated request, or a refusal */
function precheck(ctx, args) {
  const v = validateRequest(args);
  if (!v.ok) return v;
  if (v.target.kind === 'local' && ctx.fs.existsSync(v.target.path)) return refuse('VALIDATE', 'TARGET_EXISTS', EXISTS);
  if (typeof ctx.brokerUrl !== 'string' || !LOOPBACK_URL.test(ctx.brokerUrl)) return refuse('BUILD', 'BROKER_REQUEST_INVALID', 'no loopback broker configured');
  return v;
}

async function createDocument(ctx, args) {
  const v = precheck(ctx, args);
  if (!v.ok) return v;
  const built = await build(ctx, v.format);
  if (!built.ok) return built;
  const written = v.target.kind === 'local' ? writeLocal(ctx.fs, v.target.path, built.bytes) : await writeNextcloud(ctx, v.target.path, built.bytes);
  if (!written.ok) return written;
  const sha256 = crypto.createHash('sha256').update(built.bytes).digest('hex');
  return { ok: true, format: v.format, target: v.target, bytes: built.bytes.length, sha256, ...written.fields };
}

/**
 * createDocumentCreateTool({ brokerUrl, instance, timeoutMs, env, fetchImpl, readToken, fs }) -> { description, inputSchema, handler }.
 * `fetchImpl` / `readToken` / `fs`: test seams.
 */
function createDocumentCreateTool({ brokerUrl, instance = 'production', timeoutMs = 60000, env = process.env, fetchImpl = (...a) => globalThis.fetch(...a), readToken = null, fs = fsDefault } = {}) {
  const tokenPath = env.EURO_OFFICE_MCP_TOKEN_PATH || path.join(os.homedir(), '.agent-token');
  const ctx = { brokerUrl, instance, timeoutMs, fetchImpl, fs, token: readToken || (() => fs.readFileSync(tokenPath, 'utf8').trim()) };
  return {
    description: 'Create a NEW docx, xlsx or pptx with DocBuilder and write it to a local path or to a path in your Nextcloud files. '
      + 'Never replaces an existing file. A nextcloud target answers the new fileId, which the co-editing tools take for every further edit.',
    inputSchema: INPUT_SCHEMA,
    handler: args => createDocument(ctx, args),
  };
}

module.exports = { createDocumentCreateTool, validateRequest, zipPartNames, FORMATS, TARGET_KINDS };
