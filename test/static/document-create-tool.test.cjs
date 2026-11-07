'use strict';

/**
 * GH #305 `document_create` offline: the seeds, the request validation, the two target adapters and the fail-closed
 * answers, against a fake broker (records every call) and a real throwaway folder. The LIVE half is
 * test/live/gh305-document-create-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const zlib = require('node:zlib');
const { SEEDS, crc32 } = require('../../src/server/document-create-seeds.cjs');
const { createDocumentCreateTool, zipPartNames } = require('../../src/server/document-create-tool.cjs');

const BROKER = 'http://127.0.0.1:3420/api/euro-broker';
const MAIN = { docx: 'word/document.xml', xlsx: 'xl/workbook.xml', pptx: 'ppt/presentation.xml' };

/** the stored (method 0) parts of a seed zip, read through the local headers */
function storedParts(buf) {
  const out = {};
  for (let p = 0; buf.readUInt32LE(p) === 0x04034b50;) {
    const size = buf.readUInt32LE(p + 18); const n = buf.readUInt16LE(p + 26); const x = buf.readUInt16LE(p + 28);
    const data = buf.subarray(p + 30 + n + x, p + 30 + n + x + size);
    out[buf.subarray(p + 30, p + 30 + n).toString()] = { body: data.toString(), crc: buf.readUInt32LE(p + 14), data };
    p += 30 + n + x + size;
  }
  return out;
}

/** a fake broker: `answers(payload)` -> { status, body }; every payload is recorded */
function fakeBroker(answers) {
  const calls = [];
  const fetchImpl = async (url, init) => {
    const payload = JSON.parse(init.body);
    calls.push({ url, auth: init.headers.Authorization, payload });
    const a = answers(payload, calls.length);
    return { status: a.status, json: async () => a.body };
  };
  return { calls, fetchImpl };
}
const built = format => ({ status: 200, body: { ok: format !== 'docx', outcome: format === 'docx' ? 'nem-valtozott' : 'kesz', documentBase64: SEEDS[format]().toString('base64') } });
const tool = (broker, extra = {}) => createDocumentCreateTool({ brokerUrl: BROKER, fetchImpl: broker.fetchImpl, readToken: () => 'tok', ...extra });
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'gh305-'));

test('seeds: each format is a valid zip (CRC checked) with its main part, and carries no text, title or author', () => {
  assert.equal(crc32(Buffer.from('123456789')), 0xcbf43926, 'the CRC-32 check value');
  for (const format of ['docx', 'xlsx', 'pptx']) {
    const buf = SEEDS[format]();
    const names = zipPartNames(buf);
    assert.ok(names.includes('[Content_Types].xml') && names.includes(MAIN[format]), format);
    const parts = storedParts(buf);
    assert.deepEqual(Object.keys(parts), names, `${format}: every central-directory entry has a local part`);
    for (const [n, part] of Object.entries(parts)) {
      assert.equal(part.crc, crc32(part.data), `${format} ${n}: header CRC`);
      if (typeof zlib.crc32 === 'function') assert.equal(part.crc, zlib.crc32(part.data) >>> 0, `${format} ${n}: CRC against zlib`);
    }
    const all = Object.values(parts).map(x => x.body).join('');
    assert.doesNotMatch(all, /<w:t\b|<a:t\b|<v>|<dc:title|<dc:creator|docProps/, `${format}: structure only`);
    assert.deepEqual(SEEDS[format](), buf, `${format}: deterministic bytes`);
  }
  assert.match(storedParts(SEEDS.xlsx())['[Content_Types].xml'].body, /spreadsheetml\.worksheet\+xml/, 'the joined worksheet content type');
});

test('validation: bad format, extension, path shape, identity key or unknown field -> BAD_REQUEST before any broker call', async () => {
  const b = fakeBroker(() => assert.fail('no broker call expected'));
  const t = tool(b);
  const bad = [
    null, [], {}, { format: 'odt', target: { kind: 'local', path: '/tmp/a.odt' } },
    { format: 'docx', target: { kind: 'local', path: '/tmp/a.xlsx' } },
    { format: 'docx', target: { kind: 'local', path: 'rel/a.docx' } },
    { format: 'docx', target: { kind: 'nextcloud', path: '/abs/a.docx' } },
    { format: 'docx', target: { kind: 'nextcloud', path: 'a/../b.docx' } },
    { format: 'docx', target: { kind: 'nextcloud', path: 'a//b.docx' } },
    { format: 'docx', target: { kind: 'nextcloud', path: 'a\\b.docx' } },
    { format: 'docx', target: { kind: 's3', path: 'a.docx' } },
    { format: 'docx', target: { kind: 'nextcloud', path: 'a.docx', user: 'someone' } },
    { format: 'docx', target: { kind: 'nextcloud', path: 'a.docx', overwrite: true } },
    { format: 'docx', target: { kind: 'nextcloud', path: 'a.docx' }, requester: 'x' },
  ];
  for (const args of bad) {
    const r = await t.handler(args);
    assert.equal(r.ok, false, JSON.stringify(args));
    assert.equal(r.code, 'BAD_REQUEST', JSON.stringify(args));
  }
  assert.equal(b.calls.length, 0);
});

test('local: creates the file from the DocBuilder bytes; an existing path is refused before DocBuilder, and never replaced', async () => {
  const dir = tmp();
  try {
    for (const format of ['docx', 'xlsx', 'pptx']) {
      const b = fakeBroker(p => built(format === 'docx' && p.op === 'ds_job' ? 'docx' : format));
      const target = path.join(dir, `new.${format}`);
      const r = await tool(b).handler({ format, target: { kind: 'local', path: target } });
      assert.equal(r.ok, true, JSON.stringify(r));
      assert.deepEqual(fs.readFileSync(target), SEEDS[format](), `${format}: the returned bytes are the file`);
      assert.equal(b.calls.length, 1);
      const job = b.calls[0].payload;
      assert.equal(job.op, 'ds_job');
      assert.equal(job.returnDoc, true);
      assert.match(job.script, /builder\.OpenFile\("__DOC_URL__"\)/);
      assert.match(job.script, new RegExp(`SaveFile\\("${format}"`));
      assert.deepEqual(Buffer.from(job.documentBase64, 'base64'), SEEDS[format](), `${format}: the format's own seed is sent`);
      assert.equal(b.calls[0].auth, 'Bearer tok');
      // the same path again: refused, no DocBuilder job, the file unchanged
      const again = fakeBroker(() => assert.fail('no broker call for an existing target'));
      fs.writeFileSync(target, 'KEEP');
      const r2 = await tool(again).handler({ format, target: { kind: 'local', path: target } });
      assert.equal(r2.code, 'TARGET_EXISTS');
      assert.equal(fs.readFileSync(target, 'utf8'), 'KEEP');
      assert.equal(again.calls.length, 0);
    }
    const missing = await tool(fakeBroker(() => built('docx'))).handler({ format: 'docx', target: { kind: 'local', path: path.join(dir, 'no-such-dir', 'a.docx') } });
    assert.equal(missing.code, 'TARGET_PARENT_MISSING');
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

test('local: a file appearing between the check and the write is still not replaced (exclusive create)', async () => {
  const dir = tmp();
  try {
    const target = path.join(dir, 'race.docx');
    const b = fakeBroker(() => { fs.writeFileSync(target, 'OTHER'); return built('docx'); });
    const r = await tool(b).handler({ format: 'docx', target: { kind: 'local', path: target } });
    assert.equal(r.code, 'TARGET_EXISTS');
    assert.equal(fs.readFileSync(target, 'utf8'), 'OTHER');
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

test('nextcloud: a create-only PUT (no overwrite, no etag, no identity) then the fileId; 412 / 409 / 403 / broker refusal each answer their own code', async () => {
  const ok = fakeBroker(p => (p.op === 'ds_job' ? built('xlsx') : p.method === 'PUT' ? { status: 200, body: { ok: true, status: 201 } } : { status: 200, body: { ok: true, fileid: 1407739 } }));
  const r = await tool(ok).handler({ format: 'xlsx', target: { kind: 'nextcloud', path: 'Shared/new.xlsx' } });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.fileId, '1407739');
  assert.deepEqual(ok.calls.map(c => [c.payload.op, c.payload.method || null]), [['ds_job', null], ['nc_webdav', 'PUT'], ['nc_webdav', 'PROPFIND']]);
  const put = ok.calls[1].payload;
  assert.deepEqual(Object.keys(put).sort(), ['dataBase64', 'instance', 'method', 'op', 'path', 'root']);
  assert.equal(put.path, 'Shared/new.xlsx');
  assert.deepEqual(Buffer.from(put.dataBase64, 'base64'), SEEDS.xlsx());
  assert.equal(ok.calls[2].payload.propfindKind, 'fileid');

  const cases = [
    [{ status: 200, body: { ok: false, status: 412 } }, 'TARGET_EXISTS'],
    [{ status: 200, body: { ok: false, status: 409 } }, 'TARGET_PARENT_MISSING'],
    [{ status: 200, body: { ok: false, status: 403 } }, 'ACCESS_DENIED'],
    [{ status: 200, body: { ok: false, status: 507 } }, 'TARGET_WRITE_FAILED'],
    [{ status: 403, body: { code: 'EURO_BROKER_DENY_NC_WEBDAV_PUT_SCOPE' } }, 'CONTROLLED_ACCESS_ERROR'],
    [{ status: 502, body: {} }, 'BROKER_UNAVAILABLE'],
  ];
  for (const [answer, code] of cases) {
    const b = fakeBroker(p => (p.op === 'ds_job' ? built('docx') : answer));
    const out = await tool(b).handler({ format: 'docx', target: { kind: 'nextcloud', path: 'a.docx' } });
    assert.equal(out.ok, false);
    assert.equal(out.code, code, JSON.stringify(answer));
    assert.equal(b.calls.length, 2, 'no fileId lookup after a failed PUT');
  }
  const noId = fakeBroker(p => (p.op === 'ds_job' ? built('docx') : p.method === 'PUT' ? { status: 200, body: { ok: true, status: 201 } } : { status: 502, body: {} }));
  const created = await tool(noId).handler({ format: 'docx', target: { kind: 'nextcloud', path: 'a.docx' } });
  assert.equal(created.ok, true);
  assert.equal(created.fileId, null);
  assert.match(created.fileIdUnavailable, /created/);
});

test('BUILD: the returned bytes decide, not the job flag -- no output, non-zip, or another format\'s package fail closed; nothing is written', async () => {
  const dir = tmp();
  try {
    const target = path.join(dir, 'x.docx');
    const cases = [
      [{ status: 200, body: { ok: true, outcome: 'kesz' } }, 'DOCBUILDER_NO_OUTPUT'],
      [{ status: 200, body: { ok: true, documentBase64: Buffer.from('not a zip at all, long enough').toString('base64') } }, 'DOCBUILDER_OUTPUT_INVALID'],
      [{ status: 200, body: { ok: true, documentBase64: SEEDS.xlsx().toString('base64') } }, 'DOCBUILDER_OUTPUT_INVALID'],
      [{ status: 403, body: { code: 'EURO_BROKER_DENY_PRINCIPAL' } }, 'CONTROLLED_ACCESS_ERROR'],
      [{ status: 500, body: { code: 'EURO_BROKER_JOB_FAILED' } }, 'BROKER_UNAVAILABLE'],
    ];
    for (const [answer, code] of cases) {
      const r = await tool(fakeBroker(() => answer)).handler({ format: 'docx', target: { kind: 'local', path: target } });
      assert.equal(r.code, code, JSON.stringify(answer));
      assert.equal(r.stage, 'BUILD');
      assert.equal(fs.existsSync(target), false);
    }
    // POS control for the same gate: docx ok:false "nem-valtozott" WITH a docx package is a success
    const r = await tool(fakeBroker(() => built('docx'))).handler({ format: 'docx', target: { kind: 'local', path: target } });
    assert.equal(r.ok, true);
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});

test('no token or no loopback broker -> refused, no network call', async () => {
  const b = fakeBroker(() => assert.fail('no call'));
  const noToken = await tool(b, { readToken: () => '' }).handler({ format: 'docx', target: { kind: 'nextcloud', path: 'a.docx' } });
  assert.equal(noToken.code, 'CONTROLLED_ACCESS_ERROR');
  const remote = await tool(b, { brokerUrl: 'https://broker.example/api' }).handler({ format: 'docx', target: { kind: 'nextcloud', path: 'a.docx' } });
  assert.equal(remote.code, 'BROKER_REQUEST_INVALID');
  assert.equal(b.calls.length, 0);
});
