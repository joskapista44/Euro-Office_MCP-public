'use strict';

/**
 * M12 reference broker -- OFFLINE acceptance (owner decision D4 = B, owner channel 2026-10-12).
 *
 * Exercises the SPEC's acceptance list (M12-REFERENCE-BROKER-CONTRACT.md §7, items 1-9) with the
 * product's REAL client (`createBrokerConfigSource`) talking, over real loopback HTTP, to the
 * reference broker (`reference-broker/server.cjs`, also real) in front of a fake Nextcloud (the
 * only fake in the chain). No fetch-level stubbing of the broker itself: both servers are real
 * `node:http` servers on ephemeral loopback ports.
 *
 * Clean-room: written from the spec and the product's own client/config-source modules only.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const { loadConfig, createServer } = require('../../reference-broker/server.cjs');
const { createBrokerConfigSource, ERRORS, validateTarget } = require('../../src/server/broker-config-source.cjs');

// ---- fixtures -------------------------------------------------------------------------------------------------

/** The measured LIVE key structure (product_ocs_config), synthetic values, bound to `fileId`. */
function editConfig(fileId, { edit = true, mode = 'edit' } = {}) {
  return {
    document: { fileType: 'docx', key: `KEY-${fileId}`, permissions: { edit, protect: true, changeHistory: true }, title: 'fixture.docx',
      url: 'https://ds.example/download', referenceData: { fileKey: fileId, instanceId: 'instfixture' }, info: { favorite: false } },
    documentType: 'word',
    editorConfig: { mode, lang: 'hu', region: 'hu-HU', callbackUrl: 'https://nc.example/callback', user: { id: 'instfixture_u', name: 'u', image: '' },
      createUrl: 'https://nc.example/create', customization: { goback: { url: 'https://nc.example/' }, compactHeader: true, feedback: false, uiTheme: 'default' } },
    documentServerUrl: 'https://ds.example/', iat: 1, exp: 2, token: 'ds-jwt-FIXTURE',
  };
}
const ocsEnvelope = (config) => ({ ocs: { meta: { status: 'ok', statuscode: 100, message: 'OK' }, data: config } });

/** A listening ephemeral-port fake Nextcloud; `setHandler` decides the answer per call, `calls` records every request. */
function fakeNextcloud() {
  const calls = [];
  let handler = ({ fileId }) => ({ status: 200, body: ocsEnvelope(editConfig(fileId)) });
  const server = http.createServer((req, res) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      const authz = req.headers.authorization || '';
      const m = /^Basic (.+)$/.exec(authz);
      const [user, pass] = m ? Buffer.from(m[1], 'base64').toString('utf8').split(':') : [null, null];
      const fm = /\/config\/([^/?]+)/.exec(req.url);
      const call = { user, pass, fileId: fm ? fm[1] : null, url: req.url, authzHeader: authz, apiReqHeader: req.headers['ocs-apirequest'] };
      calls.push(call);
      const outcome = handler(call);
      if (outcome.noResponse) return; // never answers -- exercises the broker's own timeout
      if (outcome.destroySocket) { req.socket.destroy(); return; }
      const text = outcome.raw !== undefined ? outcome.raw : JSON.stringify(outcome.body);
      res.writeHead(outcome.status, { 'Content-Type': 'application/json' });
      res.end(text);
    });
  });
  return { server, calls, setHandler: (fn) => { handler = fn; } };
}

/**
 * Starts a fake Nextcloud + a reference broker in front of it, with the given requester specs
 * (`{ id, nextcloudUser, appPassword }`; `appPassword: null` writes no app-password file at all).
 * Returns `{ source(id), nc, brokerUrl, cleanup() }`. Everything lives under a disposable tmp dir.
 */
async function harness(requesterSpecs, { brokerTimeoutMs = 2000 } = {}) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ref-broker-test-'));
  const tokens = {};
  const requesters = [];
  for (const spec of requesterSpecs) {
    const token = `tok-${spec.id}-${crypto.randomBytes(8).toString('hex')}`;
    tokens[spec.id] = token;
    const pwFile = path.join(tmp, `${spec.id}.app-password`);
    if (spec.appPassword !== null) fs.writeFileSync(pwFile, spec.appPassword === undefined ? 'app-pw-fixture' : spec.appPassword, { mode: 0o600 });
    requesters.push({
      id: spec.id,
      tokenSha256: crypto.createHash('sha256').update(token, 'utf8').digest('hex'),
      nextcloudUser: spec.nextcloudUser || spec.id,
      appPasswordFile: pwFile,
    });
  }

  const nc = fakeNextcloud();
  await new Promise((resolve) => nc.server.listen(0, '127.0.0.1', resolve));
  const ncPort = nc.server.address().port;

  const config = loadConfig({
    listen: { host: '127.0.0.1', port: 0 },
    path: '/api/euro-broker',
    instances: { production: { nextcloudUrl: `http://127.0.0.1:${ncPort}` } },
    requesters,
    timeoutMs: brokerTimeoutMs,
  });
  const logLines = [];
  const broker = createServer({ config, log: (fields) => logLines.push(fields) });
  await new Promise((resolve) => broker.listen(0, '127.0.0.1', resolve));
  const brokerPort = broker.address().port;
  const brokerUrl = `http://127.0.0.1:${brokerPort}${config.path}`;

  const source = (id, extra = {}) => createBrokerConfigSource({ brokerUrl, instance: 'production', readToken: () => tokens[id], timeoutMs: 5000, ...extra });

  async function cleanup() {
    if (broker.listening) await new Promise((resolve) => broker.close(resolve));
    if (nc.server.listening) await new Promise((resolve) => nc.server.close(resolve));
    fs.rmSync(tmp, { recursive: true, force: true });
  }
  return { source, tokens, nc, brokerUrl, logLines, cleanup, tmp };
}

// ---- 1: own READ/WRITE against an edit-capable config; WRITE against a view-only config ------------------------

test('1a: own READ -> config, edit-capable config passed through unmodified', async () => {
  const h = await harness([{ id: 'alice' }]);
  try {
    const r = await h.source('alice').requestEditorConfig('42', 'READ');
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.config, editConfig('42'));
  } finally { await h.cleanup(); }
});

test('1b: own WRITE against an edit-capable config -> config', async () => {
  const h = await harness([{ id: 'alice' }]);
  try {
    const r = await h.source('alice').requestEditorConfig('42', 'WRITE');
    assert.equal(r.ok, true, JSON.stringify(r));
  } finally { await h.cleanup(); }
});

test('1c: WRITE against a view-only config -> ACCESS_DENIED; READ against the SAME config still passes through', async () => {
  const h = await harness([{ id: 'alice' }]);
  h.nc.setHandler(({ fileId }) => ({ status: 200, body: ocsEnvelope(editConfig(fileId, { edit: false, mode: 'view' })) }));
  try {
    const w = await h.source('alice').requestEditorConfig('42', 'WRITE');
    assert.deepEqual([w.ok, w.code], [false, ERRORS.ACCESS_DENIED]);
    const r = await h.source('alice').requestEditorConfig('42', 'READ');
    assert.equal(r.ok, true, 'READ is not narrowed even though the config is view-only');
    assert.equal(r.config.editorConfig.mode, 'view', 'the mode is echoed unmodified, not rewritten');
  } finally { await h.cleanup(); }
});

test('1d: a config with no ocs envelope at all is accepted (the body itself is the config)', async () => {
  const h = await harness([{ id: 'alice' }]);
  h.nc.setHandler(({ fileId }) => ({ status: 200, body: editConfig(fileId) }));
  try {
    const r = await h.source('alice').requestEditorConfig('42', 'READ');
    assert.equal(r.ok, true, JSON.stringify(r));
  } finally { await h.cleanup(); }
});

// ---- 2: Nextcloud 404 and 403 -> the SAME ACCESS_DENIED answer --------------------------------------------------

test('2: Nextcloud 404 and 403 both map to ACCESS_DENIED (missing == not allowed, no probing)', async () => {
  const h = await harness([{ id: 'alice' }]);
  try {
    for (const status of [404, 403]) {
      h.nc.setHandler(() => ({ status, body: { ocs: { meta: { statuscode: status } } } }));
      const r = await h.source('alice').requestEditorConfig('42', 'READ');
      assert.deepEqual([r.ok, r.code], [false, ERRORS.ACCESS_DENIED], `HTTP ${status}`);
    }
    // the same verdict when the OCS envelope itself carries the failing statuscode on an HTTP 200
    for (const metaCode of [403, 404, 997, 998]) {
      h.nc.setHandler(() => ({ status: 200, body: { ocs: { meta: { statuscode: metaCode } } } }));
      const r = await h.source('alice').requestEditorConfig('42', 'READ');
      assert.deepEqual([r.ok, r.code], [false, ERRORS.ACCESS_DENIED], `meta.statuscode ${metaCode}`);
    }
    assert.equal(h.nc.calls.length, 6, 'exactly one upstream call per request -- no retry on a denial, with any other credential');
  } finally { await h.cleanup(); }
});

// ---- 3: two requesters, each only ever reaches Nextcloud with its OWN credential --------------------------------

test('3: alice and bob each reach Nextcloud with only their own nextcloudUser/app-password', async () => {
  const h = await harness([{ id: 'alice', appPassword: 'alice-secret' }, { id: 'bob', nextcloudUser: 'robert', appPassword: 'bob-secret' }]);
  try {
    await h.source('alice').requestEditorConfig('1', 'READ');
    await h.source('bob').requestEditorConfig('2', 'READ');
    assert.equal(h.nc.calls.length, 2, 'one upstream call per request, no retry');
    assert.deepEqual([h.nc.calls[0].user, h.nc.calls[0].pass], ['alice', 'alice-secret']);
    assert.deepEqual([h.nc.calls[1].user, h.nc.calls[1].pass], ['robert', 'bob-secret']);
  } finally { await h.cleanup(); }
});

// ---- 4: identity-bearing body keys -> 400 ------------------------------------------------------------------------

test('4: a body carrying user/uid/requester (bypassing the client) is refused 400 BAD_BODY; the client itself never sends one', async () => {
  const h = await harness([{ id: 'alice' }]);
  try {
    for (const extra of [{ user: 'bob' }, { uid: 'bob' }, { requester: 'bob' }, { credential: 'x' }]) {
      const res = await fetch(h.brokerUrl, {
        method: 'POST', headers: { Authorization: `Bearer ${h.tokens.alice}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ op: 'product_ocs_config', instance: 'production', fileId: '42', requestedMode: 'READ', ...extra }),
      });
      const json = await res.json();
      assert.deepEqual([res.status, json], [400, { ok: false, code: 'PRODUCT_OCS_CONFIG_BAD_BODY' }], JSON.stringify(extra));
    }
    assert.equal(h.nc.calls.length, 0, 'none of the malformed bodies ever reached Nextcloud');
    // the product's OWN second layer (validateTarget) refuses the same identity keys before any broker call at all
    for (const k of ['user', 'uid', 'requester']) assert.equal(validateTarget({ fileId: '42', [k]: 'bob' }, 'READ').ok, false, k);
  } finally { await h.cleanup(); }
});

// ---- 5: unknown/missing token, missing app-password file -> CONTROLLED_ACCESS_ERROR -------------------------------

test('5a: an unknown bearer token -> CONTROLLED_ACCESS_ERROR (broker 401 AUTHZ_DENY_NO_AUTH)', async () => {
  const h = await harness([{ id: 'alice' }]);
  try {
    const r = await h.source('alice', { readToken: () => 'tok-not-registered-anywhere' }).requestEditorConfig('42', 'READ');
    assert.deepEqual([r.ok, r.code], [false, ERRORS.CONTROLLED_ACCESS_ERROR]);
    assert.equal(h.nc.calls.length, 0);
  } finally { await h.cleanup(); }
});

test('5b: a requester with a missing/empty app-password file -> CONTROLLED_ACCESS_ERROR, no Nextcloud call', async () => {
  const h = await harness([{ id: 'alice', appPassword: null }]);
  try {
    const r = await h.source('alice').requestEditorConfig('42', 'READ');
    assert.deepEqual([r.ok, r.code], [false, ERRORS.CONTROLLED_ACCESS_ERROR]);
    assert.equal(h.nc.calls.length, 0, 'no Nextcloud call is made without a usable app password');
  } finally { await h.cleanup(); }
});

// ---- 6: Nextcloud down / 500 / timeout -> BROKER_UNAVAILABLE -----------------------------------------------------

test('6a: Nextcloud down (connection refused) -> BROKER_UNAVAILABLE', async () => {
  const h = await harness([{ id: 'alice' }]);
  await new Promise((resolve) => h.nc.server.close(resolve));
  try {
    const r = await h.source('alice').requestEditorConfig('42', 'READ');
    assert.deepEqual([r.ok, r.code], [false, ERRORS.BROKER_UNAVAILABLE]);
  } finally { await h.cleanup(); }
});

test('6b: Nextcloud 500 -> BROKER_UNAVAILABLE', async () => {
  const h = await harness([{ id: 'alice' }]);
  h.nc.setHandler(() => ({ status: 500, body: { ocs: { meta: { statuscode: 500 } } } }));
  try {
    const r = await h.source('alice').requestEditorConfig('42', 'READ');
    assert.deepEqual([r.ok, r.code], [false, ERRORS.BROKER_UNAVAILABLE]);
    assert.equal(h.nc.calls.length, 1, 'no retry on an upstream 500');
  } finally { await h.cleanup(); }
});

test('6c: Nextcloud never answers -> the broker\'s own upstream timeout -> BROKER_UNAVAILABLE', async () => {
  const h = await harness([{ id: 'alice' }], { brokerTimeoutMs: 150 });
  h.nc.setHandler(() => ({ noResponse: true }));
  try {
    const r = await h.source('alice').requestEditorConfig('42', 'READ');
    assert.deepEqual([r.ok, r.code], [false, ERRORS.BROKER_UNAVAILABLE]);
  } finally { await h.cleanup(); }
});

test('6d: NO_CREDENTIAL_FALLBACK -- a failing Nextcloud call for alice is never retried with bob\'s credential, even though bob\'s would succeed', async () => {
  const h = await harness([{ id: 'alice', appPassword: 'alice-secret' }, { id: 'bob', nextcloudUser: 'robert', appPassword: 'bob-secret' }]);
  h.nc.setHandler(({ user, fileId }) => (user === 'alice' ? { status: 500, body: { ocs: { meta: { statuscode: 500 } } } } : { status: 200, body: ocsEnvelope(editConfig(fileId)) }));
  try {
    const r = await h.source('alice').requestEditorConfig('42', 'READ');
    assert.deepEqual([r.ok, r.code], [false, ERRORS.BROKER_UNAVAILABLE]);
    assert.deepEqual(h.nc.calls.map((c) => c.user), ['alice'], 'exactly one call, with alice\'s own credential -- never a second call, never bob\'s');
  } finally { await h.cleanup(); }
});

// ---- fail-closed hygiene (review): unmapped Nextcloud status / meta.statuscode never fall through -----

test('6e: an unmapped Nextcloud HTTP status (not 200/401/403/404/5xx) never falls through as a success -- isolated from the meta.statuscode gate (no envelope at all)', async () => {
  const h = await harness([{ id: 'alice' }]);
  try {
    for (const status of [302, 400, 410]) {
      // no `ocs` envelope at all, so an unwrapped body would otherwise look like a (malformed) success -- only the
      // raw HTTP status itself must stop it, isolating this from the separate meta.statuscode check (6f).
      h.nc.setHandler(({ fileId }) => ({ status, body: editConfig(fileId) }));
      const r = await h.source('alice').requestEditorConfig('42', 'READ');
      assert.deepEqual([r.ok, r.code], [false, ERRORS.BROKER_UNAVAILABLE], `HTTP ${status}`);
    }
  } finally { await h.cleanup(); }
});

test('6f: an OCS meta.statuscode outside 100/200 and the named deny codes never falls through as a success', async () => {
  const h = await harness([{ id: 'alice' }]);
  h.nc.setHandler(({ fileId }) => ({ status: 200, body: { ocs: { meta: { statuscode: 996 }, data: editConfig(fileId) } } }));
  try {
    const r = await h.source('alice').requestEditorConfig('42', 'READ');
    assert.deepEqual([r.ok, r.code], [false, ERRORS.BROKER_UNAVAILABLE]);
  } finally { await h.cleanup(); }
});

// ---- Cache-Control: no-store on every answer (the 200 answer carries the editor config's DS token) ----------------

test('Cache-Control: no-store on every broker answer, success and refusal alike', async () => {
  const h = await harness([{ id: 'alice' }]);
  try {
    const okRes = await fetch(h.brokerUrl, { method: 'POST', headers: { Authorization: `Bearer ${h.tokens.alice}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ op: 'product_ocs_config', instance: 'production', fileId: '42', requestedMode: 'READ' }) });
    assert.equal(okRes.headers.get('cache-control'), 'no-store');
    const badRes = await fetch(h.brokerUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ op: 'product_ocs_config', instance: 'production', fileId: '42', requestedMode: 'READ' }) });
    assert.equal(badRes.headers.get('cache-control'), 'no-store');
    const notFoundRes = await fetch(h.brokerUrl.replace('/api/euro-broker', '/other'), { method: 'POST' });
    assert.equal(notFoundRes.headers.get('cache-control'), 'no-store');
  } finally { await h.cleanup(); }
});

// ---- CLI entry point: node reference-broker/server.cjs --config <path> ---------------------------------------------

test('CLI: no --config -> usage on stderr, exit 78, no listener', async () => {
  const { execFile } = require('node:child_process');
  const bin = path.join(__dirname, '..', '..', 'reference-broker', 'server.cjs');
  const result = await new Promise((resolve) => execFile(process.execPath, [bin], (error, stdout, stderr) => resolve({ code: error ? error.code : 0, stdout, stderr })));
  assert.equal(result.code, 78);
  assert.match(result.stderr, /--config/);
  assert.equal(result.stdout, '');
});

test('CLI: an invalid config file -> a one-line reason on stderr, exit 78, no listener', async () => {
  const { execFile } = require('node:child_process');
  const bin = path.join(__dirname, '..', '..', 'reference-broker', 'server.cjs');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ref-broker-cli-'));
  const cfgPath = path.join(tmp, 'bad.json');
  fs.writeFileSync(cfgPath, JSON.stringify({ listen: { host: '0.0.0.0', port: 0 }, path: '/x', instances: {}, requesters: [] }));
  try {
    const result = await new Promise((resolve) => execFile(process.execPath, [bin, '--config', cfgPath], (error, stdout, stderr) => resolve({ code: error ? error.code : 0, stdout, stderr })));
    assert.equal(result.code, 78);
    assert.match(result.stderr, /invalid config/);
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
});

test('CLI: a valid config starts a listener and prints its bound address to stderr', async () => {
  const { spawn } = require('node:child_process');
  const bin = path.join(__dirname, '..', '..', 'reference-broker', 'server.cjs');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ref-broker-cli-ok-'));
  const pwFile = path.join(tmp, 'a.pw'); fs.writeFileSync(pwFile, 'pw');
  const cfgPath = path.join(tmp, 'ok.json');
  fs.writeFileSync(cfgPath, JSON.stringify({ listen: { host: '127.0.0.1', port: 0 }, path: '/api/euro-broker', instances: { production: { nextcloudUrl: 'https://nc.example' } }, requesters: [{ id: 'a', tokenSha256: 'a'.repeat(64), nextcloudUser: 'a', appPasswordFile: pwFile }], timeoutMs: 1000 }));
  const child = spawn(process.execPath, [bin, '--config', cfgPath]);
  let stderr = '';
  child.stderr.on('data', (c) => { stderr += c.toString(); });
  try {
    await new Promise((resolve, reject) => {
      const t = setTimeout(() => reject(new Error(`no startup line within 3s; stderr so far: ${stderr}`)), 3000);
      const check = () => { if (/listening on 127\.0\.0\.1:\d+\/api\/euro-broker/.test(stderr)) { clearTimeout(t); resolve(); } };
      child.stderr.on('data', check);
      check();
    });
  } finally {
    child.kill();
    fs.rmSync(tmp, { recursive: true, force: true });
  }
});

// ---- 7: refuses to start on a non-loopback host or an unknown config key -----------------------------------------

test('7: refuses to start on a non-loopback listen.host', () => {
  assert.throws(() => loadConfig({ listen: { host: '0.0.0.0', port: 3421 }, path: '/api/euro-broker', instances: { production: { nextcloudUrl: 'https://nc.example' } }, requesters: [], timeoutMs: 1000 }), /loopback/);
});

test('7: refuses to start on an unknown config key, at every level', () => {
  const base = () => ({ listen: { host: '127.0.0.1', port: 0 }, path: '/api/euro-broker', instances: { production: { nextcloudUrl: 'https://nc.example' } }, requesters: [], timeoutMs: 1000 });
  assert.throws(() => loadConfig({ ...base(), extra: 1 }), /unknown config key/);
  assert.throws(() => loadConfig({ ...base(), listen: { host: '127.0.0.1', port: 0, extra: 1 } }), /unknown config key/);
  assert.throws(() => loadConfig({ ...base(), instances: { production: { nextcloudUrl: 'https://nc.example', extra: 1 } } }), /unknown config key/);
  assert.throws(() => loadConfig({ ...base(), requesters: [{ id: 'a', tokenSha256: 'a'.repeat(64), nextcloudUser: 'a', appPasswordFile: '/x', extra: 1 }] }), /unknown config key/);
});

// ---- 8: secret scan -- no token, app password or Basic header in any response body or log line --------------------

test('8: secret scan -- the bearer token, the app password and a Basic header never appear in any response body or log line', async () => {
  const h = await harness([{ id: 'alice', appPassword: 'the-app-password-FIXTURE' }]);
  try {
    const bodies = [];
    const call = async (fileIdOrBody, mode) => {
      if (typeof fileIdOrBody === 'object') {
        const res = await fetch(h.brokerUrl, { method: 'POST', headers: { Authorization: `Bearer ${h.tokens.alice}`, 'Content-Type': 'application/json' }, body: JSON.stringify(fileIdOrBody) });
        bodies.push(await res.text());
        return;
      }
      bodies.push(JSON.stringify(await h.source('alice').requestEditorConfig(fileIdOrBody, mode)));
    };
    await call('42', 'READ');
    await call('42', 'WRITE');
    await call({ op: 'bogus' });
    await call({ op: 'product_ocs_config', instance: 'production', fileId: '42', requestedMode: 'READ', user: 'bob' });
    h.nc.setHandler(() => ({ status: 403, body: { ocs: { meta: { statuscode: 403 } } } }));
    await call('42', 'READ');
    h.nc.setHandler(() => ({ status: 500, body: {} }));
    await call('42', 'READ');
    await h.source('alice', { readToken: () => 'tok-unknown' }).requestEditorConfig('42', 'READ').then((r) => bodies.push(JSON.stringify(r)));

    const haystack = bodies.join('\n') + '\n' + JSON.stringify(h.logLines);
    assert.equal(haystack.includes(h.tokens.alice), false, 'bearer token leaked');
    assert.equal(haystack.includes('the-app-password-FIXTURE'), false, 'app password leaked');
    assert.equal(/Basic [A-Za-z0-9+/=]/.test(haystack), false, 'a Basic header leaked');
    for (const line of h.logLines) assert.deepEqual(Object.keys(line).sort(), ['code', 'fileId', 'requester', 'requestedMode', 'status'].sort(), JSON.stringify(line));
  } finally { await h.cleanup(); }
});

// ---- request-shape coverage (spec §3, beyond the numbered 1-9 but exercised by the same request handler) --------

test('request shape: oversize body, non-JSON body, wrong op/instance/fileId/requestedMode, other method/path all 400/404', async () => {
  const h = await harness([{ id: 'alice' }]);
  try {
    const okBody = { op: 'product_ocs_config', instance: 'production', fileId: '42', requestedMode: 'READ' };
    const post = (body, headers = {}) => fetch(h.brokerUrl, { method: 'POST', headers: { Authorization: `Bearer ${h.tokens.alice}`, 'Content-Type': 'application/json', ...headers }, body });

    let r = await post('x'.repeat(5000));
    assert.equal(r.status, 400);
    r = await post('{not json');
    assert.equal(r.status, 400);
    r = await post(JSON.stringify({ ...okBody, op: 'other_op' }));
    assert.deepEqual([r.status, (await r.json()).code], [400, 'EURO_BROKER_DENY_UNKNOWN_OP']);
    r = await post(JSON.stringify({ ...okBody, instance: 'staging' }));
    assert.deepEqual([r.status, (await r.json()).code], [400, 'EURO_BROKER_DENY_INSTANCE']);
    r = await post(JSON.stringify({ ...okBody, fileId: 'abc' }));
    assert.deepEqual([r.status, (await r.json()).code], [400, 'PRODUCT_OCS_CONFIG_BAD_BODY']);
    r = await post(JSON.stringify({ ...okBody, requestedMode: 'DELETE' }));
    assert.deepEqual([r.status, (await r.json()).code], [400, 'PRODUCT_OCS_CONFIG_BAD_BODY']);
    r = await fetch(h.brokerUrl.replace('/api/euro-broker', '/other'), { method: 'POST', body: JSON.stringify(okBody) });
    assert.equal(r.status, 404);
    r = await fetch(h.brokerUrl, { method: 'GET' });
    assert.equal(r.status, 404);
    assert.equal(h.nc.calls.length, 0, 'none of these malformed requests ever reached Nextcloud');
  } finally { await h.cleanup(); }
});

test('auth shape: no Authorization header -> 401 AUTHZ_DENY_NO_AUTH', async () => {
  const h = await harness([{ id: 'alice' }]);
  try {
    const r = await fetch(h.brokerUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ op: 'product_ocs_config', instance: 'production', fileId: '42', requestedMode: 'READ' }) });
    assert.deepEqual([r.status, (await r.json()).code], [401, 'AUTHZ_DENY_NO_AUTH']);
  } finally { await h.cleanup(); }
});
