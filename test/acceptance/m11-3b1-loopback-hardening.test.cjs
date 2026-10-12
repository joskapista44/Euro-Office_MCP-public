'use strict';

/**
 * M11.3b.1 -- LOOPBACK CONFIG HOST HARDENING (owner decision LOOPBACK_HARDENING=REQUIRED, owner channel 2026-10-09).
 *
 * The editor page embeds the full editor config, including the Document Server JWT. Proven here, with real HTTP requests
 * against the real host (src/drivers/eurooffice/config-host.cjs) and through both real drivers:
 *   EVIDENCE   the pre-hardening host (verbatim, test/fixtures/legacy-config-host.cjs) served the JWT on ANY path;
 *   VALID      the session's own URL (the one the driver navigates to) serves the page with the config;
 *   DENY       no token, a wrong token, a previous session's token, a closed session's token, a foreign Host / Origin,
 *              any method but GET/HEAD -- never the config, never a CORS header;
 *   ENTROPY    256-bit tokens, all distinct;
 *   NO LEAK    the token never reaches the driver's error text (a navigation error carries the URL -> redacted);
 *   LIFECYCLE  both drivers still open/close as before; the host is dead after close.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');

const { startConfigHost, TOKEN_BYTES } = require('../../src/drivers/eurooffice/config-host.cjs');
const { startHost: legacyStartHost } = require('../fixtures/legacy-config-host.cjs');
const docx = require('../../src/drivers/eurooffice/docx-driver.cjs');
const xlsx = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

const JWT = 'eyJ.FIXTURE-DS-JWT.sig';
const CONFIG = Object.freeze({ documentServerUrl: 'https://ds.example/', document: { key: 'KEY-1302148', referenceData: { fileKey: '1302148' } }, token: JWT });

/** A raw request with full header control (fetch forbids Host). -> { status, headers, body } | { error } */
function request(url, { method = 'GET', headers = {} } = {}) {
  const u = new URL(url);
  return new Promise(resolve => {
    const req = http.request({ host: u.hostname, port: u.port, path: `${u.pathname}${u.search}`, method, headers, agent: false }, res => {
      let body = '';
      res.on('data', d => { body += d; });
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body }));
    });
    req.on('error', e => resolve({ error: e.code || String(e) }));
    req.end();
  });
}
const origin = url => new URL(url).origin;
const at = (url, path) => `${origin(url)}${path}`;
const tokenOf = url => /\/s\/([0-9a-f]+)\//.exec(url)[1];
const noConfig = r => assert.equal(String(r.body || '').includes(JWT), false, 'no JWT / config in the answer');
const noCors = r => assert.equal(Object.keys(r.headers || {}).some(h => h.startsWith('access-control-')), false, 'no CORS header');

// ---- evidence: the pre-hardening behaviour ---------------------------------------------------------------------------

test('EVIDENCE (pre-hardening): the legacy host served the full config incl. the JWT on ANY path, to ANY client', async () => {
  const h = await legacyStartHost(CONFIG);
  try {
    for (const p of ['/', '/anything', '/s/0000/']) {
      const r = await request(at(h.url, p));
      assert.equal(r.status, 200, p);
      assert.ok(r.body.includes(JWT), `the legacy host leaks the JWT on ${p}`);
    }
  } finally { await new Promise(r => h.server.close(r)); }
});

// ---- the hardened host ------------------------------------------------------------------------------------------------

test('VALID_SESSION_TOKEN: the session URL serves the editor page with the config; no-store, no-referrer, loopback only', async () => {
  const h = await startConfigHost(CONFIG);
  try {
    assert.match(h.url, /^http:\/\/127\.0\.0\.1:\d+\/s\/[0-9a-f]{64}\/$/);
    const r = await request(h.url);
    assert.equal(r.status, 200);
    assert.ok(r.body.includes(JWT), 'the authorised editor still gets its config');
    assert.ok(r.body.includes('new DocsAPI.DocEditor'), 'the page contract is unchanged');
    assert.equal(r.headers['cache-control'], 'no-store');
    assert.equal(r.headers['referrer-policy'], 'no-referrer', 'api.js is loaded without a Referer carrying the token');
    assert.ok(r.body.includes('<meta name="referrer" content="no-referrer">'));
    noCors(r);
    const head = await request(h.url, { method: 'HEAD' });
    assert.deepEqual([head.status, head.body], [200, '']);
    // loopback only: the same port on a non-loopback address of this host is not listening
    const ext = Object.values(require('node:os').networkInterfaces()).flat().find(i => i && i.family === 'IPv4' && !i.internal);
    if (ext) {
      const u = new URL(h.url);
      const r2 = await request(`http://${ext.address}:${u.port}${u.pathname}`);
      assert.ok(r2.error, `not reachable on ${ext.address}: ${JSON.stringify(r2).slice(0, 80)}`);
    }
  } finally { await h.close(); }
});

test('NO_TOKEN / WRONG_TOKEN / other paths: 404, never the config (CONFIG_WITHOUT_AUTH and JWT_WITHOUT_AUTH = INACCESSIBLE)', async () => {
  const h = await startConfigHost(CONFIG);
  const t = tokenOf(h.url);
  const wrong = (t[0] === 'a' ? 'b' : 'a') + t.slice(1);
  try {
    for (const p of ['/', '/s/', '/s', `/s/${wrong}/`, `/s/${t}`, `/s/${t}/x`, `/s/${t}/?q=1`, `/S/${t}/`, `/s/${t.slice(0, 32)}/`, '/favicon.ico']) {
      const r = await request(at(h.url, p));
      assert.equal(r.status, 404, p);
      noConfig(r);
      noCors(r);
    }
  } finally { await h.close(); }
});

test('PREVIOUS_SESSION_TOKEN: a token of one session opens nothing on another session host', async () => {
  const a = await startConfigHost(CONFIG);
  const b = await startConfigHost(CONFIG);
  try {
    const r = await request(at(b.url, new URL(a.url).pathname));
    assert.equal(r.status, 404);
    noConfig(r);
  } finally { await a.close(); await b.close(); }
});

test('EXPIRED_TOKEN: after close() the session URL is dead (no listener) and its token never answers again', async () => {
  const h = await startConfigHost(CONFIG);
  assert.equal((await request(h.url)).status, 200, 'positive control while the session lives');
  await h.close();
  const r = await request(h.url);
  assert.ok(r.error || r.status === 404, JSON.stringify(r));
  noConfig(r);
});

test('Host / Origin / method: a foreign Host (DNS rebinding) or Origin, and any write method are refused -- no CORS opening', async () => {
  const h = await startConfigHost(CONFIG);
  const port = new URL(h.url).port;
  try {
    for (const host of ['localhost', `localhost:${port}`, 'attacker.example', `127.0.0.1:${Number(port) + 1}`]) {
      const r = await request(h.url, { headers: { Host: host } });
      assert.equal(r.status, 403, host);
      noConfig(r);
    }
    for (const o of ['http://attacker.example', 'null', `http://localhost:${port}`]) {
      const r = await request(h.url, { headers: { Origin: o } });
      assert.equal(r.status, 403, o);
      noConfig(r);
      noCors(r);
    }
    const same = await request(h.url, { headers: { Origin: origin(h.url) } });
    assert.equal(same.status, 200, 'the page\'s own origin is not refused');
    for (const m of ['POST', 'PUT', 'DELETE', 'OPTIONS', 'PATCH']) {
      const r = await request(h.url, { method: m, headers: { Origin: 'http://attacker.example', 'Access-Control-Request-Method': 'GET' } });
      assert.equal(r.status, 405, m);
      noConfig(r);
      noCors(r);
    }
  } finally { await h.close(); }
});

test('TOKEN_ENTROPY: 256 bits from crypto.randomBytes, every session distinct', async () => {
  assert.equal(TOKEN_BYTES, 32);
  const seen = new Set();
  for (let i = 0; i < 100; i += 1) {
    const h = await startConfigHost(CONFIG);
    const t = tokenOf(h.url);
    assert.equal(t.length, 64);
    seen.add(t);
    await h.close();
  }
  assert.equal(seen.size, 100);
});

test('redact(): the token is stripped from any text that carries the URL', async () => {
  const h = await startConfigHost(CONFIG);
  try {
    const t = tokenOf(h.url);
    const msg = `page.goto: net::ERR_ABORTED at ${h.url}\n  navigating to "${h.url}"`;
    assert.equal(h.redact(msg).includes(t), false);
    assert.ok(h.redact(msg).includes('/s/[redacted]/'));
  } finally { await h.close(); }
});

// ---- through both real drivers ----------------------------------------------------------------------------------------

/** A fake Playwright whose page.goto performs a REAL request to the host URL the driver navigates to. */
function browserFake({ gotoThrows = false } = {}) {
  const seen = { url: null, page: null, closedProbe: null };
  const page = {
    on() {}, waitForTimeout: async () => {}, frames: () => [], evaluate: async () => ({ ok: true }),
    goto: async url => {
      seen.url = url;
      seen.page = await request(url);
      if (gotoThrows) throw new Error(`page.goto: net::ERR_CONNECTION_RESET at ${url}`);
    },
    waitForFunction: async () => { throw new Error('editor never ready (fixture)'); },
  };
  const browser = { newPage: async () => page, close: async () => {} };
  return { seen, loadPlaywright: () => ({ ok: true, pw: { chromium: { launch: async () => browser } } }) };
}
const configSource = { kind: 'BROKER', getEditorConfig: async () => JSON.parse(JSON.stringify(CONFIG)) };

for (const [name, mod, ctor] of [['DOCX', docx, 'createEuroOfficeDocxDriver'], ['XLSX', xlsx, 'createEuroOfficeXlsxDriver']]) {
  test(`${name} driver: navigates to the token URL and gets its config; every error detail is token-free; the host is dead after the session`, async () => {
    for (const gotoThrows of [false, true]) {
      const f = browserFake({ gotoThrows });
      const d = mod[ctor]({ configSource, loadPlaywright: f.loadPlaywright, timeoutMs: 50, pollMs: 5, closeAckTimeoutMs: 50 });
      const o = await d.open({ fileId: '1302148', operation: 'READ' });
      assert.equal(o.ok, false, 'the fixture editor never becomes ready');
      assert.match(f.seen.url, /^http:\/\/127\.0\.0\.1:\d+\/s\/[0-9a-f]{64}\/$/, 'the driver navigates to the token URL');
      assert.equal(f.seen.page.status, 200);
      assert.ok(f.seen.page.body.includes(JWT), 'VALID_EDITOR_CONFIG: the session browser gets the config');
      assert.equal(JSON.stringify(o).includes(tokenOf(f.seen.url)), false, `TOKEN_IN_ERRORS=NO (${o.detail})`);
      if (gotoThrows) assert.match(String(o.detail), /\[redacted\]/, 'the navigation error is kept, redacted');
      const after = await request(f.seen.url);
      assert.ok(after.error || after.status === 404, 'SESSION_CLEANUP: the host is gone with the session');
      noConfig(after);
    }
  });
}

test('the host module logs nothing and the drivers build no other loopback server', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const root = path.join(__dirname, '..', '..', 'src', 'drivers', 'eurooffice');
  assert.equal(/console\.|process\.std(out|err)/.test(fs.readFileSync(path.join(root, 'config-host.cjs'), 'utf8')), false);
  for (const f of ['docx-driver.cjs', 'xlsx-driver.cjs']) {
    const src = fs.readFileSync(path.join(root, f), 'utf8');
    assert.equal(/createServer|\.listen\(/.test(src), false, `${f}: no own HTTP server`);
    assert.match(src, /const startHost = startConfigHost;/);
  }
});
