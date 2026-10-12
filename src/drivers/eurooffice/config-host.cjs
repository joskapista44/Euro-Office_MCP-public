'use strict';

/**
 * The loopback editor-config host of ONE editor session, shared by the DOCX and XLSX drivers (M11.3b.1, owner decision
 * LOOPBACK_HARDENING=REQUIRED, owner channel 2026-10-09).
 *
 * The page embeds the full editor config -- including the Document Server JWT -- so it must not be readable by anyone
 * but the session's own browser. Before M11.3b.1 the host answered EVERY request on 127.0.0.1:<port> with that page, so
 * any local process (any UID on the host) could read the config while the session lived.
 *
 * Now:
 *   - a fresh 256-bit random token per session (crypto.randomBytes); the page is served ONLY at /s/<token>/;
 *   - no token / a wrong token / a previous session's token / any other path -> 404 (one uniform answer, no oracle);
 *   - only GET/HEAD; the Host header must be exactly 127.0.0.1:<port> (no DNS-rebinding path); a foreign Origin is
 *     refused; no CORS header is ever sent;
 *   - Cache-Control: no-store; Referrer-Policy: no-referrer (+ the meta tag), so the token never reaches the Document
 *     Server through a Referer header when the page loads api.js;
 *   - close() invalidates the token at once (every later request -> 404) and then stops listening;
 *   - the token is never logged; redact() strips it from any text that could carry the URL (e.g. a navigation error).
 */

const http = require('node:http');
const crypto = require('node:crypto');

const TOKEN_BYTES = 32; // 256 bits
const DENY = Object.freeze({ 'content-type': 'text/plain; charset=utf-8', 'cache-control': 'no-store' });
const PAGE = Object.freeze({ 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store', 'referrer-policy': 'no-referrer',
  'x-content-type-options': 'nosniff', 'x-frame-options': 'DENY' });

/** The editor page (unchanged contract: window.__ready / __error / __config / __docEditor). */
function editorPage(config) {
  const api = `${config.documentServerUrl}web-apps/apps/api/documents/api.js?shardKey=${encodeURIComponent(config.document.key)}`;
  const safe = JSON.stringify(config).replace(/</g, '\\u003c');
  return `<!doctype html><html><head><meta charset="utf-8"><meta name="referrer" content="no-referrer"><script src="${api}"></script></head><body><div id="editor" style="width:1200px;height:800px"></div><script>window.__ready=false;window.__error=null;window.__config=${safe};window.__config.events={onDocumentReady:function(){window.__ready=true}};try{window.__docEditor=new DocsAPI.DocEditor('editor',window.__config)}catch(e){window.__error=String(e&&e.stack||e)}</script></body></html>`;
}

const sameText = (a, b) => {
  const x = Buffer.from(String(a));
  const y = Buffer.from(String(b));
  return x.length === y.length && crypto.timingSafeEqual(x, y);
};

/** The refusal status of a request, or 0 when it may have the page: method, then Host, then Origin, then the token path. */
function refusalOf(req, { origin, path, active }) {
  if (req.method !== 'GET' && req.method !== 'HEAD') return 405;
  if (req.headers.host !== origin.slice('http://'.length)) return 403;
  if (req.headers.origin !== undefined && req.headers.origin !== origin) return 403;
  return active && sameText(req.url, path) ? 0 : 404;
}

/** startConfigHost(config) -> Promise<{ url, close(), redact(text) }>. Listens on 127.0.0.1 only, ephemeral port. */
function startConfigHost(config) {
  const html = editorPage(config);
  const token = crypto.randomBytes(TOKEN_BYTES).toString('hex');
  const path = `/s/${token}/`;
  let active = true;
  let origin = null;
  const deny = (res, code) => { res.writeHead(code, DENY); res.end(code === 404 ? 'not found' : 'refused'); };
  const server = http.createServer((req, res) => {
    const code = refusalOf(req, { origin, path, active });
    if (code) { deny(res, code); return; }
    res.writeHead(200, PAGE);
    res.end(req.method === 'HEAD' ? undefined : html);
  });
  const redact = text => String(text).split(token).join('[redacted]');
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      origin = `http://127.0.0.1:${server.address().port}`;
      resolve({
        url: `${origin}${path}`,
        redact,
        close: () => {
          active = false; // the token is dead before the socket closes
          if (typeof server.closeAllConnections === 'function') server.closeAllConnections();
          return new Promise(r => server.close(() => r()));
        },
      });
    });
  });
}

module.exports = { startConfigHost, editorPage, TOKEN_BYTES };
