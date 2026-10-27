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

/**
 * The editor page (unchanged contract: window.__ready / __error / __config / __docEditor).
 *
 * GitHub #77: before this, only onDocumentReady was wired -- the open() wait (`window.__ready === true ||
 * window.__error`) was blind to onError/onWarning/onInfo for 90s if the constructor itself never threw. Now:
 *   onError   -> window.__error = { kind:'onError', code, description } (the constructor's own catch still sets
 *               __error to a plain string on a THROWN error; the wait predicate does not care which shape)
 *   onWarning -> pushed to window.__warnings[] -- does NOT set __ready/__error, the wait is unaffected on purpose
 *               (a warning is not a terminal state; config-host.cjs's own driver-side reader is what surfaces it)
 *   onInfo    -> window.__info = { mode } -- same: does not affect the wait
 */
function editorPage(config) {
  const api = `${config.documentServerUrl}web-apps/apps/api/documents/api.js?shardKey=${encodeURIComponent(config.document.key)}`;
  const safe = JSON.stringify(config).replace(/</g, '\\u003c');
  const events = "{onDocumentReady:function(){window.__ready=true},"
    + "onError:function(e){window.__error={kind:'onError',code:e&&e.data&&e.data.errorCode,description:e&&e.data&&e.data.errorDescription}},"
    + "onWarning:function(e){window.__warnings.push({code:e&&e.data&&e.data.errorCode,description:e&&e.data&&e.data.errorDescription})},"
    + "onInfo:function(e){window.__info={mode:e&&e.data&&e.data.mode}}}";
  return `<!doctype html><html><head><meta charset="utf-8"><meta name="referrer" content="no-referrer"><script src="${api}"></script></head><body><div id="editor" style="width:1200px;height:800px"></div><script>window.__ready=false;window.__error=null;window.__warnings=[];window.__info=null;window.__config=${safe};window.__config.events=${events};try{window.__docEditor=new DocsAPI.DocEditor('editor',window.__config)}catch(e){window.__error=String(e&&e.stack||e)}</script></body></html>`;
}

/**
 * GitHub #77: the page-side read of what DS reported, for a driver's open() to call back after its ready/error
 * wait times out (one evaluate() call; null if the page cannot be read at all, e.g. it already closed/crashed --
 * kept DISTINCT from "nothing happened", which is a different fact).
 */
async function readDsOpenState(page) {
  return page.evaluate(() => ({ error: window.__error || null, warnings: window.__warnings || [], info: window.__info || null })).catch(() => null);
}

/**
 * GitHub #77: the OPEN_FAILED detail for a timed-out onReady/onError wait -- names the DS's own error code /
 * description when onError actually fired, what onWarning/onInfo recorded even though neither stops the wait, or
 * that literally no DS event arrived in the whole window ('no-ds-event', distinct from 'ds-state-unreadable').
 */
function openWaitFailureDetail(dsState) {
  if (dsState === null) return 'ds-state-unreadable: window.__error/__warnings/__info could not be read after the timeout';
  const { error, warnings, info } = dsState;
  if (error) {
    if (error && typeof error === 'object' && error.kind === 'onError') return `DS onError: code ${error.code}, ${error.description}`;
    return `the editor constructor threw: ${error}`;
  }
  if ((warnings && warnings.length) || info) return `no-ds-ready-or-error: DS reported onWarning/onInfo only: ${JSON.stringify({ warnings, info })}`;
  return 'no-ds-event: no onDocumentReady/onError/onWarning/onInfo observed in the wait window';
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

module.exports = { startConfigHost, editorPage, TOKEN_BYTES, readDsOpenState, openWaitFailureDetail };
