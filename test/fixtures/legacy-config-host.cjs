'use strict';

/**
 * EVIDENCE ONLY (M11.3b.1): the loopback config host as it was in BOTH drivers before the hardening -- copied verbatim
 * from main 10d3312b (src/drivers/eurooffice/{docx,xlsx}-driver.cjs startHost; the two were identical). It answers every
 * request on any path with the full editor config (incl. the Document Server JWT). Never imported by src/; kept so the
 * vulnerable behaviour stays reproducible next to the fixed one (test/acceptance/m11-3b1-loopback-hardening.test.cjs).
 */

const http = require('node:http');

function startHost(config) {
  const api = `${config.documentServerUrl}web-apps/apps/api/documents/api.js?shardKey=${encodeURIComponent(config.document.key)}`;
  const safe = JSON.stringify(config).replace(/</g, '\\u003c');
  const html = `<!doctype html><html><head><meta charset="utf-8"><script src="${api}"></script></head><body><div id="editor" style="width:1200px;height:800px"></div><script>window.__ready=false;window.__error=null;window.__config=${safe};window.__config.events={onDocumentReady:function(){window.__ready=true}};try{window.__docEditor=new DocsAPI.DocEditor('editor',window.__config)}catch(e){window.__error=String(e&&e.stack||e)}</script></body></html>`;
  return new Promise(resolve => {
    const server = http.createServer((_req, res) => { res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' }); res.end(html); });
    server.listen(0, '127.0.0.1', () => resolve({ server, url: `http://127.0.0.1:${server.address().port}/` }));
  });
}

module.exports = { startHost };
