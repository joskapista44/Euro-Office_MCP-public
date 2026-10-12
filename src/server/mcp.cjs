'use strict';

/**
 * A minimal MCP server over stdio (owner decision, owner channel 2026-10-08, point 3) -- a TRANSPORT adapter only:
 * no lifecycle or executor of its own.
 *
 * Wire format: MCP stdio transport = newline-delimited JSON-RPC 2.0 messages on stdin/stdout. stdout carries ONLY
 * protocol messages (one `write` per message); everything else goes to the logger (stderr / per-UID file).
 * Methods: initialize, notifications/initialized, ping, tools/list, tools/call. Unknown method -> -32601,
 * unparseable line -> -32700. stdin EOF or SIGTERM/SIGINT -> stop reading, finish in-flight calls, exit 0.
 *
 * No third-party MCP SDK: the product keeps zero runtime dependencies (NEW_INTERNAL_DEPENDENCY=NONE).
 */

const readline = require('node:readline');

const SUPPORTED_PROTOCOL_VERSIONS = Object.freeze(['2025-06-18', '2025-03-26', '2024-11-05']);
const SERVER_NAME = 'euro-office-product';

/**
 * createMcpServer({ tools, info, log, input, output }) -> { start(): Promise<exitReason>, stop(reason) }
 * `tools`: { [name]: { description, inputSchema, handler(args) -> Promise<object> } }
 */
function createMcpServer({ tools, info, log, input = process.stdin, output = process.stdout }) {
  let initialized = false;
  let stopping = false;
  const inFlight = new Set();
  let resolveDone;
  let rl = null;
  const done = new Promise(res => { resolveDone = res; });

  const send = msg => { if (!output.destroyed) output.write(`${JSON.stringify(msg)}\n`); };
  const reply = (id, result) => send({ jsonrpc: '2.0', id, result });
  const fail = (id, code, message) => send({ jsonrpc: '2.0', id: id === undefined ? null : id, error: { code, message } });

  async function handle(msg) {
    if (!msg || msg.jsonrpc !== '2.0' || typeof msg.method !== 'string') { fail(msg && msg.id, -32600, 'Invalid Request'); return; }
    const { id, method, params } = msg;
    const isNotification = id === undefined;
    switch (method) {
      case 'initialize': {
        const asked = params && params.protocolVersion;
        const protocolVersion = SUPPORTED_PROTOCOL_VERSIONS.includes(asked) ? asked : SUPPORTED_PROTOCOL_VERSIONS[0];
        reply(id, { protocolVersion, capabilities: { tools: { listChanged: false } }, serverInfo: { name: SERVER_NAME, version: info.version } });
        return;
      }
      case 'notifications/initialized': initialized = true; return;
      case 'ping': if (!isNotification) reply(id, {}); return;
      case 'tools/list':
        reply(id, { tools: Object.entries(tools).map(([name, t]) => ({ name, description: t.description, inputSchema: t.inputSchema })) });
        return;
      case 'tools/call': {
        const name = params && params.name;
        const tool = tools[name];
        if (!tool) { fail(id, -32602, `Unknown tool: ${name}`); return; }
        try {
          const out = await tool.handler((params && params.arguments) || {});
          reply(id, { content: [{ type: 'text', text: JSON.stringify(out, null, 2) }], structuredContent: out, isError: false });
        } catch (e) {
          log.error('tool failed', { tool: name, error: String(e && e.message || e) });
          reply(id, { content: [{ type: 'text', text: `tool ${name} failed` }], isError: true });
        }
        return;
      }
      default:
        if (!isNotification) fail(id, -32601, `Method not found: ${method}`);
    }
  }

  function stop(reason) {
    if (stopping) return;
    stopping = true;
    log.info('stopping', { reason, inFlight: inFlight.size, initialized });
    // stop reading (a signal leaves stdin open, which would keep the process alive); in-flight calls still answer
    try { if (rl) rl.close(); input.pause(); if (typeof input.destroy === 'function') input.destroy(); } catch { /* already closed */ }
    Promise.allSettled([...inFlight]).then(() => resolveDone(reason));
  }

  function start() {
    rl = readline.createInterface({ input, crlfDelay: Infinity });
    rl.on('line', line => {
      if (stopping || !line.trim()) return;
      let msg;
      try { msg = JSON.parse(line); } catch { fail(null, -32700, 'Parse error'); return; }
      const p = handle(msg).catch(e => log.error('handler error', { error: String(e && e.message || e) }));
      inFlight.add(p);
      p.finally(() => inFlight.delete(p));
    });
    rl.on('close', () => stop('stdin-eof'));
    return done;
  }

  return { start, stop };
}

module.exports = { createMcpServer, SUPPORTED_PROTOCOL_VERSIONS, SERVER_NAME };
