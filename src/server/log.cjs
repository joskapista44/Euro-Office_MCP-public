'use strict';

/**
 * Product MCP logging (owner decision, owner channel 2026-10-08, point 14).
 *
 * stdout is the MCP protocol channel and NOTHING else writes to it; logs go to stderr and to a per-UID file under
 * $XDG_STATE_HOME/euro-office-mcp/ (default ~/.local/state/euro-office-mcp/), directory 0700, file 0600.
 * Never logged: tokens, passwords, raw secret-bearing requests, document content -- secret-shaped keys and
 * credential-shaped strings are redacted, every string is truncated.
 */

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const LEVELS = { error: 0, warn: 1, info: 2, debug: 3 };
const SECRET_KEY = /pass(word)?|secret|token|apikey|api_key|authorization|cookie|private_?key/i;
const SECRET_VALUE = /\b(Bearer|Basic)\s+[A-Za-z0-9+/=._~-]{8,}|mvdk_[A-Za-z0-9_-]{20,}/g;
const MAX_STRING = 300;

function stateDir(env = process.env) {
  const base = env.XDG_STATE_HOME && path.isAbsolute(env.XDG_STATE_HOME) ? env.XDG_STATE_HOME : path.join(env.HOME || os.homedir(), '.local', 'state');
  return path.join(base, 'euro-office-mcp');
}

function redact(value, depth = 0) {
  if (depth > 6) return '[DEPTH]';
  if (typeof value === 'string') {
    const s = value.replace(SECRET_VALUE, '[REDACTED]');
    return s.length > MAX_STRING ? `${s.slice(0, MAX_STRING)}...[${s.length} chars]` : s;
  }
  if (Array.isArray(value)) return value.slice(0, 20).map(v => redact(v, depth + 1));
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, SECRET_KEY.test(k) ? '[REDACTED]' : redact(v, depth + 1)]));
  }
  return value;
}

/** createLogger({ level, env, stderr }) -> { error, warn, info, debug, file } ; file logging is best-effort. */
function createLogger({ level = 'info', env = process.env, stderr = process.stderr } = {}) {
  const max = LEVELS[level] ?? LEVELS.info;
  let file = null;
  try {
    const dir = stateDir(env);
    fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
    file = path.join(dir, `${new Date().toISOString().slice(0, 10)}.log`);
    fs.closeSync(fs.openSync(file, 'a', 0o600));
  } catch { file = null; }
  const write = (lvl, msg, fields) => {
    if (LEVELS[lvl] > max) return;
    const line = JSON.stringify({ t: new Date().toISOString(), lvl, msg: redact(String(msg)), ...(fields ? { f: redact(fields) } : {}) });
    try { stderr.write(`${line}\n`); } catch { /* stderr gone: nothing else to do */ }
    if (file) try { fs.appendFileSync(file, `${line}\n`); } catch { /* best-effort */ }
  };
  return {
    file,
    error: (m, f) => write('error', m, f), warn: (m, f) => write('warn', m, f),
    info: (m, f) => write('info', m, f), debug: (m, f) => write('debug', m, f),
  };
}

module.exports = { createLogger, redact, stateDir };
