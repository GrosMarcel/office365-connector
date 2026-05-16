'use strict';
/**
 * Centralized security helpers for the office365-connector skill.
 *
 * Goals:
 *   - Prevent path traversal via untrusted account names / Graph IDs.
 *   - Prevent URL/host injection in OAuth + Graph requests.
 *   - Atomic, mode-0600 writes for any file containing secrets or tokens.
 *   - Scrub OAuth/Graph error bodies before they end up in logs.
 *   - Strip terminal escape sequences from anything we print.
 *   - Hard timeouts and response-size caps on every outbound HTTPS request.
 */

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const https = require('https');

const UUID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
const ACCOUNT_NAME_RE = /^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$/;
const GRAPH_ID_RE = /^[A-Za-z0-9_=+\-/]{1,512}$/;

const ALLOWED_HOSTS = new Set([
  'graph.microsoft.com',
  'login.microsoftonline.com'
]);

const MAX_RESPONSE_BYTES = 8 * 1024 * 1024;
const DEFAULT_TIMEOUT_MS = 30000;

function assertAccountName(name) {
  if (typeof name !== 'string' || !ACCOUNT_NAME_RE.test(name)) {
    throw new Error('Invalid account name. Allowed: letters, digits, "-", "_", max 64 chars, must start with a letter or digit.');
  }
  return name;
}

function assertUuid(value, label) {
  if (typeof value !== 'string' || !UUID_RE.test(value)) {
    throw new Error(`Invalid ${label}: expected a UUID.`);
  }
  return value;
}

function assertGraphId(value, label) {
  if (typeof value !== 'string' || !GRAPH_ID_RE.test(value)) {
    throw new Error(`Invalid ${label}: contains forbidden characters.`);
  }
  return value;
}

function assertEmail(value, label) {
  if (typeof value !== 'string' || value.length > 320 || !/^[^\s<>"'\\/]+@[^\s<>"'\\/]+\.[^\s<>"'\\/]+$/.test(value)) {
    throw new Error(`Invalid ${label}: not a valid email address.`);
  }
  return value;
}

function assertSafeChildPath(baseDir, untrustedName) {
  assertAccountName(untrustedName);
  const resolvedBase = path.resolve(baseDir);
  const resolved = path.resolve(resolvedBase, `${untrustedName}.json`);
  if (resolved !== path.join(resolvedBase, `${untrustedName}.json`)) {
    throw new Error('Path traversal detected in account name.');
  }
  return resolved;
}

/**
 * Atomically write `data` to `filePath` with mode 0600.
 * - Writes to a sibling temp file then renames (atomic on POSIX).
 * - chmods explicitly even when the file already exists.
 * - Verifies the parent directory is mode 0700.
 */
function secureWriteFileSync(filePath, data) {
  const dir = path.dirname(filePath);
  fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  try { fs.chmodSync(dir, 0o700); } catch (_) { /* best effort */ }

  const tmp = `${filePath}.${process.pid}.${crypto.randomBytes(6).toString('hex')}.tmp`;
  const fd = fs.openSync(tmp, 'wx', 0o600);
  try {
    fs.writeSync(fd, data);
    try { fs.fsyncSync(fd); } catch (_) { /* not all FS support fsync */ }
  } finally {
    fs.closeSync(fd);
  }
  try {
    fs.renameSync(tmp, filePath);
    fs.chmodSync(filePath, 0o600);
  } catch (e) {
    try { fs.unlinkSync(tmp); } catch (_) { /* ignore */ }
    throw e;
  }
}

function secureReadJsonSync(filePath) {
  if (!fs.existsSync(filePath)) return null;
  const stat = fs.statSync(filePath);
  if ((stat.mode & 0o077) !== 0) {
    try { fs.chmodSync(filePath, 0o600); } catch (_) { /* best effort */ }
  }
  const raw = fs.readFileSync(filePath, 'utf8');
  try {
    return JSON.parse(raw);
  } catch (_) {
    throw new Error(`Failed to parse JSON file: ${filePath}`);
  }
}

const ANSI_ESC_RE = /\[[0-9;?]*[ -\/]*[@-~]/g;
const C0_C1_RE = /[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]/g;

function sanitizeForTerminal(value) {
  if (value == null) return '';
  const s = String(value);
  return s.replace(ANSI_ESC_RE, '').replace(C0_C1_RE, '');
}

const TOKEN_FIELD_RE = /("?(?:access_token|refresh_token|id_token|device_code|client_secret|user_code)"?\s*[:=]\s*"?)[^"\s,}]+/gi;

function scrubSecrets(value) {
  if (value == null) return '';
  return String(value).replace(TOKEN_FIELD_RE, '$1[REDACTED]');
}

function safeErrorMessage(prefix, status, parsedOrRaw) {
  let msg;
  if (parsedOrRaw && typeof parsedOrRaw === 'object') {
    msg = parsedOrRaw.error_description
      || (parsedOrRaw.error && (parsedOrRaw.error.message || parsedOrRaw.error))
      || `HTTP ${status}`;
  } else {
    msg = `HTTP ${status}`;
  }
  return `${prefix} ${scrubSecrets(msg)}`;
}

function assertAllowedHost(hostname) {
  if (!ALLOWED_HOSTS.has(hostname)) {
    throw new Error(`Refusing to contact unexpected host: ${hostname}`);
  }
}

/**
 * Make a JSON HTTPS request with:
 *   - host allow-listing,
 *   - hard timeout,
 *   - response size cap,
 *   - scrubbed error messages.
 *
 * Returns the parsed JSON body on 2xx. Rejects with a safe Error otherwise.
 */
function httpsJsonRequest(targetUrl, options = {}) {
  return new Promise((resolve, reject) => {
    let urlObj;
    try {
      urlObj = new URL(targetUrl);
    } catch (_) {
      return reject(new Error('Invalid request URL'));
    }
    if (urlObj.protocol !== 'https:') {
      return reject(new Error('Refusing non-HTTPS request'));
    }
    try {
      assertAllowedHost(urlObj.hostname);
    } catch (e) {
      return reject(e);
    }

    const reqOptions = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: options.method || 'GET',
      headers: Object.assign({ 'Accept': 'application/json' }, options.headers || {}),
      timeout: options.timeoutMs || DEFAULT_TIMEOUT_MS
    };

    const req = https.request(reqOptions, (res) => {
      let received = 0;
      const chunks = [];
      res.on('data', (chunk) => {
        received += chunk.length;
        if (received > MAX_RESPONSE_BYTES) {
          res.destroy(new Error('Response too large'));
          return;
        }
        chunks.push(chunk);
      });
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        let parsed = null;
        if (raw.length > 0) {
          try { parsed = JSON.parse(raw); } catch (_) { parsed = null; }
        }
        if (res.statusCode >= 200 && res.statusCode < 300) {
          if (options.acceptEmpty && parsed === null) {
            return resolve({ success: true, statusCode: res.statusCode });
          }
          return resolve(parsed != null ? parsed : { success: true, statusCode: res.statusCode });
        }
        reject(new Error(safeErrorMessage(`HTTP ${res.statusCode}:`, res.statusCode, parsed)));
      });
      res.on('error', (err) => reject(new Error(scrubSecrets(err.message))));
    });

    req.on('timeout', () => {
      req.destroy(new Error('Request timed out'));
    });
    req.on('error', (err) => reject(new Error(scrubSecrets(err.message))));

    if (options.body != null) {
      req.write(options.body);
    }
    req.end();
  });
}

/**
 * Parse `--account=NAME` from argv with validation.
 * Returns null when absent.
 */
function parseAccountFlag(argv) {
  const arg = argv.find((a) => typeof a === 'string' && a.startsWith('--account='));
  if (!arg) return null;
  const value = arg.slice('--account='.length);
  if (!value) throw new Error('--account requires a value');
  return assertAccountName(value);
}

/**
 * Read a single line from stdin without echo, for prompting secrets.
 * Falls back to plain readline if TTY raw mode is unavailable.
 */
function readSecretFromStdin(prompt) {
  return new Promise((resolve, reject) => {
    process.stderr.write(prompt);
    const stdin = process.stdin;
    let value = '';
    const isTTY = stdin.isTTY === true;
    if (isTTY) {
      try { stdin.setRawMode(true); } catch (_) { /* not always possible */ }
    }
    stdin.resume();
    stdin.setEncoding('utf8');

    function cleanup() {
      stdin.removeListener('data', onData);
      if (isTTY) {
        try { stdin.setRawMode(false); } catch (_) { /* ignore */ }
      }
      stdin.pause();
      process.stderr.write('\n');
    }

    function onData(chunk) {
      for (let i = 0; i < chunk.length; i++) {
        const ch = chunk[i];
        if (ch === '\n' || ch === '\r' || ch === '') {
          cleanup();
          return resolve(value);
        }
        if (ch === '') {
          cleanup();
          return reject(new Error('Aborted'));
        }
        if (ch === '' || ch === '\b') {
          value = value.slice(0, -1);
          continue;
        }
        value += ch;
      }
    }

    stdin.on('data', onData);
  });
}

module.exports = {
  assertAccountName,
  assertUuid,
  assertGraphId,
  assertEmail,
  assertSafeChildPath,
  assertAllowedHost,
  secureWriteFileSync,
  secureReadJsonSync,
  sanitizeForTerminal,
  scrubSecrets,
  safeErrorMessage,
  httpsJsonRequest,
  parseAccountFlag,
  readSecretFromStdin,
  MAX_RESPONSE_BYTES,
  DEFAULT_TIMEOUT_MS
};
