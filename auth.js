#!/usr/bin/env node
'use strict';
/**
 * Microsoft Graph OAuth 2.0 Device Code Flow Authentication
 * Handles token acquisition, refresh, and storage.
 *
 * Security:
 *   - Host-pinned, timeout-bounded HTTPS requests (no SSRF, no hangs).
 *   - tenantId / clientId validated as UUIDs before URL composition.
 *   - Tokens written atomically with mode 0600, parent dir 0700.
 *   - Error messages strip access_token / refresh_token / device_code / client_secret
 *     before they ever reach stderr or callers.
 *   - `token` CLI subcommand requires an explicit --confirm flag to avoid accidental
 *     terminal exposure of bearer tokens.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { getAccount } = require('./accounts.js');
const {
  assertUuid,
  secureWriteFileSync,
  secureReadJsonSync,
  httpsJsonRequest,
  parseAccountFlag,
  scrubSecrets
} = require('./security.js');

const SCOPES = [
  'User.Read',
  'Mail.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.Read',
  'Calendars.ReadWrite',
  'Contacts.Read',
  'Contacts.ReadWrite',
  'offline_access'
].join(' ');

const PENDING_PREFIX = '.pending-';

function getAccountConfig(accountName) {
  try {
    return getAccount(accountName);
  } catch (error) {
    if (accountName) throw error;
    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
      throw new Error('No account configured and no credentials in environment. Run: node accounts.js add <name> ...');
    }
    assertUuid(tenantId, 'AZURE_TENANT_ID');
    assertUuid(clientId, 'AZURE_CLIENT_ID');

    const home = os.homedir();
    if (!home) {
      throw new Error('Unable to resolve home directory; refusing to write tokens.');
    }
    return {
      name: 'legacy',
      tenantId,
      clientId,
      clientSecret,
      tokenPath: path.join(home, '.hermes', 'auth', 'microsoft-graph.json')
    };
  }
}

function authorityFor(accountConfig) {
  assertUuid(accountConfig.tenantId, 'tenantId');
  return `https://login.microsoftonline.com/${accountConfig.tenantId}`;
}

async function requestDeviceCode(accountConfig) {
  assertUuid(accountConfig.clientId, 'clientId');
  const url = `${authorityFor(accountConfig)}/oauth2/v2.0/devicecode`;
  const body = new URLSearchParams({
    client_id: accountConfig.clientId,
    scope: SCOPES
  }).toString();

  return httpsJsonRequest(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(body)
    },
    body
  });
}

async function pollForToken(deviceCode, accountConfig) {
  assertUuid(accountConfig.clientId, 'clientId');
  const url = `${authorityFor(accountConfig)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
    client_id: accountConfig.clientId,
    device_code: deviceCode
  }).toString();

  return httpsJsonRequest(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(body)
    },
    body
  });
}

async function refreshAccessToken(refreshToken, accountConfig) {
  assertUuid(accountConfig.clientId, 'clientId');
  const url = `${authorityFor(accountConfig)}/oauth2/v2.0/token`;

  // Device-code flow is a *public* client; never send client_secret on refresh.
  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: accountConfig.clientId,
    refresh_token: refreshToken,
    scope: SCOPES
  }).toString();

  return httpsJsonRequest(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(body)
    },
    body
  });
}

function saveTokens(tokens, accountConfig) {
  if (!tokens || typeof tokens.access_token !== 'string' || typeof tokens.expires_in !== 'number') {
    throw new Error('Invalid token response from authority');
  }
  const data = {
    access_token: tokens.access_token,
    refresh_token: tokens.refresh_token,
    expires_at: Date.now() + (tokens.expires_in * 1000),
    scope: tokens.scope
  };

  secureWriteFileSync(accountConfig.tokenPath, JSON.stringify(data, null, 2));
  return data;
}

function loadTokens(accountConfig) {
  try {
    return secureReadJsonSync(accountConfig.tokenPath);
  } catch (e) {
    // Never log the raw file contents — could include the token itself.
    console.error('Failed to load tokens (file unreadable or corrupted).');
    return null;
  }
}

async function getAccessToken(accountName = null) {
  const accountConfig = getAccountConfig(accountName);
  const tokens = loadTokens(accountConfig);

  if (!tokens) {
    throw new Error(`Not authenticated for account "${accountConfig.name}". Run authentication first.`);
  }

  if (tokens.expires_at < Date.now() + (5 * 60 * 1000)) {
    console.error('Token expired, refreshing...');
    if (!tokens.refresh_token) {
      throw new Error('No refresh token available; please re-authenticate.');
    }
    const refreshed = await refreshAccessToken(tokens.refresh_token, accountConfig);
    return saveTokens(refreshed, accountConfig).access_token;
  }

  return tokens.access_token;
}

/**
 * Per-account pending-flow state file. The stored device_code is itself a
 * credential — file mode 0600 (enforced by secureWriteFileSync) prevents
 * other users on the same host from racing the polling endpoint.
 */
function getPendingPath(accountConfig) {
  const dir = path.dirname(accountConfig.tokenPath);
  return path.join(dir, `${PENDING_PREFIX}${accountConfig.name}.json`);
}

/**
 * Step 1 of a non-blocking device-code flow: request a device code, persist
 * the pending state, and return the prompt info IMMEDIATELY. The caller is
 * expected to surface verification_uri + user_code to the human, then loop
 * on pollAuthenticate() until the flow resolves.
 */
async function beginAuthenticate(accountConfig) {
  const existing = loadTokens(accountConfig);
  if (existing && existing.expires_at > Date.now()) {
    return { status: 'already_authenticated', account: accountConfig.name, expires_at_ms: existing.expires_at };
  }

  const dc = await requestDeviceCode(accountConfig);
  if (!dc || !dc.device_code) {
    throw new Error('Authority did not return a device code');
  }

  const interval = Math.max(1, parseInt(dc.interval, 10) || 5);
  const expiresIn = parseInt(dc.expires_in, 10) || 600;
  const pending = {
    device_code: dc.device_code,
    interval,
    expires_at: Math.floor(Date.now() / 1000) + expiresIn,
    account: accountConfig.name
  };
  secureWriteFileSync(getPendingPath(accountConfig), JSON.stringify(pending));

  return {
    status: 'pending_started',
    account: accountConfig.name,
    verification_uri: dc.verification_uri,
    user_code: dc.user_code,
    expires_in: expiresIn,
    poll_interval_s: interval
  };
}

/**
 * Step 2 of a non-blocking device-code flow: read the pending state and do
 * EXACTLY ONE poll of the token endpoint. Returns a single status snapshot.
 * Caller is expected to re-run periodically until status leaves 'pending'.
 */
async function pollAuthenticate(accountConfig) {
  const pendingPath = getPendingPath(accountConfig);
  const pending = secureReadJsonSync(pendingPath);

  if (!pending) {
    const existing = loadTokens(accountConfig);
    if (existing && existing.expires_at > Date.now()) {
      return { status: 'already_authenticated', account: accountConfig.name, expires_at_ms: existing.expires_at };
    }
    return { status: 'no_pending_flow', account: accountConfig.name, message: 'Run login first.' };
  }

  if (Math.floor(Date.now() / 1000) >= pending.expires_at) {
    try { fs.unlinkSync(pendingPath); } catch (_) { /* ignore */ }
    return { status: 'expired', account: accountConfig.name };
  }

  try {
    const tokenResponse = await pollForToken(pending.device_code, accountConfig);
    const saved = saveTokens(tokenResponse, accountConfig);
    try { fs.unlinkSync(pendingPath); } catch (_) { /* ignore */ }
    return { status: 'authenticated', account: accountConfig.name, expires_at_ms: saved.expires_at };
  } catch (error) {
    const msg = scrubSecrets(error.message || '');
    if (msg.includes('authorization_pending') || msg.includes('AADSTS70016')) {
      return { status: 'pending', account: accountConfig.name, next_poll_in_s: pending.interval };
    }
    if (msg.includes('authorization_declined')) {
      try { fs.unlinkSync(pendingPath); } catch (_) { /* ignore */ }
      return { status: 'declined', account: accountConfig.name };
    }
    if (msg.includes('expired_token') || msg.includes('AADSTS70019')) {
      try { fs.unlinkSync(pendingPath); } catch (_) { /* ignore */ }
      return { status: 'expired', account: accountConfig.name };
    }
    if (msg.includes('slow_down')) {
      return { status: 'pending', account: accountConfig.name, next_poll_in_s: pending.interval * 2 };
    }
    throw error;
  }
}

/**
 * DNS + TLS reachability check for the two hosts the plugin contacts. Helps
 * tell apart "network is broken inside this container" from "flow logic is
 * broken" when login appears to hang.
 */
async function diagnose() {
  const dns = require('dns').promises;
  const tls = require('tls');
  const results = {};
  for (const host of ['login.microsoftonline.com', 'graph.microsoft.com']) {
    const entry = {};
    try {
      const { address } = await dns.lookup(host);
      entry.ip = address;
      entry.dns = 'ok';
    } catch (e) {
      entry.dns = `error: ${e.message}`;
      results[host] = entry;
      continue;
    }
    try {
      await new Promise((resolve, reject) => {
        const socket = tls.connect(
          { host, port: 443, servername: host },
          () => { socket.end(); resolve(); }
        );
        socket.setTimeout(5000, () => { socket.destroy(); reject(new Error('TLS connect timeout')); });
        socket.on('error', reject);
      });
      entry.tls = 'ok';
    } catch (e) {
      entry.tls = `error: ${e.message}`;
    }
    results[host] = entry;
  }
  return { hosts: results };
}

async function authenticate(accountName = null) {
  const accountConfig = getAccountConfig(accountName);

  console.error(`Starting Microsoft Graph authentication for account "${accountConfig.name}"...\n`);

  const existingTokens = loadTokens(accountConfig);
  if (existingTokens && existingTokens.expires_at > Date.now()) {
    console.error('✅ Already authenticated! Token is valid.');
    return existingTokens;
  }

  console.error('Requesting device code...');
  const deviceCodeResponse = await requestDeviceCode(accountConfig);

  if (!deviceCodeResponse || !deviceCodeResponse.device_code) {
    throw new Error('Authority did not return a device code');
  }

  console.error('\n' + '='.repeat(60));
  console.error('📱 AUTHENTICATION REQUIRED');
  console.error('='.repeat(60));
  console.error(`\n1. Open this URL: ${deviceCodeResponse.verification_uri}`);
  console.error(`\n2. Enter this code: ${deviceCodeResponse.user_code}`);
  console.error(`\n3. Sign in with your Microsoft account`);
  console.error(`\n4. Approve the requested permissions`);
  console.error('\n' + '='.repeat(60) + '\n');
  console.error('Waiting for authentication...');

  const interval = Math.max(1, deviceCodeResponse.interval || 5) * 1000;
  const expiresAt = Date.now() + (deviceCodeResponse.expires_in * 1000);

  while (Date.now() < expiresAt) {
    await new Promise((resolve) => setTimeout(resolve, interval));

    try {
      const tokenResponse = await pollForToken(deviceCodeResponse.device_code, accountConfig);
      const saved = saveTokens(tokenResponse, accountConfig);

      console.error('\n✅ Authentication successful!');
      console.error(`Token expires: ${new Date(saved.expires_at).toLocaleString()}`);
      console.error(`Tokens saved to: ${accountConfig.tokenPath}\n`);

      return saved;
    } catch (error) {
      const msg = scrubSecrets(error.message || '');
      if (msg.includes('authorization_pending') || msg.includes('AADSTS70016')) {
        process.stderr.write('.');
      } else if (msg.includes('authorization_declined')) {
        throw new Error('User declined authorization');
      } else if (msg.includes('expired_token') || msg.includes('AADSTS70019')) {
        throw new Error('Device code expired - please try again');
      } else if (msg.includes('slow_down')) {
        await new Promise((resolve) => setTimeout(resolve, interval));
      } else {
        throw new Error(msg);
      }
    }
  }

  throw new Error('Authentication timed out');
}

if (require.main === module) {
  (async () => {
    const command = process.argv[2];
    let accountName;
    try {
      accountName = parseAccountFlag(process.argv);
    } catch (e) {
      console.error('❌', e.message);
      process.exit(1);
    }

    if (command === 'login') {
      // Default (non-blocking) behavior:
      //   - First call: requests a device code, writes pending state, prints
      //     the prompt, exits 0. The caller surfaces the code to the human.
      //   - Subsequent calls: poll once and return current status.
      // This is what makes the flow survive short LLM tool timeouts. Pass
      // --blocking to keep the legacy single-call behavior.
      const blocking = process.argv.includes('--blocking');

      try {
        if (blocking) {
          await authenticate(accountName);
          process.exit(0);
        }

        const accountConfig = getAccountConfig(accountName);
        const pendingPath = getPendingPath(accountConfig);
        let result;
        if (fs.existsSync(pendingPath)) {
          result = await pollAuthenticate(accountConfig);
        } else {
          result = await beginAuthenticate(accountConfig);
        }
        // Structured JSON to stdout for callers; pretty hint to stderr for humans.
        process.stdout.write(JSON.stringify(result, null, 2) + '\n');
        if (result.status === 'pending_started') {
          console.error(`\n📱 Open ${result.verification_uri} and enter code ${result.user_code}.`);
          console.error(`   Re-run \`node auth.js login --account=${accountConfig.name}\` every ${result.poll_interval_s}s until authenticated.\n`);
        } else if (result.status === 'pending') {
          console.error(`\n⏳ Still waiting for the browser step. Retry in ${result.next_poll_in_s}s.\n`);
        } else if (result.status === 'authenticated' || result.status === 'already_authenticated') {
          console.error('\n✅ Authenticated.\n');
        } else if (result.status === 'expired') {
          console.error('\n❌ Device code expired before completion. Run login again to get a fresh code.\n');
        } else if (result.status === 'declined') {
          console.error('\n❌ User declined the authorization request.\n');
        }
        const exitMap = {
          authenticated: 0,
          already_authenticated: 0,
          pending_started: 0,
          pending: 3,         // distinct exit code so callers know to loop
          expired: 4,
          declined: 4,
          no_pending_flow: 0
        };
        process.exit(exitMap[result.status] != null ? exitMap[result.status] : 1);
      } catch (err) {
        console.error('\n❌ Authentication failed:', scrubSecrets(err.message));
        process.exit(1);
      }
    } else if (command === 'diag') {
      try {
        const result = await diagnose();
        process.stdout.write(JSON.stringify(result, null, 2) + '\n');
        process.exit(0);
      } catch (err) {
        console.error('❌', scrubSecrets(err.message));
        process.exit(1);
      }
    } else if (command === 'status') {
      try {
        const accountConfig = getAccountConfig(accountName);
        const tokens = loadTokens(accountConfig);

        console.log(`Account: ${accountConfig.name}`);

        if (!tokens) {
          console.log('Status: ❌ Not authenticated');
          process.exit(1);
        }

        const expired = tokens.expires_at < Date.now();
        console.log(`Status: ${expired ? '⚠️  Expired' : '✅ Valid'}`);
        console.log(`Expires: ${new Date(tokens.expires_at).toLocaleString()}`);
        console.log(`Scopes: ${tokens.scope}`);
        process.exit(expired ? 1 : 0);
      } catch (err) {
        console.error('❌ Error:', scrubSecrets(err.message));
        process.exit(1);
      }
    } else if (command === 'token') {
      // Bearer tokens are credentials. Require explicit opt-in to print one to stdout.
      const confirmed = process.argv.includes('--confirm');
      if (!confirmed) {
        console.error('Refusing to print bearer token to stdout without --confirm.');
        console.error('Add --confirm to acknowledge that the token will appear in your terminal/log.');
        process.exit(2);
      }
      try {
        const token = await getAccessToken(accountName);
        process.stdout.write(token + '\n');
        process.exit(0);
      } catch (err) {
        console.error('❌ Failed to get token:', scrubSecrets(err.message));
        process.exit(1);
      }
    } else {
      console.log('Usage:');
      console.log('  node auth.js login [--account=name]                  - Non-blocking auth: returns user_code, then polls on each rerun.');
      console.log('                                                          Exit codes: 0=authenticated/started, 3=pending (rerun), 4=expired/declined.');
      console.log('  node auth.js login [--account=name] --blocking       - Legacy blocking flow (NOT recommended, may exceed LLM timeouts).');
      console.log('  node auth.js status [--account=name]                 - Check authentication status.');
      console.log('  node auth.js token  [--account=name] --confirm       - Print current access token (sensitive).');
      console.log('  node auth.js diag                                    - DNS + TLS reachability check for Microsoft endpoints.');
      console.log('\nIf --account is not specified, the default account is used.');
      console.log('\nDevice-code flow (non-blocking):');
      console.log('  1. Run `node auth.js login --account=work`. Prints verification URL + code.');
      console.log('  2. Complete the browser step in a normal session.');
      console.log('  3. Re-run the same command every ~5s until exit code is 0.');
      process.exit(1);
    }
  })();
}

module.exports = {
  authenticate,
  beginAuthenticate,
  pollAuthenticate,
  diagnose,
  getAccessToken,
  getPendingPath,
  loadTokens,
  saveTokens,
  refreshAccessToken
};
