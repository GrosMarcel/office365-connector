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
      try {
        await authenticate(accountName);
      } catch (err) {
        console.error('\n❌ Authentication failed:', scrubSecrets(err.message));
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
      console.log('  node auth.js login [--account=name]              - Authenticate with Microsoft');
      console.log('  node auth.js status [--account=name]             - Check authentication status');
      console.log('  node auth.js token [--account=name] --confirm    - Print current access token (sensitive)');
      console.log('\nIf --account is not specified, the default account is used.');
      process.exit(1);
    }
  })();
}

module.exports = {
  authenticate,
  getAccessToken,
  loadTokens,
  saveTokens,
  refreshAccessToken
};
