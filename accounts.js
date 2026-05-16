#!/usr/bin/env node
'use strict';
/**
 * Office 365 Multi-Account Management
 * Handles configuration and switching between multiple Microsoft 365 accounts.
 *
 * Security:
 *   - Account names are validated against a strict allow-list (no path traversal).
 *   - tenantId / clientId must be UUIDs; email format is validated.
 *   - Config and token files are written atomically with mode 0600.
 *   - Client secrets are NEVER read from argv (visible via `ps` / shell history);
 *     they must be supplied via the AZURE_CLIENT_SECRET env var or interactive prompt.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const {
  assertAccountName,
  assertUuid,
  assertEmail,
  assertSafeChildPath,
  secureWriteFileSync,
  secureReadJsonSync,
  readSecretFromStdin
} = require('./security.js');

const HOME = os.homedir();
if (!HOME) {
  console.error('❌ Unable to resolve home directory; refusing to proceed.');
  process.exit(1);
}
const ACCOUNTS_CONFIG_PATH = path.join(HOME, '.hermes', 'auth', 'office365-accounts.json');
const ACCOUNTS_DIR = path.join(HOME, '.hermes', 'auth', 'office365');

function ensureDirectories() {
  fs.mkdirSync(path.dirname(ACCOUNTS_CONFIG_PATH), { recursive: true, mode: 0o700 });
  fs.mkdirSync(ACCOUNTS_DIR, { recursive: true, mode: 0o700 });
  try { fs.chmodSync(path.dirname(ACCOUNTS_CONFIG_PATH), 0o700); } catch (_) { /* best effort */ }
  try { fs.chmodSync(ACCOUNTS_DIR, 0o700); } catch (_) { /* best effort */ }
}

function loadAccounts() {
  ensureDirectories();
  const data = secureReadJsonSync(ACCOUNTS_CONFIG_PATH);
  if (!data) return { default: null, accounts: {} };
  if (typeof data !== 'object' || !data.accounts || typeof data.accounts !== 'object') {
    return { default: null, accounts: {} };
  }
  return data;
}

function saveAccounts(config) {
  ensureDirectories();
  secureWriteFileSync(ACCOUNTS_CONFIG_PATH, JSON.stringify(config, null, 2));
}

function getAccountTokenPath(accountName) {
  return assertSafeChildPath(ACCOUNTS_DIR, accountName);
}

function addAccount(name, tenantId, clientId, clientSecret, options = {}) {
  assertAccountName(name);
  assertUuid(tenantId, 'tenantId');
  assertUuid(clientId, 'clientId');
  if (typeof clientSecret !== 'string' || clientSecret.length < 8 || clientSecret.length > 512) {
    throw new Error('Invalid client secret');
  }
  if (options.email) assertEmail(options.email, 'email');
  if (options.description && (typeof options.description !== 'string' || options.description.length > 256)) {
    throw new Error('Description too long');
  }

  const config = loadAccounts();

  config.accounts[name] = {
    tenantId,
    clientId,
    clientSecret,
    email: options.email || null,
    description: options.description || null,
    addedAt: (config.accounts[name] && config.accounts[name].addedAt) || new Date().toISOString()
  };

  if (!config.default) config.default = name;

  saveAccounts(config);
  return config.accounts[name];
}

function removeAccount(name) {
  assertAccountName(name);
  const config = loadAccounts();

  if (!config.accounts[name]) {
    throw new Error(`Account "${name}" not found`);
  }

  const tokenPath = getAccountTokenPath(name);
  if (fs.existsSync(tokenPath)) {
    fs.unlinkSync(tokenPath);
  }

  delete config.accounts[name];

  if (config.default === name) {
    config.default = Object.keys(config.accounts)[0] || null;
  }

  saveAccounts(config);
}

function setDefault(name) {
  assertAccountName(name);
  const config = loadAccounts();

  if (!config.accounts[name]) {
    throw new Error(`Account "${name}" not found`);
  }

  config.default = name;
  saveAccounts(config);
}

function getAccount(name) {
  const config = loadAccounts();

  if (!name) name = config.default;
  if (!name) throw new Error('No account specified and no default account set');

  assertAccountName(name);

  const account = config.accounts[name];
  if (!account) throw new Error(`Account "${name}" not found`);

  return Object.assign({ name }, account, { tokenPath: getAccountTokenPath(name) });
}

function listAccounts() {
  const config = loadAccounts();
  return {
    default: config.default,
    accounts: Object.keys(config.accounts).map((name) => Object.assign(
      { name, isDefault: name === config.default },
      // Never include the client secret in list output.
      { tenantId: config.accounts[name].tenantId,
        clientId: config.accounts[name].clientId,
        email: config.accounts[name].email,
        description: config.accounts[name].description,
        addedAt: config.accounts[name].addedAt }
    ))
  };
}

function importLegacy() {
  // v1 of this skill stored a single token file at ~/.hermes/auth/microsoft-graph.json.
  // Pull it into the v2.x per-account layout if found.
  const legacyTokenPath = path.join(HOME, '.hermes', 'auth', 'microsoft-graph.json');

  if (!fs.existsSync(legacyTokenPath)) return null;

  const config = loadAccounts();
  if (Object.keys(config.accounts).length > 0) return null;

  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    console.error('⚠️  Legacy token file found but no credentials in environment');
    return null;
  }

  addAccount('primary', tenantId, clientId, clientSecret, {
    description: 'Imported from legacy setup'
  });

  const newTokenPath = getAccountTokenPath('primary');
  fs.copyFileSync(legacyTokenPath, newTokenPath);
  try { fs.chmodSync(newTokenPath, 0o600); } catch (_) { /* best effort */ }
  fs.unlinkSync(legacyTokenPath);

  console.error('✅ Imported legacy account as "primary"');
  return 'primary';
}

async function obtainSecret() {
  if (process.env.AZURE_CLIENT_SECRET && process.env.AZURE_CLIENT_SECRET.length >= 8) {
    return process.env.AZURE_CLIENT_SECRET;
  }
  if (!process.stdin.isTTY) {
    throw new Error('Client secret required. Set AZURE_CLIENT_SECRET env var or run in an interactive terminal.');
  }
  const secret = await readSecretFromStdin('Client secret (hidden): ');
  if (!secret || secret.length < 8) throw new Error('Client secret too short or empty.');
  return secret;
}

if (require.main === module) {
  (async () => {
    const command = process.argv[2];
    const args = process.argv.slice(3);

    try {
      if (command === 'list') {
        const result = listAccounts();
        console.log('📧 Office 365 Accounts:\n');
        if (result.accounts.length === 0) {
          console.log('No accounts configured.');
          console.log('\nAdd an account with:');
          console.log('  AZURE_CLIENT_SECRET=... node accounts.js add <name> <tenant-id> <client-id> [email] [description]');
          console.log('  (or omit the env var to be prompted interactively)');
        } else {
          result.accounts.forEach((acc) => {
            const defaultMarker = acc.isDefault ? ' [DEFAULT]' : '';
            console.log(`${acc.name}${defaultMarker}`);
            if (acc.email) console.log(`  Email: ${acc.email}`);
            if (acc.description) console.log(`  Description: ${acc.description}`);
            console.log(`  Added: ${new Date(acc.addedAt).toLocaleDateString()}`);
            console.log('');
          });
        }
      } else if (command === 'add') {
        if (args.length < 3) {
          console.error('Usage: AZURE_CLIENT_SECRET=... node accounts.js add <name> <tenant-id> <client-id> [email] [description]');
          console.error('       (omit the env var to be prompted for the secret without echo)');
          console.error('NOTE: passing the secret on the command line is refused — it would leak via `ps` and shell history.');
          process.exit(1);
        }

        const [name, tenantId, clientId, email, description] = args;
        const secret = await obtainSecret();
        addAccount(name, tenantId, clientId, secret, { email, description });
        console.log(`✅ Added account "${name}"`);
      } else if (command === 'remove') {
        if (args.length < 1) {
          console.error('Usage: node accounts.js remove <name>');
          process.exit(1);
        }
        removeAccount(args[0]);
        console.log(`✅ Removed account "${args[0]}"`);
      } else if (command === 'default') {
        if (args.length < 1) {
          console.error('Usage: node accounts.js default <name>');
          process.exit(1);
        }
        setDefault(args[0]);
        console.log(`✅ Set "${args[0]}" as default account`);
      } else if (command === 'import-legacy') {
        const result = importLegacy();
        if (!result) console.log('No legacy setup found or already imported');
      } else {
        console.log('Office 365 Multi-Account Management\n');
        console.log('Commands:');
        console.log('  list                      - List all accounts');
        console.log('  add <name> <tenant> <client> [email] [desc]');
        console.log('                              (secret via AZURE_CLIENT_SECRET env var or interactive prompt)');
        console.log('  remove <name>             - Remove account');
        console.log('  default <name>            - Set default account');
        console.log('  import-legacy             - Import from single-account setup');
        process.exit(1);
      }
    } catch (error) {
      console.error('❌ Error:', error.message);
      process.exit(1);
    }
  })();
}

module.exports = {
  loadAccounts,
  saveAccounts,
  addAccount,
  removeAccount,
  setDefault,
  getAccount,
  listAccounts,
  getAccountTokenPath,
  importLegacy
};
