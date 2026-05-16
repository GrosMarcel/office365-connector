# Changelog

All notable changes to this project will be documented in this file.

## [3.0.0] - 2026-05-16

### Changed — Hermes Agent plugin port

Re-packaged the skill as a Hermes Agent plugin following the [official plugin
guide](https://hermes-agent.nousresearch.com/docs/guides/build-a-hermes-plugin).

**New plugin layout:**
- `plugin.yaml` — manifest declaring name, version, 17 provided tools, and
  required environment variables.
- `__init__.py` — `register(ctx)` entry point that binds every tool against
  the central Hermes tool registry.
- `schemas.py` — OpenAI function-calling schemas for each tool.
- `tools.py` — handlers returning JSON strings per the Hermes contract.
- `_security.py` / `_accounts.py` / `_auth.py` / `_graph.py` — Python port
  of the Node.js modules.

**Runtime changes:**
- Node.js CLI scripts replaced by Python 3 modules (stdlib only).
- Operations are now exposed as model-callable tools rather than shell
  commands; the agent invokes them directly.

**Preserved from v2.0.0:**
- Multi-account model, per-account token isolation, default-account
  selection, legacy import path.
- Storage layout under `~/.hermes/auth/`.
- Security posture: UUID/email/Graph-ID validation, host-pinned HTTPS,
  atomic 0600 token writes, ANSI/C0/C1 sanitization, secret scrubbing.
- Backward-compatible env-var fallback (`AZURE_TENANT_ID` / `AZURE_CLIENT_ID`
  / `AZURE_CLIENT_SECRET`).

**Not yet ported in v3.0.0:**
- Contact tools (read / search / write). Tracked for v3.1.

### Migration from v2.0.0

The on-disk format is unchanged. To switch:

1. Install the plugin: `cp -r office365-connector ~/.hermes/plugins/` then
   `hermes plugins enable office365-connector`.
2. Existing accounts and tokens at `~/.hermes/auth/office365-accounts.json`
   and `~/.hermes/auth/office365/` are picked up automatically.
3. The legacy Node.js scripts (`accounts.js`, `auth.js`, etc.) are no
   longer used by the agent — they remain in the directory for reference
   and can be removed at your convenience.

## [2.0.0] - 2026-02-09

### Added - Multi-Account Support

**New Features:**
- Multi-account configuration system
- Per-account authentication and token management
- Account switching via `--account=<name>` flag
- Default account selection
- Legacy single-account import tool
- Comprehensive multi-account documentation

**New Files:**
- `accounts.js` - Account management CLI
- `MULTI-ACCOUNT.md` - Complete usage guide
- `CREDITS.md` - Attribution and acknowledgments
- `CHANGELOG.md` - This file

**Enhanced Files:**
- `auth.js` - Now supports multiple accounts with `--account=` flag
- `email.js` - Multi-account email operations
- `calendar.js` - Multi-account calendar operations
- `send-email.js` - Send from specific accounts
- `cancel-event.js` - Cancel events from specific accounts

### Changed

- Token storage moved from single file to per-account files in `~/.hermes-agent/auth/office365/`
- Account configuration stored in `~/.hermes-agent/auth/office365-accounts.json`
- All CLI scripts now accept `--account=<name>` parameter

### Maintained

- Full backward compatibility with v1.0.0
- Environment variable fallback (AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)
- All original functionality preserved
- No breaking changes for existing single-account users

### Migration

Existing users can:
1. Continue using without changes (environment variables still work)
2. Import existing setup: `node accounts.js import-legacy`
3. Add additional accounts: `node accounts.js add <name> ...`

## [1.0.0] - Original Release

Original Office 365 Connector skill with single-account support.

**Features:**
- OAuth 2.0 Device Code Flow authentication
- Email operations (read, search, send, reply)
- Calendar operations (read events, cancel)
- Contact operations (read, search)
- Automatic token refresh
- Azure App Registration setup guide
