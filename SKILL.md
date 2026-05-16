---
name: office365-connector
description: Hermes Agent plugin that exposes Office 365 / Outlook email, calendar, and account management as model-callable tools. Backed by Microsoft Graph with OAuth 2.0 device-code authentication, per-account token isolation, and full multi-account support.
---

# Office 365 Connector — Hermes Agent Plugin

## Overview

This plugin gives **Hermes Agent** native access to **Office 365 / Outlook**: read and send email, read and modify calendar events, and manage one or more Microsoft 365 identities from a single agent. Every operation is exposed as a tool the model can call directly — no shell-outs, no CLI wrappers.

It is the v3.0.0 rewrite of the original `office365-connector` skill: the Node.js CLI scripts have been replaced with a Python package that conforms to the [Hermes plugin format](https://hermes-agent.nousresearch.com/docs/guides/build-a-hermes-plugin) — a `plugin.yaml` manifest, a `register(ctx)` entry point in `__init__.py`, OpenAI-style function schemas in `schemas.py`, and JSON-returning handlers in `tools.py`.

**Attribution:** v2.0.0 multi-account work by **Matthew Gordon** ([matt@workandthrive.ai](mailto:matt@workandthrive.ai)). v3.0.0 Hermes port keeps the same model. See [CREDITS.md](CREDITS.md).

## What's New in v3.0.0 — Hermes Plugin Port

- **Hermes plugin layout** — `plugin.yaml`, `__init__.py`, `schemas.py`, `tools.py` per the developer guide
- **Python rewrite** — all runtime code ported from Node.js to Python 3 (stdlib only, no external deps)
- **Model-callable tools** — 17 registered tools (account management, auth, email read/write, calendar read/cancel) the LLM can invoke directly
- **Same security posture** — UUID/email/Graph-ID validation, host-pinned HTTPS, atomic 0600 token writes, ANSI/secret scrubbing — all carried over
- **Same on-disk layout** — accounts at `~/.hermes/auth/office365-accounts.json`, tokens at `~/.hermes/auth/office365/<name>.json`

See [CHANGELOG.md](CHANGELOG.md) for the full version history.

## Installation

Pick the channel that fits your setup. All three end with Hermes calling
`register(ctx)` and exposing the 17 tools.

### 1. `pip install` from GitHub (recommended for remote / Docker)

Once the plugin lives in a GitHub repo, a single command installs it into
any Hermes runtime — local, VPS, Docker, Hostinger, NixOS — and Hermes
picks it up automatically via the `hermes_agent.plugins` entry point. No
file copy, no `enable` step.

```bash
pip install "git+https://github.com/<owner>/office365-connector.git@v3.0.0"
```

Inside a Docker container:

```bash
docker exec hermes pip install \
  "git+https://github.com/<owner>/office365-connector.git@v3.0.0"
docker restart hermes
```

To pin to a tag, swap `@v3.0.0` for any commit / tag / branch ref.

### 2. `pip install` from PyPI

Once published (`python -m build && twine upload dist/*`):

```bash
pip install hermes-plugin-office365-connector
```

This is the cleanest "store-style" install — no GitHub URL, no `enable`
step, version-pinnable like any other Python package.

### 3. Local directory in `~/.hermes/plugins/`

For development or air-gapped hosts:

```bash
mkdir -p ~/.hermes/plugins
cp -r office365-connector-2.0.0 ~/.hermes/plugins/office365-connector
hermes plugins enable office365-connector
```

Hermes auto-discovers user-level plugins on next start. Per the docs,
`~/.hermes/plugins/<name>` overrides a pip-installed plugin of the same
name — handy for hot-patching without re-publishing.

### Docker / Hostinger one-shot

Mount a persistent volume for the OAuth tokens, then pip-install from
GitHub at build or runtime:

```yaml
services:
  hermes:
    image: <your-hermes-image>
    volumes:
      - hermes-auth:/root/.hermes/auth
    environment:
      AZURE_TENANT_ID: "${AZURE_TENANT_ID}"
      AZURE_CLIENT_ID: "${AZURE_CLIENT_ID}"
      AZURE_CLIENT_SECRET: "${AZURE_CLIENT_SECRET}"
    command: >
      sh -c "pip install --no-cache-dir
             git+https://github.com/<owner>/office365-connector.git@v3.0.0
             && exec hermes"
volumes:
  hermes-auth:
```

Secrets go in a `.env` next to the compose file (chmod 600, never
committed). Tokens persist across container restarts because the
per-account JSONs live in the `hermes-auth` volume.

## Plugin Layout

```
office365-connector/
├── plugin.yaml              # Hermes manifest: name, version, tool + env declarations
├── __init__.py              # register(ctx) — registers all 17 tools at startup
├── schemas.py               # OpenAI function-calling schemas for every tool
├── tools.py                 # Handler functions: (args, **kwargs) -> JSON string
├── _security.py             # UUID/email/Graph-ID validation, host-pinned HTTPS, atomic writes
├── _accounts.py             # Multi-account config (load / add / remove / default / list)
├── _auth.py                 # OAuth 2.0 device-code flow + token refresh
├── _graph.py                # Microsoft Graph: email read/send, calendar read/cancel
├── references/              # Setup guide + permissions reference
├── MULTI-ACCOUNT.md
├── CHANGELOG.md
└── CREDITS.md
```

## Registered Tools

All tools accept an optional `account` argument; omitted means "use the default account."

### Account management
- `office365_accounts_list` — list accounts and the default
- `office365_accounts_add` — register a new account (client secret read from `AZURE_CLIENT_SECRET`, never from args)
- `office365_accounts_remove` — delete an account and its stored tokens
- `office365_accounts_set_default` — mark an account as default
- `office365_accounts_import_legacy` — migrate a v1 single-account setup

### Authentication
- `office365_auth_login` — run the OAuth 2.0 device-code flow; returns the verification URL + user code
- `office365_auth_status` — report whether the account is authenticated and token expiry

### Email
- `office365_email_recent` — most recent inbox messages
- `office365_email_search` — full-text $search over the mailbox
- `office365_email_from_sender` — messages from a given sender
- `office365_email_read` — fetch a full message by Graph ID
- `office365_email_send` — send an email (to/cc/replyTo/HTML)
- `office365_email_reply` — inline reply to a message

### Calendar
- `office365_calendar_today` — today's events
- `office365_calendar_week` — this week's events
- `office365_calendar_range` — events between arbitrary start and end datetimes
- `office365_calendar_cancel_event` — cancel an event with an optional comment

## Tool Contract

Every handler follows the Hermes contract from the [plugin guide](https://hermes-agent.nousresearch.com/docs/guides/build-a-hermes-plugin):

```python
def handler(args: dict, **kwargs) -> str:
    """args: parameters the LLM passed. Returns a JSON string — always, even on error."""
```

Success payload:

```json
{"ok": true, "result": ...}
```

Failure payload (with secrets scrubbed):

```json
{"ok": false, "error": "..."}
```

## Prerequisites

Before adding the first account, complete an Azure App Registration to obtain:

1. **Tenant ID** — Azure AD tenant identifier (UUID)
2. **Client ID** — application (client) ID (UUID)
3. **Client Secret** — application secret value (8–512 chars)

Required delegated permissions: `User.Read`, `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `Contacts.Read`, `Contacts.ReadWrite`, `offline_access`. See [references/setup-guide.md](references/setup-guide.md) and [references/permissions.md](references/permissions.md).

## First-Run Walkthrough

```text
You: Connect my work Office 365 account.
Agent → office365_accounts_add(name="work", tenant_id=..., client_id=..., email="me@work.com")
       (with AZURE_CLIENT_SECRET present in env)
Agent → office365_auth_login(account="work")
Agent: "Open https://microsoft.com/devicelogin and enter code ABC-DEF-123."
You: [completes browser flow]
Agent: "Authenticated. Token expires in 1 hour; refresh is automatic."

You: What's on my calendar today?
Agent → office365_calendar_today(account="work")
Agent: "You have 3 events today: ..."
```

## Configuration

Accounts and tokens live under `~/.hermes/`:

- `~/.hermes/auth/office365-accounts.json` — non-secret account metadata + the registered client secret
- `~/.hermes/auth/office365/<name>.json` — per-account access and refresh tokens (mode 0600)

Single-account legacy users can still rely on environment variables — `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` — which the auth path uses as a fallback when no account is registered.

Hermes-style env wiring:

```yaml
env:
  vars:
    AZURE_TENANT_ID: "your-tenant-id"
    AZURE_CLIENT_ID: "your-client-id"
    AZURE_CLIENT_SECRET: "your-client-secret"
```

## Authentication Flow

OAuth 2.0 Device Code Flow, run inside `office365_auth_login`:

1. Plugin requests a device code from `login.microsoftonline.com`.
2. Tool returns the verification URL and short user code to the agent, which surfaces both to the human.
3. Plugin polls the token endpoint at the authority-supplied interval.
4. On success, access + refresh tokens are written atomically (mode 0600) to the per-account file.
5. Subsequent tool calls refresh the access token transparently whenever it's within 5 minutes of expiry.

## Security Posture (carried over from v2)

- UUIDs validated before being concatenated into OAuth URLs (no host/path injection)
- Hostnames pinned to `graph.microsoft.com` and `login.microsoftonline.com`
- 30 s timeout and 8 MB response cap on every outbound HTTPS request
- All printed strings (subjects, names, previews) stripped of ANSI / C0 / C1 control codes
- All error messages scrubbed of `access_token` / `refresh_token` / `client_secret` / `device_code` / `user_code` before being returned to the model
- Client secret never accepted via tool args — only via `AZURE_CLIENT_SECRET`
- Account names validated against a strict allow-list (no path traversal into the token directory)

## Rate Limits

Microsoft Graph: 130,000 requests/hour per app; per-user limits vary. The plugin surfaces 429 responses verbatim so the agent can back off — automatic retry can be added in a future revision.

## Troubleshooting

**"Not authenticated for account 'x'"** — call `office365_auth_login` for that account.

**"No account specified and no default account set"** — call `office365_accounts_set_default` or pass `account="..."` explicitly.

**"AADSTS700016 / 65001 / 700082"** — verify the Azure App Registration, ensure consent has been granted, or re-run `office365_auth_login`.

**"403 Forbidden"** — check that the delegated permissions in [references/permissions.md](references/permissions.md) are granted and consented in Azure.

## Limitations

- Attachments: max 4 MB per attachment (Graph limit)
- Email recipients: max 500 per message (also enforced client-side)
- Calendar events: limited to ~1,095 days into the future
- The plugin does **not** currently expose contact operations (read/search/write) as Hermes tools — calendar + email only in v3.0.0. Contact tools will land in v3.1.

## Reference Documentation

- [MULTI-ACCOUNT.md](MULTI-ACCOUNT.md) — multi-account usage notes (conceptually unchanged from v2)
- [CHANGELOG.md](CHANGELOG.md) — version history
- [CREDITS.md](CREDITS.md) — attribution
- [references/setup-guide.md](references/setup-guide.md) — Azure App Registration walkthrough
- [references/permissions.md](references/permissions.md) — permission reference

## Hermes Resources

- [Plugin Guide](https://hermes-agent.nousresearch.com/docs/guides/build-a-hermes-plugin)
- [Tools Runtime](https://hermes-agent.nousresearch.com/docs/developer-guide/tools-runtime)
- [Architecture Overview](https://hermes-agent.nousresearch.com/docs/developer-guide/architecture)

## Microsoft Resources

- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/api/overview)
- [Delegated vs Application Permissions](https://learn.microsoft.com/en-us/graph/auth/auth-concepts)
- [Throttling](https://learn.microsoft.com/en-us/graph/throttling)

## License

Maintains the original skill's licensing. See [CREDITS.md](CREDITS.md).
