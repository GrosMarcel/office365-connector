"""Standalone CLI for the office365-connector Hermes plugin.

Run with:  python -m office365_connector <command> [args]

Provided so a token file can be produced from any machine that can reach
login.microsoftonline.com — useful when the Hermes runtime is in a
container that hits LLM tool-call timeouts during the device-code flow,
or for emergency token refresh during an outage.

Commands:
  accounts list
  accounts add <name> <tenant-id> <client-id> [email] [description]
                       (AZURE_CLIENT_SECRET must be set in the environment)
  login [--account=NAME]
  status [--account=NAME]
  diag
"""

from __future__ import annotations

import json
import os
import sys

from . import _accounts, _auth
from ._security import scrub_secrets


def _parse_account_flag(argv: list[str]) -> tuple[str | None, list[str]]:
    account = None
    rest = []
    for a in argv:
        if a.startswith("--account="):
            account = a.split("=", 1)[1]
        else:
            rest.append(a)
    return account, rest


def _cmd_accounts(rest: list[str]) -> int:
    if not rest:
        print("Usage: accounts {list|add|remove|default|import-legacy} ...", file=sys.stderr)
        return 1
    sub = rest[0]
    args = rest[1:]
    if sub == "list":
        print(json.dumps(_accounts.list_accounts(), indent=2))
        return 0
    if sub == "add":
        if len(args) < 3:
            print(
                "Usage: accounts add <name> <tenant-id> <client-id> [email] [description]\n"
                "       AZURE_CLIENT_SECRET must be set in the environment.",
                file=sys.stderr,
            )
            return 1
        secret = os.environ.get("AZURE_CLIENT_SECRET")
        if not secret or len(secret) < 8:
            print("AZURE_CLIENT_SECRET env var is required.", file=sys.stderr)
            return 1
        name, tenant, client = args[0], args[1], args[2]
        email = args[3] if len(args) > 3 else None
        description = args[4] if len(args) > 4 else None
        _accounts.add_account(name, tenant, client, secret, email=email, description=description)
        print(f"Added account: {name}")
        return 0
    if sub == "remove":
        if not args:
            print("Usage: accounts remove <name>", file=sys.stderr)
            return 1
        _accounts.remove_account(args[0])
        print(f"Removed: {args[0]}")
        return 0
    if sub == "default":
        if not args:
            print("Usage: accounts default <name>", file=sys.stderr)
            return 1
        _accounts.set_default(args[0])
        print(f"Default: {args[0]}")
        return 0
    if sub == "import-legacy":
        result = _accounts.import_legacy()
        print(f"Imported: {result}" if result else "No legacy setup found.")
        return 0
    print(f"Unknown accounts subcommand: {sub}", file=sys.stderr)
    return 1


def _cmd_login(account: str | None) -> int:
    def _on_prompt(uri, code, expires_in):
        print("\n" + "=" * 60)
        print("  Open this URL in a browser:")
        print(f"    {uri}")
        print(f"  Enter this code: {code}")
        print(f"  Code expires in: {expires_in} seconds")
        print("=" * 60 + "\n  Waiting for you to complete the browser step...")

    try:
        result = _auth.authenticate(account, on_prompt=_on_prompt)
        print(json.dumps(result, indent=2))
        return 0
    except Exception as e:
        print(f"Auth failed: {scrub_secrets(str(e))}", file=sys.stderr)
        return 1


def _cmd_status(account: str | None) -> int:
    print(json.dumps(_auth.auth_status(account), indent=2))
    return 0


def _cmd_diag() -> int:
    print(json.dumps(_auth.diagnose(), indent=2))
    return 0


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0

    cmd = argv[0]
    account, rest = _parse_account_flag(argv[1:])

    try:
        if cmd == "accounts":
            return _cmd_accounts(rest)
        if cmd == "login":
            return _cmd_login(account)
        if cmd == "status":
            return _cmd_status(account)
        if cmd == "diag":
            return _cmd_diag()
        print(f"Unknown command: {cmd}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {scrub_secrets(str(e))}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
