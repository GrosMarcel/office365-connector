"""Office 365 multi-account management for the Hermes plugin.

Account config lives at ~/.hermes/auth/office365-accounts.json
Per-account tokens live at ~/.hermes/auth/office365/<name>.json
"""

from __future__ import annotations

import contextlib
import datetime as _dt
import json
import os
import shutil
from typing import Any

from ._security import (
    assert_account_name,
    assert_email,
    assert_safe_child_path,
    assert_uuid,
    secure_read_json,
    secure_write_file,
)

HOME = os.path.expanduser("~")
if not HOME or HOME == "~":
    raise RuntimeError("Unable to resolve home directory; refusing to proceed.")

ACCOUNTS_CONFIG_PATH = os.path.join(HOME, ".hermes", "auth", "office365-accounts.json")
ACCOUNTS_DIR = os.path.join(HOME, ".hermes", "auth", "office365")


def _ensure_directories() -> None:
    os.makedirs(os.path.dirname(ACCOUNTS_CONFIG_PATH), mode=0o700, exist_ok=True)
    os.makedirs(ACCOUNTS_DIR, mode=0o700, exist_ok=True)
    with contextlib.suppress(OSError):
        os.chmod(os.path.dirname(ACCOUNTS_CONFIG_PATH), 0o700)
    with contextlib.suppress(OSError):
        os.chmod(ACCOUNTS_DIR, 0o700)


def load_accounts() -> dict[str, Any]:
    _ensure_directories()
    data = secure_read_json(ACCOUNTS_CONFIG_PATH)
    if not data:
        return {"default": None, "accounts": {}}
    if not isinstance(data, dict) or not isinstance(data.get("accounts"), dict):
        return {"default": None, "accounts": {}}
    return data


def save_accounts(config: dict[str, Any]) -> None:
    _ensure_directories()
    secure_write_file(ACCOUNTS_CONFIG_PATH, json.dumps(config, indent=2))


def get_account_token_path(name: str) -> str:
    return assert_safe_child_path(ACCOUNTS_DIR, name)


def add_account(
    name: str,
    tenant_id: str,
    client_id: str,
    client_secret: str,
    email: str | None = None,
    description: str | None = None,
) -> dict[str, Any]:
    assert_account_name(name)
    assert_uuid(tenant_id, "tenantId")
    assert_uuid(client_id, "clientId")
    if not isinstance(client_secret, str) or not (8 <= len(client_secret) <= 512):
        raise ValueError("Invalid client secret")
    if email:
        assert_email(email, "email")
    if description is not None:
        if not isinstance(description, str) or len(description) > 256:
            raise ValueError("Description too long")

    config = load_accounts()
    existing = config["accounts"].get(name) or {}
    config["accounts"][name] = {
        "tenantId": tenant_id,
        "clientId": client_id,
        "clientSecret": client_secret,
        "email": email,
        "description": description,
        "addedAt": existing.get("addedAt") or _dt.datetime.utcnow().isoformat() + "Z",
    }

    if not config.get("default"):
        config["default"] = name

    save_accounts(config)
    return config["accounts"][name]


def remove_account(name: str) -> None:
    assert_account_name(name)
    config = load_accounts()
    if name not in config["accounts"]:
        raise ValueError(f'Account "{name}" not found')

    token_path = get_account_token_path(name)
    if os.path.exists(token_path):
        with contextlib.suppress(OSError):
            os.unlink(token_path)

    del config["accounts"][name]
    if config.get("default") == name:
        remaining = list(config["accounts"].keys())
        config["default"] = remaining[0] if remaining else None

    save_accounts(config)


def set_default(name: str) -> None:
    assert_account_name(name)
    config = load_accounts()
    if name not in config["accounts"]:
        raise ValueError(f'Account "{name}" not found')
    config["default"] = name
    save_accounts(config)


def get_account(name: str | None = None) -> dict[str, Any]:
    config = load_accounts()
    if not name:
        name = config.get("default")
    if not name:
        raise ValueError("No account specified and no default account set")

    assert_account_name(name)
    account = config["accounts"].get(name)
    if not account:
        raise ValueError(f'Account "{name}" not found')

    return {
        "name": name,
        "tenantId": account["tenantId"],
        "clientId": account["clientId"],
        "clientSecret": account["clientSecret"],
        "email": account.get("email"),
        "description": account.get("description"),
        "tokenPath": get_account_token_path(name),
    }


def list_accounts() -> dict[str, Any]:
    config = load_accounts()
    default = config.get("default")
    return {
        "default": default,
        "accounts": [
            {
                "name": name,
                "isDefault": name == default,
                "tenantId": acc["tenantId"],
                "clientId": acc["clientId"],
                "email": acc.get("email"),
                "description": acc.get("description"),
                "addedAt": acc.get("addedAt"),
            }
            for name, acc in config["accounts"].items()
        ],
    }


def import_legacy() -> str | None:
    """Migrate from v1 single-account layout (~/.hermes/auth/microsoft-graph.json).

    Requires AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET in the env
    so we can reconstruct the per-account record. Token file is moved, not copied.
    """
    legacy_token_path = os.path.join(HOME, ".hermes", "auth", "microsoft-graph.json")
    if not os.path.exists(legacy_token_path):
        return None

    config = load_accounts()
    if config["accounts"]:
        return None

    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not (tenant_id and client_id and client_secret):
        return None

    add_account(
        "primary",
        tenant_id,
        client_id,
        client_secret,
        description="Imported from legacy setup",
    )

    new_token_path = get_account_token_path("primary")
    shutil.copy2(legacy_token_path, new_token_path)
    with contextlib.suppress(OSError):
        os.chmod(new_token_path, 0o600)
    with contextlib.suppress(OSError):
        os.unlink(legacy_token_path)

    return "primary"
