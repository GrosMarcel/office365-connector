"""Microsoft Graph OAuth 2.0 Device Code Flow Authentication for Hermes.

Tokens written atomically at mode 0600 under
~/.hermes/auth/office365/<account>.json. Refresh is automatic.
"""

from __future__ import annotations

import json
import os
import time
import urllib.parse
from typing import Any

from ._accounts import get_account
from ._security import (
    assert_uuid,
    https_json_request,
    scrub_secrets,
    secure_read_json,
    secure_write_file,
)

SCOPES = " ".join(
    [
        "User.Read",
        "Mail.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "Calendars.Read",
        "Calendars.ReadWrite",
        "Contacts.Read",
        "Contacts.ReadWrite",
        "offline_access",
    ]
)


def _get_account_config(account_name: str | None) -> dict[str, Any]:
    try:
        return get_account(account_name)
    except Exception:
        if account_name:
            raise
        tenant_id = os.environ.get("AZURE_TENANT_ID")
        client_id = os.environ.get("AZURE_CLIENT_ID")
        client_secret = os.environ.get("AZURE_CLIENT_SECRET")
        if not (tenant_id and client_id and client_secret):
            raise RuntimeError(
                "No account configured and no credentials in environment. "
                "Use office365_accounts_add to register an account."
            ) from None
        assert_uuid(tenant_id, "AZURE_TENANT_ID")
        assert_uuid(client_id, "AZURE_CLIENT_ID")
        home = os.path.expanduser("~")
        return {
            "name": "legacy",
            "tenantId": tenant_id,
            "clientId": client_id,
            "clientSecret": client_secret,
            "tokenPath": os.path.join(home, ".hermes", "auth", "microsoft-graph.json"),
        }


def _authority_for(cfg: dict[str, Any]) -> str:
    assert_uuid(cfg["tenantId"], "tenantId")
    return f"https://login.microsoftonline.com/{cfg['tenantId']}"


def _request_device_code(cfg: dict[str, Any]) -> dict[str, Any]:
    assert_uuid(cfg["clientId"], "clientId")
    url = f"{_authority_for(cfg)}/oauth2/v2.0/devicecode"
    body = urllib.parse.urlencode({"client_id": cfg["clientId"], "scope": SCOPES})
    return https_json_request(
        url,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        body=body,
    )


def _poll_for_token(device_code: str, cfg: dict[str, Any]) -> dict[str, Any]:
    assert_uuid(cfg["clientId"], "clientId")
    url = f"{_authority_for(cfg)}/oauth2/v2.0/token"
    body = urllib.parse.urlencode(
        {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": cfg["clientId"],
            "device_code": device_code,
        }
    )
    return https_json_request(
        url,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        body=body,
    )


def _refresh_access_token(refresh_token: str, cfg: dict[str, Any]) -> dict[str, Any]:
    assert_uuid(cfg["clientId"], "clientId")
    url = f"{_authority_for(cfg)}/oauth2/v2.0/token"
    # Device-code flow is a public client — never send client_secret on refresh.
    body = urllib.parse.urlencode(
        {
            "grant_type": "refresh_token",
            "client_id": cfg["clientId"],
            "refresh_token": refresh_token,
            "scope": SCOPES,
        }
    )
    return https_json_request(
        url,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        body=body,
    )


def _save_tokens(tokens: dict[str, Any], cfg: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(tokens, dict) or not isinstance(tokens.get("access_token"), str):
        raise RuntimeError("Invalid token response from authority")
    if not isinstance(tokens.get("expires_in"), (int, float)):
        raise RuntimeError("Invalid token response from authority")

    data = {
        "access_token": tokens["access_token"],
        "refresh_token": tokens.get("refresh_token"),
        "expires_at": int(time.time() * 1000) + int(tokens["expires_in"]) * 1000,
        "scope": tokens.get("scope"),
    }
    secure_write_file(cfg["tokenPath"], json.dumps(data, indent=2))
    return data


def _load_tokens(cfg: dict[str, Any]):
    try:
        return secure_read_json(cfg["tokenPath"])
    except Exception:
        return None


def get_access_token(account_name: str | None = None) -> str:
    cfg = _get_account_config(account_name)
    tokens = _load_tokens(cfg)
    if not tokens:
        raise RuntimeError(
            f'Not authenticated for account "{cfg["name"]}". '
            f"Run office365_auth_login first."
        )

    now_ms = int(time.time() * 1000)
    if tokens["expires_at"] < now_ms + 5 * 60 * 1000:
        if not tokens.get("refresh_token"):
            raise RuntimeError("No refresh token available; please re-authenticate.")
        refreshed = _refresh_access_token(tokens["refresh_token"], cfg)
        return _save_tokens(refreshed, cfg)["access_token"]

    return tokens["access_token"]


def authenticate(account_name: str | None = None, on_prompt=None) -> dict[str, Any]:
    """Run the OAuth 2.0 device code flow.

    `on_prompt(verification_uri, user_code, expires_in)` is called once the
    user code is available so callers (e.g. the LLM tool handler) can surface
    it to the user. Returns the saved token record on success.
    """
    cfg = _get_account_config(account_name)

    existing = _load_tokens(cfg)
    if existing and existing["expires_at"] > int(time.time() * 1000):
        return {"status": "already_authenticated", "expires_at": existing["expires_at"]}

    dc = _request_device_code(cfg)
    if not isinstance(dc, dict) or not dc.get("device_code"):
        raise RuntimeError("Authority did not return a device code")

    if on_prompt:
        on_prompt(dc.get("verification_uri"), dc.get("user_code"), dc.get("expires_in"))

    interval = max(1, int(dc.get("interval") or 5))
    deadline = time.time() + int(dc.get("expires_in") or 600)

    while time.time() < deadline:
        time.sleep(interval)
        try:
            token_response = _poll_for_token(dc["device_code"], cfg)
            saved = _save_tokens(token_response, cfg)
            return {
                "status": "authenticated",
                "account": cfg["name"],
                "expires_at": saved["expires_at"],
                "token_path": cfg["tokenPath"],
            }
        except RuntimeError as e:
            msg = scrub_secrets(str(e))
            if "authorization_pending" in msg or "AADSTS70016" in msg:
                continue
            if "authorization_declined" in msg:
                raise RuntimeError("User declined authorization") from None
            if "expired_token" in msg or "AADSTS70019" in msg:
                raise RuntimeError("Device code expired - please try again") from None
            if "slow_down" in msg:
                time.sleep(interval)
                continue
            raise

    raise RuntimeError("Authentication timed out")


def auth_status(account_name: str | None = None) -> dict[str, Any]:
    cfg = _get_account_config(account_name)
    tokens = _load_tokens(cfg)
    if not tokens:
        return {"account": cfg["name"], "authenticated": False}

    now_ms = int(time.time() * 1000)
    expired = tokens["expires_at"] < now_ms
    return {
        "account": cfg["name"],
        "authenticated": not expired,
        "expired": expired,
        "expires_at_ms": tokens["expires_at"],
        "scope": tokens.get("scope"),
    }
