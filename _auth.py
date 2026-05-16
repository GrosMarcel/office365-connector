"""Microsoft Graph OAuth 2.0 Device Code Flow Authentication for Hermes.

Tokens written atomically at mode 0600 under
~/.hermes/auth/office365/<account>.json. Refresh is automatic.

The device-code flow is split into a non-blocking `begin` + `poll` pair so
that LLM tool-call timeouts (typically 30-120 s) don't kill the flow while
the human completes the browser step (which can take minutes). The blocking
`authenticate()` is kept for the standalone CLI.
"""

from __future__ import annotations

import contextlib
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

PENDING_PREFIX = ".pending-"

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


def _pending_path(cfg: dict[str, Any]) -> str:
    """Per-account state file for an in-flight device-code flow.

    Mode 0600 because the stored device_code is itself a credential — anyone
    who can read it can race the polling endpoint.
    """
    token_dir = os.path.dirname(cfg["tokenPath"])
    return os.path.join(token_dir, f"{PENDING_PREFIX}{cfg['name']}.json")


def begin_authenticate(account_name: str | None = None) -> dict[str, Any]:
    """Start the device-code flow and return the prompt info immediately.

    The returned `user_code` / `verification_uri` must be surfaced to the
    human. Then poll with `poll_authenticate()` until the human completes
    the browser step. Each call to this function invalidates any previously
    pending flow for the same account.
    """
    cfg = _get_account_config(account_name)

    existing = _load_tokens(cfg)
    if existing and existing["expires_at"] > int(time.time() * 1000):
        return {
            "status": "already_authenticated",
            "account": cfg["name"],
            "expires_at_ms": existing["expires_at"],
        }

    dc = _request_device_code(cfg)
    if not isinstance(dc, dict) or not dc.get("device_code"):
        raise RuntimeError("Authority did not return a device code")

    interval = max(1, int(dc.get("interval") or 5))
    expires_in = int(dc.get("expires_in") or 600)

    pending = {
        "device_code": dc["device_code"],
        "interval": interval,
        "expires_at": int(time.time()) + expires_in,
        "account": cfg["name"],
    }
    secure_write_file(_pending_path(cfg), json.dumps(pending))

    return {
        "status": "pending",
        "account": cfg["name"],
        "verification_uri": dc.get("verification_uri"),
        "user_code": dc.get("user_code"),
        "expires_in": expires_in,
        "poll_interval_s": interval,
        "instructions": (
            f"Open {dc.get('verification_uri')} and enter code "
            f"{dc.get('user_code')}. Then call office365_auth_login_poll "
            f"every ~{interval}s until status is 'authenticated'."
        ),
    }


def poll_authenticate(account_name: str | None = None) -> dict[str, Any]:
    """Do a single non-blocking poll of the token endpoint.

    Returns one of:
      - {"status": "authenticated", ...}
      - {"status": "pending", "next_poll_in_s": N}
      - {"status": "expired"}        # device code TTL elapsed
      - {"status": "declined"}       # user clicked Cancel
      - {"status": "no_pending_flow"}# nothing in progress for this account

    Side effect: on success, writes the token file and removes the pending
    state file. On terminal failure, removes the pending state file.
    """
    cfg = _get_account_config(account_name)
    pending_path = _pending_path(cfg)
    pending = secure_read_json(pending_path)

    if not pending:
        existing = _load_tokens(cfg)
        if existing and existing["expires_at"] > int(time.time() * 1000):
            return {
                "status": "already_authenticated",
                "account": cfg["name"],
                "expires_at_ms": existing["expires_at"],
            }
        return {
            "status": "no_pending_flow",
            "account": cfg["name"],
            "message": "Call office365_auth_login first to start the flow.",
        }

    if int(time.time()) >= int(pending["expires_at"]):
        with contextlib.suppress(OSError):
            os.unlink(pending_path)
        return {"status": "expired", "account": cfg["name"]}

    try:
        token_response = _poll_for_token(pending["device_code"], cfg)
        saved = _save_tokens(token_response, cfg)
        with contextlib.suppress(OSError):
            os.unlink(pending_path)
        return {
            "status": "authenticated",
            "account": cfg["name"],
            "expires_at_ms": saved["expires_at"],
        }
    except RuntimeError as e:
        msg = scrub_secrets(str(e))
        if "authorization_pending" in msg or "AADSTS70016" in msg:
            return {
                "status": "pending",
                "account": cfg["name"],
                "next_poll_in_s": int(pending["interval"]),
            }
        if "authorization_declined" in msg:
            with contextlib.suppress(OSError):
                os.unlink(pending_path)
            return {"status": "declined", "account": cfg["name"]}
        if "expired_token" in msg or "AADSTS70019" in msg:
            with contextlib.suppress(OSError):
                os.unlink(pending_path)
            return {"status": "expired", "account": cfg["name"]}
        if "slow_down" in msg:
            return {
                "status": "pending",
                "account": cfg["name"],
                "next_poll_in_s": int(pending["interval"]) * 2,
            }
        raise


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
    pending = secure_read_json(_pending_path(cfg))
    pending_state: dict[str, Any] | None = None
    if pending:
        pending_state = {
            "expires_at": int(pending["expires_at"]),
            "expires_in_s": max(0, int(pending["expires_at"]) - int(time.time())),
        }

    if not tokens:
        return {
            "account": cfg["name"],
            "authenticated": False,
            "pending_flow": pending_state,
        }

    now_ms = int(time.time() * 1000)
    expired = tokens["expires_at"] < now_ms
    return {
        "account": cfg["name"],
        "authenticated": not expired,
        "expired": expired,
        "expires_at_ms": tokens["expires_at"],
        "scope": tokens.get("scope"),
        "pending_flow": pending_state,
    }


def diagnose() -> dict[str, Any]:
    """Network reachability check for the two pinned hosts.

    Useful for diagnosing the "tool call times out" case — distinguishes
    DNS failure, TCP/TLS failure, and HTTP-layer failure inside containers
    where outbound networking is restricted.
    """
    import socket
    import ssl
    import urllib.error
    import urllib.request

    results: dict[str, Any] = {}
    for host in ("login.microsoftonline.com", "graph.microsoft.com"):
        entry: dict[str, Any] = {}
        try:
            entry["ip"] = socket.gethostbyname(host)
            entry["dns"] = "ok"
        except Exception as e:
            entry["dns"] = f"error: {e}"
            results[host] = entry
            continue
        try:
            ctx = ssl.create_default_context()
            req = urllib.request.Request(f"https://{host}/", method="HEAD")
            urllib.request.urlopen(req, timeout=5, context=ctx).read(0)
            entry["https"] = "ok"
        except urllib.error.HTTPError as e:
            # Any HTTP response means TLS + transport are fine.
            entry["https"] = f"ok (http {e.code})"
        except Exception as e:
            entry["https"] = f"error: {e}"
        results[host] = entry
    return {"hosts": results}
