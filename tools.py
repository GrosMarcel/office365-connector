"""Tool handlers for the office365-connector Hermes plugin.

Every handler obeys the Hermes contract:
  - accepts `args: dict` (parameters from the model) plus `**kwargs`
  - returns a JSON-encoded string (always — on success and on error)
"""

from __future__ import annotations

import json
import os
from typing import Any

from . import _accounts, _auth, _graph
from ._security import scrub_secrets


def _ok(payload: Any) -> str:
    return json.dumps({"ok": True, "result": payload}, ensure_ascii=False, default=str)


def _err(exc: Exception) -> str:
    return json.dumps(
        {"ok": False, "error": scrub_secrets(str(exc) or exc.__class__.__name__)},
        ensure_ascii=False,
    )


def _account(args: dict) -> str | None:
    value = args.get("account")
    if value in (None, ""):
        return None
    return value


# -------- accounts --------

def accounts_list(args: dict, **_kwargs) -> str:
    try:
        return _ok(_accounts.list_accounts())
    except Exception as e:
        return _err(e)


def accounts_add(args: dict, **_kwargs) -> str:
    try:
        secret = os.environ.get("AZURE_CLIENT_SECRET")
        if not secret or len(secret) < 8:
            raise RuntimeError(
                "AZURE_CLIENT_SECRET must be set in the environment before calling "
                "office365_accounts_add. Secrets are never accepted as tool arguments."
            )
        result = _accounts.add_account(
            name=args["name"],
            tenant_id=args["tenant_id"],
            client_id=args["client_id"],
            client_secret=secret,
            email=args.get("email"),
            description=args.get("description"),
        )
        return _ok(
            {
                "name": args["name"],
                "email": result.get("email"),
                "description": result.get("description"),
                "addedAt": result.get("addedAt"),
            }
        )
    except Exception as e:
        return _err(e)


def accounts_remove(args: dict, **_kwargs) -> str:
    try:
        _accounts.remove_account(args["name"])
        return _ok({"removed": args["name"]})
    except Exception as e:
        return _err(e)


def accounts_set_default(args: dict, **_kwargs) -> str:
    try:
        _accounts.set_default(args["name"])
        return _ok({"default": args["name"]})
    except Exception as e:
        return _err(e)


def accounts_import_legacy(args: dict, **_kwargs) -> str:
    try:
        result = _accounts.import_legacy()
        if not result:
            return _ok({"imported": False, "reason": "no legacy setup found or already imported"})
        return _ok({"imported": True, "as": result})
    except Exception as e:
        return _err(e)


# -------- auth --------

def auth_login(args: dict, **_kwargs) -> str:
    """Start the device-code flow and return immediately with the prompt info.

    Non-blocking: the LLM should surface the user_code + verification_uri to
    the human, then call `office365_auth_login_poll` until status is
    `authenticated` (or `expired` / `declined`).
    """
    try:
        return _ok(_auth.begin_authenticate(_account(args)))
    except Exception as e:
        return _err(e)


def auth_login_poll(args: dict, **_kwargs) -> str:
    try:
        return _ok(_auth.poll_authenticate(_account(args)))
    except Exception as e:
        return _err(e)


def auth_status(args: dict, **_kwargs) -> str:
    try:
        return _ok(_auth.auth_status(_account(args)))
    except Exception as e:
        return _err(e)


def auth_diag(args: dict, **_kwargs) -> str:
    try:
        return _ok(_auth.diagnose())
    except Exception as e:
        return _err(e)


# -------- email read --------

def email_recent(args: dict, **_kwargs) -> str:
    try:
        emails = _graph.get_recent(args.get("count", 10), _account(args))
        return _ok([_graph.summarize_email(e) for e in emails])
    except Exception as e:
        return _err(e)


def email_search(args: dict, **_kwargs) -> str:
    try:
        emails = _graph.search_emails(args["query"], args.get("count", 10), _account(args))
        return _ok([_graph.summarize_email(e) for e in emails])
    except Exception as e:
        return _err(e)


def email_from_sender(args: dict, **_kwargs) -> str:
    try:
        emails = _graph.get_from_sender(args["sender"], args.get("count", 10), _account(args))
        return _ok([_graph.summarize_email(e) for e in emails])
    except Exception as e:
        return _err(e)


def email_read(args: dict, **_kwargs) -> str:
    try:
        email = _graph.get_email_by_id(args["message_id"], _account(args))
        return _ok(_graph.summarize_email(email, include_body=True))
    except Exception as e:
        return _err(e)


# -------- email write --------

def email_send(args: dict, **_kwargs) -> str:
    try:
        _graph.send_email(
            to=args["to"],
            subject=args["subject"],
            body=args["body"],
            cc=args.get("cc"),
            reply_to=args.get("reply_to"),
            html=bool(args.get("html", False)),
            account_name=_account(args),
        )
        return _ok({"sent": True})
    except Exception as e:
        return _err(e)


def email_reply(args: dict, **_kwargs) -> str:
    try:
        _graph.reply_to_email(args["message_id"], args["body"], _account(args))
        return _ok({"replied": True})
    except Exception as e:
        return _err(e)


# -------- calendar --------

def calendar_today(args: dict, **_kwargs) -> str:
    try:
        events = _graph.get_today(_account(args))
        return _ok([_graph.summarize_event(e) for e in events])
    except Exception as e:
        return _err(e)


def calendar_week(args: dict, **_kwargs) -> str:
    try:
        events = _graph.get_week(_account(args))
        return _ok([_graph.summarize_event(e) for e in events])
    except Exception as e:
        return _err(e)


def calendar_range(args: dict, **_kwargs) -> str:
    try:
        events = _graph.get_events(args["start"], args["end"], _account(args))
        return _ok([_graph.summarize_event(e) for e in events])
    except Exception as e:
        return _err(e)


def calendar_cancel_event(args: dict, **_kwargs) -> str:
    try:
        _graph.cancel_event(args["event_id"], args.get("comment", "") or "", _account(args))
        return _ok({"cancelled": True})
    except Exception as e:
        return _err(e)
