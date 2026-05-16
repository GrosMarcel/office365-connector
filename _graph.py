"""Microsoft Graph email + calendar operations for the Hermes plugin."""

from __future__ import annotations

import datetime as _dt
import json
import re
import urllib.parse
from typing import Any

from ._auth import get_access_token
from ._security import (
    assert_email,
    assert_graph_id,
    https_json_request,
    sanitize_for_terminal,
)

GRAPH_BASE = "https://graph.microsoft.com"


def _graph_get(path_and_query: str, access_token: str) -> dict[str, Any]:
    return https_json_request(
        GRAPH_BASE + path_and_query,
        method="GET",
        headers={"Authorization": f"Bearer {access_token}"},
    )


def _graph_post(path_and_query: str, access_token: str, json_body: dict) -> dict[str, Any]:
    body = json.dumps(json_body)
    return https_json_request(
        GRAPH_BASE + path_and_query,
        method="POST",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        body=body,
        accept_empty=True,
    )


def _sanitize_search_query(q: str) -> str:
    if not isinstance(q, str):
        raise ValueError("Search query must be a string")
    if len(q) > 1000:
        raise ValueError("Search query too long")
    cleaned = re.sub(r'[\x00-\x1f"]', " ", q).strip()
    if not cleaned:
        raise ValueError("Empty search query")
    return cleaned


def _clamp_top(top, default: int = 10) -> int:
    try:
        val = int(top)
    except (TypeError, ValueError):
        val = default
    return max(1, min(100, val))


# -------- email read --------

def search_emails(query: str, top: int = 10, account_name: str | None = None) -> list[dict]:
    safe_q = _sanitize_search_query(query)
    safe_top = _clamp_top(top)
    token = get_access_token(account_name)
    encoded = urllib.parse.quote(safe_q, safe="")
    url = (
        f'/v1.0/me/messages?$search="{encoded}"'
        f"&$top={safe_top}&$orderby=receivedDateTime desc"
    )
    response = _graph_get(url, token)
    return response.get("value", []) if isinstance(response, dict) else []


def get_from_sender(sender: str, top: int = 10, account_name: str | None = None) -> list[dict]:
    safe_sender = _sanitize_search_query(sender)
    safe_top = _clamp_top(top)
    token = get_access_token(account_name)
    encoded = urllib.parse.quote(safe_sender, safe="")
    url = (
        f'/v1.0/me/messages?$search="from:{encoded}"'
        f"&$top={safe_top}&$orderby=receivedDateTime desc"
    )
    response = _graph_get(url, token)
    return response.get("value", []) if isinstance(response, dict) else []


def get_recent(top: int = 10, account_name: str | None = None) -> list[dict]:
    safe_top = _clamp_top(top)
    token = get_access_token(account_name)
    url = f"/v1.0/me/messages?$top={safe_top}&$orderby=receivedDateTime desc"
    response = _graph_get(url, token)
    return response.get("value", []) if isinstance(response, dict) else []


def get_email_by_id(email_id: str, account_name: str | None = None) -> dict:
    assert_graph_id(email_id, "emailId")
    token = get_access_token(account_name)
    url = f"/v1.0/me/messages/{urllib.parse.quote(email_id, safe='')}"
    return _graph_get(url, token)


def summarize_email(email: dict, include_body: bool = False) -> dict[str, Any]:
    """Return a compact, safe-to-display representation of a message."""
    from_part = (email.get("from") or {}).get("emailAddress") or {}
    summary: dict[str, Any] = {
        "id": email.get("id"),
        "subject": sanitize_for_terminal(email.get("subject") or "(no subject)"),
        "from": {
            "name": sanitize_for_terminal(from_part.get("name") or ""),
            "address": sanitize_for_terminal(from_part.get("address") or ""),
        },
        "receivedDateTime": email.get("receivedDateTime"),
        "isRead": email.get("isRead"),
        "preview": sanitize_for_terminal((email.get("bodyPreview") or "")[:300]),
    }
    if include_body:
        body = (email.get("body") or {}).get("content") or email.get("bodyPreview") or ""
        plain = re.sub(r"<[^>]*>", "", body).replace("&nbsp;", " ").strip()
        summary["body"] = sanitize_for_terminal(plain)
    return summary


# -------- email write --------

def _to_recipient_list(addresses, label: str) -> list[dict]:
    if isinstance(addresses, str):
        addresses = [addresses]
    if not addresses:
        raise ValueError(f"{label}: at least one recipient required")
    if len(addresses) > 500:
        raise ValueError(f"{label}: too many recipients (max 500)")
    return [{"emailAddress": {"address": assert_email(a, label)}} for a in addresses]


def send_email(
    to,
    subject: str,
    body: str,
    cc=None,
    reply_to: str | None = None,
    html: bool = False,
    account_name: str | None = None,
) -> dict:
    if not isinstance(subject, str) or len(subject) > 998:
        raise ValueError("Invalid subject")
    if not isinstance(body, str):
        raise ValueError("Invalid body")
    if len(body) > 10 * 1024 * 1024:
        raise ValueError("Body too large (>10 MB)")

    token = get_access_token(account_name)
    message: dict[str, Any] = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML" if html else "Text",
                "content": body,
            },
            "toRecipients": _to_recipient_list(to, "to"),
        }
    }
    if cc:
        message["message"]["ccRecipients"] = _to_recipient_list(cc, "cc")
    if reply_to:
        message["message"]["replyTo"] = [
            {"emailAddress": {"address": assert_email(reply_to, "replyTo")}}
        ]
    return _graph_post("/v1.0/me/sendMail", token, message)


def reply_to_email(message_id: str, body: str, account_name: str | None = None) -> dict:
    assert_graph_id(message_id, "messageId")
    if not isinstance(body, str):
        raise ValueError("Invalid body")
    if len(body) > 10 * 1024 * 1024:
        raise ValueError("Body too large (>10 MB)")

    token = get_access_token(account_name)
    url = f"/v1.0/me/messages/{urllib.parse.quote(message_id, safe='')}/reply"
    return _graph_post(url, token, {"comment": body})


# -------- calendar --------

def _parse_when(value) -> _dt.datetime:
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, str):
        try:
            parsed = _dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError as e:
            raise ValueError(f"Invalid datetime: {value!r}") from e
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=_dt.timezone.utc)
        return parsed
    raise ValueError("Datetime must be ISO 8601 string")


def get_events(
    start: _dt.datetime | str,
    end: _dt.datetime | str,
    account_name: str | None = None,
) -> list[dict]:
    start_dt = _parse_when(start)
    end_dt = _parse_when(end)
    token = get_access_token(account_name)
    start_iso = start_dt.astimezone(_dt.timezone.utc).isoformat().replace("+00:00", "Z")
    end_iso = end_dt.astimezone(_dt.timezone.utc).isoformat().replace("+00:00", "Z")
    url = (
        f"/v1.0/me/calendarview"
        f"?startDateTime={urllib.parse.quote(start_iso, safe='')}"
        f"&endDateTime={urllib.parse.quote(end_iso, safe='')}"
        f"&$orderby=start/dateTime&$top=50"
    )
    response = _graph_get(url, token)
    return response.get("value", []) if isinstance(response, dict) else []


def get_today(account_name: str | None = None) -> list[dict]:
    now = _dt.datetime.now(_dt.timezone.utc)
    start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + _dt.timedelta(days=1) - _dt.timedelta(seconds=1)
    return get_events(start, end, account_name)


def get_week(account_name: str | None = None) -> list[dict]:
    now = _dt.datetime.now(_dt.timezone.utc)
    start = (now - _dt.timedelta(days=now.weekday())).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    end = start + _dt.timedelta(days=7)
    return get_events(start, end, account_name)


def summarize_event(event: dict) -> dict[str, Any]:
    start = (event.get("start") or {}).get("dateTime")
    end = (event.get("end") or {}).get("dateTime")
    organizer = (event.get("organizer") or {}).get("emailAddress") or {}
    organizer_addr = organizer.get("address")

    attendees = []
    for a in event.get("attendees") or []:
        ea = a.get("emailAddress") or {}
        if ea.get("address") != organizer_addr:
            attendees.append(
                {
                    "name": sanitize_for_terminal(ea.get("name") or ""),
                    "address": sanitize_for_terminal(ea.get("address") or ""),
                }
            )

    return {
        "id": event.get("id"),
        "subject": sanitize_for_terminal(event.get("subject") or "(no subject)"),
        "isAllDay": bool(event.get("isAllDay")),
        "start": start,
        "end": end,
        "location": sanitize_for_terminal(
            ((event.get("location") or {}).get("displayName") or "")
        ),
        "organizer": {
            "name": sanitize_for_terminal(organizer.get("name") or ""),
            "address": sanitize_for_terminal(organizer_addr or ""),
        },
        "attendees": attendees[:20],
    }


def cancel_event(event_id: str, comment: str = "", account_name: str | None = None) -> dict:
    assert_graph_id(event_id, "eventId")
    if not isinstance(comment, str) or len(comment) > 4000:
        raise ValueError("Invalid comment (string, max 4000 chars)")
    token = get_access_token(account_name)
    url = f"/v1.0/me/events/{urllib.parse.quote(event_id, safe='')}/cancel"
    return _graph_post(url, token, {"comment": comment})
