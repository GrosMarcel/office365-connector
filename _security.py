"""Security helpers for the office365-connector Hermes plugin.

Goals:
  - Prevent path traversal via untrusted account names / Graph IDs.
  - Prevent URL/host injection in OAuth + Graph requests.
  - Atomic, mode-0600 writes for any file containing secrets or tokens.
  - Scrub OAuth/Graph error bodies before they end up in logs.
  - Strip terminal escape sequences from anything we print.
  - Hard timeouts and response-size caps on every outbound HTTPS request.
"""

from __future__ import annotations

import contextlib
import json
import os
import re
import ssl
import tempfile
import urllib.error
import urllib.parse
import urllib.request

UUID_RE = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
ACCOUNT_NAME_RE = re.compile(r"^[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}$")
GRAPH_ID_RE = re.compile(r"^[A-Za-z0-9_=+\-/]{1,512}$")
EMAIL_RE = re.compile(r"^[^\s<>\"'\\/]+@[^\s<>\"'\\/]+\.[^\s<>\"'\\/]+$")

ALLOWED_HOSTS = frozenset({"graph.microsoft.com", "login.microsoftonline.com"})
MAX_RESPONSE_BYTES = 8 * 1024 * 1024
DEFAULT_TIMEOUT_S = 30


def assert_account_name(name: str) -> str:
    if not isinstance(name, str) or not ACCOUNT_NAME_RE.match(name):
        raise ValueError(
            'Invalid account name. Allowed: letters, digits, "-", "_", '
            "max 64 chars, must start with a letter or digit."
        )
    return name


def assert_uuid(value: str, label: str) -> str:
    if not isinstance(value, str) or not UUID_RE.match(value):
        raise ValueError(f"Invalid {label}: expected a UUID.")
    return value


def assert_graph_id(value: str, label: str) -> str:
    if not isinstance(value, str) or not GRAPH_ID_RE.match(value):
        raise ValueError(f"Invalid {label}: contains forbidden characters.")
    return value


def assert_email(value: str, label: str) -> str:
    if not isinstance(value, str) or len(value) > 320 or not EMAIL_RE.match(value):
        raise ValueError(f"Invalid {label}: not a valid email address.")
    return value


def assert_safe_child_path(base_dir: str, untrusted_name: str) -> str:
    assert_account_name(untrusted_name)
    resolved_base = os.path.realpath(base_dir)
    target = os.path.join(resolved_base, f"{untrusted_name}.json")
    resolved = os.path.realpath(target)
    if os.path.dirname(resolved) != resolved_base:
        raise ValueError("Path traversal detected in account name.")
    return target


def secure_write_file(file_path: str, data: str) -> None:
    parent = os.path.dirname(file_path)
    os.makedirs(parent, mode=0o700, exist_ok=True)
    with contextlib.suppress(OSError):
        os.chmod(parent, 0o700)

    fd, tmp = tempfile.mkstemp(prefix=os.path.basename(file_path) + ".", dir=parent)
    try:
        with os.fdopen(fd, "w") as f:
            f.write(data)
            with contextlib.suppress(OSError):
                os.fsync(f.fileno())
        os.chmod(tmp, 0o600)
        os.replace(tmp, file_path)
        os.chmod(file_path, 0o600)
    except Exception:
        with contextlib.suppress(OSError):
            os.unlink(tmp)
        raise


def secure_read_json(file_path: str):
    if not os.path.exists(file_path):
        return None
    st = os.stat(file_path)
    if (st.st_mode & 0o077) != 0:
        with contextlib.suppress(OSError):
            os.chmod(file_path, 0o600)
    with open(file_path, "r", encoding="utf-8") as f:
        raw = f.read()
    if not raw:
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"Failed to parse JSON file: {file_path}") from e


ANSI_ESC_RE = re.compile(r"\x1b\[[0-9;?]*[ -/]*[@-~]")
C0_C1_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")


def sanitize_for_terminal(value) -> str:
    if value is None:
        return ""
    return C0_C1_RE.sub("", ANSI_ESC_RE.sub("", str(value)))


TOKEN_FIELD_RE = re.compile(
    r'("?(?:access_token|refresh_token|id_token|device_code|client_secret|user_code)"?\s*[:=]\s*"?)[^"\s,}]+',
    re.IGNORECASE,
)


def scrub_secrets(value) -> str:
    if value is None:
        return ""
    return TOKEN_FIELD_RE.sub(r"\1[REDACTED]", str(value))


def _safe_error_message(prefix: str, status: int, parsed) -> str:
    if isinstance(parsed, dict):
        if "error_description" in parsed:
            msg = parsed["error_description"]
        elif "error" in parsed:
            err = parsed["error"]
            msg = err.get("message", str(err)) if isinstance(err, dict) else str(err)
        else:
            msg = f"HTTP {status}"
    else:
        msg = f"HTTP {status}"
    return f"{prefix} {scrub_secrets(msg)}"


def assert_allowed_host(hostname: str) -> None:
    if hostname not in ALLOWED_HOSTS:
        raise ValueError(f"Refusing to contact unexpected host: {hostname}")


def https_json_request(
    target_url: str,
    method: str = "GET",
    headers: dict | None = None,
    body=None,
    accept_empty: bool = False,
    timeout: int = DEFAULT_TIMEOUT_S,
):
    parsed = urllib.parse.urlparse(target_url)
    if parsed.scheme != "https":
        raise ValueError("Refusing non-HTTPS request")
    assert_allowed_host(parsed.hostname or "")

    req_headers = {"Accept": "application/json"}
    if headers:
        req_headers.update(headers)

    data = body.encode("utf-8") if isinstance(body, str) else body
    req = urllib.request.Request(target_url, data=data, headers=req_headers, method=method)
    ctx = ssl.create_default_context()

    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            raw = resp.read(MAX_RESPONSE_BYTES + 1)
            if len(raw) > MAX_RESPONSE_BYTES:
                raise RuntimeError("Response too large")
            text = raw.decode("utf-8") if raw else ""
            if text:
                try:
                    return json.loads(text)
                except json.JSONDecodeError:
                    return {"success": True, "statusCode": resp.status, "raw": text}
            return {"success": True, "statusCode": resp.status}
    except urllib.error.HTTPError as e:
        try:
            err_raw = e.read(MAX_RESPONSE_BYTES + 1)
        except Exception:
            err_raw = b""
        err_text = err_raw.decode("utf-8", errors="replace") if err_raw else ""
        parsed_body = None
        if err_text:
            try:
                parsed_body = json.loads(err_text)
            except json.JSONDecodeError:
                parsed_body = None
        raise RuntimeError(_safe_error_message(f"HTTP {e.code}:", e.code, parsed_body)) from None
    except urllib.error.URLError as e:
        raise RuntimeError(scrub_secrets(f"Request failed: {e.reason}")) from None
