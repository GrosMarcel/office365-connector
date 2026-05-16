"""OpenAI function-calling schemas for the office365-connector Hermes plugin.

Every tool accepts an optional `account` parameter to select among configured
identities; omitted means "use the default account".
"""

_ACCOUNT_PROP = {
    "account": {
        "type": "string",
        "description": "Optional account name. Omit to use the default account.",
    }
}


ACCOUNTS_LIST = {
    "name": "office365_accounts_list",
    "description": (
        "List all configured Office 365 / Outlook accounts and indicate which is the default."
    ),
    "parameters": {"type": "object", "properties": {}, "required": []},
}

ACCOUNTS_ADD = {
    "name": "office365_accounts_add",
    "description": (
        "Register a new Office 365 account. The client secret MUST be provided via "
        "the AZURE_CLIENT_SECRET environment variable; it is never accepted as an argument "
        "to avoid leaking through tool-call logs."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "name": {
                "type": "string",
                "description": 'Short account label (e.g. "work", "personal"). Letters/digits/_/-, max 64 chars.',
            },
            "tenant_id": {
                "type": "string",
                "description": "Azure AD tenant ID (UUID).",
            },
            "client_id": {
                "type": "string",
                "description": "Azure App Registration client ID (UUID).",
            },
            "email": {
                "type": "string",
                "description": "Optional email address associated with the account.",
            },
            "description": {
                "type": "string",
                "description": "Optional human-readable description (max 256 chars).",
            },
        },
        "required": ["name", "tenant_id", "client_id"],
    },
}

ACCOUNTS_REMOVE = {
    "name": "office365_accounts_remove",
    "description": "Remove a configured Office 365 account and delete its stored tokens.",
    "parameters": {
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "Account name to remove."},
        },
        "required": ["name"],
    },
}

ACCOUNTS_SET_DEFAULT = {
    "name": "office365_accounts_set_default",
    "description": "Set the default Office 365 account used when no explicit account is given.",
    "parameters": {
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "Account name to mark as default."},
        },
        "required": ["name"],
    },
}

ACCOUNTS_IMPORT_LEGACY = {
    "name": "office365_accounts_import_legacy",
    "description": (
        "Migrate a v1 single-account token file (~/.hermes/auth/microsoft-graph.json) "
        "into the multi-account layout. Requires AZURE_TENANT_ID / AZURE_CLIENT_ID / "
        "AZURE_CLIENT_SECRET in the environment."
    ),
    "parameters": {"type": "object", "properties": {}, "required": []},
}

AUTH_LOGIN = {
    "name": "office365_auth_login",
    "description": (
        "Start the OAuth 2.0 device-code authentication flow. Returns IMMEDIATELY "
        "(non-blocking) with a verification URL and user code that the human must "
        "enter in a browser. After surfacing the prompt to the user, call "
        "office365_auth_login_poll every ~5 seconds (or whatever poll_interval_s "
        "the response indicates) until status is 'authenticated', 'expired', or "
        "'declined'. This split avoids LLM tool-call timeouts that would otherwise "
        "kill a multi-minute blocking flow."
    ),
    "parameters": {
        "type": "object",
        "properties": {**_ACCOUNT_PROP},
        "required": [],
    },
}

AUTH_LOGIN_POLL = {
    "name": "office365_auth_login_poll",
    "description": (
        "Poll the token endpoint once to check whether the human has completed "
        "the in-flight device-code flow. Call this after office365_auth_login, "
        "repeatedly, at the interval reported by that call. Returns a status of "
        "'authenticated', 'pending', 'expired', 'declined', or 'no_pending_flow'."
    ),
    "parameters": {
        "type": "object",
        "properties": {**_ACCOUNT_PROP},
        "required": [],
    },
}

AUTH_STATUS = {
    "name": "office365_auth_status",
    "description": (
        "Report whether the Office 365 account is currently authenticated, when "
        "the access token expires, and whether a device-code flow is currently "
        "in progress."
    ),
    "parameters": {
        "type": "object",
        "properties": {**_ACCOUNT_PROP},
        "required": [],
    },
}

AUTH_DIAG = {
    "name": "office365_auth_diag",
    "description": (
        "Diagnose outbound network reachability to login.microsoftonline.com "
        "and graph.microsoft.com from the current runtime. Use this when "
        "auth_login appears to hang or fail — it distinguishes DNS, TLS, and "
        "HTTP-layer failures, which is the first thing to check inside a "
        "container."
    ),
    "parameters": {"type": "object", "properties": {}, "required": []},
}

EMAIL_RECENT = {
    "name": "office365_email_recent",
    "description": "List the most recent messages in the user's inbox, newest first.",
    "parameters": {
        "type": "object",
        "properties": {
            "count": {
                "type": "integer",
                "description": "Number of messages to return (1-100).",
                "default": 10,
            },
            **_ACCOUNT_PROP,
        },
        "required": [],
    },
}

EMAIL_SEARCH = {
    "name": "office365_email_search",
    "description": "Full-text search the user's mailbox via Microsoft Graph $search.",
    "parameters": {
        "type": "object",
        "properties": {
            "query": {"type": "string", "description": "Search query."},
            "count": {
                "type": "integer",
                "description": "Maximum results to return (1-100).",
                "default": 10,
            },
            **_ACCOUNT_PROP,
        },
        "required": ["query"],
    },
}

EMAIL_FROM_SENDER = {
    "name": "office365_email_from_sender",
    "description": "List messages from a specific sender address or display name.",
    "parameters": {
        "type": "object",
        "properties": {
            "sender": {"type": "string", "description": "Sender email or name."},
            "count": {
                "type": "integer",
                "description": "Maximum results (1-100).",
                "default": 10,
            },
            **_ACCOUNT_PROP,
        },
        "required": ["sender"],
    },
}

EMAIL_READ = {
    "name": "office365_email_read",
    "description": "Fetch the full body of a specific message by its Graph ID.",
    "parameters": {
        "type": "object",
        "properties": {
            "message_id": {"type": "string", "description": "Graph message ID."},
            **_ACCOUNT_PROP,
        },
        "required": ["message_id"],
    },
}

EMAIL_SEND = {
    "name": "office365_email_send",
    "description": "Send an email from the user's Office 365 mailbox.",
    "parameters": {
        "type": "object",
        "properties": {
            "to": {
                "anyOf": [
                    {"type": "string"},
                    {"type": "array", "items": {"type": "string"}},
                ],
                "description": "Single recipient or list of recipients.",
            },
            "subject": {"type": "string", "description": "Email subject (max 998 chars)."},
            "body": {"type": "string", "description": "Email body."},
            "cc": {
                "anyOf": [
                    {"type": "string"},
                    {"type": "array", "items": {"type": "string"}},
                ],
                "description": "Optional CC recipients.",
            },
            "reply_to": {"type": "string", "description": "Optional Reply-To address."},
            "html": {
                "type": "boolean",
                "description": "If true, the body is treated as HTML; otherwise plain text.",
                "default": False,
            },
            **_ACCOUNT_PROP,
        },
        "required": ["to", "subject", "body"],
    },
}

EMAIL_REPLY = {
    "name": "office365_email_reply",
    "description": "Reply inline to an existing message identified by its Graph ID.",
    "parameters": {
        "type": "object",
        "properties": {
            "message_id": {"type": "string", "description": "Graph message ID to reply to."},
            "body": {"type": "string", "description": "Reply body."},
            **_ACCOUNT_PROP,
        },
        "required": ["message_id", "body"],
    },
}

CALENDAR_TODAY = {
    "name": "office365_calendar_today",
    "description": "List today's calendar events for the user.",
    "parameters": {
        "type": "object",
        "properties": {**_ACCOUNT_PROP},
        "required": [],
    },
}

CALENDAR_WEEK = {
    "name": "office365_calendar_week",
    "description": "List this week's calendar events for the user.",
    "parameters": {
        "type": "object",
        "properties": {**_ACCOUNT_PROP},
        "required": [],
    },
}

CALENDAR_RANGE = {
    "name": "office365_calendar_range",
    "description": "List calendar events between an arbitrary start and end datetime.",
    "parameters": {
        "type": "object",
        "properties": {
            "start": {
                "type": "string",
                "description": "ISO 8601 datetime (e.g. 2026-05-16T00:00:00Z).",
            },
            "end": {
                "type": "string",
                "description": "ISO 8601 datetime, must be after start.",
            },
            **_ACCOUNT_PROP,
        },
        "required": ["start", "end"],
    },
}

CALENDAR_CANCEL_EVENT = {
    "name": "office365_calendar_cancel_event",
    "description": "Cancel an existing calendar event, sending an optional comment to attendees.",
    "parameters": {
        "type": "object",
        "properties": {
            "event_id": {"type": "string", "description": "Graph event ID."},
            "comment": {
                "type": "string",
                "description": "Optional comment to attendees (max 4000 chars).",
                "default": "",
            },
            **_ACCOUNT_PROP,
        },
        "required": ["event_id"],
    },
}


ALL = [
    ACCOUNTS_LIST,
    ACCOUNTS_ADD,
    ACCOUNTS_REMOVE,
    ACCOUNTS_SET_DEFAULT,
    ACCOUNTS_IMPORT_LEGACY,
    AUTH_LOGIN,
    AUTH_LOGIN_POLL,
    AUTH_STATUS,
    AUTH_DIAG,
    EMAIL_RECENT,
    EMAIL_SEARCH,
    EMAIL_FROM_SENDER,
    EMAIL_READ,
    EMAIL_SEND,
    EMAIL_REPLY,
    CALENDAR_TODAY,
    CALENDAR_WEEK,
    CALENDAR_RANGE,
    CALENDAR_CANCEL_EVENT,
]
