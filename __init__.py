"""office365-connector — Hermes Agent plugin entry point.

Hermes calls `register(ctx)` exactly once at startup. We register one tool per
exposed capability against the central tool registry.
"""

from . import schemas, tools

_TOOLSET = "office365"


def register(ctx) -> None:
    bindings = [
        (schemas.ACCOUNTS_LIST, tools.accounts_list),
        (schemas.ACCOUNTS_ADD, tools.accounts_add),
        (schemas.ACCOUNTS_REMOVE, tools.accounts_remove),
        (schemas.ACCOUNTS_SET_DEFAULT, tools.accounts_set_default),
        (schemas.ACCOUNTS_IMPORT_LEGACY, tools.accounts_import_legacy),
        (schemas.AUTH_LOGIN, tools.auth_login),
        (schemas.AUTH_STATUS, tools.auth_status),
        (schemas.EMAIL_RECENT, tools.email_recent),
        (schemas.EMAIL_SEARCH, tools.email_search),
        (schemas.EMAIL_FROM_SENDER, tools.email_from_sender),
        (schemas.EMAIL_READ, tools.email_read),
        (schemas.EMAIL_SEND, tools.email_send),
        (schemas.EMAIL_REPLY, tools.email_reply),
        (schemas.CALENDAR_TODAY, tools.calendar_today),
        (schemas.CALENDAR_WEEK, tools.calendar_week),
        (schemas.CALENDAR_RANGE, tools.calendar_range),
        (schemas.CALENDAR_CANCEL_EVENT, tools.calendar_cancel_event),
    ]

    for schema, handler in bindings:
        ctx.register_tool(
            name=schema["name"],
            toolset=_TOOLSET,
            schema=schema,
            handler=handler,
        )
