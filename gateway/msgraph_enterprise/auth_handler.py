"""App-only Microsoft Graph credential loading for gateway callers."""

from __future__ import annotations

import re
from typing import TypedDict
from urllib.parse import urlsplit

from agent.secret_scope import get_secret


class GraphCredentials(TypedDict):
    tenant_id: str
    client_id: str
    client_secret: str
    scope: str
    authority_url: str


DEFAULT_SCOPE = "https://graph.microsoft.com/.default"
DEFAULT_AUTHORITY_URL = "https://login.microsoftonline.com"
_GUID_RE = re.compile(
    r"^[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$"
)


def _is_resource_identifier(value: str) -> bool:
    """Check for a GUID or an absolute resource URI without query/fragment."""
    if _GUID_RE.fullmatch(value):
        return True
    try:
        parsed = urlsplit(value)
        return bool(
            parsed.scheme
            and re.fullmatch(r"[A-Za-z][A-Za-z0-9+.-]*", parsed.scheme)
            and parsed.netloc
            and parsed.hostname
            and not parsed.username
            and not parsed.password
            and not parsed.query
            and not parsed.fragment
        )
    except ValueError:
        return False


def _normalize_scope(scope: str) -> str:
    """Normalize a client-credentials resource into its v2 ``/.default`` scope."""
    if not isinstance(scope, str):
        raise ValueError("MSGRAPH_SCOPE must be a resource identifier or a /.default scope.")
    scope = scope.strip()
    if not scope:
        return DEFAULT_SCOPE
    if (
        "/.default/" in scope.lower()
        or any(character.isspace() or ord(character) < 32 for character in scope)
    ):
        raise ValueError("MSGRAPH_SCOPE must be a resource identifier or a /.default scope.")

    if scope.endswith("/.default"):
        resource = scope[: -len("/.default")]
        if _is_resource_identifier(resource):
            return scope
    elif _GUID_RE.fullmatch(scope):
        return DEFAULT_SCOPE
    elif _is_resource_identifier(scope):
        return f"{scope.rstrip('/')}/.default"

    raise ValueError("MSGRAPH_SCOPE must be a resource identifier or a /.default scope.")


def get_graph_credentials() -> GraphCredentials | None:
    """Return normalized credentials, or ``None`` when not configured.

    ``get_secret`` preserves the gateway's profile isolation and fails closed
    when called outside a profile scope in multiplex mode.
    """
    tenant_id = (get_secret("MSGRAPH_TENANT_ID", "") or "").strip()
    client_id = (get_secret("MSGRAPH_CLIENT_ID", "") or "").strip()
    client_secret = (get_secret("MSGRAPH_CLIENT_SECRET", "") or "").strip()
    if not tenant_id or not client_id or not client_secret:
        return None

    configured_scope = get_secret("MSGRAPH_SCOPE", DEFAULT_SCOPE) or DEFAULT_SCOPE
    scope = _normalize_scope(configured_scope)
    authority_url = (
        get_secret("MSGRAPH_AUTHORITY_URL", DEFAULT_AUTHORITY_URL)
        or DEFAULT_AUTHORITY_URL
    ).strip()
    return {
        "tenant_id": tenant_id,
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": scope,
        "authority_url": authority_url.rstrip("/"),
    }
