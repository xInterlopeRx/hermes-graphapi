"""Shared endpoint argument validation and response formatting."""

from __future__ import annotations

import json
from typing import Any
from urllib.parse import quote

from agent.secret_scope import get_secret
from .client_manager import MicrosoftGraphClient, MicrosoftGraphClientError


def graph_client_and_user(args: dict[str, Any]) -> tuple[MicrosoftGraphClient, str]:
    client = MicrosoftGraphClient()
    user_id = str(args.get("user_id") or get_secret("MSGRAPH_USER_ID", "") or "").strip()
    if not user_id:
        raise MicrosoftGraphClientError(
            "user_id is required for app-only Microsoft Graph requests (or set MSGRAPH_USER_ID)."
        )
    return client, quote(user_id, safe="")


def segment(value: Any) -> str:
    return quote(str(value), safe="")


def error_result(exc: Exception) -> str:
    # Client errors are deliberately sanitized; never serialize credentials,
    # request bodies, token responses, or raw server payloads to a tool result.
    if isinstance(exc, MicrosoftGraphClientError):
        message = str(exc)
    else:
        message = "Microsoft Graph request failed."
    return json.dumps({"error": message})
