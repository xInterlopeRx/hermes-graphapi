"""Fetch a single mailbox message."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result, graph_client_and_user, segment


async def get_message(args: dict[str, Any], **kwargs: Any) -> str:
    message_id = str(args.get("message_id") or "").strip()
    if not message_id:
        return json.dumps({"error": "message_id is required."})
    try:
        client, user = graph_client_and_user(args)
        path = f"users/{user}/messages/{segment(message_id)}"
        params = {"$select": args["select"]} if args.get("select") else None
        return json.dumps(client.get_json(path, params=params))
    except Exception as exc:
        return error_result(exc)
