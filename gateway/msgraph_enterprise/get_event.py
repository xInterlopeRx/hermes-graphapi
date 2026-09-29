"""Fetch a single Microsoft Graph event."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result, graph_client_and_user, segment


async def get_event(args: dict[str, Any], **kwargs: Any) -> str:
    event_id = str(args.get("event_id") or "").strip()
    if not event_id:
        return json.dumps({"error": "event_id is required."})
    try:
        client, user = graph_client_and_user(args)
        path = f"users/{user}/events/{segment(event_id)}"
        params = {"$select": args["select"]} if args.get("select") else None
        return json.dumps(client.get_json(path, params=params))
    except Exception as exc:
        return error_result(exc)
