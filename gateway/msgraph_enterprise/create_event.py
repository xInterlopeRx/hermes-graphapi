"""Create a Microsoft Graph event."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result
from ._write_helpers import client_user, required_payload, segment


async def create_event(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        event = required_payload(args, "event")
        client, user = client_user(args)
        calendar_id = str(args.get("calendar_id") or "").strip()
        path = (
            f"users/{user}/calendars/{segment(calendar_id)}/events"
            if calendar_id else f"users/{user}/events"
        )
        return json.dumps(client.post_json(path, event))
    except Exception as exc:
        return error_result(exc)
