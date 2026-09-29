"""Delete a Microsoft Graph event."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result
from ._write_helpers import client_user, required_id, segment


async def delete_event(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        event_id = required_id(args, "event_id")
        client, user = client_user(args)
        calendar_id = str(args.get("calendar_id") or "").strip()
        path = (
            f"users/{user}/calendars/{segment(calendar_id)}/events/{event_id}"
            if calendar_id else f"users/{user}/events/{event_id}"
        )
        return json.dumps(client.delete_json(path))
    except Exception as exc:
        return error_result(exc)
