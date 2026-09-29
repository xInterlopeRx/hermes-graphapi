"""Create a draft Microsoft Graph message."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result
from ._write_helpers import client_user, required_payload, segment


async def create_draft(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        message = required_payload(args, "message")
        client, user = client_user(args)
        folder_id = str(args.get("folder_id") or "").strip()
        path = (
            f"users/{user}/mailFolders/{segment(folder_id)}/messages"
            if folder_id else f"users/{user}/messages"
        )
        return json.dumps(client.post_json(path, message))
    except Exception as exc:
        return error_result(exc)
