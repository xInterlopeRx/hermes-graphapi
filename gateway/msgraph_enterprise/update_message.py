"""Update a Microsoft Graph message."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result
from ._write_helpers import client_user, required_id, required_payload


async def update_message(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        message_id = required_id(args, "message_id")
        message = required_payload(args, "message")
        client, user = client_user(args)
        path = f"users/{user}/messages/{message_id}"
        return json.dumps(client.patch_json(path, message))
    except Exception as exc:
        return error_result(exc)
