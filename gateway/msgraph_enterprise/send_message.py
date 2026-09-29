"""Send an explicit Microsoft Graph message."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result
from ._write_helpers import client_user, required_payload


async def send_message(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        message = required_payload(args, "message")
        client, user = client_user(args)
        payload = {
            "message": message,
            "saveToSentItems": args.get("save_to_sent_items", True),
        }
        return json.dumps(client.post_json(f"users/{user}/sendMail", payload))
    except Exception as exc:
        return error_result(exc)
