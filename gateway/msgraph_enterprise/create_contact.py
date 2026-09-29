"""Create a Microsoft Graph contact."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result
from ._write_helpers import client_user, required_payload, segment


async def create_contact(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        contact = required_payload(args, "contact")
        client, user = client_user(args)
        folder_id = str(args.get("folder_id") or "").strip()
        path = (
            f"users/{user}/contactFolders/{segment(folder_id)}/contacts"
            if folder_id else f"users/{user}/contacts"
        )
        return json.dumps(client.post_json(path, contact))
    except Exception as exc:
        return error_result(exc)
