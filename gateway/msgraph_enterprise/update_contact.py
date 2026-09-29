"""Update a Microsoft Graph contact."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result
from ._write_helpers import client_user, required_id, required_payload


async def update_contact(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        contact_id = required_id(args, "contact_id")
        contact = required_payload(args, "contact")
        client, user = client_user(args)
        path = f"users/{user}/contacts/{contact_id}"
        return json.dumps(client.patch_json(path, contact))
    except Exception as exc:
        return error_result(exc)
