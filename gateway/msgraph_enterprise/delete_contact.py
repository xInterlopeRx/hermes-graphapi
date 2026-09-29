"""Delete a Microsoft Graph contact."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result
from ._write_helpers import client_user, required_id


async def delete_contact(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        contact_id = required_id(args, "contact_id")
        client, user = client_user(args)
        return json.dumps(client.delete_json(f"users/{user}/contacts/{contact_id}"))
    except Exception as exc:
        return error_result(exc)
