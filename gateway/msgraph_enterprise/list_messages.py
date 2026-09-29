"""List messages in a mailbox or folder."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result, graph_client_and_user, segment


async def list_messages(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        client, user = graph_client_and_user(args)
        folder_id = str(args.get("folder_id") or "").strip()
        path = (
            f"users/{user}/mailFolders/{segment(folder_id)}/messages"
            if folder_id else f"users/{user}/messages"
        )
        params: dict[str, Any] = {}
        for arg, query in (("filter", "$filter"), ("search", "$search"), ("select", "$select"), ("orderby", "$orderby"), ("top", "$top")):
            if args.get(arg) not in (None, ""):
                params[query] = args[arg]
        headers = {"ConsistencyLevel": "eventual"} if args.get("search") else None
        if args.get("next_link"):
            path, params = str(args["next_link"]), {}
        payload = (
            client.get_paginated(path, params, headers)
            if args.get("paginate", True)
            else client.get_json(path, params, headers)
        )
        return json.dumps(payload)
    except Exception as exc:
        return error_result(exc)
