"""List contacts in a mailbox, optionally matching an email address."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result, graph_client_and_user, segment


async def list_contacts(args: dict[str, Any], **kwargs: Any) -> str:
    try:
        client, user = graph_client_and_user(args)
        folder_id = str(args.get("folder_id") or "").strip()
        path = (
            f"users/{user}/contactFolders/{segment(folder_id)}/contacts"
            if folder_id else f"users/{user}/contacts"
        )
        params: dict[str, Any] = {}
        filters = []
        if args.get("email"):
            email = str(args["email"]).replace("'", "''")
            filters.append(f"emailAddresses/any(a:a/address eq '{email}')")
        for arg, query in (("filter", "$filter"), ("select", "$select"), ("orderby", "$orderby"), ("top", "$top")):
            if args.get(arg) not in (None, ""):
                if arg == "filter":
                    filters.append(str(args[arg]))
                else:
                    params[query] = args[arg]
        if filters:
            params["$filter"] = " and ".join(f"({value})" for value in filters)
        if args.get("next_link"):
            path, params = str(args["next_link"]), {}
        payload = (
            client.get_paginated(path, params)
            if args.get("paginate", True)
            else client.get_json(path, params)
        )
        return json.dumps(payload)
    except Exception as exc:
        return error_result(exc)
