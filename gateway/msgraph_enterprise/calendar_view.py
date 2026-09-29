"""List calendar events in a specified time window."""

from __future__ import annotations

import json
from typing import Any

from ._helpers import error_result, graph_client_and_user, segment


async def calendar_view(args: dict[str, Any], **kwargs: Any) -> str:
    start = args.get("start_date_time")
    end = args.get("end_date_time")
    if not start or not end:
        return json.dumps({"error": "start_date_time and end_date_time are required."})
    try:
        client, user = graph_client_and_user(args)
        calendar_id = str(args.get("calendar_id") or "").strip()
        path = (
            f"users/{user}/calendars/{segment(calendar_id)}/calendarView"
            if calendar_id else f"users/{user}/calendarView"
        )
        params: dict[str, Any] = {"startDateTime": start, "endDateTime": end}
        for arg, query in (("top", "$top"), ("select", "$select"), ("orderby", "$orderby"), ("filter", "$filter")):
            if args.get(arg) not in (None, ""):
                params[query] = args[arg]
        headers = {}
        if args.get("timezone"):
            headers["Prefer"] = f'outlook.timezone="{args["timezone"]}"'
        if args.get("next_link"):
            path, params = str(args["next_link"]), None
        payload = (
            client.get_paginated(path, params, headers)
            if args.get("paginate", True)
            else client.get_json(path, params, headers)
        )
        return json.dumps(payload)
    except Exception as exc:
        return error_result(exc)
