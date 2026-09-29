"""Shared validation and request helpers for explicit Graph write gates."""

from __future__ import annotations

from typing import Any

from ._helpers import graph_client_and_user, segment
from .client_manager import MicrosoftGraphClient, MicrosoftGraphClientError


def required_id(args: dict[str, Any], name: str) -> str:
    value = str(args.get(name) or "").strip()
    if not value:
        raise MicrosoftGraphClientError(f"{name} is required.")
    return segment(value)


def required_payload(args: dict[str, Any], name: str) -> dict[str, Any]:
    value = args.get(name)
    if not isinstance(value, dict) or not value:
        raise MicrosoftGraphClientError(f"{name} must be a non-empty JSON object.")
    return value


def client_user(args: dict[str, Any]) -> tuple[MicrosoftGraphClient, str]:
    return graph_client_and_user(args)
