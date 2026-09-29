"""Compatibility import; implementation lives in gateway.msgraph_enterprise."""

from gateway.msgraph_enterprise.client_manager import (
    MicrosoftGraphClient,
    MicrosoftGraphClientError,
)

__all__ = ["MicrosoftGraphClient", "MicrosoftGraphClientError"]
