"""Compatibility import; implementation lives in gateway.msgraph_enterprise."""

from gateway.msgraph_enterprise.auth_handler import get_graph_credentials

__all__ = ["get_graph_credentials"]
