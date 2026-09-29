"""Test shim for the Hermes secret-scope accessor."""


def get_secret(key: str, default=None):
    """Return defaults so focused tests never need Hermes or real secrets."""
    return default
