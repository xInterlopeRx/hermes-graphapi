"""Synchronous app-only Microsoft Graph client for gateway endpoint helpers."""

from __future__ import annotations

import time
import threading
from typing import Any
from urllib.parse import quote, urlparse

import requests

from .auth_handler import get_graph_credentials


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_TIMEOUT = (5.0, 30.0)
TOKEN_EXPIRY_SKEW = 60.0
_TOKEN_CACHE: dict[tuple[str, str, str, str], tuple[str, float]] = {}
_TOKEN_CACHE_LOCK = threading.RLock()


class MicrosoftGraphClientError(RuntimeError):
    """Safe-to-return Graph failure without credentials or response bodies."""


class MicrosoftGraphClient:
    """Acquire/caches an app-only token and makes bounded Graph requests."""

    def __init__(
        self,
        *,
        credentials: dict[str, str] | None = None,
        session: requests.Session | None = None,
        timeout: float | tuple[float, float] = DEFAULT_TIMEOUT,
        base_url: str = GRAPH_BASE_URL,
    ) -> None:
        self.credentials = credentials if credentials is not None else get_graph_credentials()
        if not self.credentials:
            raise MicrosoftGraphClientError(
                "Microsoft Graph is not configured; set MSGRAPH_TENANT_ID, "
                "MSGRAPH_CLIENT_ID, and MSGRAPH_CLIENT_SECRET."
            )
        self.session = session if session is not None else requests.Session()
        self.timeout = timeout
        self.base_url = base_url.rstrip("/")
        # These attributes remain useful to callers that inject a token in
        # isolated tests; normally token state is shared in _TOKEN_CACHE.
        self.token: str | None = None
        self._token_expires_at = 0.0
        self._token_from_shared_cache = False

    @classmethod
    def clear_token_cache(cls) -> None:
        """Clear process-wide access tokens (primarily for isolated tests)."""
        with _TOKEN_CACHE_LOCK:
            _TOKEN_CACHE.clear()

    def _token_cache_key(self) -> tuple[str, str, str, str]:
        return (
            self.credentials["authority_url"].rstrip("/"),
            self.credentials["tenant_id"],
            self.credentials["client_id"],
            self.credentials.get("scope", "https://graph.microsoft.com/.default"),
        )

    def _refresh_token(self) -> str:
        url = (
            f"{self.credentials['authority_url'].rstrip('/')}/"
            f"{quote(self.credentials['tenant_id'].strip('/'), safe='')}/oauth2/v2.0/token"
        )
        data = {
            "client_id": self.credentials["client_id"],
            "scope": self.credentials.get("scope", "https://graph.microsoft.com/.default"),
            "client_secret": self.credentials["client_secret"],
            "grant_type": "client_credentials",
        }
        try:
            response = self.session.post(url, data=data, timeout=self.timeout)
            response.raise_for_status()
            payload = response.json()
        except MicrosoftGraphClientError:
            raise
        except Exception as exc:
            status = getattr(getattr(exc, "response", None), "status_code", None)
            suffix = f" (HTTP {status})" if status is not None else ""
            raise MicrosoftGraphClientError(
                f"Microsoft Graph token request failed{suffix}."
            ) from None

        token = payload.get("access_token") if isinstance(payload, dict) else None
        if not isinstance(token, str) or not token:
            raise MicrosoftGraphClientError(
                "Microsoft Graph token response did not contain an access token."
            )
        try:
            lifetime = max(0, int(payload.get("expires_in", 3600)))
        except (TypeError, ValueError):
            lifetime = 3600
        expires_at = time.monotonic() + lifetime
        self.token = token
        self._token_expires_at = expires_at
        self._token_from_shared_cache = True
        with _TOKEN_CACHE_LOCK:
            _TOKEN_CACHE[self._token_cache_key()] = (token, expires_at)
        return token

    def _get_token(self) -> str:
        # Serialize lookup and refresh so concurrent requests in this process
        # don't all exchange credentials for the same token simultaneously.
        with _TOKEN_CACHE_LOCK:
            now = time.monotonic()
            key = self._token_cache_key()
            cached = _TOKEN_CACHE.get(key)
            if cached and cached[1] > now + TOKEN_EXPIRY_SKEW:
                self.token, self._token_expires_at = cached
                self._token_from_shared_cache = True
                return self.token
            # Preserve explicit instance injection used by callers/tests, but
            # do not let it override an expired shared cache entry.
            if (
                not cached
                and not self._token_from_shared_cache
                and self.token
                and self._token_expires_at > now + TOKEN_EXPIRY_SKEW
            ):
                return self.token
            return self._refresh_token()

    def _request_json(
        self,
        method: str,
        path: str,
        *,
        payload: dict[str, Any] | None = None,
    ) -> Any:
        url = self._resolve_url(path)
        request_headers = {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }
        try:
            response = getattr(self.session, method.lower())(
                url,
                json=payload,
                headers=request_headers,
                timeout=self.timeout,
            )
            response.raise_for_status()
            try:
                return response.json()
            except (ValueError, TypeError):
                # Graph uses 202/204 responses for successful send/delete.
                return {}
        except MicrosoftGraphClientError:
            raise
        except Exception as exc:
            status = getattr(getattr(exc, "response", None), "status_code", None)
            suffix = f" (HTTP {status})" if status is not None else ""
            raise MicrosoftGraphClientError(
                f"Microsoft Graph request failed{suffix}."
            ) from None

    def post_json(self, path: str, payload: dict[str, Any]) -> Any:
        return self._request_json("POST", path, payload=payload)

    def patch_json(self, path: str, payload: dict[str, Any]) -> Any:
        return self._request_json("PATCH", path, payload=payload)

    def delete_json(self, path: str) -> Any:
        return self._request_json("DELETE", path)

    def get_json(
        self,
        path: str,
        params: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
    ) -> Any:
        url = self._resolve_url(path)
        request_headers = {"Authorization": f"Bearer {self._get_token()}"}
        if headers:
            request_headers.update(headers)
        try:
            response = self.session.get(
                url,
                params=params,
                headers=request_headers,
                timeout=self.timeout,
            )
            response.raise_for_status()
            return response.json()
        except MicrosoftGraphClientError:
            raise
        except Exception as exc:
            status = getattr(getattr(exc, "response", None), "status_code", None)
            suffix = f" (HTTP {status})" if status is not None else ""
            raise MicrosoftGraphClientError(
                f"Microsoft Graph request failed{suffix}."
            ) from None

    def get_paginated(
        self,
        path: str,
        params: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
        *,
        max_pages: int = 100,
    ) -> dict[str, Any]:
        """Fetch and combine Graph collection pages, validating next links."""
        url = self._resolve_url(path)
        page_params = dict(params or {})
        values: list[Any] = []
        first_payload: dict[str, Any] | None = None
        for _ in range(max_pages):
            payload = self.get_json(url, params=page_params or None, headers=headers)
            if not isinstance(payload, dict):
                raise MicrosoftGraphClientError("Microsoft Graph collection response was invalid.")
            if first_payload is None:
                first_payload = {key: value for key, value in payload.items() if key not in ("value", "@odata.nextLink")}
            items = payload.get("value")
            if isinstance(items, list):
                values.extend(items)
            next_link = payload.get("@odata.nextLink")
            if not next_link:
                result = first_payload or {}
                result["value"] = values
                return result
            next_url = self._resolve_url(str(next_link))
            if urlparse(next_url).netloc.lower() != urlparse(self.base_url).netloc.lower():
                raise MicrosoftGraphClientError("Microsoft Graph returned an untrusted pagination URL.")
            url, page_params = next_url, {}
        raise MicrosoftGraphClientError("Microsoft Graph pagination exceeded the page limit.")

    def _resolve_url(self, path_or_url: str) -> str:
        if path_or_url.startswith("https://"):
            parsed = urlparse(path_or_url)
            if parsed.netloc.lower() != urlparse(self.base_url).netloc.lower():
                raise MicrosoftGraphClientError("Refusing a Graph URL outside the configured service.")
            return path_or_url
        if path_or_url.startswith("http://"):
            raise MicrosoftGraphClientError("Refusing an insecure Graph URL.")
        return f"{self.base_url}/{path_or_url.lstrip('/')}"
