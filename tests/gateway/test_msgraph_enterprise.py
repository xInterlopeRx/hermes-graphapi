"""Behavior tests for the app-only Microsoft Graph gateway package."""

from __future__ import annotations

import asyncio
import json
from urllib.parse import parse_qs, urlparse

import pytest
import requests

from gateway.msgraph_enterprise import auth_handler, calendar_view, get_event
from gateway.msgraph_enterprise import get_message, list_contacts, list_messages
from gateway.msgraph_enterprise.client_manager import (
    MicrosoftGraphClient,
    MicrosoftGraphClientError,
)


CREDS = {
    "tenant_id": "tenant-id",
    "client_id": "client-id",
    "client_secret": "never-print-this-secret",
    "scope": "https://graph.microsoft.com/.default",
    "authority_url": "https://login.microsoftonline.com",
}


@pytest.fixture(autouse=True)
def clear_graph_token_cache():
    MicrosoftGraphClient.clear_token_cache()
    yield
    MicrosoftGraphClient.clear_token_cache()


class FakeResponse:
    def __init__(self, payload, status_code=200):
        self.payload = payload
        self.status_code = status_code

    def raise_for_status(self):
        if self.status_code >= 400:
            response = requests.Response()
            response.status_code = self.status_code
            raise requests.HTTPError(response=response)

    def json(self):
        return self.payload


class FakeSession:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def post(self, url, **kwargs):
        self.calls.append(("POST", url, kwargs))
        return self.responses.pop(0)

    def patch(self, url, **kwargs):
        self.calls.append(("PATCH", url, kwargs))
        return self.responses.pop(0)

    def delete(self, url, **kwargs):
        self.calls.append(("DELETE", url, kwargs))
        return self.responses.pop(0)

    def get(self, url, **kwargs):
        self.calls.append(("GET", url, kwargs))
        return self.responses.pop(0)


def test_credentials_are_normalized_and_optional(monkeypatch):
    values = {
        "MSGRAPH_TENANT_ID": " tenant ",
        "MSGRAPH_CLIENT_ID": " client ",
        "MSGRAPH_CLIENT_SECRET": " secret ",
    }
    monkeypatch.setattr(auth_handler, "get_secret", lambda key, default=None: values.get(key, default))
    credentials = auth_handler.get_graph_credentials()
    assert credentials["tenant_id"] == "tenant"
    assert credentials["scope"] == "https://graph.microsoft.com/.default"
    values.clear()
    assert auth_handler.get_graph_credentials() is None


@pytest.mark.parametrize(
    ("configured_scope", "expected_scope"),
    [
        ("00000000-0000-0000-0000-000000000000", auth_handler.DEFAULT_SCOPE),
        ("api://graph-resource", "api://graph-resource/.default"),
        ("https://resource.example", "https://resource.example/.default"),
        ("https://resource.example/.default", "https://resource.example/.default"),
        ("https://graph.microsoft.com/.default", "https://graph.microsoft.com/.default"),
    ],
)
def test_credentials_normalize_resource_scope(configured_scope, expected_scope, monkeypatch):
    values = {
        "MSGRAPH_TENANT_ID": "tenant",
        "MSGRAPH_CLIENT_ID": "client",
        "MSGRAPH_CLIENT_SECRET": "secret",
        "MSGRAPH_SCOPE": configured_scope,
    }
    monkeypatch.setattr(auth_handler, "get_secret", lambda key, default=None: values.get(key, default))

    assert auth_handler.get_graph_credentials()["scope"] == expected_scope


@pytest.mark.parametrize("configured_scope", ["not a scope", "https://graph.microsoft.com/.default/extra", "https://graph.microsoft.com?resource=other"])
def test_credentials_reject_malformed_scope_without_echoing_value(configured_scope, monkeypatch):
    values = {
        "MSGRAPH_TENANT_ID": "tenant",
        "MSGRAPH_CLIENT_ID": "client",
        "MSGRAPH_CLIENT_SECRET": "secret",
        "MSGRAPH_SCOPE": configured_scope,
    }
    monkeypatch.setattr(auth_handler, "get_secret", lambda key, default=None: values.get(key, default))

    with pytest.raises(ValueError) as exc:
        auth_handler.get_graph_credentials()
    assert str(exc.value) == "MSGRAPH_SCOPE must be a resource identifier or a /.default scope."
    assert configured_scope not in str(exc.value)


def test_empty_scope_uses_default_scope(monkeypatch):
    values = {
        "MSGRAPH_TENANT_ID": "tenant",
        "MSGRAPH_CLIENT_ID": "client",
        "MSGRAPH_CLIENT_SECRET": "secret",
        "MSGRAPH_SCOPE": "  ",
    }
    monkeypatch.setattr(auth_handler, "get_secret", lambda key, default=None: values.get(key, default))

    assert auth_handler.get_graph_credentials()["scope"] == auth_handler.DEFAULT_SCOPE


def test_token_request_uses_app_only_form_timeout_and_reuses_token():
    session = FakeSession([
        FakeResponse({"access_token": "mock-token", "expires_in": 3600}),
        FakeResponse({"ok": True}),
        FakeResponse({"ok": True}),
    ])
    client = MicrosoftGraphClient(credentials=CREDS, session=session)

    assert client.get_json("users/alice@example.com/messages") == {"ok": True}
    assert client.get_json("users/alice@example.com/messages") == {"ok": True}
    token_call, first_get, second_get = session.calls
    assert token_call[0] == "POST"
    assert token_call[1] == "https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token"
    assert token_call[2]["data"] == {
        "client_id": "client-id",
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CREDS["client_secret"],
        "grant_type": "client_credentials",
    }
    assert token_call[2]["timeout"] == (5.0, 30.0)
    assert first_get[2]["headers"]["Authorization"] == "Bearer mock-token"
    assert len([call for call in session.calls if call[0] == "POST"]) == 1


def test_token_cache_is_shared_between_clients_with_same_identity():
    first_session = FakeSession([FakeResponse({"access_token": "cached-token", "expires_in": 3600})])
    second_session = FakeSession([])
    first = MicrosoftGraphClient(credentials=CREDS, session=first_session)
    second = MicrosoftGraphClient(credentials=CREDS, session=second_session)

    assert first._get_token() == "cached-token"
    assert second._get_token() == "cached-token"
    assert len(first_session.calls) == 1
    assert second_session.calls == []


def test_token_cache_refreshes_inside_expiry_skew(monkeypatch):
    now = [100.0]
    monkeypatch.setattr("gateway.msgraph_enterprise.client_manager.time.monotonic", lambda: now[0])
    session = FakeSession([
        FakeResponse({"access_token": "first-token", "expires_in": 90}),
        FakeResponse({"access_token": "second-token", "expires_in": 3600}),
    ])
    client = MicrosoftGraphClient(credentials=CREDS, session=session)

    assert client._get_token() == "first-token"
    now[0] = 131.0  # expiry is within the 60-second refresh skew
    assert client._get_token() == "second-token"
    assert [call[0] for call in session.calls] == ["POST", "POST"]


@pytest.mark.parametrize(
    "changed_field,changed_value",
    [
        ("authority_url", "https://login.example"),
        ("tenant_id", "other-tenant"),
        ("client_id", "other-client"),
        ("scope", "https://other.example/.default"),
    ],
)
def test_token_cache_separates_authority_tenant_client_and_scope(changed_field, changed_value):
    first_session = FakeSession([FakeResponse({"access_token": "first-token", "expires_in": 3600})])
    second_session = FakeSession([FakeResponse({"access_token": "second-token", "expires_in": 3600})])
    first = MicrosoftGraphClient(credentials=CREDS, session=first_session)
    changed = {**CREDS, changed_field: changed_value}
    second = MicrosoftGraphClient(credentials=changed, session=second_session)

    assert first._get_token() == "first-token"
    assert second._get_token() == "second-token"
    assert len(first_session.calls) == len(second_session.calls) == 1


@pytest.mark.parametrize(
    ("method", "path", "payload"),
    [
        ("post_json", "users/alice/messages", {"subject": "new"}),
        ("patch_json", "users/alice/messages/m1", {"isRead": True}),
        ("delete_json", "users/alice/messages/m1", None),
    ],
)
def test_write_methods_use_mocked_session_with_bearer_timeout_and_payload(method, path, payload):
    session = FakeSession([FakeResponse({"ok": True})])
    client = MicrosoftGraphClient(credentials=CREDS, session=session)
    client.token = "injected-test-token"
    client._token_expires_at = 9_999_999_999

    result = getattr(client, method)(path, payload) if payload is not None else client.delete_json(path)

    assert result == {"ok": True}
    verb, url, kwargs = session.calls[0]
    assert verb == {"post_json": "POST", "patch_json": "PATCH", "delete_json": "DELETE"}[method]
    assert url.endswith(path)
    assert kwargs["headers"]["Authorization"] == "Bearer injected-test-token"
    assert kwargs["timeout"] == (5.0, 30.0)
    assert kwargs["json"] == payload


def test_token_request_uses_graph_default_for_bare_guid_scope(monkeypatch):
    values = {
        "MSGRAPH_TENANT_ID": "tenant",
        "MSGRAPH_CLIENT_ID": "client",
        "MSGRAPH_CLIENT_SECRET": "secret",
        "MSGRAPH_SCOPE": "00000000-0000-0000-0000-000000000000",
    }
    monkeypatch.setattr(auth_handler, "get_secret", lambda key, default=None: values.get(key, default))
    session = FakeSession([FakeResponse({"access_token": "mock-token", "expires_in": 3600})])
    client = MicrosoftGraphClient(session=session)

    assert client._get_token() == "mock-token"
    assert session.calls[0][2]["data"]["scope"] == auth_handler.DEFAULT_SCOPE


def test_http_failures_are_sanitized_and_have_timeout():
    session = FakeSession([FakeResponse({"error": "sensitive response"}, 401)])
    client = MicrosoftGraphClient(credentials=CREDS, session=session)
    with pytest.raises(MicrosoftGraphClientError) as exc:
        client.get_json("users/alice/messages")
    assert "HTTP 401" in str(exc.value)
    assert CREDS["client_secret"] not in str(exc.value)
    assert "sensitive response" not in str(exc.value)
    assert session.calls[0][2]["timeout"] == (5.0, 30.0)


def test_token_failure_does_not_expose_secret_or_response_body():
    session = FakeSession([FakeResponse({"error": "invalid_client", "error_description": CREDS["client_secret"]}, 401)])
    client = MicrosoftGraphClient(credentials=CREDS, session=session)
    with pytest.raises(MicrosoftGraphClientError) as exc:
        client.get_json("users/alice/messages")
    assert "HTTP 401" in str(exc.value)
    assert CREDS["client_secret"] not in str(exc.value)
    assert "invalid_client" not in str(exc.value)


def test_pagination_combines_pages_and_does_not_reapply_initial_params():
    session = FakeSession([
        FakeResponse({"value": [{"id": "1"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/a/messages?$skiptoken=next"}),
        FakeResponse({"value": [{"id": "2"}]}),
    ])
    client = MicrosoftGraphClient(credentials=CREDS, session=session)
    client.token = "mock-token"
    client._token_expires_at = 9_999_999_999
    result = client.get_paginated("users/a/messages", {"$top": 1})
    assert result == {"value": [{"id": "1"}, {"id": "2"}]}
    assert session.calls[0][2]["params"] == {"$top": 1}
    assert session.calls[1][2]["params"] is None
    assert session.calls[1][1].endswith("$skiptoken=next")


def test_pagination_rejects_foreign_next_link():
    session = FakeSession([
        FakeResponse({"value": [], "@odata.nextLink": "https://attacker.example/steal"}),
    ])
    client = MicrosoftGraphClient(credentials=CREDS, session=session)
    client.token = "mock-token"
    client._token_expires_at = 9_999_999_999
    with pytest.raises(MicrosoftGraphClientError, match="outside the configured service"):
        client.get_paginated("users/a/messages")


def test_missing_credentials_smoke_does_not_make_network_call(monkeypatch):
    monkeypatch.setattr(auth_handler, "get_secret", lambda key, default=None: default)
    assert auth_handler.get_graph_credentials() is None
    with pytest.raises(MicrosoftGraphClientError, match="not configured") as exc:
        MicrosoftGraphClient()
    assert CREDS["client_secret"] not in str(exc.value)


def test_local_smoke_checks_configuration_without_network_or_secret_output():
    """A local installation probe that never contacts Graph."""
    credentials = auth_handler.get_graph_credentials()
    if credentials is None:
        with pytest.raises(MicrosoftGraphClientError, match="not configured"):
            MicrosoftGraphClient()
        return

    client = MicrosoftGraphClient()
    assert client.credentials["tenant_id"]
    assert client.credentials["client_id"]
    assert client.token is None


def _run(handler, args):
    return json.loads(asyncio.run(handler(args)))


@pytest.mark.parametrize(
    ("handler", "args", "path", "params", "headers"),
    [
        (
            calendar_view.calendar_view,
            {"user_id": "alice@example.com", "start_date_time": "2026-01-01T00:00:00Z", "end_date_time": "2026-01-02T00:00:00Z", "calendar_id": "Team Calendar", "top": 4, "timezone": "UTC"},
            "/users/alice%40example.com/calendars/Team%20Calendar/calendarView",
            {"startDateTime": "2026-01-01T00:00:00Z", "endDateTime": "2026-01-02T00:00:00Z", "$top": 4},
            {"Prefer": 'outlook.timezone="UTC"'},
        ),
        (
            list_messages.list_messages,
            {"user_id": "alice@example.com", "folder_id": "Inbox", "filter": "isRead eq false", "search": '"hello"', "top": 3},
            "/users/alice%40example.com/mailFolders/Inbox/messages",
            {"$filter": "isRead eq false", "$search": '"hello"', "$top": 3},
            {"ConsistencyLevel": "eventual"},
        ),
        (
            get_message.get_message,
            {"user_id": "alice@example.com", "message_id": "m/1", "select": "id,subject"},
            "/users/alice%40example.com/messages/m%2F1",
            {"$select": "id,subject"},
            {},
        ),
        (
            get_event.get_event,
            {"user_id": "alice@example.com", "event_id": "e 1"},
            "/users/alice%40example.com/events/e%201",
            None,
            {},
        ),
        (
            list_contacts.list_contacts,
            {"user_id": "alice@example.com", "email": "o'connor@example.com", "select": "id,displayName"},
            "/users/alice%40example.com/contacts",
            {"$filter": "(emailAddresses/any(a:a/address eq 'o''connor@example.com'))", "$select": "id,displayName"},
            {},
        ),
    ],
)
def test_endpoint_routes_and_odata(handler, args, path, params, headers, monkeypatch):
    calls = []

    class Client:
        def get_json(self, actual_path, params=None, headers=None):
            calls.append((actual_path, params, headers or {}))
            return {"value": []}

        def get_paginated(self, actual_path, params=None, headers=None):
            return self.get_json(actual_path, params, headers)

    import importlib

    endpoint_module = importlib.import_module(handler.__module__)
    monkeypatch.setattr(endpoint_module, "graph_client_and_user", lambda _args: (Client(), "alice%40example.com"))
    assert _run(handler, args) == {"value": []}
    assert calls == [(path.lstrip("/"), params, headers)]


def test_invalid_input_returns_without_constructing_client(monkeypatch):
    from gateway.msgraph_enterprise import _helpers

    monkeypatch.setattr(_helpers, "graph_client_and_user", lambda _args: pytest.fail("unexpected client"))
    result = _run(calendar_view.calendar_view, {"start_date_time": "2026-01-01"})
    assert "start_date_time and end_date_time" in result["error"]
