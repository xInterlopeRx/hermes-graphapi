"""Mocked behavior tests for explicit Microsoft Graph write endpoints."""

from __future__ import annotations

import asyncio
import importlib
import json

import pytest

from gateway.msgraph_enterprise import (
    create_contact,
    create_draft,
    create_event,
    delete_contact,
    delete_event,
    delete_message,
    send_message,
    update_contact,
    update_event,
    update_message,
)
from gateway.msgraph_enterprise.client_manager import MicrosoftGraphClient
from gateway.msgraph_enterprise.client_manager import MicrosoftGraphClientError


class FakeResponse:
    def __init__(self, payload=None, status_code=200):
        self.payload = payload
        self.status_code = status_code

    def raise_for_status(self):
        if self.status_code >= 400:
            import requests

            response = requests.Response()
            response.status_code = self.status_code
            raise requests.HTTPError(response=response)

    def json(self):
        if self.payload is None:
            raise ValueError("no content")
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


@pytest.fixture(autouse=True)
def clear_graph_token_cache():
    MicrosoftGraphClient.clear_token_cache()
    yield
    MicrosoftGraphClient.clear_token_cache()


def run(handler, args):
    return json.loads(asyncio.run(handler(args)))


@pytest.mark.parametrize(
    ("module", "handler", "args", "method", "path", "request_payload"),
    [
        (
            create_event,
            create_event.create_event,
            {"event": {"subject": "Review"}},
            "post_json",
            "users/alice%40example.com/events",
            {"subject": "Review"},
        ),
        (
            create_event,
            create_event.create_event,
            {"calendar_id": "Team/42", "event": {"subject": "Review"}},
            "post_json",
            "users/alice%40example.com/calendars/Team%2F42/events",
            {"subject": "Review"},
        ),
        (
            update_event,
            update_event.update_event,
            {"event_id": "event/1", "event": {"location": {"displayName": "Room"}}},
            "patch_json",
            "users/alice%40example.com/events/event%2F1",
            {"location": {"displayName": "Room"}},
        ),
        (
            update_event,
            update_event.update_event,
            {"calendar_id": "Team Calendar", "event_id": "e1", "event": {"isAllDay": True}},
            "patch_json",
            "users/alice%40example.com/calendars/Team%20Calendar/events/e1",
            {"isAllDay": True},
        ),
        (
            delete_event,
            delete_event.delete_event,
            {"event_id": "e1"},
            "delete_json",
            "users/alice%40example.com/events/e1",
            None,
        ),
        (
            delete_event,
            delete_event.delete_event,
            {"calendar_id": "c1", "event_id": "e1"},
            "delete_json",
            "users/alice%40example.com/calendars/c1/events/e1",
            None,
        ),
        (
            create_contact,
            create_contact.create_contact,
            {"contact": {"displayName": "Ada"}},
            "post_json",
            "users/alice%40example.com/contacts",
            {"displayName": "Ada"},
        ),
        (
            create_contact,
            create_contact.create_contact,
            {"folder_id": "People/Team", "contact": {"displayName": "Ada"}},
            "post_json",
            "users/alice%40example.com/contactFolders/People%2FTeam/contacts",
            {"displayName": "Ada"},
        ),
        (
            update_contact,
            update_contact.update_contact,
            {"contact_id": "c/1", "contact": {"givenName": "Ada"}},
            "patch_json",
            "users/alice%40example.com/contacts/c%2F1",
            {"givenName": "Ada"},
        ),
        (
            delete_contact,
            delete_contact.delete_contact,
            {"contact_id": "c1"},
            "delete_json",
            "users/alice%40example.com/contacts/c1",
            None,
        ),
        (
            create_draft,
            create_draft.create_draft,
            {"message": {"subject": "Draft"}},
            "post_json",
            "users/alice%40example.com/messages",
            {"subject": "Draft"},
        ),
        (
            create_draft,
            create_draft.create_draft,
            {"folder_id": "Drafts/2", "message": {"subject": "Draft"}},
            "post_json",
            "users/alice%40example.com/mailFolders/Drafts%2F2/messages",
            {"subject": "Draft"},
        ),
        (
            update_message,
            update_message.update_message,
            {"message_id": "m/1", "message": {"subject": "Updated"}},
            "patch_json",
            "users/alice%40example.com/messages/m%2F1",
            {"subject": "Updated"},
        ),
        (
            delete_message,
            delete_message.delete_message,
            {"message_id": "m1"},
            "delete_json",
            "users/alice%40example.com/messages/m1",
            None,
        ),
        (
            send_message,
            send_message.send_message,
            {"message": {"subject": "Hello", "body": {"content": "Hi"}}},
            "post_json",
            "users/alice%40example.com/sendMail",
            {"message": {"subject": "Hello", "body": {"content": "Hi"}}, "saveToSentItems": True},
        ),
        (
            send_message,
            send_message.send_message,
            {"message": {"subject": "Hello"}, "save_to_sent_items": False},
            "post_json",
            "users/alice%40example.com/sendMail",
            {"message": {"subject": "Hello"}, "saveToSentItems": False},
        ),
    ],
)
def test_write_routes_and_payloads(monkeypatch, module, handler, args, method, path, request_payload):
    calls = []

    class Client:
        def post_json(self, actual_path, payload):
            calls.append(("post_json", actual_path, payload))
            return {"ok": True}

        def patch_json(self, actual_path, payload):
            calls.append(("patch_json", actual_path, payload))
            return {"ok": True}

        def delete_json(self, actual_path):
            calls.append(("delete_json", actual_path, None))
            return {}

    monkeypatch.setattr(module, "client_user", lambda _args: (Client(), "alice%40example.com"))
    assert run(handler, args) == ({"ok": True} if method != "delete_json" else {})
    assert calls == [(method, path, request_payload)]


@pytest.mark.parametrize(
    ("handler", "args", "required_text"),
    [
        (create_event.create_event, {}, "event must be a non-empty JSON object"),
        (update_event.update_event, {"event": {"subject": "x"}}, "event_id is required"),
        (delete_event.delete_event, {}, "event_id is required"),
        (create_contact.create_contact, {"contact": []}, "contact must be a non-empty JSON object"),
        (update_contact.update_contact, {"contact_id": "c1", "contact": {}}, "contact must be a non-empty JSON object"),
        (delete_contact.delete_contact, {}, "contact_id is required"),
        (create_draft.create_draft, {}, "message must be a non-empty JSON object"),
        (update_message.update_message, {"message_id": "m1"}, "message must be a non-empty JSON object"),
        (delete_message.delete_message, {}, "message_id is required"),
        (send_message.send_message, {}, "message must be a non-empty JSON object"),
    ],
)
def test_invalid_write_args_fail_before_client_construction(monkeypatch, handler, args, required_text):
    import gateway.msgraph_enterprise._helpers as helpers

    monkeypatch.setattr(helpers, "graph_client_and_user", lambda _args: pytest.fail("unexpected client"))
    endpoint_module = importlib.import_module(handler.__module__)
    monkeypatch.setattr(endpoint_module, "client_user", lambda _args: pytest.fail("unexpected client"))
    result = run(handler, args)
    assert required_text in result["error"]


def test_write_http_failures_are_sanitized_and_do_not_include_request_body():
    secret_body = "private message body"
    session = FakeSession([
        FakeResponse({"access_token": "mock-token", "expires_in": 3600}),
        FakeResponse({"error": secret_body}, status_code=403),
    ])
    client = MicrosoftGraphClient(
        credentials={
            "tenant_id": "test-tenant",
            "client_id": "test-client",
            "client_secret": "test-secret",
            "scope": "https://graph.microsoft.com/.default",
            "authority_url": "https://login.microsoftonline.com",
        },
        session=session,
    )
    with pytest.raises(Exception) as exc:
        client.post_json("users/u/sendMail", {"message": {"body": secret_body}})
    assert "HTTP 403" in str(exc.value)
    assert secret_body not in str(exc.value)
    assert "test-secret" not in str(exc.value)
    assert session.calls[-1][2]["timeout"] == (5.0, 30.0)


def test_endpoint_failure_json_omits_request_body_and_server_details(monkeypatch):
    private_text = "confidential body and server diagnostic"

    class FailingClient:
        def post_json(self, path, payload):
            raise MicrosoftGraphClientError("Microsoft Graph request failed (HTTP 403).")

    monkeypatch.setattr(create_draft, "client_user", lambda _args: (FailingClient(), "user"))
    result = run(create_draft.create_draft, {"message": {"body": private_text}})
    assert result == {"error": "Microsoft Graph request failed (HTTP 403)."}
    assert private_text not in json.dumps(result)


def test_existing_read_gate_performs_no_implicit_writes(monkeypatch):
    from gateway.msgraph_enterprise import get_event

    calls = []

    class ReadOnlyClient:
        def get_json(self, path, params=None, headers=None):
            calls.append(("GET", path))
            return {"id": "e1"}

        def post_json(self, *args, **kwargs):
            pytest.fail("read endpoint attempted a write")

        def patch_json(self, *args, **kwargs):
            pytest.fail("read endpoint attempted a write")

        def delete_json(self, *args, **kwargs):
            pytest.fail("read endpoint attempted a write")

    import gateway.msgraph_enterprise.get_event as module

    monkeypatch.setattr(module, "graph_client_and_user", lambda _args: (ReadOnlyClient(), "u"))
    assert run(get_event.get_event, {"event_id": "e1"}) == {"id": "e1"}
    assert calls == [("GET", "users/u/events/e1")]


def test_compatibility_package_reexports_write_handlers():
    for name in (
        "create_event", "update_event", "delete_event", "create_contact",
        "update_contact", "delete_contact", "create_draft", "update_message",
        "delete_message", "send_message",
    ):
        compatibility = importlib.import_module(f"gateway.platforms.msgraph_enterprise.{name}")
        canonical = importlib.import_module(f"gateway.msgraph_enterprise.{name}")
        assert getattr(compatibility, name) is getattr(canonical, name)
