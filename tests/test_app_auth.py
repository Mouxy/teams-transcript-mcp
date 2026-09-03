"""Cloud application-token acquisition tests."""

from __future__ import annotations

from types import SimpleNamespace


def test_application_graph_token_uses_client_credentials(monkeypatch):
    from transcript_sync.cloud import app_auth

    calls = []

    class FakeApplication:
        def acquire_token_for_client(self, scopes):
            calls.append(scopes)
            return {"access_token": "application-graph-token"}

    monkeypatch.setattr(app_auth, "TENANT_ID", "tenant-id")
    monkeypatch.setattr(app_auth, "CLIENT_ID", "client-id")
    monkeypatch.setattr(app_auth, "CERT_PEM_PATH", "configured")
    monkeypatch.setattr(app_auth, "_application", FakeApplication())

    assert app_auth.graph_token() == "application-graph-token"
    assert calls == [["https://graph.microsoft.com/.default"]]


def test_cloud_context_uses_app_token_and_validated_caller_oid(monkeypatch):
    from transcript_sync.cloud import server

    monkeypatch.setattr(server.app_auth, "graph_token", lambda: "app-token")
    request = SimpleNamespace(
        scope={
            "entra_claims": {
                "oid": "validated-caller-oid",
                "preferred_username": "caller@example.com",
            },
            "entra_token": "connector-user-token",
        }
    )
    ctx = SimpleNamespace(request_context=SimpleNamespace(request=request))

    fetch_context = server._ctx_for(ctx)

    assert fetch_context.token == "app-token"
    assert fetch_context.graph_user_id == "validated-caller-oid"
    assert fetch_context.user_email == "caller@example.com"


def test_missing_certificate_is_mapped_to_server_not_authorized(monkeypatch, tmp_path):
    from transcript_sync.cloud import app_auth, server

    monkeypatch.setattr(app_auth, "TENANT_ID", "tenant-id")
    monkeypatch.setattr(app_auth, "CLIENT_ID", "client-id")
    monkeypatch.setattr(app_auth, "CERT_PEM_PATH", str(tmp_path / "missing.pem"))
    monkeypatch.setattr(app_auth, "_application", None)

    result = server._application_guard(app_auth.graph_token)

    assert result == {
        "status": "error",
        "error_code": "server_not_authorized",
        "message": "'Transcript Sync Cloud' cannot authenticate to Microsoft Graph.",
    }
