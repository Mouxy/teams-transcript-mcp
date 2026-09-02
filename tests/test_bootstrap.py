"""Tests for public Entra bootstrap helpers. No tenant calls."""

from __future__ import annotations

import argparse
import base64
import importlib.util
import stat
import sys
from pathlib import Path

import pytest

from transcript_sync.cloud import obo

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "create_cloud_app.py"
SPEC = importlib.util.spec_from_file_location("create_cloud_app", SCRIPT)
assert SPEC and SPEC.loader
create_cloud_app = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(create_cloud_app)

LOCAL_SCRIPT = ROOT / "scripts" / "create_entra_app.py"
LOCAL_SPEC = importlib.util.spec_from_file_location("create_entra_app_test", LOCAL_SCRIPT)
assert LOCAL_SPEC and LOCAL_SPEC.loader
create_entra_app = importlib.util.module_from_spec(LOCAL_SPEC)
LOCAL_SPEC.loader.exec_module(create_entra_app)


def test_server_url_is_normalised_to_https_origin():
    assert (
        create_cloud_app.normalise_server_url("https://mcp.example.com/")
        == "https://mcp.example.com"
    )


@pytest.mark.parametrize(
    "value",
    [
        "http://mcp.example.com",
        "https://mcp.example.com/mcp",
        "https://mcp.example.com?unexpected=true",
        "https://mcp.example.com#fragment",
        "https://user@mcp.example.com",
        "https://user:password@mcp.example.com",
        "https://mcp.example.com:notaport",
        "https://mcp.example.com:",
        "https://mcp.example.com:0",
        "https://mcp.example.com:65536",
        "https://[::1",
        "https://@",
        "mcp.example.com",
    ],
)
def test_server_url_rejects_non_https_origin(value):
    with pytest.raises(argparse.ArgumentTypeError):
        create_cloud_app.normalise_server_url(value)


def test_cloud_app_registers_both_claude_callback_domains():
    assert create_cloud_app.CLAUDE_CALLBACKS == [
        "https://claude.ai/api/mcp/auth_callback",
        "https://claude.com/api/mcp/auth_callback",
    ]


def test_exposed_scope_is_generic_and_user_delegated():
    scope = create_cloud_app.scope_definition("scope-id")
    assert scope["id"] == "scope-id"
    assert scope["value"] == "access_as_user"
    assert scope["type"] == "User"
    assert scope["isEnabled"] is True


def test_existing_named_cloud_app_requires_explicit_client_id(monkeypatch):
    calls = []

    def fake_graph_call(token, method, path, body=None, params=None):
        calls.append((method, path, body, params))
        return {"value": [{"id": "object-id", "appId": "22222222-3333-4444-5555-666666666666"}]}

    monkeypatch.setattr(create_cloud_app, "graph_call", fake_graph_call)
    with pytest.raises(
        SystemExit,
        match="--app-client-id 22222222-3333-4444-5555-666666666666",
    ):
        create_cloud_app.resolve_application("token", "Transcript Sync Cloud", None)
    assert all(method == "GET" for method, *_ in calls)


def test_explicit_cloud_client_id_selects_exact_app(monkeypatch):
    seen = {}

    def fake_graph_call(token, method, path, body=None, params=None):
        seen.update(params or {})
        return {"value": [{"id": "object-id", "appId": "22222222-3333-4444-5555-666666666666"}]}

    monkeypatch.setattr(create_cloud_app, "graph_call", fake_graph_call)
    app, created = create_cloud_app.resolve_application(
        "token", "Transcript Sync Cloud", "22222222-3333-4444-5555-666666666666"
    )
    assert app == {"id": "object-id", "appId": "22222222-3333-4444-5555-666666666666"}
    assert created is False
    assert seen["$filter"] == "appId eq '22222222-3333-4444-5555-666666666666'"


def test_existing_certificate_is_verified_by_thumbprint():
    class FakeCertificate:
        @staticmethod
        def fingerprint(algorithm):
            return b"certificate-thumbprint"

    expected = base64.b64encode(b"certificate-thumbprint").decode()
    credentials = [{"customKeyIdentifier": expected}]
    assert create_cloud_app.certificate_is_registered(
        FakeCertificate(), credentials
    ) is True
    assert create_cloud_app.certificate_is_registered(
        FakeCertificate(), [{"customKeyIdentifier": "different"}]
    ) is False


def test_existing_certificate_accepts_graph_hex_thumbprint():
    class FakeCertificate:
        @staticmethod
        def fingerprint(algorithm):
            return b"\x03\xf3\x09\x04"

    assert create_cloud_app.certificate_is_registered(
        FakeCertificate(), [{"customKeyIdentifier": "03F30904"}]
    ) is True


def test_existing_named_local_app_requires_explicit_client_id(monkeypatch):
    calls = []

    def fake_graph_call(token, method, path, body=None, params=None):
        calls.append((method, path, body, params))
        return {"value": [{
            "id": "object-id",
            "appId": "33333333-4444-5555-6666-777777777777",
        }]}

    monkeypatch.setattr(create_entra_app, "graph_call", fake_graph_call)
    with pytest.raises(
        SystemExit,
        match="--app-client-id 33333333-4444-5555-6666-777777777777",
    ):
        create_entra_app.resolve_application(
            "token", "Transcript Sync Local", None, []
        )
    assert all(method == "GET" for method, *_ in calls)


def test_explicit_local_client_id_converges_exact_app(monkeypatch):
    calls = []

    def fake_graph_call(token, method, path, body=None, params=None):
        calls.append((method, path, body, params))
        if method == "GET":
            return {"value": [{
                "id": "object-id",
                "appId": "33333333-4444-5555-6666-777777777777",
            }]}
        return None

    monkeypatch.setattr(create_entra_app, "graph_call", fake_graph_call)
    app, created = create_entra_app.resolve_application(
        "token",
        "Transcript Sync Local",
        "33333333-4444-5555-6666-777777777777",
        [{"id": "scope", "type": "Scope"}],
    )
    assert app["id"] == "object-id"
    assert created is False
    assert any(method == "PATCH" for method, *_ in calls)


def test_cloud_bootstrap_creates_private_obo_pem_and_uploads_certificate(
    monkeypatch, tmp_path
):
    calls = []
    graph_scopes = [
        {"value": name, "id": f"scope-{index}"}
        for index, name in enumerate(create_cloud_app.DELEGATED_SCOPES)
    ]

    def fake_graph_call(token, method, path, body=None, params=None):
        calls.append((method, path, body, params))
        if method == "GET" and path == "/servicePrincipals":
            if params and params.get("$select") == "oauth2PermissionScopes":
                return {"value": [{"oauth2PermissionScopes": graph_scopes}]}
            return {"value": []}
        if method == "GET" and path == "/applications":
            return {"value": []}
        if method == "POST" and path == "/applications":
            return {
                "id": "object-id",
                "appId": "44444444-5555-6666-7777-888888888888",
            }
        if method == "GET" and path == "/applications/object-id":
            return {"api": {}, "web": {}, "keyCredentials": []}
        return None

    certificate_out = tmp_path / "cloud-cert.pem"
    monkeypatch.setattr(create_cloud_app, "caller_token", lambda *args: "token")
    monkeypatch.setattr(create_cloud_app, "graph_call", fake_graph_call)
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "create_cloud_app.py",
            "--tenant",
            "11111111-2222-3333-4444-555555555555",
            "--caller-client-id",
            "55555555-6666-7777-8888-999999999999",
            "--caller-pem",
            str(tmp_path / "caller.pem"),
            "--certificate-out",
            str(certificate_out),
        ],
    )

    create_cloud_app.main()

    assert certificate_out.exists()
    assert stat.S_IMODE(certificate_out.stat().st_mode) == 0o600
    pem = certificate_out.read_text()
    assert "BEGIN " + "PRIVATE KEY" in pem
    assert "BEGIN CERTIFICATE" in pem
    uploaded_credentials = [
        body["keyCredentials"]
        for method, path, body, _ in calls
        if method == "PATCH" and body and "keyCredentials" in body
    ]
    assert len(uploaded_credentials) == 1
    assert uploaded_credentials[0][0]["displayName"] == "transcript-sync-cloud-obo"

    monkeypatch.setattr(obo, "CERT_PEM_PATH", pem)
    credential = obo._client_credential()
    assert credential["private_key"] == pem
    assert len(credential["thumbprint"]) == 40


def test_obo_pem_source_accepts_multiline_secret_content(tmp_path):
    pem = (
        "-----BEGIN " + "PRIVATE KEY-----\nexample\n-----END "
        + "PRIVATE KEY-----\n"
    )
    assert obo._pem_bytes(pem) == pem.encode()

    pem_path = tmp_path / "credential.pem"
    pem_path.write_text(pem)
    assert obo._pem_bytes(str(pem_path)) == pem.encode()
