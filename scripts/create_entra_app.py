#!/usr/bin/env python3
"""Create an idempotent delegated public-client app for local Transcript Sync.

Uses an existing certificate-authenticated Graph caller (Application.ReadWrite.All)
to create the app registration with public-client flows and the delegated Graph
scopes the MCP server needs. Scope IDs are looked up dynamically from the Graph
resource service principal — nothing hardcoded.

Admin consent is completed separately in the Microsoft Entra admin centre.

Never prints private keys or tokens.
"""
from __future__ import annotations

import argparse
import base64
import json
import sys
import time
import uuid
from pathlib import Path
from typing import Any, cast

import requests
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding

GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
DEFAULT_APP_NAME = "Transcript Sync Local"
DELEGATED_SCOPES = [
    "Calendars.Read",
    "OnlineMeetings.Read",
    "OnlineMeetingTranscript.Read.All",
    "OnlineMeetingArtifact.Read.All",
]


def b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode()


def caller_token(tenant: str, client_id: str, pem_path: Path) -> str:
    pem = pem_path.read_bytes()
    key = serialization.load_pem_private_key(pem, password=None)
    cert = x509.load_pem_x509_certificates(pem)[0]
    now = int(time.time())
    header = {"alg": "RS256", "typ": "JWT", "x5t": b64url(cert.fingerprint(hashes.SHA1()))}
    payload = {
        "aud": f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        "iss": client_id,
        "sub": client_id,
        "jti": str(uuid.uuid4()),
        "nbf": now - 60,
        "exp": now + 600,
    }
    signing_input = f"{b64url(json.dumps(header).encode())}.{b64url(json.dumps(payload).encode())}"
    signature = key.sign(signing_input.encode(), padding.PKCS1v15(), hashes.SHA256())
    response = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": client_id,
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "client_assertion": f"{signing_input}.{b64url(signature)}",
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["access_token"]


def graph_call(token: str, method: str, path: str, body=None, params=None):
    last = None
    for attempt in range(10):
        response = requests.request(
            method,
            f"https://graph.microsoft.com/v1.0{path}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=body,
            params=params,
            timeout=60,
        )
        last = response
        retryable = response.status_code == 404 or (
            response.status_code == 400 and "NoBackingApplicationObject" in response.text
        )
        if response.status_code < 400:
            return response.json() if response.content else None
        if retryable and attempt < 9:
            time.sleep(5)
            continue
        break
    raise RuntimeError(f"Graph {method} {path}: HTTP {last.status_code}: {last.text[:600]}")


def _odata_literal(value: str) -> str:
    return value.replace("'", "''")


def resolve_application(
    token: str,
    app_name: str,
    app_client_id: str | None,
    resource_access: list[dict],
) -> tuple[dict[str, Any], bool]:
    """Resolve by immutable client ID or safely create a new local app."""
    if app_client_id:
        try:
            uuid.UUID(app_client_id)
        except ValueError as exc:
            raise SystemExit("--app-client-id must be an application UUID") from exc
        matches = cast(dict[str, Any], graph_call(
            token,
            "GET",
            "/applications",
            params={
                "$filter": f"appId eq '{_odata_literal(app_client_id)}'",
                "$select": "id,appId,displayName",
            },
        )).get("value", [])
        if len(matches) != 1:
            raise SystemExit(
                f"Expected exactly one app with client ID {app_client_id}; "
                f"found {len(matches)}."
            )
        app = matches[0]
        created = False
    else:
        matches = cast(dict[str, Any], graph_call(
            token,
            "GET",
            "/applications",
            params={
                "$filter": f"displayName eq '{_odata_literal(app_name)}'",
                "$select": "id,appId,displayName",
            },
        )).get("value", [])
        if len(matches) > 1:
            raise SystemExit(
                f"Refusing: {len(matches)} apps named '{app_name}'. "
                "Select the intended app with --app-client-id."
            )
        if matches:
            raise SystemExit(
                f"An app named '{app_name}' already exists. Refusing to mutate "
                f"it by name; rerun with --app-client-id {matches[0]['appId']}."
            )
        app = cast(dict[str, Any], graph_call(token, "POST", "/applications", body={
            "displayName": app_name,
            "signInAudience": "AzureADMyOrg",
        }))
        created = True

    graph_call(token, "PATCH", f"/applications/{app['id']}", body={
        "isFallbackPublicClient": True,
        "publicClient": {"redirectUris": ["http://localhost"]},
        "requiredResourceAccess": [
            {"resourceAppId": GRAPH_APP_ID, "resourceAccess": resource_access}
        ],
    })
    return app, created


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tenant", required=True, help="Microsoft Entra tenant ID")
    parser.add_argument("--app-name", default=DEFAULT_APP_NAME)
    parser.add_argument(
        "--app-client-id",
        help="Immutable client ID of an existing app to converge",
    )
    parser.add_argument("--caller-client-id", required=True, help="Client ID of the cert-auth caller app")
    parser.add_argument("--caller-pem", required=True, type=Path, help="PEM (cert+key) for the caller app")
    args = parser.parse_args()
    app_name = args.app_name

    token = caller_token(args.tenant, args.caller_client_id, args.caller_pem)

    # 1. Resolve delegated scope IDs dynamically from the Graph resource SP.
    graph_sp = graph_call(
        token, "GET", "/servicePrincipals",
        params={"$filter": f"appId eq '{GRAPH_APP_ID}'",
                "$select": "oauth2PermissionScopes"},
    )["value"][0]
    scope_ids = {s["value"]: s["id"] for s in graph_sp["oauth2PermissionScopes"]}
    missing = [s for s in DELEGATED_SCOPES if s not in scope_ids]
    if missing:
        sys.exit(f"Scopes not found on Graph resource: {missing}")
    resource_access = [
        {"id": scope_ids[s], "type": "Scope"} for s in DELEGATED_SCOPES
    ]

    # 2. Find or create the app registration, then converge only an explicitly
    # identified existing app.
    app, created = resolve_application(
        token, app_name, args.app_client_id, resource_access
    )
    if created:
        print(f"Created app: {app['appId']}")
    else:
        print(f"App already exists: {app['appId']} — ensuring config is current.")

    # 3. Ensure a service principal exists (needed for admin consent + sign-in).
    sps = graph_call(
        token, "GET", "/servicePrincipals",
        params={"$filter": f"appId eq '{app['appId']}'", "$select": "id"},
    ).get("value", [])
    if not sps:
        graph_call(token, "POST", "/servicePrincipals", body={"appId": app["appId"]})
        print("Created service principal.")

    print("\n=== Next step (tenant administrator) ===")
    print(
        f"In Entra admin centre, open Enterprise applications → {app_name} "
        "→ Permissions → Grant admin consent, then verify the granted scopes."
    )
    print("\nThen configure your local MCP client with:")
    print(f"  TRANSCRIPT_SYNC_CLIENT_ID={app['appId']}")
    print(f"  TRANSCRIPT_SYNC_TENANT_ID={args.tenant}")


if __name__ == "__main__":
    main()
