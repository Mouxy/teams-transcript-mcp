#!/usr/bin/env python3
"""Create or update the Entra app used by the cloud Transcript Sync server.

The script uses an existing certificate-authenticated Microsoft Graph caller
with Application.ReadWrite.All to manage a single-tenant confidential client.
It configures delegated Graph permissions, an exposed access_as_user scope,
Claude connector callbacks and an OBO certificate.

Run it once before deployment, then again with --server-url after Azure assigns
the Container App URL. Client-secret creation is explicit because Microsoft
only returns the value once.
"""
from __future__ import annotations

import argparse
import base64
import datetime as dt
import sys
import uuid
from pathlib import Path
from typing import Any, cast
from urllib.parse import urlsplit

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa

sys.path.insert(0, str(Path(__file__).parent))
from create_entra_app import caller_token, graph_call

GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
DEFAULT_APP_NAME = "Transcript Sync Cloud"
DELEGATED_SCOPES = [
    "Calendars.Read",
    "OnlineMeetings.Read",
    "OnlineMeetingTranscript.Read.All",
    "OnlineMeetingArtifact.Read.All",
]
CLAUDE_CALLBACKS = [
    "https://claude.ai/api/mcp/auth_callback",
    "https://claude.com/api/mcp/auth_callback",
]
DEFAULT_CERT_OUT = Path.home() / ".transcript-sync" / "cloud-cert.pem"


def normalise_server_url(value: str | None) -> str | None:
    """Return an HTTPS origin without a trailing slash."""
    if value is None:
        return None
    value = value.rstrip("/")
    try:
        parsed = urlsplit(value)
        port = parsed.port
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            "--server-url is not a valid HTTPS origin"
        ) from exc
    if (
        parsed.scheme != "https"
        or not parsed.netloc
        or parsed.netloc.endswith(":")
        or port == 0
        or parsed.hostname is None
        or parsed.username is not None
        or parsed.password is not None
        or parsed.path
        or parsed.query
        or parsed.fragment
    ):
        raise argparse.ArgumentTypeError(
            "--server-url must be an HTTPS origin, for example "
            "https://transcript-sync.example.com"
        )
    return value


def scope_definition(scope_id: str) -> dict:
    return {
        "id": scope_id,
        "adminConsentDescription": "Access Teams transcripts as the signed-in user",
        "adminConsentDisplayName": "Access Teams transcripts",
        "userConsentDescription": "Access Teams transcripts as you",
        "userConsentDisplayName": "Access Teams transcripts",
        "value": "access_as_user",
        "type": "User",
        "isEnabled": True,
    }


def _odata_literal(value: str) -> str:
    return value.replace("'", "''")


def resolve_application(
    token: str, app_name: str, app_client_id: str | None
) -> tuple[dict[str, Any], bool]:
    """Resolve an explicitly identified app or create a new named app.

    A display-name match is never mutated without a follow-up run that names
    its immutable client ID.
    """
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
        return matches[0], False

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
            f"An app named '{app_name}' already exists. Refusing to mutate it "
            f"by name; rerun with --app-client-id {matches[0]['appId']}."
        )
    app = cast(dict[str, Any], graph_call(token, "POST", "/applications", body={
        "displayName": app_name,
        "signInAudience": "AzureADMyOrg",
    }))
    return app, True


def load_certificate_identity(pem: bytes) -> x509.Certificate:
    """Load a combined private-key/certificate PEM and verify they match."""
    private_key = serialization.load_pem_private_key(pem, password=None)
    certificates = x509.load_pem_x509_certificates(pem)
    if not certificates:
        raise ValueError("combined PEM contains no certificate")
    certificate = certificates[0]
    encoding = serialization.Encoding.DER
    public_format = serialization.PublicFormat.SubjectPublicKeyInfo
    private_public = private_key.public_key().public_bytes(encoding, public_format)
    certificate_public = certificate.public_key().public_bytes(
        encoding, public_format
    )
    if private_public != certificate_public:
        raise ValueError("private key does not match the certificate")
    return certificate


def certificate_is_registered(
    certificate: x509.Certificate, key_credentials: list[dict]
) -> bool:
    fingerprint = certificate.fingerprint(hashes.SHA1())
    # Graph has returned customKeyIdentifier in both documented base64 and
    # uppercase hexadecimal forms across tenants. Accept either representation.
    expected = {
        base64.b64encode(fingerprint).decode(),
        fingerprint.hex().upper(),
    }
    return any(
        credential.get("customKeyIdentifier") in expected
        for credential in key_credentials
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tenant", required=True, help="Microsoft Entra tenant ID")
    parser.add_argument("--app-name", default=DEFAULT_APP_NAME)
    parser.add_argument(
        "--app-client-id",
        help="Immutable client ID of an existing app to converge",
    )
    parser.add_argument("--server-url", type=normalise_server_url)
    parser.add_argument("--caller-client-id", required=True)
    parser.add_argument("--caller-pem", required=True, type=Path)
    parser.add_argument("--certificate-out", type=Path, default=DEFAULT_CERT_OUT)
    parser.add_argument(
        "--create-client-secret",
        action="store_true",
        help="Create and print a one-year connector secret (shown once)",
    )
    args = parser.parse_args()
    cert_out = args.certificate_out.expanduser().resolve()
    if cert_out.exists() and not args.app_client_id:
        sys.exit(
            f"Certificate path already exists: {cert_out}. Refusing to guess "
            "which Entra app owns it; pass --app-client-id for that app or "
            "choose a new --certificate-out path."
        )

    token = caller_token(args.tenant, args.caller_client_id, args.caller_pem)

    graph_sp = graph_call(token, "GET", "/servicePrincipals", params={
        "$filter": f"appId eq '{GRAPH_APP_ID}'", "$select": "oauth2PermissionScopes",
    })["value"][0]
    scope_ids = {s["value"]: s["id"] for s in graph_sp["oauth2PermissionScopes"]}
    missing = [scope for scope in DELEGATED_SCOPES if scope not in scope_ids]
    if missing:
        sys.exit(f"Scopes not found on Graph resource: {missing}")
    resource_access = [
        {"id": scope_ids[scope], "type": "Scope"} for scope in DELEGATED_SCOPES
    ]

    app, created = resolve_application(
        token, args.app_name, args.app_client_id
    )
    if created:
        print(f"Created app: {app['appId']}")
    else:
        print(f"App already exists: {app['appId']} — ensuring config is current.")
    if not created and not cert_out.exists():
        sys.exit(
            f"Certificate path does not exist for app {app['appId']}: {cert_out}. "
            "Refusing to generate a replacement implicitly. Restore the matching "
            "PEM or use a deliberate certificate-rotation procedure."
        )

    app_obj_id = app["id"]
    details = cast(dict[str, Any], graph_call(
        token,
        "GET",
        f"/applications/{app_obj_id}",
        params={"$select": "id,appId,api,web,keyCredentials"},
    ))
    # Shortly after a key write, GET /applications/{object-id} can return an
    # empty keyCredentials array while the filtered collection endpoint already
    # has the committed key. Use the latter as the credential source of truth.
    credential_matches = cast(dict[str, Any], graph_call(
        token,
        "GET",
        "/applications",
        params={
            "$filter": f"appId eq '{_odata_literal(app['appId'])}'",
            "$select": "id,keyCredentials",
        },
    )).get("value", [])
    registered_credentials = (
        credential_matches[0].get("keyCredentials", [])
        if len(credential_matches) == 1
        else details.get("keyCredentials", [])
    )
    existing_certificate = None
    if cert_out.exists():
        try:
            existing_certificate = load_certificate_identity(cert_out.read_bytes())
        except (TypeError, ValueError) as exc:
            sys.exit(f"Invalid combined certificate PEM at {cert_out}: {exc}")
        if not certificate_is_registered(
            existing_certificate, registered_credentials
        ):
            sys.exit(
                f"The certificate at {cert_out} is not registered on app "
                f"{app['appId']}. Refusing to deploy a mismatched OBO credential."
            )
    current_scopes = details.get("api", {}).get("oauth2PermissionScopes", [])
    access_scope = next(
        (scope for scope in current_scopes if scope.get("value") == "access_as_user"),
        None,
    )
    access_scope_id = access_scope["id"] if access_scope else str(uuid.uuid4())
    identifier_uri = args.server_url or f"api://{app['appId']}"

    graph_call(token, "PATCH", f"/applications/{app_obj_id}", body={
        "identifierUris": [identifier_uri],
        "api": {
            "requestedAccessTokenVersion": 2,
            "oauth2PermissionScopes": [scope_definition(access_scope_id)],
        },
        "web": {"redirectUris": CLAUDE_CALLBACKS},
        "requiredResourceAccess": [
            {"resourceAppId": GRAPH_APP_ID, "resourceAccess": resource_access}
        ],
    })

    service_principals = graph_call(token, "GET", "/servicePrincipals", params={
        "$filter": f"appId eq '{app['appId']}'", "$select": "id",
    }).get("value", [])
    if not service_principals:
        graph_call(token, "POST", "/servicePrincipals", body={"appId": app["appId"]})
        print("Created service principal.")

    if existing_certificate is not None:
        print(f"Certificate verified against app {app['appId']}: {cert_out}")
    else:
        key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        subject = x509.Name([
            x509.NameAttribute(x509.oid.NameOID.COMMON_NAME, args.app_name)
        ])
        cert = (
            x509.CertificateBuilder()
            .subject_name(subject)
            .issuer_name(subject)
            .public_key(key.public_key())
            .serial_number(x509.random_serial_number())
            .not_valid_before(dt.datetime.now(dt.UTC) - dt.timedelta(minutes=5))
            .not_valid_after(dt.datetime.now(dt.UTC) + dt.timedelta(days=730))
            .sign(key, hashes.SHA256())
        )
        cert_out.parent.mkdir(parents=True, exist_ok=True)
        cert_out.write_bytes(
            key.private_bytes(
                serialization.Encoding.PEM,
                serialization.PrivateFormat.PKCS8,
                serialization.NoEncryption(),
            ) + cert.public_bytes(serialization.Encoding.PEM)
        )
        cert_out.chmod(0o600)
        graph_call(token, "PATCH", f"/applications/{app_obj_id}", body={
            "keyCredentials": [{
                "type": "AsymmetricX509Cert",
                "usage": "Verify",
                "displayName": "transcript-sync-cloud-obo",
                "key": base64.b64encode(
                    cert.public_bytes(serialization.Encoding.DER)
                ).decode(),
            }],
        })
        print(f"Certificate created and uploaded; PEM saved to {cert_out} (0600).")

    secret_text = None
    if args.create_client_secret:
        secret_response = cast(dict[str, Any], graph_call(
            token,
            "POST",
            f"/applications/{app_obj_id}/addPassword",
            body={"passwordCredential": {
                "displayName": "claude-connector",
                "endDateTime": (
                    dt.datetime.now(dt.UTC) + dt.timedelta(days=365)
                ).strftime("%Y-%m-%dT%H:%M:%SZ"),
            }},
        ))
        secret_text = secret_response["secretText"]

    print("\n=== Tenant administrator action ===")
    print(
        f"In Entra admin centre, open Enterprise applications → {args.app_name} "
        "→ Permissions → Grant admin consent, then verify the granted scopes."
    )
    print("Set Assignment required? to Yes and assign only the intended pilot users/group.")
    print("\n=== Connector configuration ===")
    print("  Server URL:      https://<your-host>/mcp")
    print(f"  Client ID:       {app['appId']}")
    if secret_text:
        print(f"  Client secret:   {secret_text}   <-- shown once; store securely")
    else:
        print("  Client secret:   not created; rerun with --create-client-secret if needed")
    print(f"  Scope:           {identifier_uri}/access_as_user")
    for callback in CLAUDE_CALLBACKS:
        print(f"  Callback:        {callback}")
    if not args.server_url:
        print(
            "\nAfter deployment, rerun this script with --server-url set to the "
            "public HTTPS origin. This is required for Claude's OAuth resource matching."
        )
    print(f"\nContainer env: TRANSCRIPT_SYNC_CLOUD_CLIENT_ID={app['appId']}")
    print(f"               TRANSCRIPT_SYNC_TENANT_ID={args.tenant}")
    print(f"               TRANSCRIPT_SYNC_CLOUD_CERT_PEM={cert_out}")


if __name__ == "__main__":
    main()
