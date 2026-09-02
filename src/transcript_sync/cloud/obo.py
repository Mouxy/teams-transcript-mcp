"""On-Behalf-Of exchange: user token for this API -> Graph token.

Per-user MSAL confidential-client instances with per-user in-memory token
caches. Process-local by design: at scale-to-zero pilot size this is one
container instance; a cold start simply re-runs OBO (Claude's connector holds
the refresh token and re-authenticates silently).
"""

from __future__ import annotations

import os
from pathlib import Path

import msal
from cryptography import x509
from cryptography.hazmat.primitives import hashes

GRAPH_SCOPES = [
    "Calendars.Read",
    "OnlineMeetings.Read",
    "OnlineMeetingTranscript.Read.All",
    "OnlineMeetingArtifact.Read.All",
]

TENANT_ID = os.environ.get("TRANSCRIPT_SYNC_TENANT_ID", "")
CLIENT_ID = os.environ.get("TRANSCRIPT_SYNC_CLOUD_CLIENT_ID", "")
CERT_PEM_PATH = os.environ.get("TRANSCRIPT_SYNC_CLOUD_CERT_PEM", "")


class OboError(Exception):
    """OBO exchange failed (consent, CA, or token problem upstream)."""


_apps: dict[str, msal.ConfidentialClientApplication] = {}
_caches: dict[str, msal.SerializableTokenCache] = {}


def _pem_bytes(source: str) -> bytes:
    """Read a file path or return PEM content supplied through an env secret."""
    if "-----BEGIN" in source:
        return source.encode()
    return Path(source).expanduser().read_bytes()


def _client_credential() -> dict:
    # Accepts a file path OR the PEM content itself (Container Apps delivers
    # secrets via secretref env vars).
    pem = _pem_bytes(CERT_PEM_PATH)
    cert = x509.load_pem_x509_certificates(pem)[0]
    thumbprint = cert.fingerprint(hashes.SHA1()).hex()
    return {"private_key": pem.decode(), "thumbprint": thumbprint}


def _app_for(oid: str) -> msal.ConfidentialClientApplication:
    app = _apps.get(oid)
    if app is None:
        cache = _caches.setdefault(oid, msal.SerializableTokenCache())
        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=_client_credential(),
            token_cache=cache,
        )
        _apps[oid] = app
    return app


def graph_token(oid: str, user_assertion: str) -> str:
    """Exchange the connector's user token for a Graph token via OBO."""
    if not (TENANT_ID and CLIENT_ID and CERT_PEM_PATH):
        raise OboError("cloud OBO is not configured (tenant/client/cert env vars)")
    result = _app_for(oid).acquire_token_on_behalf_of(
        user_assertion=user_assertion, scopes=GRAPH_SCOPES
    )
    if "access_token" not in result:
        raise OboError(
            f"OBO failed: {result.get('error')}: {result.get('error_description')}"
        )
    return result["access_token"]
