"""Certificate-authenticated Microsoft Graph application token."""

from __future__ import annotations

import os
from pathlib import Path

import msal
from cryptography import x509
from cryptography.hazmat.primitives import hashes

GRAPH_DEFAULT_SCOPE = ["https://graph.microsoft.com/.default"]
TENANT_ID = os.environ.get("TRANSCRIPT_SYNC_TENANT_ID", "")
CLIENT_ID = os.environ.get("TRANSCRIPT_SYNC_CLOUD_CLIENT_ID", "")
CERT_PEM_PATH = os.environ.get("TRANSCRIPT_SYNC_CLOUD_CERT_PEM", "")


class ApplicationTokenError(Exception):
    """Application-token acquisition failed."""


def _pem_bytes(source: str) -> bytes:
    if "-----BEGIN" in source:
        return source.encode()
    return Path(source).expanduser().read_bytes()


def _client_credential() -> dict:
    pem = _pem_bytes(CERT_PEM_PATH)
    cert = x509.load_pem_x509_certificates(pem)[0]
    return {
        "private_key": pem.decode(),
        "thumbprint": cert.fingerprint(hashes.SHA1()).hex(),
    }


_application: msal.ConfidentialClientApplication | None = None


def _app() -> msal.ConfidentialClientApplication:
    global _application
    if _application is None:
        _application = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=_client_credential(),
        )
    return _application


def graph_token() -> str:
    """Acquire an app-only Graph token from the app's consented roles."""
    if not (TENANT_ID and CLIENT_ID and CERT_PEM_PATH):
        raise ApplicationTokenError(
            "cloud application authentication is not configured"
        )
    try:
        result = _app().acquire_token_for_client(scopes=GRAPH_DEFAULT_SCOPE) or {}
    except ApplicationTokenError:
        raise
    except Exception as exc:  # isolate the external authentication boundary
        raise ApplicationTokenError(
            f"application token setup failed: {type(exc).__name__}"
        ) from exc
    if "access_token" not in result:
        raise ApplicationTokenError(
            f"application token failed: {result.get('error')}: "
            f"{str(result.get('error_description') or '')[:200]}"
        )
    return result["access_token"]
