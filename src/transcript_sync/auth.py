"""Delegated Graph auth: interactive browser sign-in, Keychain token cache."""

from __future__ import annotations

import os
import sys

import keyring
import keyring.errors
import msal

TENANT_ID = os.environ.get("TRANSCRIPT_SYNC_TENANT_ID", "")
CLIENT_ID = os.environ.get("TRANSCRIPT_SYNC_CLIENT_ID", "")

SCOPES = [
    "Calendars.Read",
    "OnlineMeetings.Read",
    "OnlineMeetingTranscript.Read.All",
    "OnlineMeetingArtifact.Read.All",  # attendance reports (strict mode)
]

KEYRING_SERVICE = "transcript-sync-graph"
KEYRING_ACCOUNT = f"msal-cache-{TENANT_ID}-{CLIENT_ID or 'unset'}"
_LEGACY_KEYRING_ACCOUNT = f"msal-cache-{TENANT_ID}"  # pre-client-id key


class KeychainError(Exception):
    """macOS Keychain could not be read/written — not the same as 'no cache'."""


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    try:
        blob = keyring.get_password(KEYRING_SERVICE, KEYRING_ACCOUNT)
        if blob is None:
            # One-time migration from the legacy tenant-only key.
            legacy = keyring.get_password(KEYRING_SERVICE, _LEGACY_KEYRING_ACCOUNT)
            if legacy:
                blob = legacy
                keyring.set_password(KEYRING_SERVICE, KEYRING_ACCOUNT, legacy)
                keyring.delete_password(KEYRING_SERVICE, _LEGACY_KEYRING_ACCOUNT)
    except keyring.errors.KeyringError as exc:
        raise KeychainError(f"Keychain read failed: {exc}") from exc
    if blob:
        cache.deserialize(blob)
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        try:
            keyring.set_password(KEYRING_SERVICE, KEYRING_ACCOUNT, cache.serialize())
        except keyring.errors.KeyringError as exc:
            raise KeychainError(f"Keychain write failed: {exc}") from exc


def _app() -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache]:
    if not (TENANT_ID and CLIENT_ID):
        raise RuntimeError(
            "TRANSCRIPT_SYNC_TENANT_ID and TRANSCRIPT_SYNC_CLIENT_ID must both be set. "
            "Create the Entra app first, then export its tenant and client IDs."
        )
    cache = _load_cache()
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )
    return app, cache


def _account(app: msal.PublicClientApplication):
    accounts = app.get_accounts()
    if not accounts:
        return None
    if len(accounts) > 1:
        print(
            f"WARNING: {len(accounts)} cached accounts; using "
            f"{accounts[0].get('username')}. sign_out clears all.",
            file=sys.stderr,
        )
    return accounts[0]


def get_token() -> str:
    """Silent token if possible; otherwise interactive browser sign-in
    (authorization-code flow to a localhost listener)."""
    app, cache = _app()
    account = _account(app)
    result = app.acquire_token_silent(SCOPES, account=account) if account else None
    if not result:
        print(
            "\n=== Transcript Sync: opening your browser for "
            "Microsoft sign-in ===",
            file=sys.stderr, flush=True,
        )
        result = app.acquire_token_interactive(
            scopes=SCOPES,
            prompt="select_account",  # force account picker — default browser
        )                             # profile may hold the wrong tenant's session
    _save_cache(cache)
    if "access_token" not in result:
        raise RuntimeError(
            f"Auth failed: {result.get('error')}: {result.get('error_description')}"
        )
    return result["access_token"]


def auth_status() -> dict:
    """Granular state: cache presence AND whether a token can be silently acquired."""
    try:
        app, _ = _app()
    except KeychainError as exc:
        return {
            "client_id": CLIENT_ID,
            "tenant_id": TENANT_ID,
            "state": "keychain_error",
            "signed_in": False,
            "account": None,
            "error": str(exc),
            "scopes": SCOPES,
        }
    account = _account(app)
    if not account:
        state = "no_cached_account"
    else:
        silent = app.acquire_token_silent(SCOPES, account=account)
        state = "token_valid" if silent and "access_token" in silent else "interaction_required"
    return {
        "client_id": CLIENT_ID,
        "tenant_id": TENANT_ID,
        "state": state,
        "signed_in": account is not None,
        "account": account.get("username") if account else None,
        "scopes": SCOPES,
    }


def sign_out() -> bool:
    try:
        keyring.delete_password(KEYRING_SERVICE, KEYRING_ACCOUNT)
    except keyring.errors.PasswordDeleteError:
        pass  # nothing stored — already signed out
    return True
