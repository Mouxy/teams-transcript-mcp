"""JWT validation tests — self-signed RSA keys, no network."""

from __future__ import annotations

import base64
import time

import jwt as pyjwt
import pytest
from cryptography.hazmat.primitives.asymmetric import rsa

from transcript_sync.cloud.entra_auth import EntraTokenValidator, TokenValidationError

TENANT = "11111111-2222-3333-4444-555555555555"
CLIENT = "cloud-app-id"
ISSUER = f"https://login.microsoftonline.com/{TENANT}/v2.0"

_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)


def _b64url(n: int, length: int) -> str:
    return base64.urlsafe_b64encode(n.to_bytes(length, "big")).rstrip(b"=").decode()


def _jwks(kid="test-key"):
    pub = _key.public_key().public_numbers()
    return {"keys": [{
        "kty": "RSA", "use": "sig", "kid": kid, "alg": "RS256",
        "n": _b64url(pub.n, (pub.n.bit_length() + 7) // 8),
        "e": _b64url(pub.e, (pub.e.bit_length() + 7) // 8),
        "x5c": ["unused"],
    }]}


def _validator(monkeypatch):
    """Real _refresh_keys path, with the JWKS GET mocked to our test key."""
    v = EntraTokenValidator(TENANT, CLIENT)

    class FakeResp:
        def raise_for_status(self):
            pass

        def json(self):
            return _jwks()

    monkeypatch.setattr(
        "transcript_sync.cloud.entra_auth.requests.get",
        lambda *a, **k: FakeResp(),
    )
    return v


def _token(overrides: dict | None = None, exp_delta: int = 3600) -> str:
    now = int(time.time())
    claims = {
        "iss": ISSUER, "aud": f"api://{CLIENT}", "iat": now, "nbf": now - 60,
        "exp": now + exp_delta, "oid": "user-oid-123",
        "preferred_username": "user@example.com",
        "scp": "access_as_user",
    }
    claims.update(overrides or {})
    return pyjwt.encode(claims, _key, algorithm="RS256",
                        headers={"kid": "test-key"})


def test_valid_token_returns_claims(monkeypatch):
    v = _validator(monkeypatch)
    claims = v.validate(_token())
    assert claims["oid"] == "user-oid-123"
    assert claims["preferred_username"] == "user@example.com"


def test_wrong_audience_rejected(monkeypatch):
    v = _validator(monkeypatch)
    with pytest.raises(TokenValidationError):
        v.validate(_token({"aud": "api://someone-else"}))


def test_wrong_issuer_rejected(monkeypatch):
    v = _validator(monkeypatch)
    with pytest.raises(TokenValidationError):
        v.validate(_token({"iss": "https://evil.example/v2.0"}))


def test_expired_rejected(monkeypatch):
    v = _validator(monkeypatch)
    with pytest.raises(TokenValidationError):
        v.validate(_token(exp_delta=-3600))


def test_missing_oid_rejected(monkeypatch):
    v = _validator(monkeypatch)
    token = pyjwt.encode(
        {"iss": ISSUER, "aud": f"api://{CLIENT}", "iat": int(time.time()),
         "exp": int(time.time()) + 3600},
        _key, algorithm="RS256", headers={"kid": "test-key"})
    with pytest.raises(TokenValidationError):
        v.validate(token)


def test_missing_or_wrong_delegated_scope_rejected(monkeypatch):
    v = _validator(monkeypatch)
    with pytest.raises(TokenValidationError, match="access_as_user"):
        v.validate(_token({"scp": ""}))
    with pytest.raises(TokenValidationError, match="access_as_user"):
        v.validate(_token({"scp": "unrelated.scope"}))


def test_server_url_is_the_only_runtime_audience(monkeypatch):
    server_url = "https://mcp.example.com"
    v = EntraTokenValidator(TENANT, CLIENT, server_url=server_url)

    class FakeResp:
        def raise_for_status(self):
            pass

        def json(self):
            return _jwks()

    monkeypatch.setattr(
        "transcript_sync.cloud.entra_auth.requests.get",
        lambda *args, **kwargs: FakeResp(),
    )
    claims = v.validate(_token({"aud": server_url}))
    assert claims["aud"] == server_url
    with pytest.raises(TokenValidationError):
        v.validate(_token({"aud": f"api://{CLIENT}"}))
