"""Entra ID bearer-token validation for the cloud MCP surface.

Validates v2.0 access tokens issued for this API. Microsoft Entra v2 access
tokens use the API's application (client) ID GUID as the `aud` claim, even when
the OAuth resource/scope is an HTTPS Application ID URI.
The cloud server uses the validated oid to pin application-only calendar calls
to that caller. The incoming token never reaches Microsoft Graph.
"""

from __future__ import annotations

import time
import uuid

import jwt
import requests
from jwt import PyJWK

ISSUER = "https://login.microsoftonline.com/{tenant}/v2.0"
JWKS_URL = "https://login.microsoftonline.com/{tenant}/discovery/v2.0/keys"


class TokenValidationError(Exception):
    """Incoming bearer token failed validation."""


class EntraTokenValidator:
    def __init__(self, tenant_id: str, client_id: str, server_url: str = "",
                 jwks_ttl: int = 3600):
        self.tenant_id = tenant_id
        self.issuer = ISSUER.format(tenant=tenant_id)
        self.audiences = (client_id,)
        self.required_scope = "access_as_user"
        self._jwks_url = JWKS_URL.format(tenant=tenant_id)
        self._keys: dict[str, PyJWK] = {}
        self._keys_fetched = 0.0
        self._jwks_ttl = jwks_ttl

    def _refresh_keys(self) -> None:
        response = requests.get(self._jwks_url, timeout=30)
        response.raise_for_status()
        self._keys = {
            k["kid"]: PyJWK.from_dict(k) for k in response.json().get("keys", [])
        }
        self._keys_fetched = time.time()

    def _key_for(self, kid: str):
        if time.time() - self._keys_fetched > self._jwks_ttl or kid not in self._keys:
            self._refresh_keys()
        key = self._keys.get(kid)
        if key is None:
            raise TokenValidationError(f"unknown signing key: {kid}")
        return key

    def validate(self, token: str) -> dict:
        """Returns claims on success; raises TokenValidationError otherwise."""
        try:
            header = jwt.get_unverified_header(token)
            if header.get("alg") != "RS256":
                raise TokenValidationError(f"unexpected alg: {header.get('alg')}")
            key = self._key_for(header["kid"])
            claims = jwt.decode(
                token,
                key=key.key,
                algorithms=["RS256"],
                audience=self.audiences,
                issuer=self.issuer,
                options={
                    "require": ["exp", "iat", "iss", "aud", "oid", "scp"]
                },
            )
            scopes = set(str(claims.get("scp", "")).split())
            if self.required_scope not in scopes:
                raise TokenValidationError(
                    f"required delegated scope missing: {self.required_scope}"
                )
            try:
                claims["oid"] = str(uuid.UUID(str(claims["oid"])))
            except (AttributeError, TypeError, ValueError) as exc:
                raise TokenValidationError("oid must be a GUID") from exc
        except TokenValidationError:
            raise
        except jwt.ExpiredSignatureError as exc:
            raise TokenValidationError("token expired") from exc
        except jwt.InvalidTokenError as exc:
            raise TokenValidationError(str(exc)) from exc
        return claims
