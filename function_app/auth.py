"""Validates the add-in's Nested App Authentication token and exchanges it
for a Microsoft Graph token via OAuth2 on-behalf-of (OBO).

Two Entra app registrations are involved (see ../Setup):
  - the ADD-IN app (public/SPA) -- the task pane acquires a token for this
    backend's exposed scope and sends it as a Bearer header.
  - the BACKEND app (confidential, this Function) -- validates that token,
    then uses its own client secret to exchange it for a Graph token scoped
    to the signed-in user (OBO), so Graph sees exactly what they can access.
"""
from __future__ import annotations

import os
from functools import lru_cache

import jwt
import msal
from jwt import PyJWKClient

TENANT_ID = os.environ["TENANT_ID"]
BACKEND_CLIENT_ID = os.environ["BACKEND_CLIENT_ID"]
BACKEND_CLIENT_SECRET = os.environ["BACKEND_CLIENT_SECRET"]
BACKEND_APP_ID_URI = os.environ.get("BACKEND_APP_ID_URI", f"api://{BACKEND_CLIENT_ID}")

_AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
_ISSUER = f"https://login.microsoftonline.com/{TENANT_ID}/v2.0"
_JWKS_URL = f"https://login.microsoftonline.com/{TENANT_ID}/discovery/v2.0/keys"


class AuthError(Exception):
    """Bearer token missing/invalid, or the OBO exchange failed."""


@lru_cache(maxsize=1)
def _jwks_client() -> PyJWKClient:
    return PyJWKClient(_JWKS_URL)


@lru_cache(maxsize=1)
def _confidential_app() -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        BACKEND_CLIENT_ID,
        authority=_AUTHORITY,
        client_credential=BACKEND_CLIENT_SECRET,
    )


def bearer_token(auth_header: str | None) -> str:
    """Pulls the raw token out of an `Authorization: Bearer <token>` header."""
    if not auth_header or not auth_header.lower().startswith("bearer "):
        raise AuthError("Missing Authorization: Bearer <token> header")
    token = auth_header.split(" ", 1)[1].strip()
    if not token:
        raise AuthError("Empty bearer token")
    return token


def validate_incoming_token(auth_header: str | None) -> dict:
    """Verifies signature, audience, issuer and expiry. Returns the claims."""
    token = bearer_token(auth_header)
    try:
        signing_key = _jwks_client().get_signing_key_from_jwt(token)
        claims = jwt.decode(
            token,
            signing_key.key,
            algorithms=["RS256"],
            audience=BACKEND_APP_ID_URI,
            issuer=_ISSUER,
        )
    except jwt.PyJWTError as e:
        raise AuthError(f"Invalid token: {e}") from e
    return claims


def get_graph_token_on_behalf_of(auth_header: str | None, scopes: list[str]) -> str:
    """Validates the incoming NAA token, then OBO-exchanges it for Graph."""
    validate_incoming_token(auth_header)  # raises AuthError if invalid
    user_assertion = bearer_token(auth_header)
    result = _confidential_app().acquire_token_on_behalf_of(
        user_assertion=user_assertion, scopes=scopes
    )
    if "access_token" not in result:
        raise AuthError(
            "On-behalf-of exchange failed: "
            + result.get("error_description", result.get("error", "unknown error"))
        )
    return result["access_token"]
