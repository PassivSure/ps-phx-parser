"""Shared-secret bearer auth for /parse.

If PHX_PARSER_AUTH_TOKEN is set in the environment, /parse requires an
`Authorization: Bearer <token>` header and compares in constant time. If the
env var is unset, auth is disabled (a warning is logged on startup) — useful
for local dev, tests, and CI but unsafe for any publicly reachable deploy.
"""

from __future__ import annotations

import logging
import os
import secrets
from typing import Annotated

from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer

AUTH_TOKEN_ENV = "PHX_PARSER_AUTH_TOKEN"

_bearer_scheme = HTTPBearer(auto_error=False)
_logger = logging.getLogger(__name__)


def _expected_token() -> str | None:
    token = os.environ.get(AUTH_TOKEN_ENV)
    return token if token else None


def log_startup_auth_state() -> None:
    if _expected_token() is None:
        _logger.warning(
            "%s is not set — /parse is unauthenticated. "
            "Set this env var before exposing the service publicly.",
            AUTH_TOKEN_ENV,
        )


def require_auth(
    credentials: Annotated[
        HTTPAuthorizationCredentials | None, Depends(_bearer_scheme)
    ] = None,
) -> None:
    expected = _expected_token()
    if expected is None:
        return

    if credentials is None or credentials.scheme.lower() != "bearer":
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing bearer token",
            headers={"WWW-Authenticate": "Bearer"},
        )

    if not secrets.compare_digest(credentials.credentials, expected):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid bearer token",
            headers={"WWW-Authenticate": "Bearer"},
        )
