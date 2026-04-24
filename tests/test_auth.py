"""Auth tests for /parse.

Existing /parse tests (tests/test_parse_endpoint.py) don't set
PHX_PARSER_AUTH_TOKEN, so they cover the auth-disabled path. This file covers
the auth-enabled path.
"""

import httpx
import pytest
import respx
from fastapi.testclient import TestClient

from app.auth import AUTH_TOKEN_ENV
from app.main import app

client = TestClient(app)

WORKBOOK_URL = "https://example.com/signed/phpp.xlsx"
TOKEN = "test-secret-token-123"


@pytest.fixture
def auth_enabled(monkeypatch):
    monkeypatch.setenv(AUTH_TOKEN_ENV, TOKEN)


def test_health_is_open_even_with_auth_enabled(auth_enabled):
    response = client.get("/health")
    assert response.status_code == 200


def test_versions_is_open_even_with_auth_enabled(auth_enabled):
    response = client.get("/versions")
    assert response.status_code == 200


def test_parse_rejects_missing_header_when_auth_enabled(auth_enabled):
    response = client.post(
        "/parse", json={"url": WORKBOOK_URL, "phpp_version": "EN_10_6IP"}
    )
    assert response.status_code == 401
    assert response.json()["detail"] == "Missing bearer token"
    assert response.headers["WWW-Authenticate"] == "Bearer"


def test_parse_rejects_wrong_token_when_auth_enabled(auth_enabled):
    response = client.post(
        "/parse",
        headers={"Authorization": "Bearer wrong-token"},
        json={"url": WORKBOOK_URL, "phpp_version": "EN_10_6IP"},
    )
    assert response.status_code == 401
    assert response.json()["detail"] == "Invalid bearer token"


def test_parse_rejects_non_bearer_scheme_when_auth_enabled(auth_enabled):
    response = client.post(
        "/parse",
        headers={"Authorization": f"Basic {TOKEN}"},
        json={"url": WORKBOOK_URL, "phpp_version": "EN_10_6IP"},
    )
    assert response.status_code in (401, 403)


@respx.mock
def test_parse_accepts_correct_bearer_token(auth_enabled, workbook_bytes):
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes)
    )

    response = client.post(
        "/parse",
        headers={"Authorization": f"Bearer {TOKEN}"},
        json={"url": WORKBOOK_URL, "phpp_version": "EN_10_6IP"},
    )
    assert response.status_code == 200
    assert response.json() == {"phpp_version": "EN_10_6IP", "num_of_units": 42}


@respx.mock
def test_parse_auth_disabled_when_env_unset(monkeypatch, workbook_bytes):
    monkeypatch.delenv(AUTH_TOKEN_ENV, raising=False)
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes)
    )

    response = client.post(
        "/parse", json={"url": WORKBOOK_URL, "phpp_version": "EN_10_6IP"}
    )
    assert response.status_code == 200
