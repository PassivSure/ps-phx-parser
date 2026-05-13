"""Tests for the /parse-wufi stub endpoint.

The endpoint is a placeholder until BldgTyp shares a real WUFI Passive
XML sample. These tests pin the contract: route exists, requires auth,
returns 501 with a clear message.
"""

from fastapi.testclient import TestClient

from app.main import app

client = TestClient(app)

WUFI_URL = "https://example.com/signed/model.xml"


def test_parse_wufi_returns_501_until_implemented():
    response = client.post("/parse-wufi", json={"url": WUFI_URL})
    assert response.status_code == 501
    assert "not yet implemented" in response.json()["detail"].lower()


def test_parse_wufi_rejects_missing_url():
    response = client.post("/parse-wufi", json={})
    assert response.status_code == 422
