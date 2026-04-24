import httpx
import respx
from fastapi.testclient import TestClient

from app.main import app

client = TestClient(app)

WORKBOOK_URL = "https://example.com/signed/phpp.xlsx"


def test_versions_endpoint_lists_bundled_shapes():
    response = client.get("/versions")
    assert response.status_code == 200
    body = response.json()
    assert "EN_10_6IP" in body["versions"]


@respx.mock
def test_parse_success(workbook_bytes):
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes)
    )

    response = client.post(
        "/parse", json={"url": WORKBOOK_URL, "phpp_version": "EN_10_6IP"}
    )
    assert response.status_code == 200
    assert response.json() == {"phpp_version": "EN_10_6IP", "num_of_units": 42}


@respx.mock
def test_parse_fetch_failure_returns_502(workbook_bytes):
    respx.get(WORKBOOK_URL).mock(return_value=httpx.Response(404))

    response = client.post(
        "/parse", json={"url": WORKBOOK_URL, "phpp_version": "EN_10_6IP"}
    )
    assert response.status_code == 502
    assert "Failed to fetch" in response.json()["detail"]


@respx.mock
def test_parse_unknown_version_returns_422(workbook_bytes):
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes)
    )

    response = client.post(
        "/parse", json={"url": WORKBOOK_URL, "phpp_version": "XX_99_9"}
    )
    assert response.status_code == 422
    assert "Unknown phpp_version" in response.json()["detail"]


@respx.mock
def test_parse_oversized_workbook_returns_413(monkeypatch):
    # shrink the cap so we don't have to mock 50 MiB
    monkeypatch.setattr("app.main.MAX_WORKBOOK_BYTES", 100)

    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=b"x" * 101)
    )

    response = client.post(
        "/parse", json={"url": WORKBOOK_URL, "phpp_version": "EN_10_6IP"}
    )
    assert response.status_code == 413


def test_parse_missing_url_returns_422():
    response = client.post("/parse", json={"phpp_version": "EN_10_6IP"})
    assert response.status_code == 422


# Note: phpp_version is optional — auto-detection is covered by
# tests/test_detect_endpoint.py::test_parse_auto_detects_when_phpp_version_omitted
