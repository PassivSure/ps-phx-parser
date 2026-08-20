import httpx
import respx
from fastapi.testclient import TestClient

from app.main import app
from tests.conftest import parse_and_wait

client = TestClient(app)

WORKBOOK_URL = "https://example.com/signed/phpp.xlsx"


@respx.mock
def test_detect_version_success(workbook_bytes_with_version):
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes_with_version)
    )

    response = client.post("/detect-version", json={"url": WORKBOOK_URL})

    assert response.status_code == 200
    assert response.json() == {
        "phpp_version": "EN_10_6IP",
        "raw_version": "10.6 IP",
        "language": "EN",
    }


@respx.mock
def test_detect_version_422_on_unsupported_file(workbook_bytes):
    # fixture has no Data sheet
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes)
    )

    response = client.post("/detect-version", json={"url": WORKBOOK_URL})

    assert response.status_code == 422
    assert "No Data" in response.json()["detail"]


@respx.mock
def test_detect_version_502_on_fetch_failure():
    respx.get(WORKBOOK_URL).mock(return_value=httpx.Response(404))

    response = client.post("/detect-version", json={"url": WORKBOOK_URL})

    assert response.status_code == 502


@respx.mock
def test_parse_auto_detects_when_phpp_version_omitted(workbook_bytes_with_version):
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes_with_version)
    )

    settled = parse_and_wait(client, WORKBOOK_URL)  # no phpp_version

    assert settled["status"] == "done", settled.get("detail")
    assert settled["result"]["parser"]["phpp_version"] == "EN_10_6IP"


@respx.mock
def test_parse_422_when_auto_detect_fails(workbook_bytes):
    # No Data sheet → detection fails
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes)
    )

    # Auto-detection needs the workbook, so it cannot be validated before the
    # 202 the way an explicit phpp_version hint can. A detection failure is
    # therefore a failed job rather than a synchronous 422.
    settled = parse_and_wait(client, WORKBOOK_URL)

    assert settled["status"] == "failed"
    assert "No Data" in settled["detail"]


@respx.mock
def test_parse_explicit_version_overrides_detection(workbook_bytes_with_version):
    # workbook has Data sheet saying 10.6 IP but caller asks for EN_10_6
    # explicit wins (caller's problem if it mismatches)
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes_with_version)
    )

    settled = parse_and_wait(client, WORKBOOK_URL, phpp_version="EN_10_6")

    assert settled["status"] == "done", settled.get("detail")
    assert settled["result"]["parser"]["phpp_version"] == "EN_10_6"
