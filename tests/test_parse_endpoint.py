import time

import httpx
import respx
from fastapi.testclient import TestClient

from app.main import app
from tests.conftest import poll_until_settled as _poll
from tests.conftest import start_parse as _start

client = TestClient(app)

WORKBOOK_URL = "https://example.com/signed/phpp.xlsx"

def poll_until_settled(job_id: str) -> dict:
    return _poll(client, job_id)


def start_parse(**payload) -> str:
    return _start(client, WORKBOOK_URL, **payload)


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

    job_id = start_parse(phpp_version="EN_10_6IP")
    settled = poll_until_settled(job_id)
    assert settled["status"] == "done", settled.get("detail")
    body = settled["result"]
    assert body["schema_version"] == "1.0.0"
    assert body["parser"]["name"] == "ps-phx-parser"
    assert body["parser"]["phpp_version"] == "EN_10_6IP"
    assert body["parser"]["version"].startswith("ps-phx-parser/")
    assert "+PHX-" in body["parser"]["version"]
    assert body["parser"]["parsed_at"]


@respx.mock
def test_parse_fetch_failure_surfaces_as_a_failed_job(workbook_bytes):
    """A fetch failure used to be a synchronous 502. It now happens after the
    202, so it surfaces as a failed job rather than a response status."""
    respx.get(WORKBOOK_URL).mock(return_value=httpx.Response(404))

    job_id = start_parse(phpp_version="EN_10_6IP")
    settled = poll_until_settled(job_id)
    assert settled["status"] == "failed"
    assert "Failed to fetch" in settled["detail"]
    assert settled["result"] is None


@respx.mock
def test_parse_unknown_version_hint_is_still_a_synchronous_422(workbook_bytes):
    """An explicit bad hint is checked before the 202: it is list membership with
    no I/O, so there is no reason to make the caller poll to learn it is wrong."""
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes)
    )

    response = client.post(
        "/parse", json={"url": WORKBOOK_URL, "phpp_version": "XX_99_9"}
    )
    assert response.status_code == 422
    assert "Unknown phpp_version" in response.json()["detail"]


def test_status_for_unknown_job_id_is_404():
    response = client.get("/parse/does-not-exist")
    assert response.status_code == 404


@respx.mock
def test_parse_returns_before_the_work_finishes(workbook_bytes):
    """The point of the change: the 202 must not wait on the parse. If this ever
    starts reporting `done` on the first read, the work has moved back onto the
    request path and the 30s router cap applies again."""
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes)
    )

    job_id = start_parse(phpp_version="EN_10_6IP")
    first_read = client.get(f"/parse/{job_id}").json()
    assert first_read["status"] == "pending"
    assert first_read["result"] is None

    assert poll_until_settled(job_id)["status"] == "done"


@respx.mock
def test_parse_oversized_workbook_fails_the_job(monkeypatch):
    """The size cap used to be a synchronous 413. It is enforced after the
    workbook is fetched, which now happens on the parse thread, so it surfaces
    as a failed job."""
    # shrink the cap so we don't have to mock 50 MiB
    monkeypatch.setattr("app.main.MAX_WORKBOOK_BYTES", 100)

    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=b"x" * 101)
    )

    job_id = start_parse(phpp_version="EN_10_6IP")
    settled = poll_until_settled(job_id)
    assert settled["status"] == "failed"
    assert "exceeds" in settled["detail"]


def test_parse_missing_url_returns_422():
    response = client.post("/parse", json={"phpp_version": "EN_10_6IP"})
    assert response.status_code == 422


# Note: phpp_version is optional — auto-detection is covered by
# tests/test_detect_endpoint.py::test_parse_auto_detects_when_phpp_version_omitted


@respx.mock
def test_service_stays_responsive_while_a_parse_runs(monkeypatch, workbook_bytes):
    """The reason parses run on a thread rather than the event loop.

    This service has a single uvicorn worker. If the ~46s CPU-bound parse ran on
    the loop, /health and every other request would block for its duration — the
    dyno would look dead while working normally. Slow the parse down and assert
    the service still answers promptly.

    If this ever starts failing, the work has moved back onto the event loop.
    """
    import app.main as main_module

    real_parse = main_module._parse_sync

    def slow_parse(body, hint):
        time.sleep(1.5)
        return real_parse(body, hint)

    monkeypatch.setattr(main_module, "_parse_sync", slow_parse)
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes)
    )

    job_id = start_parse(phpp_version="EN_10_6IP")

    started = time.monotonic()
    health = client.get("/health")
    elapsed = time.monotonic() - started

    assert health.status_code == 200
    assert elapsed < 0.5, f"event loop blocked: /health took {elapsed:.2f}s"
    assert client.get(f"/parse/{job_id}").json()["status"] == "pending"
