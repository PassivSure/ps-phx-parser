"""ps-phx-parser FastAPI entry point."""

import os
import pathlib
import tomllib
from collections.abc import Mapping
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from datetime import UTC, datetime
from importlib.metadata import version as package_version
from typing import Any, Literal

import httpx
from fastapi import Depends, FastAPI, HTTPException
from judoscale.asgi.middleware import FastAPIRequestQueueTimeMiddleware
from pydantic import BaseModel, HttpUrl

from app import jobs
from app.auth import log_startup_auth_state, require_auth
from app.parser import ParseError, available_versions, parse_workbook
from app.version_detection import DetectionError, detect_version


class WorkbookFetchError(Exception):
    """Fetch/size failure raised off the request path, where HTTPException has
    no meaning — there is no response to attach a status to."""


SCHEMA_VERSION = "1.0.0"
PARSER_NAME = "ps-phx-parser"


def _read_parser_pkg_version() -> str:
    pyproject = pathlib.Path(__file__).resolve().parent.parent / "pyproject.toml"
    with pyproject.open("rb") as f:
        return tomllib.load(f)["project"]["version"]


# e.g. "ps-phx-parser/0.1.0+PHX-1.56.51" — lets the consumer correlate
# parser regressions to a release.
PARSER_VERSION = f"{PARSER_NAME}/{_read_parser_pkg_version()}+PHX-{package_version('PHX')}"


@asynccontextmanager
async def lifespan(_app: FastAPI):
    log_startup_auth_state()
    yield


app = FastAPI(
    title="ps-phx-parser",
    description="Headless PHPP reader using PHX field mappings + openpyxl.",
    version="0.1.0",
    lifespan=lifespan,
)

# Request-queue-time reporting for Judoscale autoscaling.
#
# The Heroku add-on exports RAILS_AUTOSCALE_URL, but judoscale's Python package
# reads only JUDOSCALE_URL (judoscale/core/config.py). Left unhandled it installs
# cleanly, logs "Not activated - no API URL provided" and reports nothing — a
# silent no-op that looks like a working integration.
#
# Passed as extra_config rather than by setting os.environ["JUDOSCALE_URL"]
# first: judoscale builds its config singleton at IMPORT time, so an assignment
# anywhere below the import above is already too late. The middleware calls
# judoconfig.update(extra_config) before checking is_enabled, which is the
# supported way in.
def judoscale_api_url(env: Mapping[str, str] = os.environ) -> str | None:
    """Resolve the reporting URL, preferring judoscale's own variable.

    Takes `env` so this is testable without mutating the process environment or
    reloading this module — other suites hold references to `app`, and reloading
    it swaps that object out from under them.
    """
    return env.get("JUDOSCALE_URL") or env.get("RAILS_AUTOSCALE_URL")


JUDOSCALE_API_URL = judoscale_api_url()

app.add_middleware(
    FastAPIRequestQueueTimeMiddleware,
    extra_config={"API_BASE_URL": JUDOSCALE_API_URL} if JUDOSCALE_API_URL else {},
)

FETCH_TIMEOUT_SECONDS = 30.0
MAX_WORKBOOK_BYTES = 50 * 1024 * 1024  # 50 MiB; real PHPPs are ~5-15 MiB


class ParseRequest(BaseModel):
    url: HttpUrl
    phpp_version: str | None = None


class ParserMeta(BaseModel):
    name: Literal["ps-phx-parser"] = PARSER_NAME
    version: str
    phpp_version: str
    parsed_at: str
    file_hash_sha256: str | None = None


class Measurement(BaseModel):
    """Pydantic model mirrors what the parser actually emits today. The
    schema (output.schema.json) is the broader contract — fields like
    `source` may be absent or null. As later P2.x tickets emit more
    fields, expand this model in lockstep."""

    value: float | None
    unit: str
    source: str | None = None


class PeakLoad(BaseModel):
    """value_per_area / unit_per_area aren't emitted yet — P2.2 reads only
    the total peak load. They'll be added when we wire TFA-aware divisors."""

    value: float | None
    unit: str
    source: str | None = None


class PeakLoads(BaseModel):
    heating: PeakLoad
    cooling: PeakLoad


class Kpis(BaseModel):
    """KPI subtree. P2.2 fills tfa/heating_demand/cooling_demand/peak_loads/
    source_eui/pe_demand. site_eui is deferred (range-scan across PER end-uses)
    and will be added to this model when its follow-up ticket lands."""

    tfa: Measurement
    heating_demand: Measurement
    cooling_demand: Measurement
    source_eui: Measurement
    pe_demand: Measurement
    peak_loads: PeakLoads


class EnvelopeComponent(BaseModel):
    """One opaque surface (wall/roof/floor/door) or thermal bridge. P2.3
    emits component/label/area_ft2 for opaques and label/length/psi for
    thermal bridges. U/R-values come from the R-Values lookup in P2.3.b."""

    component: str
    label: str
    area_ft2: float | None = None
    length_ft: float | None = None
    psi_value_Btuh_ftF: float | None = None
    source_area: str | None = None


class Airtightness(BaseModel):
    n50_ach: float | None
    source: str | None = None


class Envelope(BaseModel):
    components: list[EnvelopeComponent]
    airtightness: Airtightness


class Organization(BaseModel):
    label: str
    name: str


class ProjectInfo(BaseModel):
    """Project metadata. P2.5 emits everything except ``verification_complete``,
    which currently hard-codes False (it needs a heuristic over Verification's
    certifier-filled cells — filed as a follow-up). The Mapper-side merge
    rule (verification KPIs supersede model KPIs when ``verification_complete``)
    therefore behaves as if the project is design-stage until the heuristic
    lands."""

    project_name: str | None = None
    postal_code: str | None = None
    location_string: str | None = None
    occupancy_type: str | None = None
    verification_complete: bool
    organizations: list[Organization]


class ParseResponse(BaseModel):
    """v1.0.0 envelope. Phase 2 fills in the optional subtrees as the
    coverage tickets ship (P2.2 kpis, P2.3 envelope, P2.4 hvac, P2.5 project).

    Schema: schema/output.schema.json. Bump SCHEMA_VERSION when the contract
    breaks; the Ruby Mapper must update in lockstep.
    """

    schema_version: Literal["1.0.0"] = SCHEMA_VERSION
    parser: ParserMeta
    kpis: Kpis | None = None
    envelope: Envelope | None = None
    project_info: ProjectInfo | None = None


class ParseAccepted(BaseModel):
    """202 body. The parse takes ~46s against a 30s router cap, so /parse hands
    back a job id immediately and the caller polls /parse/{job_id}."""

    job_id: str
    status: Literal["pending"] = "pending"


class ParseStatus(BaseModel):
    """GET /parse/{job_id}. ``result`` is populated only when status == "done";
    ``detail`` carries the reason only when status == "failed"."""

    job_id: str
    status: Literal["pending", "done", "failed"]
    detail: str | None = None
    result: ParseResponse | None = None


class WufiParseRequest(BaseModel):
    """Request for the /parse-wufi endpoint. WUFI Passive XML / .wpx files
    don't carry a version stem in the PHPP sense — PHX.from_WUFI_XML picks
    the right reader based on XML content."""

    url: HttpUrl


class DetectVersionRequest(BaseModel):
    url: HttpUrl


class DetectVersionResponse(BaseModel):
    phpp_version: str
    raw_version: str
    language: str


async def _fetch_workbook(url: str) -> bytes:
    async with httpx.AsyncClient(timeout=FETCH_TIMEOUT_SECONDS) as client:
        try:
            response = await client.get(url)
            response.raise_for_status()
        except httpx.HTTPError as exc:
            raise HTTPException(
                status_code=502,
                detail=f"Failed to fetch workbook: {exc}",
            ) from exc

    body = response.content
    if len(body) > MAX_WORKBOOK_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"Workbook exceeds {MAX_WORKBOOK_BYTES} bytes",
        )
    return body


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/versions")
async def versions() -> dict[str, list[str]]:
    return {"versions": available_versions()}


@app.post(
    "/detect-version",
    response_model=DetectVersionResponse,
    dependencies=[Depends(require_auth)],
)
async def detect_version_endpoint(req: DetectVersionRequest) -> DetectVersionResponse:
    body = await _fetch_workbook(str(req.url))
    try:
        detected = detect_version(body)
    except DetectionError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc

    return DetectVersionResponse(
        phpp_version=detected.shape_stem,
        raw_version=detected.raw,
        language=detected.language,
    )


# Parses run on a thread pool, NOT as asyncio tasks.
#
# Two reasons, and the second is the one that matters. First, the work is
# CPU-bound openpyxl: on the event loop it would block every other request for
# ~46s on this service's single uvicorn worker. Second, an asyncio task only
# progresses while the loop is being serviced — which couples "does the parse
# run" to "is anything else driving the loop". A thread has no such dependency,
# so the work completes whether or not another request arrives.
#
# max_workers=1 is deliberate: openpyxl holds the whole workbook in memory
# (~300 MB for a 27 MB file), so two concurrent parses would risk the dyno's
# memory quota. Parses queue instead. At current volume — a handful of uploads a
# day — the queue is almost always empty.
_PARSE_POOL = ThreadPoolExecutor(max_workers=1, thread_name_prefix="phx-parse")


def _fetch_workbook_sync(url: str) -> bytes:
    """Blocking fetch, for use inside the parse thread.

    The async `_fetch_workbook` is still used by /detect-version, which is fast
    enough to stay on the request path.
    """
    with httpx.Client(timeout=FETCH_TIMEOUT_SECONDS) as client:
        try:
            response = client.get(url)
            response.raise_for_status()
        except httpx.HTTPError as exc:
            raise WorkbookFetchError(f"Failed to fetch workbook: {exc}") from exc

    body = response.content
    if len(body) > MAX_WORKBOOK_BYTES:
        raise WorkbookFetchError(
            f"Workbook exceeds {MAX_WORKBOOK_BYTES} bytes: {len(body)}"
        )
    return body


def _parse_sync(body: bytes, version_hint: str | None) -> tuple[str, dict]:
    """The whole CPU-bound half, run in a worker thread.

    BOTH steps open the workbook with openpyxl, so neither may run on the event
    loop: this service has a single uvicorn worker, and a ~46s parse on the loop
    makes it unresponsive to /health and every other request for the duration.
    """
    version = version_hint or detect_version(body).shape_stem
    return version, parse_workbook(body, version)


def _run_parse(job_id: str, url: str, version_hint: str | None) -> None:
    """Fetch, parse, and record the outcome against ``job_id``.

    Runs on the parse thread. Never raises: this is detached from any request,
    so an escaping exception would leave a job stuck `pending` forever rather
    than producing a response anyone sees. Every exit path writes a terminal
    status.
    """
    try:
        body = _fetch_workbook_sync(url)
    except WorkbookFetchError as exc:
        jobs.fail(job_id, str(exc))
        return

    try:
        version, result = _parse_sync(body, version_hint)
    except (ParseError, DetectionError) as exc:
        jobs.fail(job_id, str(exc))
        return
    except Exception as exc:
        jobs.fail(job_id, f"{type(exc).__name__}: {exc}")
        return

    jobs.finish(
        job_id,
        ParseResponse(
            parser=ParserMeta(
                version=PARSER_VERSION,
                phpp_version=version,
                parsed_at=datetime.now(UTC).isoformat(),
            ),
            kpis=result.get("kpis"),
            envelope=result.get("envelope"),
            project_info=result.get("project_info"),
        ).model_dump(),
    )


@app.post(
    "/parse",
    status_code=202,
    response_model=ParseAccepted,
    dependencies=[Depends(require_auth)],
)
async def parse(req: ParseRequest) -> ParseAccepted:
    """Accept a parse and return immediately; poll /parse/{job_id} for the result.

    Asynchronous because a real workbook takes ~46s and Heroku's router caps
    time-to-first-byte at 30s — see app/jobs.py for the measurement.

    An explicitly-supplied phpp_version is validated HERE rather than in the
    worker: the check is a list membership with no I/O, and returning 422
    synchronously keeps an obviously-wrong request a fast error instead of a
    poll that ends in failure. An auto-detected version cannot be validated this
    way, since detecting it requires the workbook.
    """
    if req.phpp_version is not None and req.phpp_version not in available_versions():
        raise HTTPException(
            status_code=422,
            detail=(
                f"Unknown phpp_version '{req.phpp_version}'. "
                f"Available: {available_versions()}"
            ),
        )

    job_id = jobs.create()
    _PARSE_POOL.submit(_run_parse, job_id, str(req.url), req.phpp_version)
    return ParseAccepted(job_id=job_id)


@app.get(
    "/parse/{job_id}",
    response_model=ParseStatus,
    dependencies=[Depends(require_auth)],
)
async def parse_status(job_id: str) -> ParseStatus:
    job = jobs.get(job_id)
    if job is None:
        raise HTTPException(
            status_code=404,
            detail="Unknown or expired job_id",
        )
    return ParseStatus(
        job_id=job_id,
        status=job.status,
        detail=job.detail,
        result=job.result,
    )


@app.post(
    "/parse-wufi",
    dependencies=[Depends(require_auth)],
)
async def parse_wufi(req: WufiParseRequest) -> dict[str, Any]:
    """Parse a WUFI Passive XML / .wpx file via PHX.from_WUFI_XML.

    STUB (2026-05-12): returns HTTP 501 until BldgTyp shares a real
    sanitized WUFI Passive XML sample we can validate against. The
    Rails-side WufiExtractJob is wired up and will see this 501 as a
    clean failure (artifact marked failed, no Performance created).

    Implementation path once we have a sample file:

    1. body = await _fetch_workbook(str(req.url))
    2. parsed_xml = PHX.from_WUFI_XML.read_wufi_xml_file.xml_to_dict(...)
    3. phx_project = PHX.from_WUFI_XML.phx_converter.convert_xml_to_phx(...)
    4. envelope = walk PhxProject -> v1.x schema envelope (kpis, envelope,
       project_info, hvac_equipment, ventilation). The existing modules in
       app.kpis / app.envelope / app.project_info are workbook-shaped;
       we'll likely need parallel modules in app.wufi.* that walk PhxProject
       directly instead of openpyxl cells.
    5. Return the envelope alongside ParserMeta (set phpp_version to e.g.
       "WUFI_PASSIVE" or extract from the model itself).

    Open research questions to resolve with the sample file (and BldgTyp):
    - Sign conventions (PHPP returns null on negative cooling; WUFI?)
    - Verification flow — how is PHIUS-certifier-signed data represented?
    - PHX object graph parity vs PHPP localization path
    - Equipment hierarchy mapping (HP/Boiler/Compact + named ranges)
    """
    raise HTTPException(
        status_code=501,
        detail=(
            "WUFI Passive parsing is not yet implemented. "
            "Endpoint reserved; implementation pending a real sample file."
        ),
    )
