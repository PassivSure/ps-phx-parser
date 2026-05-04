"""ps-phx-parser FastAPI entry point."""

import pathlib
import tomllib
from contextlib import asynccontextmanager
from datetime import UTC, datetime
from importlib.metadata import version as package_version
from typing import Literal

import httpx
from fastapi import Depends, FastAPI, HTTPException
from pydantic import BaseModel, HttpUrl

from app.auth import log_startup_auth_state, require_auth
from app.parser import ParseError, available_versions, parse_workbook
from app.version_detection import DetectionError, detect_version

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


@app.post(
    "/parse",
    response_model=ParseResponse,
    dependencies=[Depends(require_auth)],
)
async def parse(req: ParseRequest) -> ParseResponse:
    body = await _fetch_workbook(str(req.url))

    version = req.phpp_version
    if version is None:
        try:
            version = detect_version(body).shape_stem
        except DetectionError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc

    try:
        result = parse_workbook(body, version)
    except ParseError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc

    return ParseResponse(
        parser=ParserMeta(
            version=PARSER_VERSION,
            phpp_version=version,
            parsed_at=datetime.now(UTC).isoformat(),
        ),
        kpis=result.get("kpis"),
        envelope=result.get("envelope"),
        project_info=result.get("project_info"),
    )
