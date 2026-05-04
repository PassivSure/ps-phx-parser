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


class ParseResponse(BaseModel):
    """v1.0.0 envelope. Phase 2 will populate kpis/envelope/hvac_equipment/etc.

    Schema: schema/output.schema.json. Bump SCHEMA_VERSION when the contract
    breaks; the Ruby Mapper must update in lockstep.
    """

    schema_version: Literal["1.0.0"] = SCHEMA_VERSION
    parser: ParserMeta


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
        # parse_workbook still returns the MVP placeholder fields; Phase 2
        # tickets (P2.2+) will expand it. The endpoint contract is the v1
        # envelope below — placeholder fields are dropped from the response.
        parse_workbook(body, version)
    except ParseError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc

    return ParseResponse(
        parser=ParserMeta(
            version=PARSER_VERSION,
            phpp_version=version,
            parsed_at=datetime.now(UTC).isoformat(),
        ),
    )
