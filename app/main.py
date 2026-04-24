"""ps-phx-parser FastAPI entry point."""

from contextlib import asynccontextmanager

import httpx
from fastapi import Depends, FastAPI, HTTPException
from pydantic import BaseModel, HttpUrl

from app.auth import log_startup_auth_state, require_auth
from app.parser import ParseError, available_versions, parse_workbook


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
    phpp_version: str


class ParseResponse(BaseModel):
    phpp_version: str
    num_of_units: int | float | str | None


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/versions")
async def versions() -> dict[str, list[str]]:
    return {"versions": available_versions()}


@app.post("/parse", response_model=ParseResponse, dependencies=[Depends(require_auth)])
async def parse(req: ParseRequest) -> ParseResponse:
    async with httpx.AsyncClient(timeout=FETCH_TIMEOUT_SECONDS) as client:
        try:
            response = await client.get(str(req.url))
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

    try:
        result = parse_workbook(body, req.phpp_version)
    except ParseError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc

    return ParseResponse(**result)
