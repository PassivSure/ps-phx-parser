"""ps-phx-parser FastAPI entry point.

Starts as a scaffold with only /health. The /parse endpoint lands in PAS-59.
"""

from fastapi import FastAPI

app = FastAPI(
    title="ps-phx-parser",
    description="Headless PHPP reader using PHX field mappings + openpyxl.",
    version="0.1.0",
)


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}
