"""Core PHPP parser: load a PHX shape, open a workbook, read fields.

Uses only `PHX.PHPP.phpp_localization.shape_model` (pydantic) and openpyxl.
Never instantiates `PHPPConnection` or touches xlwings.

P2.2 adds the `kpis` subtree (see app.kpis). Subsequent P2.x tickets add
envelope, hvac_equipment, project_info.
"""

from __future__ import annotations

import pathlib
from functools import lru_cache
from io import BytesIO
from typing import Any

from openpyxl import load_workbook
from PHX.PHPP.phpp_localization import shape_model

from app.envelope import read_envelope
from app.kpis import read_kpis

SHAPES_DIR = pathlib.Path(shape_model.__file__).parent


class ParseError(Exception):
    """Raised when the workbook cannot be parsed with the requested shape."""


def available_versions() -> list[str]:
    """Shape filenames (stems) bundled with the installed PHX package."""
    return sorted(p.stem for p in SHAPES_DIR.glob("*.json"))


@lru_cache(maxsize=16)
def load_shape(version: str) -> shape_model.PhppShape:
    """Load and cache a PhppShape by its filename stem (e.g. 'EN_10_6IP')."""
    path = SHAPES_DIR / f"{version}.json"
    if not path.is_file():
        raise ParseError(
            f"Unknown phpp_version {version!r}. Available: {available_versions()}"
        )
    return shape_model.PhppShape.model_validate_json(path.read_bytes())


def parse_workbook(workbook_bytes: bytes, version: str) -> dict[str, Any]:
    """Read a PHPP workbook's bytes and extract the v1.0.0 subtrees we cover.

    Currently emits ``kpis`` (P2.2). The endpoint wraps this in the
    schema-versioned envelope and adds the ``parser`` block.

    Raises ``ParseError`` on unknown version. Per-field read failures bubble
    up as ``None`` values within their subtree.
    """
    shape = load_shape(version)
    wb = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        return {
            "kpis": read_kpis(wb, shape),
            "envelope": read_envelope(wb, shape),
        }
    finally:
        wb.close()
