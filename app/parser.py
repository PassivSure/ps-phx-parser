"""Core PHPP parser: load a PHX shape, open a workbook, read fields.

Uses only `PHX.PHPP.phpp_localization.shape_model` (pydantic) and openpyxl.
Never instantiates `PHPPConnection` or touches xlwings.

MVP reads a single field: `VERIFICATION.num_of_units`. PAS-62 expands coverage.
"""

from __future__ import annotations

import pathlib
from functools import lru_cache
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet
from PHX.PHPP.phpp_localization import shape_model

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


def _find_locator_row(
    ws: Worksheet, locator_col: str, locator_string: str
) -> int | None:
    col_idx = column_index_from_string(locator_col)
    needle = locator_string.strip().lower()
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=col_idx).value
        if val is None:
            continue
        if str(val).strip().lower() == needle:
            return row
    return None


def _read_field(ws: Worksheet, item: shape_model.VerificationInputItem) -> Any:
    row = _find_locator_row(ws, item.locator_col, item.locator_string)
    if row is None:
        return None
    target_row = row + item.input_row_offset
    target_col = column_index_from_string(item.input_column)
    return ws.cell(row=target_row, column=target_col).value


def _sheet_by_name(wb, target_name: str) -> Worksheet:
    needle = target_name.strip().lower()
    for name in wb.sheetnames:
        if name.strip().lower() == needle:
            return wb[name]
    raise ParseError(
        f"Sheet {target_name!r} not found in workbook. "
        f"Available: {wb.sheetnames[:10]}"
    )


def parse_workbook(workbook_bytes: bytes, version: str) -> dict[str, Any]:
    """Read a PHPP workbook's bytes and extract the MVP field set.

    Returns a dict shaped like:
        {"phpp_version": "EN_10_6IP", "num_of_units": 1}

    Raises ParseError on unknown version or missing target sheet.
    """
    shape = load_shape(version)

    from io import BytesIO

    wb = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        verification_ws = _sheet_by_name(wb, shape.VERIFICATION.name)
        num_of_units = _read_field(verification_ws, shape.VERIFICATION.num_of_units)
    finally:
        wb.close()

    return {
        "phpp_version": version,
        "num_of_units": num_of_units,
    }
