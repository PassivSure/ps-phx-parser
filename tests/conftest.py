"""Shared test fixtures."""

from io import BytesIO

import pytest
from openpyxl import Workbook
from PHX.PHPP.phpp_localization import shape_model

from app.parser import load_shape


@pytest.fixture
def shape_en_10_6ip() -> shape_model.PhppShape:
    return load_shape("EN_10_6IP")


def build_workbook_bytes(
    shape: shape_model.PhppShape,
    num_of_units_value: int | float | str | None = 42,
) -> bytes:
    """Build a minimal in-memory PHPP-shaped workbook for tests.

    Writes only what the MVP parser reads:
      - A sheet named after shape.VERIFICATION.name
      - shape.VERIFICATION.num_of_units locator string + value cell
    """
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(shape.VERIFICATION.name)

    item = shape.VERIFICATION.num_of_units
    locator_row = 30
    ws[f"{item.locator_col}{locator_row}"] = item.locator_string
    value_row = locator_row + item.input_row_offset
    if num_of_units_value is not None:
        ws[f"{item.input_column}{value_row}"] = num_of_units_value

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture
def workbook_bytes(shape_en_10_6ip) -> bytes:
    return build_workbook_bytes(shape_en_10_6ip, num_of_units_value=42)
