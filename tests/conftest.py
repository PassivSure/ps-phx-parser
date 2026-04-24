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
    data_sheet_version: str | None = None,
    data_sheet_pe_factor: str | None = "1-PE-factors (non-renewable) PHI Certification",
) -> bytes:
    """Build a minimal in-memory PHPP-shaped workbook for tests.

    Always writes:
      - A sheet named after shape.VERIFICATION.name
      - shape.VERIFICATION.num_of_units locator string + value cell

    Optionally writes a Data sheet for version-detection tests. Pass
    data_sheet_version='10.6 IP' (or similar) to enable it.
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

    if data_sheet_version is not None:
        data_ws = wb.create_sheet("Data")
        data_ws["A5"] = "PHPP Version"
        data_ws["B5"] = data_sheet_version
        if data_sheet_pe_factor:
            data_ws["C5"] = "Language"
            data_ws["D5"] = data_sheet_pe_factor

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture
def workbook_bytes(shape_en_10_6ip) -> bytes:
    return build_workbook_bytes(shape_en_10_6ip, num_of_units_value=42)


@pytest.fixture
def workbook_bytes_with_version(shape_en_10_6ip) -> bytes:
    return build_workbook_bytes(
        shape_en_10_6ip,
        num_of_units_value=42,
        data_sheet_version="10.6 IP",
    )
