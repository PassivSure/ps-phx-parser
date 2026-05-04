"""Ground-truth integration tests for project_info reads.

Skipped when the canonical workbook isn't present locally. Values pinned
2026-05-04 from ``~/Downloads/PHPP_V10.6_IP_Example.xlsx``.
"""

import pathlib
from io import BytesIO

import pytest
from openpyxl import load_workbook

from app.parser import load_shape
from app.project_info import read_project_info

REAL_WORKBOOK = pathlib.Path.home() / "Downloads" / "PHPP_V10.6_IP_Example.xlsx"

pytestmark = pytest.mark.skipif(
    not REAL_WORKBOOK.is_file(),
    reason=f"Canonical PHPP fixture not found at {REAL_WORKBOOK}",
)


@pytest.fixture(scope="module")
def real_info() -> dict:
    wb = load_workbook(BytesIO(REAL_WORKBOOK.read_bytes()), data_only=True, read_only=True)
    try:
        return read_project_info(wb, load_shape("EN_10_6IP"))
    finally:
        wb.close()


def test_real_project_name(real_info):
    assert real_info["project_name"] == "Hollis Montessori School"


def test_real_postal_code_quotes_stripped(real_info):
    """Verification!K7 in the canonical file is literally '"03049"'."""
    assert real_info["postal_code"] == "03049"


def test_real_location_string(real_info):
    assert real_info["location_string"] == "South Merrimack Road, Hollis, NH 03049"


def test_real_occupancy_type_raw_phx_format(real_info):
    """PHX option enum value, prefixed by its key. Mapper resolves into
    OccupancyType records."""
    assert (
        real_info["occupancy_type"]
        == "21-Non-res building: School half-days (< 7 h)"
    )


def test_real_organizations_pinned(real_info):
    """Five organizations populated in the canonical project: home owner,
    architect, mechanical engineer, energy consultant, certification body.
    Civil engineer / general contractor rows are blank in this file."""
    by_label = {o["label"]: o["name"] for o in real_info["organizations"]}
    assert by_label == {
        "Home owner": "Hollis Montessori School",
        "Architect": "Windy Hill Associates",
        "Mechanical engineer": "ZeroEnergy Design",
        "Energy consultant": "Example Energy Consultant",
        "Certification body": "Passive House Institute",
    }
