"""Ground-truth integration tests for envelope reads.

Skipped when the canonical workbook isn't present locally. Values pinned
2026-05-04 from ``~/Downloads/PHPP_V10.6_IP_Example.xlsx``.
"""

import pathlib
from io import BytesIO
from math import isclose

import pytest
from openpyxl import load_workbook

from app.envelope import read_envelope
from app.parser import load_shape

REAL_WORKBOOK = pathlib.Path.home() / "Downloads" / "PHPP_V10.6_IP_Example.xlsx"

pytestmark = pytest.mark.skipif(
    not REAL_WORKBOOK.is_file(),
    reason=f"Canonical PHPP fixture not found at {REAL_WORKBOOK}",
)


@pytest.fixture(scope="module")
def real_envelope() -> dict:
    wb = load_workbook(BytesIO(REAL_WORKBOOK.read_bytes()), data_only=True, read_only=True)
    try:
        return read_envelope(wb, load_shape("EN_10_6IP"))
    finally:
        wb.close()


def test_real_component_counts_by_type(real_envelope):
    """Pinned counts for the canonical file. If a future PHX/openpyxl bump
    or shape change breaks classification, these will catch it."""
    counts = {}
    for c in real_envelope["components"]:
        counts.setdefault(c["component"], 0)
        counts[c["component"]] += 1
    assert counts == {
        "door": 1,
        "slab_on_grade": 1,
        "roof": 4,
        "wall": 20,
        "thermal_bridge": 2,
    }


def test_real_roof_total_area(real_envelope):
    """The canonical file has 4 roof entries summing to ~9856 ft²
    (4742.6 + 4658.5 + 227.3 + 227.3)."""
    roofs = [c for c in real_envelope["components"] if c["component"] == "roof"]
    total = sum(c["area_ft2"] for c in roofs if c["area_ft2"] is not None)
    assert isclose(total, 9855.6285, rel_tol=0.001)


def test_real_slab_label_and_area(real_envelope):
    """One slab on grade — area matches the projected building footprint
    closely (it's the basement floor)."""
    slabs = [c for c in real_envelope["components"] if c["component"] == "slab_on_grade"]
    assert len(slabs) == 1
    assert slabs[0]["label"] == "Floor_9183_D"
    assert isclose(slabs[0]["area_ft2"], 9056.111, rel_tol=0.001)


def test_real_door_label(real_envelope):
    """The single door entry in this file has no area filled in (it's a
    template row); we still emit it because the description is set."""
    doors = [c for c in real_envelope["components"] if c["component"] == "door"]
    assert len(doors) == 1
    assert doors[0]["label"] == "Exterior door"


def test_real_thermal_bridges(real_envelope):
    """Two thermal bridges — slab edge & interior drain pipes."""
    tbs = [c for c in real_envelope["components"] if c["component"] == "thermal_bridge"]
    assert len(tbs) == 2
    by_label = {tb["label"]: tb for tb in tbs}
    assert "Interior drain pipes pipes" in by_label
    assert "Estimate Load Bearing Walls" in by_label
    assert isclose(
        by_label["Estimate Load Bearing Walls"]["length_ft"], 136.155, rel_tol=0.001
    )
    assert isclose(
        by_label["Estimate Load Bearing Walls"]["psi_value_Btuh_ftF"], 0.1733, rel_tol=0.001
    )


def test_real_airtightness_n50(real_envelope):
    """n50 = 0.3 ACH for this passive house example."""
    assert real_envelope["airtightness"]["n50_ach"] == 0.3
    assert real_envelope["airtightness"]["source"] == "Ventilation!M23"
