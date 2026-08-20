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
    # No "door" key: this building has no exterior door. The row PHPP seeds
    # for area group 7 carries no quantity, so its area formula yields "" and
    # PHPP totals the group as 0. It used to be counted as a door — see
    # test_no_phantom_door below.
    assert counts == {
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


def test_no_phantom_door(real_envelope):
    """This reverses a previously pinned behaviour, deliberately.

    The old assertion read: "The single door entry in this file has no area
    filled in (it's a template row); we still emit it because the description
    is set." So the template row was recognised and emitting it was a choice.

    It was the wrong one. PHPP seeds one such row per area group and totals an
    unused group as **0** — this file's own Summary block says `door: 0.0`. A
    consumer counting doors got one that does not exist, with a null area, on
    every building without an exterior door. Agreeing with the workbook means
    emitting nothing.
    """
    doors = [c for c in real_envelope["components"] if c["component"] == "door"]
    assert doors == []


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
