"""envelope.surfaces, envelope.group_totals, window orientation and bridge groups,
checked on real workbooks against PHPP's own Summary block.

ps-rails builds its envelope from these reads and gates it on the same sums, so
each test here is the check it will make: every area group's rows add up to
PHPP's total for that group, the windows facing each way add up to that
orientation's window group (2-6), and each bridge group's lengths to its total
(15-17). Skipped when the files aren't present.
"""

import pathlib
from collections import defaultdict

import pytest

from app.main import Envelope
from app.parser import parse_workbook

DOWNLOADS = pathlib.Path.home() / "Downloads"

CASES = {
    "55th_v97": ("250121_55th_IP9_741zhOb_FINAL.xlsx", "EN_9_7IP"),
    "6840_v106": ("260309 PHPP 6840 E ^th Ave Pkwy v 10.6.xlsx", "EN_10_6IP"),
    "holmes_easyph": ("2536 Holmes_easyPH_260512.xlsx", "EN_10_6"),
}

# PHPP's window groups, by the orientation word the Windows sheet gives a row.
WINDOW_GROUPS = {"North": 2, "East": 3, "South": 4, "West": 5, "Horizontal": 6}


@pytest.fixture(scope="module", params=sorted(CASES), ids=sorted(CASES))
def envelope(request):
    filename, version = CASES[request.param]
    path = DOWNLOADS / filename
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    return parse_workbook(path.read_bytes(), version)["envelope"]


def _totals(envelope):
    return {t["group"]: t["area_ft2"] for t in envelope["group_totals"] if "area_ft2" in t}


def test_every_surface_group_adds_up_to_phpps_total(envelope):
    """Groups 7-14, the zone-X group included, which `components` leaves out."""
    sums = defaultdict(float)
    for s in envelope["surfaces"]:
        sums[s["group"]] += s["area_ft2"] or 0.0
    totals = _totals(envelope)

    for group in range(7, 15):
        assert sums.get(group, 0.0) == pytest.approx(totals.get(group, 0.0), rel=1e-9, abs=1e-6), group


def test_windows_add_up_to_phpps_total_for_each_orientation(envelope):
    sums = defaultdict(float)
    windows = [c for c in envelope["components"] if c["component"] == "window"]
    for w in windows:
        sums[WINDOW_GROUPS[w["orientation"]]] += w["area_ft2"]
    totals = _totals(envelope)

    assert windows
    for group in WINDOW_GROUPS.values():
        assert sums.get(group, 0.0) == pytest.approx(totals.get(group, 0.0), rel=1e-9, abs=1e-6), group


def test_bridge_lengths_add_up_to_phpps_total_for_each_group(envelope):
    sums = defaultdict(float)
    for b in envelope["components"]:
        if b["component"] == "thermal_bridge":
            sums[b["group"]] += b["length_ft"] or 0.0
    totals = {t["group"]: t["length_ft"] for t in envelope["group_totals"] if "length_ft" in t}

    for group in (15, 16, 17):
        assert sums.get(group, 0.0) == pytest.approx(totals.get(group, 0.0), rel=1e-9, abs=1e-6), group


def test_every_used_surface_states_its_thermal_figure(envelope):
    """R on an IP workbook, U on an SI one, in the surface table's own column."""
    used = [s for s in envelope["surfaces"] if (s["area_ft2"] or 0) > 0]
    assert used
    assert all(
        (s["r_value_hr_ft2F_Btu"] or 0) > 0 or (s["u_value_w_m2k"] or 0) > 0 for s in used
    ), [s["label"] for s in used if not (s["r_value_hr_ft2F_Btu"] or s["u_value_w_m2k"])]


def test_the_envelope_survives_the_response_model(envelope):
    """pydantic drops a field the model does not declare: until the window
    fields were declared, the API sent each window's label and area only."""
    dumped = Envelope.model_validate(envelope).model_dump()

    assert dumped["components"] == [
        {**{k: None for k in dumped["components"][i]}, **c} for i, c in enumerate(envelope["components"])
    ]
    assert len(dumped["surfaces"]) == len(envelope["surfaces"])
    assert len(dumped["group_totals"]) == len(envelope["group_totals"])


def test_the_55th_door_carries_its_typed_r():
    path = DOWNLOADS / CASES["55th_v97"][0]
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    surfaces = parse_workbook(path.read_bytes(), "EN_9_7IP")["envelope"]["surfaces"]

    door = next(s for s in surfaces if s["label"] == "Exterior door")
    assert door == {**door, "group": 7, "area_ft2": pytest.approx(44.916669), "assembly_id": None,
                    "r_value_hr_ft2F_Btu": 6.7, "source_thermal": "Areas!AE40"}
