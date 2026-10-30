"""Ground truth from three real v10.6 workbooks.

Skipped when the files are absent, like the other real-workbook suites. Three
files, not one -- and they exercise different lookup paths, which a smaller
corpus would hide: the Swegon is user-defined (`01ud`, user block) and the
Brink is a certified catalog entry (`1363vs03`, catalog block). Both of those
recover moisture, so on their own they'd never exercise the `hrv` branch of
`_classify_ventilation` -- the easyPH example's user-defined unit (`01ud`,
row 13) has no humidity-recovery figure and closes that gap.
"""

import pathlib

import pytest

from app.parser import parse_workbook
from tests.test_equipment import _hvac_equipment_item_allowed_keys

DOWNLOADS = pathlib.Path.home() / "Downloads"

CASES = {
    "6840_ip": {
        "file": "260309 PHPP 6840 E ^th Ave Pkwy v 10.6.xlsx",
        "version": "EN_10_6IP",
        "vent_name": "Swegon Casa R7 Genius Sorption",
        "vent_type": "erv",          # humidity recovery 0.86
        "heat_recovery_pct": 86.0,
        "airflow_m3h": 306.924,
        "airflow_cfm": 180.65,
    },
    "holmes_si": {
        "file": "2536 Holmes_easyPH_260512.xlsx",
        "version": "EN_10_6",
        "vent_name": "Brink Climate Systems B.V. - Brink Flair 400 Enthalpie",
        "vent_type": "erv",          # humidity recovery 0.71
        "heat_recovery_pct": 84.0,
        "airflow_m3h": 319.412,
        "airflow_cfm": 188.00,
    },
    "easyph_hrv": {
        "file": "PHPP_EN_V10.6_easyPH_Example.xlsx",
        "version": "EN_10_6",
        "vent_name": "Heat recovery unit",
        "vent_type": "hrv",          # no humidity recovery figure at all
        "heat_recovery_pct": 83.0,
        "airflow_m3h": 152.0,
        "airflow_cfm": 89.46,
    },
}


@pytest.fixture(scope="module", params=sorted(CASES), ids=sorted(CASES))
def case(request):
    spec = CASES[request.param]
    path = DOWNLOADS / spec["file"]
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    result = parse_workbook(path.read_bytes(), spec["version"])
    return spec, result["hvac_equipment"]


def _vent(items):
    return next(i for i in items if i["equipment_type"] in ("hrv", "erv"))


def test_ventilation_device_name(case):
    spec, items = case
    assert _vent(items)["name"] == spec["vent_name"]


def test_ventilation_is_classified_by_humidity_recovery(case):
    """Both the 6840 and Holmes units recover moisture, so both are ERVs.
    Production stores one of them as `hrv`, which this is more accurate
    than. The easyPH example unit has no humidity-recovery figure at all,
    which is the `hrv` branch's only real-data coverage."""
    spec, items = case
    assert _vent(items)["equipment_type"] == spec["vent_type"]


def test_heat_recovery_efficiency(case):
    spec, items = case
    assert _vent(items)["heat_recovery_efficiency_pct"] == pytest.approx(
        spec["heat_recovery_pct"], abs=0.5
    )


def test_airflow_in_both_units(case):
    """The metric read is the source; cfm is derived. This pairing is what
    validated the whole approach: 306.924 m3/h -> 180.65 cfm against the 181.0
    an independent extraction recorded for the same workbook."""
    spec, items = case
    vent = _vent(items)
    assert vent["airflow_m3h"] == pytest.approx(spec["airflow_m3h"], abs=0.01)
    assert vent["airflow_cfm"] == pytest.approx(spec["airflow_cfm"], abs=0.5)


def test_every_item_conforms_to_the_schema_key_set(case):
    # Two-way, not `<=`: a subset check passes on an item that's missing a
    # required key, which is exactly the shape of bug this test exists to
    # catch. `allowed` comes from the schema itself (same helper
    # test_equipment.py uses), not a second hardcoded copy that could drift
    # from the real contract. Loops over every emitted item (not `_vent()`'s
    # hrv/erv-only filter), since that's the only assertion in this file
    # that ever touches the cooling_unit / heat_pump / dhw_heat_pump items a
    # real workbook may also emit.
    _, items = case
    allowed = _hvac_equipment_item_allowed_keys()
    assert items, "no equipment extracted"
    for item in items:
        assert set(item) == allowed, set(item) ^ allowed


# Heat-pump capacity (kW) and COP, read by hand off each file's metric HP pane
# ('HP SI' on an IP workbook): the test point with the largest capacity. The
# IP pane of the same files shows the COP x 3.412 (16.38 for 6840's 4.8), which
# is what an IP-pane read reports.
HEAT_PUMPS = {
    "6840_ip": ("260309 PHPP 6840 E ^th Ave Pkwy v 10.6.xlsx", "EN_10_6IP", {
        ("heat_pump", "Standard air-to-air heat pump"): (18.36, 4.8),
        ("dhw_heat_pump", "Standard DHW HP"): (7.44, 3.0),
    }),
    "phi_10_6_ip_example": ("PHPP_V10.6_IP_Example.xlsx", "EN_10_6IP", {
        # Its water-based unit is selected but not ticked, and its table is
        # filled in anyway: not emitted.
        ("heat_pump", "Minisplit classroom"): (1.08, 4.56),
    }),
    "easyph_example": ("PHPP_EN_V10.6_easyPH_Example.xlsx", "EN_10_6", {
        ("heat_pump", "Standard air/water heat pump"): (5.0, 4.3),
        ("dhw_heat_pump", "Standard air/water heat pump"): (5.0, 4.3),
    }),
    "holmes_si": ("2536 Holmes_easyPH_260512.xlsx", "EN_10_6", {
        # Its heating is a combined ventilation unit (93ud): HP!AF66 / AG66,
        # test point 1's maximum, the largest of its ten points.
        ("heat_pump", "easyPH heating"): (13.07, 3.91),
        ("dhw_heat_pump", "HPWH"): (1.6, 2.84),
    }),
}


@pytest.fixture(scope="module", params=sorted(HEAT_PUMPS), ids=sorted(HEAT_PUMPS))
def heat_pump_case(request):
    file, version, expected = HEAT_PUMPS[request.param]
    path = DOWNLOADS / file
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    items = parse_workbook(path.read_bytes(), version)["hvac_equipment"]
    return expected, [i for i in items if i["equipment_type"] in ("heat_pump", "dhw_heat_pump")]


def test_heat_pump_capacity_and_cop(heat_pump_case):
    expected, items = heat_pump_case
    got = {(i["equipment_type"], i["name"]): (i["capacity"], i["efficiency_value"]) for i in items}
    assert set(got) == set(expected)
    for key, (capacity, cop) in expected.items():
        assert got[key][0] == pytest.approx(capacity, abs=0.01), key
        assert got[key][1] == pytest.approx(cop, abs=0.01), key
    assert all(i["capacity_unit"] == "kW" and i["efficiency_type"] == "COP" for i in items)
