"""app.boiler on synthetic and real workbooks."""

import pathlib
from io import BytesIO

import pytest
from openpyxl import Workbook, load_workbook

from app.boiler import read_boiler
from app.parser import load_shape, parse_workbook

DOWNLOADS = pathlib.Path.home() / "Downloads"


def _boiler_workbook(
    boiler_type="10-Improved condensing gas boiler",
    fuel="30-Natural gas",
    unit="kBTU/hr",
    sheet="Boiler",
):
    """The Boiler sheet's input block as PHPP 9.7 and 10.x lay it out: labels in
    I, the "Project data" column N, units in O, user-defined values in R (set
    apart here, so a read of the wrong column fails)."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(sheet)
    ws["I21"], ws["N21"] = "Boiler type", boiler_type
    ws["I22"], ws["N22"] = "Fuel", fuel
    ws["N32"], ws["P32"], ws["R32"] = "Project data", "Standard values", "User-defined"
    ws["I33"], ws["N33"], ws["O33"], ws["P33"], ws["R33"] = "Design output", 40, unit, 51.2, 99
    ws["I34"], ws["N34"] = "Installation of boiler (Outdoor: 0, Indoor: 1)", 1
    ws["I36"], ws["N36"], ws["R36"] = "Boiler efficiency at 30% load", 0.97, 0.5
    ws["I37"], ws["N37"], ws["R37"] = "Boiler efficiency at nominal output", 0.93, 0.5
    ws["I38"], ws["N38"], ws["R38"] = "Standby heat loss boiler at 70 °C", 0.012, 0.5
    buf = BytesIO()
    wb.save(buf)
    return load_workbook(BytesIO(buf.getvalue()), data_only=True, read_only=True)


def test_reads_the_project_data_column():
    b = read_boiler(_boiler_workbook(), load_shape("EN_10_6IP"))
    assert {k: v for k, v in b.items() if k != "sources"} == {
        "boiler_type": "10-Improved condensing gas boiler",
        "fuel": "30-Natural gas",
        "design_output": 40.0,
        "design_output_unit": "kBTU/hr",
        "indoor": True,
        "efficiency_30": 0.97,
        "efficiency_100": 0.93,
        "standby_loss_70": 0.012,
    }
    assert b["sources"]["design_output"] == "Boiler!N33"
    assert b["sources"]["efficiency_30"] == "Boiler!N36"


def test_si_unit_comes_from_the_sheet():
    b = read_boiler(_boiler_workbook(unit="kW"), load_shape("EN_10_6"))
    assert b["design_output_unit"] == "kW"


@pytest.mark.parametrize("boiler_type", ["1-None", None, "  "])
def test_no_boiler_type_is_none(boiler_type):
    assert read_boiler(_boiler_workbook(boiler_type=boiler_type), load_shape("EN_10_6IP")) is None


def test_dash_fuel_is_none():
    assert read_boiler(_boiler_workbook(fuel="-"), load_shape("EN_10_6IP"))["fuel"] is None


def test_no_boiler_sheet_is_none():
    wb = Workbook()
    assert read_boiler(wb, load_shape("EN_10_6IP")) is None


@pytest.mark.parametrize(
    "file,version,expected",
    [
        # The PHI 10.6 IP example (Hollis Montessori School) with its heat
        # generator switched in Excel to an improved condensing gas boiler for
        # heating and DHW, user-defined values on, climate US0035a-Boston.
        (
            "PHPP_V10.6_IP_Example_boiler.xlsx",
            "EN_10_6IP",
            {
                "boiler_type": "10-Improved condensing gas boiler",
                "fuel": "30-Natural gas",
                "design_output": 40.0,
                "design_output_unit": "kBTU/hr",
                "indoor": True,
                "efficiency_30": 0.97,
                "efficiency_100": 0.93,
                "standby_loss_70": 0.012,
            },
        ),
        ("PHPP_V10.6_IP_Example.xlsx", "EN_10_6IP", None),
        ("PHPP_EN_V10.6_easyPH_Example.xlsx", "EN_10_6", None),
        ("250121_55th_IP9_741zhOb_FINAL.xlsx", "EN_9_7IP", None),
    ],
)
def test_real_workbooks(file, version, expected):
    path = DOWNLOADS / file
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    b = parse_workbook(path.read_bytes(), version)["boiler"]
    if expected is None:
        assert b is None
    else:
        assert {k: v for k, v in b.items() if k != "sources"} == expected


def test_boiler_rows_in_use_on_the_boiler_example():
    path = DOWNLOADS / "PHPP_V10.6_IP_Example_boiler.xlsx"
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    hg = parse_workbook(path.read_bytes(), "EN_10_6IP")["heat_generation"]
    used = {
        block: {(r["label"], r["source"]) for r in hg[block] if (r["final_energy"] or 0) > 0}
        for block in ("heating", "dhw")
    }
    assert used == {
        "heating": {("Boiler Condensing Natural gas", "PER!T23")},
        "dhw": {
            ("Boiler Condensing Natural gas", "PER!T45"),
            ("Solar thermal system", "PER!T48"),
        },
    }
