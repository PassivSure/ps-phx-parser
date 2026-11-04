"""app.easyph and the placeholder swap in app.equipment, on synthetic and real
workbooks."""

import pathlib
from io import BytesIO

import pytest
from openpyxl import Workbook, load_workbook

from app.easyph import read_easyph
from app.equipment import _item, _resolve_easyph
from app.parser import parse_workbook

DOWNLOADS = pathlib.Path.home() / "Downloads"
MARK = "❌"


def _form(
    offset=0,
    cooling="simplified",
    heating="simplified",
    dhw="detailed",
    cooling_available=True,
    power_unit="kW",
    heating_rating_col="S",
):
    """easyPH's cooling, heat pump and DHW blocks as Holmes lays them out
    (rows 172-277), moved down by +offset+ rows."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("easyPH")

    def put(row, **cols):
        for col, value in cols.items():
            ws[f"{col}{row + offset}"] = value

    put(172, D=MARK, I="Cooling units")
    put(174, O="Mechanical cooling available?", P="x" if cooling_available else None)
    put(176, O="x" if cooling == "simplified" else None, P="Simplified efficiency verification")
    put(181, P="Manufacturer", Q="Carrier")
    put(182, P="Product name", Q="27VNA324A*030*")
    put(183, P="Nominal power", Q=6.68, R=power_unit)
    put(184, P="Efficiency value", Q="SEER", R="SEER(2)")
    put(187, P="Efficiency characteristic value", R=21.5)
    put(
        189,
        O="x" if cooling == "detailed" else None,
        P="ALTERNATIVELY: Detailed proof of efficiency",
    )
    put(210, D=MARK, I="Heat generator type")

    put(225, D=MARK, I="Heat pump")
    put(227, O="Space heating")
    put(229, O="x" if heating == "simplified" else None, P="Simplified efficiency verification")
    put(233, P="Manufacturer", Q="Carrier")
    put(234, P="Product name", Q="27VNA324A*030*")
    put(235, P="Nominal power", Q=7.39, R=power_unit)
    put(236, R="SCOP", S="HSPF(2)")
    put(239, P="Efficiency characteristic value", **{heating_rating_col: 10.5})
    put(
        242,
        O="x" if heating == "detailed" else None,
        P="ALTERNATIVELY: Detailed proof of efficiency",
    )

    put(253, O="DHW")
    put(256, O="x" if dhw == "simplified" else None, P="Simplified efficiency verification")
    put(260, P="Manufacturer", Q="Rheem")
    put(261, P="Product name", Q="ProTerra")
    put(262, P="Nominal power", Q=1.6, R=power_unit)
    put(263, Q="seasonal", R="SCOP", S="HSPF(2)")
    put(266, P="Efficiency characteristic value", R=3.1)
    put(269, O="x" if dhw == "detailed" else None, P="ALTERNATIVELY: Detailed proof of efficiency")
    put(275, D=MARK, I="Boiler")

    buf = BytesIO()
    wb.save(buf)
    return load_workbook(BytesIO(buf.getvalue()), data_only=True, read_only=True)


def _placeholder(equipment_type, name):
    return _item(
        equipment_type=equipment_type,
        name=name,
        capacity=13.07,
        capacity_unit="kW",
        efficiency_value=3.9,
        efficiency_type="COP",
        source="WP_Heizungssystem_Auswahl_Lueftung",
    )


@pytest.mark.parametrize("offset", [-2, 0, 7])
def test_reads_each_block_by_its_header(offset):
    forms = read_easyph(_form(offset=offset))
    assert forms["cooling"]["form"] == "simplified"
    assert forms["heating"]["form"] == "simplified"
    assert forms["dhw"] == {"form": "detailed", "unit": None}
    heating = forms["heating"]["unit"]
    assert {
        k: heating[k] for k in ("manufacturer", "product", "nominal_power", "nominal_power_unit")
    } == {
        "manufacturer": "Carrier",
        "product": "27VNA324A*030*",
        "nominal_power": 7.39,
        "nominal_power_unit": "kW",
    }
    assert (heating["efficiency_value"], heating["efficiency_metric"]) == (10.5, "HSPF2")
    assert heating["sources"]["product"] == f"easyPH!Q{234 + offset}"
    assert (
        forms["cooling"]["unit"]["efficiency_value"],
        forms["cooling"]["unit"]["efficiency_metric"],
    ) == (21.5, "SEER2")
    assert forms["cooling"]["unit"]["nominal_power"] == 6.68


def test_metric_is_the_header_over_the_filled_column():
    unit = read_easyph(_form(heating_rating_col="R"))["heating"]["unit"]
    assert (unit["efficiency_value"], unit["efficiency_metric"]) == (10.5, "SCOP")


def test_dhw_simplified_reads_its_own_block():
    unit = read_easyph(_form(dhw="simplified"))["dhw"]["unit"]
    assert (unit["manufacturer"], unit["product"], unit["efficiency_metric"]) == (
        "Rheem",
        "ProTerra",
        "SCOP",
    )


def test_cooling_not_available_reads_no_form():
    assert read_easyph(_form(cooling_available=False))["cooling"] == {"form": None, "unit": None}


def test_no_easyph_sheet_is_none():
    assert read_easyph(Workbook()) is None


def test_placeholder_becomes_the_form_unit():
    items = [
        _placeholder("heat_pump", "easyPH heating"),
        _placeholder("cooling_unit", "easyPH cooling"),
    ]
    heat_pump, cooling = _resolve_easyph(_form(), items)
    assert heat_pump == _item(
        equipment_type="heat_pump",
        name="27VNA324A*030*",
        manufacturer="Carrier",
        capacity=7.39,
        capacity_unit="kW",
        efficiency_value=10.5,
        efficiency_type="HSPF2",
        source="easyPH!Q234; easyPH simplified form",
    )
    assert (cooling["equipment_type"], cooling["name"], cooling["efficiency_type"]) == (
        "cooling_unit",
        "27VNA324A*030*",
        "SEER2",
    )
    # ps-rails' overlay keys on the selection cell: the form's cell is none of them.
    assert cooling["source"].split(";")[0] == "easyPH!Q182"


def test_ip_nominal_power_is_btu_per_hour():
    (heat_pump,) = _resolve_easyph(
        _form(power_unit="kBTU/hr"), [_placeholder("heat_pump", "easyPH heating")]
    )
    assert (heat_pump["capacity"], heat_pump["capacity_unit"]) == (7390.0, "Btu/h")


@pytest.mark.parametrize("form", [{"heating": "detailed"}, {"heating": None}])
def test_placeholder_without_a_simplified_form_is_dropped(form):
    kept = _item(
        equipment_type="hrv", name="Heat recovery unit", source="Lueftung_Auswahl_Lueftungsgeraet"
    )
    assert _resolve_easyph(_form(**form), [kept, _placeholder("heat_pump", "easyPH heating")]) == [
        kept
    ]


def test_placeholder_in_a_workbook_with_no_form_is_dropped():
    assert _resolve_easyph(Workbook(), [_placeholder("heat_pump", "easyPH heating")]) == []


def test_real_units_pass_through():
    real = _placeholder("heat_pump", "Standard air/water heat pump")
    assert _resolve_easyph(_form(), [real]) == [real]


def _real(file, version="EN_10_6"):
    path = DOWNLOADS / file
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    return parse_workbook(path.read_bytes(), version)["hvac_equipment"]


def _shown(units):
    return {
        (
            u["equipment_type"],
            u["name"],
            u["manufacturer"],
            u["capacity"],
            u["efficiency_value"],
            u["efficiency_type"],
        )
        for u in units
    }


def test_holmes():
    # The staging extract's Claude read (project 26): Carrier 27VNA324A*030* at
    # 7.386 kW, HSPF(2) 10.5, for heating; the same unit cools.
    assert _shown(_real("2536 Holmes_easyPH_260512.xlsx")) == {
        ("erv", "Brink Climate Systems B.V. - Brink Flair 400 Enthalpie", None, None, None, None),
        ("cooling_unit", "27VNA324A*030*", "Carrier", 6.682297772567409, 21.5, "SEER2"),
        ("heat_pump", "27VNA324A*030*", "Carrier", 7.38569753810082, 10.5, "HSPF2"),
        ("dhw_heat_pump", "HPWH", None, 1.5997656021095812, 2.840538836965976, "COP"),
    }


def test_easyph_example_uses_components_units():
    assert _shown(_real("PHPP_EN_V10.6_easyPH_Example.xlsx")) == {
        ("hrv", "Heat recovery unit", None, None, None, None),
        ("heat_pump", "Standard air/water heat pump", None, 5.0, 4.3, "COP"),
        ("dhw_heat_pump", "Standard air/water heat pump", None, 5.0, 4.3, "COP"),
    }
