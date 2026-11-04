"""app.equipment_v97 on synthetic 9.7 layouts and on two real 9.7 IP workbooks."""

import pathlib

import pytest
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from app.equipment import read_equipment
from app.parser import parse_workbook

DOWNLOADS = pathlib.Path.home() / "Downloads"


def _name(wb, name, ref):
    wb.defined_names.add(DefinedName(name, attr_text=ref))


def _v97(
    *,
    heating="2-Mini split",
    dhw="1-Tank HP",
    count=2,
    heating_share=1,
    dhw_share=1,
    slot_names=("Tank HP", "Mini split"),
    cooling_box="x",
    hr_header="Unit hHR",
):
    """The cells app.equipment_v97 reads, at 9.7's own places."""
    wb = Workbook()
    vent = wb.active
    vent.title = "Ventilation SI"
    hp = wb.create_sheet("HP SI")
    cool = wb.create_sheet("Cooling units SI")

    vent["N87"], vent["K88"], vent["N88"], vent["O88"] = hr_header, "01ud-Brink Flair 325", 0.91, 0
    vent["N65"] = 292.13
    _name(wb, "Lueftung_Auswahl_Lueftungsgeraet", "'Ventilation SI'!$K$88")
    _name(wb, "Lueftung_Auslegungsvolumenstrom", "'Ventilation SI'!$N$65")

    hp["I9"], hp["M9"] = "Covered fraction of space heating demand", heating_share
    hp["I14"], hp["M14"] = "Covered fraction of DHW demand", dhw_share
    hp["M18"], hp["J21"], hp["J36"] = count, heating, dhw
    _name(wb, "WP_Anzahl", "'HP SI'!$M$18")
    _name(wb, "WP_Heizung_Art_Waermepumpe", "'HP SI'!$J$21")
    _name(wb, "WP_Warmwasser_Art_Waermepumpe", "'HP SI'!$J$36")
    tables = {
        1: [(0, 60, 0, 1), (20, 60, 1.6, 3.3), (None, None, None, None), (15, 60, 1.2, 3.1)],
        2: [(-15, 35, 2.23, 2.47), (8.33, 35, 4.81, 3.5), (-8.3, 39, 2.64, 2.36)],
    }
    for slot, rows in tables.items():
        top = 500 + 30 * slot
        hp[f"I{top}"] = slot_names[slot - 1]
        _name(wb, f"WP_WP{slot}_Name", f"'HP SI'!$I${top}:$J${top}")
        _name(wb, f"WP_WP{slot}_Leistungsdaten", f"'HP SI'!$K${top + 5}:$N${top + 19}")
        for n, values in enumerate(rows):
            for col, value in zip("KLMN", values, strict=True):
                hp[f"{col}{top + 5 + n}"] = value

    cool["I22"], cool["P25"], cool["P26"], cool["P29"] = cooling_box, 7.0337, 1189.3, 1.46536
    for name, cell in (
        ("Kuehlgeraete_Umluft_Kuehlung_Ankreuzen", "I22"),
        ("Kuehlgeraete_Umluft_Kuehlung_Leistung", "P25"),
        ("Kuehlgeraete_Umluft_Kuehlung_Volumenstrom", "P26"),
        ("Kuehlgeraete_Umluft_Kuehlung_JAZ", "P29"),
    ):
        _name(wb, name, f"'Cooling units SI'!${cell[0]}${cell[1:]}")
    return wb


def _by_type(items):
    return {i["equipment_type"]: i for i in items}


def test_reads_every_unit():
    items = _by_type(read_equipment(_v97(), "EN_9_7IP"))
    assert set(items) == {"hrv", "cooling_unit", "heat_pump", "dhw_heat_pump"}
    vent = items["hrv"]
    assert (vent["name"], vent["heat_recovery_efficiency_pct"]) == ("Brink Flair 325", 91.0)
    assert vent["airflow_cfm"] == pytest.approx(171.94, abs=0.01)
    hp = items["heat_pump"]
    assert (hp["name"], hp["capacity"], hp["efficiency_value"]) == ("Mini split", 4.81, 3.5)
    assert hp["source"].startswith("WP_Heizung_Art_Waermepumpe -> WP_WP2; test point 2")
    cool = items["cooling_unit"]
    assert cool["name"] == "Recirculation cooling"
    assert (cool["capacity"], cool["capacity_unit"]) == (pytest.approx(7.0337), "kW")
    assert cool["efficiency_type"] == "SEER"
    assert cool["efficiency_value"] == pytest.approx(5.0, abs=0.01)


def test_a_blank_test_point_does_not_shift_the_rest():
    # resolve() drops empty cells; a flat read would pair 15 °C with the wrong COP.
    dhw = _by_type(read_equipment(_v97(), "EN_9_7IP"))["dhw_heat_pump"]
    assert (dhw["capacity"], dhw["efficiency_value"]) == (1.6, 3.3)


def test_moisture_recovery_reads_erv():
    wb = _v97()
    wb["Ventilation SI"]["O88"] = 0.71
    assert _by_type(read_equipment(wb, "EN_9_7IP"))["erv"]["heat_recovery_efficiency_pct"] == 91.0


def test_no_recovery_without_its_header():
    vent = _by_type(read_equipment(_v97(hr_header="Something else"), "EN_9_7IP"))["hrv"]
    assert vent["heat_recovery_efficiency_pct"] is None


@pytest.mark.parametrize("kwargs,absent", [
    ({"heating_share": 0}, "heat_pump"),
    ({"dhw_share": 0}, "dhw_heat_pump"),
    ({"count": 1}, "dhw_heat_pump"),
    ({"slot_names": ("Tank HP", "Heat pump, type 2")}, "heat_pump"),
    ({"heating": "Mini split"}, "heat_pump"),
    ({"cooling_box": None}, "cooling_unit"),
])
def test_a_unit_is_emitted_only_when_in_use_and_found(kwargs, absent):
    assert absent not in _by_type(read_equipment(_v97(**kwargs), "EN_9_7IP"))


# Read by hand off each file's 'HP SI', 'Ventilation SI' and 'Cooling units' panes.
# Names, heat recovery, airflow, cooling capacity and SEER match the Claude rows
# stored on staging for the same files (artifacts 8 and 74).
REAL = {
    "55th": ("250121_55th_IP9_741zhOb_FINAL.xlsx", {
        "hrv": ("Brink Flair 325", None, None, 91.0, 171.94),
        "cooling_unit": ("Recirculation cooling", 7.03, 5.0, None, 700.0),
        "heat_pump": ("Ducted Mini-split+Ducted Heat Pump", 4.81, 3.5, None, None),
        "dhw_heat_pump": ("Rheem PROTERRA65", 1.6, 3.3, None, None),
    }),
    # The Mitsubishi's test points carry COPs of 6-12.4 on both panes, and PHPP
    # calculates with them as typed.
    "17_mile": ("250901.2821 17 mile drive.IP9.7.final.xlsx", {
        "hrv": ("Brink Flair 400", None, None, 89.0, 175.72),
        "cooling_unit": ("Recirculation cooling", 2.64, 18.59, None, 350.0),
        "heat_pump": ("Mitsubishi SUZ-KA09", 3.98, 12.4, None, None),
        "dhw_heat_pump": ("Sanden GS4-45HPC", 5.85, 3.04, None, None),
    }),
}


@pytest.mark.parametrize("key", sorted(REAL))
def test_real_workbook(key):
    file, expected = REAL[key]
    path = DOWNLOADS / file
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    got = _by_type(parse_workbook(path.read_bytes(), "EN_9_7IP")["hvac_equipment"])
    assert set(got) == set(expected)
    for kind, (name, capacity, efficiency, recovery, cfm) in expected.items():
        item = got[kind]
        assert item["name"] == name, kind
        for field, value in (("capacity", capacity), ("efficiency_value", efficiency),
                             ("heat_recovery_efficiency_pct", recovery), ("airflow_cfm", cfm)):
            if value is None:
                continue
            assert item[field] == pytest.approx(value, abs=0.01), (kind, field)
