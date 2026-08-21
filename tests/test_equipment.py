import pytest
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from app.equipment import _classify_ventilation, _strip_prefix, read_equipment


def _wb(names: dict[str, object], sheet: str = "Ventilation SI"):
    """Build a workbook whose defined names point at real cells."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for i, (name, value) in enumerate(names.items(), start=2):
        ws.cell(row=i, column=1, value=value)
        wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!$A${i}"))
    return wb


class TestStripPrefix:
    # PHPP dropdown members carry an id prefix. The same prefix format already
    # bit the easyPH writer, so it is parsed rather than assumed away.
    @pytest.mark.parametrize(
        "raw,expected",
        [
            ("01ud-Swegon Casa R7 Genius Sorption", "Swegon Casa R7 Genius Sorption"),
            ("1363vs03-Brink Climate Systems B.V. - Brink Flair",
             "Brink Climate Systems B.V. - Brink Flair"),
            ("99ud-Standard air-to-air heat pump", "Standard air-to-air heat pump"),
            ("no prefix here", "no prefix here"),
            ("", None),
            (None, None),
        ],
    )
    def test_strips_only_the_leading_id(self, raw, expected):
        assert _strip_prefix(raw) == expected

    def test_keeps_internal_hyphens(self):
        # "Brink Climate Systems B.V. - Brink Flair" contains a hyphen that is
        # part of the name; splitting on the last one would truncate it.
        assert _strip_prefix("1363vs03-A - B") == "A - B"


class TestVentilationClassification:
    # Humidity recovery is the hrv/erv discriminator. Both units in the real
    # corpus recover moisture, so the hrv branch needs a synthetic case or it
    # is never exercised.
    def test_moisture_recovery_means_erv(self):
        assert _classify_ventilation(86.0) == "erv"

    def test_no_moisture_recovery_means_hrv(self):
        assert _classify_ventilation(0.0) == "hrv"

    def test_unknown_moisture_recovery_means_hrv(self):
        assert _classify_ventilation(None) == "hrv"


class TestDiscovery:
    def test_emits_nothing_for_a_non_v106_version(self):
        wb = _wb({"Lueftung_Auswahl_Lueftungsgeraet": "01ud-Unit"})
        assert read_equipment(wb, "EN_9_7IP") == []

    def test_emits_no_ventilation_device_when_none_is_selected(self):
        # The phantom-device guard: a file has a row per family whether or not
        # the building uses one.
        assert read_equipment(_wb({}), "EN_10_6") == []

    def test_emits_a_ventilation_device_when_one_is_selected(self):
        wb = _wb({"Lueftung_Auswahl_Lueftungsgeraet": "01ud-Swegon Casa R7"})
        items = read_equipment(wb, "EN_10_6")
        assert [i["name"] for i in items] == ["Swegon Casa R7"]

    def test_cooling_device_requires_its_presence_flag(self):
        selected = {"Kuehlgeraete_Kompressor_Umluft_Geraet": "01ud-Daikin 4MXTH36AVJU9"}
        assert read_equipment(_wb(selected), "EN_10_6") == []

        flagged = dict(selected, Kuehlgeraete_Umluft_Kuehlung_Ankreuzen="x")
        items = read_equipment(_wb(flagged), "EN_10_6")
        assert [i["equipment_type"] for i in items] == ["cooling_unit"]

    def test_every_emitted_item_uses_only_schema_keys(self):
        wb = _wb({"Lueftung_Auswahl_Lueftungsgeraet": "01ud-Swegon Casa R7"})
        allowed = {
            "equipment_type", "name", "manufacturer", "capacity", "capacity_unit",
            "efficiency_value", "efficiency_type", "airflow_cfm", "airflow_m3h",
            "heat_recovery_efficiency_pct", "source",
        }
        for item in read_equipment(wb, "EN_10_6"):
            assert set(item) <= allowed, set(item) - allowed
            assert item["equipment_type"] is not None
