"""Local shape corrections applied over PHX's bundled shapes.

The EN_9_7IP entry is the reason this mechanism exists; see
`app/shape_overrides.py` for what is wrong with it upstream.
"""

import pytest

from app.parser import available_versions, load_shape
from app.shape_overrides import SHAPE_OVERRIDES, apply_overrides


class TestApplyOverrides:
    def test_untouched_version_passes_through(self):
        data = {"HEATING_DEMAND": {"name": "Heating", "unit": "KBTU"}}
        assert apply_overrides("EN_10_6IP", data) == data

    @staticmethod
    def _si_shaped_input():
        """EN_9_7IP as PHX ships it, reduced to the sections the override
        targets."""
        return {
            "HEATING_DEMAND": {
                "name": "Heating SI",
                "unit": "KWH",
                "col_kWh_m2_year": "Q",
                "row_annual_demand": 77,
            },
            "COOLING_DEMAND": {"name": "Cooling SI", "unit": "KWH"},
            "HEATING_PEAK_LOAD": {"name": "Heating load SI", "unit": "W"},
            "COOLING_PEAK_LOAD": {"name": "Cooling load SI", "unit": "W"},
        }

    def test_returns_a_copy_rather_than_mutating_the_input(self):
        data = self._si_shaped_input()
        apply_overrides("EN_9_7IP", data)
        assert data["HEATING_DEMAND"]["name"] == "Heating SI"
        assert data["HEATING_PEAK_LOAD"]["unit"] == "W"

    def test_merges_into_a_section_without_dropping_its_other_fields(self):
        data = self._si_shaped_input()
        patched = apply_overrides("EN_9_7IP", data)

        assert patched["HEATING_DEMAND"]["name"] == "Heating"
        assert patched["HEATING_DEMAND"]["unit"] == "KBTU"
        # The cell coordinates are what make the repoint safe — losing them
        # would silently move every read.
        assert patched["HEATING_DEMAND"]["col_kWh_m2_year"] == "Q"
        assert patched["HEATING_DEMAND"]["row_annual_demand"] == 77

    def test_raises_when_the_targeted_section_no_longer_exists(self):
        """An override that has drifted from upstream is a bug, and a loud one
        beats silently inventing a section PHX has dropped."""
        with pytest.raises(KeyError, match="EN_9_7IP"):
            apply_overrides("EN_9_7IP", {"COOLING_DEMAND": {}})


@pytest.fixture(scope="module")
def shape():
    return load_shape("EN_9_7IP")


class TestEn97IpShape:
    """The loaded shape, after the override, reads the IP panes."""

    @pytest.mark.parametrize(
        "section,expected_name,expected_unit",
        [
            ("HEATING_DEMAND", "Heating", "KBTU"),
            ("COOLING_DEMAND", "Cooling", "KBTU"),
            ("HEATING_PEAK_LOAD", "Heating load", "BTU/HR"),
            ("COOLING_PEAK_LOAD", "Cooling load", "BTU/HR"),
        ],
    )
    def test_energy_sections_point_at_ip_sheets(
        self, shape, section, expected_name, expected_unit
    ):
        sub = getattr(shape, section)
        assert sub.name == expected_name
        assert sub.unit == expected_unit

    @pytest.mark.parametrize(
        "section", ["PER", "OVERVIEW"]
    )
    def test_deliberately_unoverridden_sections_keep_their_si_pane(
        self, shape, section
    ):
        """PER's locator matches no row on v9.7 regardless of pane, and
        OVERVIEW is self-consistent with SI units. Both are left alone on
        purpose — this pins that as a decision, not an oversight."""
        assert getattr(shape, section).name.endswith(" SI")

    def test_sibling_ip_shape_is_untouched(self):
        """EN_10_6IP already points at the IP sheets, so nothing should be
        overriding it — if this fails, the override is too broad."""
        assert "EN_10_6IP" not in SHAPE_OVERRIDES
        shape = load_shape("EN_10_6IP")
        assert shape.HEATING_DEMAND.name == "Heating"
        assert shape.HEATING_DEMAND.unit == "KBTU"

    def test_si_shape_is_untouched(self):
        """EN_10_6 is a genuine SI shape; its SI sheets are correct."""
        shape = load_shape("EN_10_6")
        assert shape.HEATING_DEMAND.unit.upper() == "KWH"


# Sections whose sheet name would reveal the same SI-pointing defect.
_PANE_SECTIONS = (
    "HEATING_DEMAND",
    "COOLING_DEMAND",
    "HEATING_PEAK_LOAD",
    "COOLING_PEAK_LOAD",
    "PER",
    "OVERVIEW",
    "AREAS",
    "UVALUES",
    "VENTILATION",
    "VERIFICATION",
)


def test_en_9_7ip_is_the_only_shape_needing_correction():
    """Sweep every shape PHX ships, not just the one sibling.

    The claim this fix rests on is that the SI-pointing is confined to
    EN_9_7IP. That was originally checked against EN_10_6IP alone, which
    skipped EN_10_4IP — the other IP shape. Asserting it across the whole set
    is what makes the claim hold, and it fails loudly if a PHX upgrade
    introduces the same defect elsewhere or fixes this one upstream.
    """
    affected = {
        version: [
            section
            for section in _PANE_SECTIONS
            # the raw upstream name, before any override
            if str(getattr(load_shape(version), section).name).endswith(" SI")
        ]
        for version in available_versions()
    }
    # EN_9_7IP's four consumed sections are corrected; PER and OVERVIEW are
    # deliberately left on their SI panes (see shape_overrides).
    assert affected["EN_9_7IP"] == ["PER", "OVERVIEW"]
    others = {v: s for v, s in affected.items() if v != "EN_9_7IP" and s}
    assert others == {}, f"another shape points at SI panes: {others}"
