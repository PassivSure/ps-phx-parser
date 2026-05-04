"""Unit tests for app.kpis on synthetic workbooks."""

from app.kpis import read_kpis
from app.parser import load_shape
from tests.conftest import KPI_FIXTURE_VALUES, build_workbook_bytes


def _kpis(workbook_bytes: bytes) -> dict:
    from io import BytesIO

    from openpyxl import load_workbook

    wb = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        return read_kpis(wb, load_shape("EN_10_6IP"))
    finally:
        wb.close()


def test_tfa_read_from_cooling_sheet(workbook_bytes_with_kpis):
    kpis = _kpis(workbook_bytes_with_kpis)
    assert kpis["tfa"]["value"] == KPI_FIXTURE_VALUES["tfa"]
    assert kpis["tfa"]["unit"] == "ft2"
    assert kpis["tfa"]["source"] == "Cooling!O8"


def test_specific_demands_read_from_demand_sheets(workbook_bytes_with_kpis):
    kpis = _kpis(workbook_bytes_with_kpis)
    assert kpis["heating_demand"]["value"] == KPI_FIXTURE_VALUES["heating_demand_specific"]
    assert kpis["heating_demand"]["unit"] == "kBtu/ft2yr"
    assert kpis["cooling_demand"]["value"] == KPI_FIXTURE_VALUES["cooling_demand_specific"]
    assert kpis["cooling_demand"]["unit"] == "kBtu/ft2yr"


def test_peak_heating_returns_max_of_two_weather_scenarios(workbook_bytes_with_kpis):
    """w1=22000, w2=24000 → design peak is the worse case (max=24000)."""
    kpis = _kpis(workbook_bytes_with_kpis)
    assert kpis["peak_loads"]["heating"]["value"] == KPI_FIXTURE_VALUES["heating_peak_w2"]
    assert kpis["peak_loads"]["heating"]["unit"] == "Btu/h"


def test_peak_cooling_negative_value_returns_null(shape_en_10_6ip):
    """Cooling-load sign gotcha: negatives are heat-loss numbers, not real
    cooling demand. Return None instead of leaking a misleading sign."""
    workbook = build_workbook_bytes(
        shape_en_10_6ip, with_kpis=True, cooling_peak_value=-500.0
    )
    kpis = _kpis(workbook)
    assert kpis["peak_loads"]["cooling"]["value"] is None


def test_per_totals_found_by_locator_string(workbook_bytes_with_kpis):
    """PER row 68 is found by `Total energy demand` prefix match in column P,
    not by hardcoded row — the table shifts when users add heating types."""
    kpis = _kpis(workbook_bytes_with_kpis)
    assert kpis["source_eui"]["value"] == KPI_FIXTURE_VALUES["per_total_v"]
    assert kpis["source_eui"]["unit"] == "kBtu/ft2yr"
    assert kpis["pe_demand"]["value"] == KPI_FIXTURE_VALUES["per_total_x"]
    assert kpis["pe_demand"]["unit"] == "kBtu/ft2yr"


def test_per_totals_locator_works_when_row_shifts(shape_en_10_6ip):
    """Author the PER total at a non-default row to prove we don't depend
    on row 68."""
    from io import BytesIO

    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    per = shape_en_10_6ip.PER
    ws = wb.create_sheet(per.name)
    custom_row = 95
    ws[f"{per.locator_col}{custom_row}"] = "Total energy demand kBTU/(ft²yr)"
    ws[f"{per.columns.per_energy}{custom_row}"] = 11.1
    ws[f"{per.columns.pe_energy}{custom_row}"] = 22.2

    # Also need a Verification sheet so parse_workbook doesn't choke
    # downstream — read_kpis itself doesn't, but cleaner this way.
    wb.create_sheet(shape_en_10_6ip.VERIFICATION.name)

    buf = BytesIO()
    wb.save(buf)

    kpis = _kpis(buf.getvalue())
    assert kpis["source_eui"]["value"] == 11.1
    assert kpis["pe_demand"]["value"] == 22.2
    assert kpis["source_eui"]["source"] == f"PER!V{custom_row}"
