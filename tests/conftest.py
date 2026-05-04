"""Shared test fixtures."""

from io import BytesIO

import pytest
from openpyxl import Workbook
from PHX.PHPP.phpp_localization import shape_model

from app.parser import load_shape

# Synthetic ground-truth values for the KPI fixture. Real-file ground truth
# lives in tests/test_kpis_real_workbook.py and is sampled from
# ~/Downloads/PHPP_V10.6_IP_Example.xlsx.
KPI_FIXTURE_VALUES = {
    "tfa": 2400.0,
    "heating_demand_specific": 4.2,
    "cooling_demand_specific": 2.1,
    "heating_peak_w1": 22000.0,
    "heating_peak_w2": 24000.0,  # max should win
    "cooling_peak_w1": 18000.0,
    "cooling_peak_w2": 18000.0,
    "per_total_v": 7.3,  # PER source EUI
    "per_total_x": 15.2,  # PE demand
}


@pytest.fixture
def shape_en_10_6ip() -> shape_model.PhppShape:
    return load_shape("EN_10_6IP")


def build_workbook_bytes(
    shape: shape_model.PhppShape,
    num_of_units_value: int | float | str | None = 42,
    data_sheet_version: str | None = None,
    data_sheet_pe_factor: str | None = "1-PE-factors (non-renewable) PHI Certification",
    with_kpis: bool = False,
    cooling_peak_value: float | None = None,
) -> bytes:
    """Build a minimal in-memory PHPP-shaped workbook for tests.

    Always writes:
      - A sheet named after shape.VERIFICATION.name
      - shape.VERIFICATION.num_of_units locator string + value cell

    Optionally writes a Data sheet for version-detection tests. Pass
    data_sheet_version='10.6 IP' (or similar) to enable it.

    Optionally writes the four KPI sheets (Heating, Cooling, Heating load,
    Cooling load, PER) with values from KPI_FIXTURE_VALUES — pass
    with_kpis=True. Pass cooling_peak_value to override the cooling peak
    (e.g. -50 to test the negative-sign gotcha).
    """
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(shape.VERIFICATION.name)
    item = shape.VERIFICATION.num_of_units
    locator_row = 30
    ws[f"{item.locator_col}{locator_row}"] = item.locator_string
    value_row = locator_row + item.input_row_offset
    if num_of_units_value is not None:
        ws[f"{item.input_column}{value_row}"] = num_of_units_value

    if data_sheet_version is not None:
        data_ws = wb.create_sheet("Data")
        data_ws["A5"] = "PHPP Version"
        data_ws["B5"] = data_sheet_version
        if data_sheet_pe_factor:
            data_ws["C5"] = "Language"
            data_ws["D5"] = data_sheet_pe_factor

    if with_kpis:
        # Heating sheet: annual demand at col_kWh_m2_year x row_annual_demand
        hd = shape.HEATING_DEMAND
        h_ws = wb.create_sheet(hd.name)
        h_ws[f"{hd.col_kWh_m2_year}{hd.row_annual_demand}"] = (
            KPI_FIXTURE_VALUES["heating_demand_specific"]
        )

        # Cooling sheet — TFA at address_tfa, sensible demand
        cd = shape.COOLING_DEMAND
        c_ws = wb.create_sheet(cd.name)
        c_ws[cd.address_tfa] = KPI_FIXTURE_VALUES["tfa"]
        c_ws[f"{cd.col_kWh_m2_year}{cd.row_annual_sensible_demand}"] = (
            KPI_FIXTURE_VALUES["cooling_demand_specific"]
        )

        # Heating load sheet — w1 + w2 at row_total_load
        hp = shape.HEATING_PEAK_LOAD
        hp_ws = wb.create_sheet(hp.name)
        hp_ws[f"{hp.col_weather_1}{hp.row_total_load}"] = KPI_FIXTURE_VALUES["heating_peak_w1"]
        hp_ws[f"{hp.col_weather_2}{hp.row_total_load}"] = KPI_FIXTURE_VALUES["heating_peak_w2"]

        # Cooling load sheet
        cp = shape.COOLING_PEAK_LOAD
        cp_ws = wb.create_sheet(cp.name)
        cp_w1 = (
            cooling_peak_value
            if cooling_peak_value is not None
            else KPI_FIXTURE_VALUES["cooling_peak_w1"]
        )
        cp_w2 = (
            cooling_peak_value
            if cooling_peak_value is not None
            else KPI_FIXTURE_VALUES["cooling_peak_w2"]
        )
        cp_ws[f"{cp.col_weather_1}{cp.row_total_sensible_load}"] = cp_w1
        cp_ws[f"{cp.col_weather_2}{cp.row_total_sensible_load}"] = cp_w2

        # PER sheet — locator-string-driven row scan looks for "Total energy demand"
        per = shape.PER
        per_ws = wb.create_sheet(per.name)
        total_row = 68  # arbitrary; the parser finds it by string
        per_ws[f"{per.locator_col}{total_row}"] = "Total energy demand kBTU/(ft²yr)"
        per_ws[f"{per.columns.per_energy}{total_row}"] = KPI_FIXTURE_VALUES["per_total_v"]
        per_ws[f"{per.columns.pe_energy}{total_row}"] = KPI_FIXTURE_VALUES["per_total_x"]

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture
def workbook_bytes(shape_en_10_6ip) -> bytes:
    return build_workbook_bytes(shape_en_10_6ip, num_of_units_value=42)


@pytest.fixture
def workbook_bytes_with_version(shape_en_10_6ip) -> bytes:
    return build_workbook_bytes(
        shape_en_10_6ip,
        num_of_units_value=42,
        data_sheet_version="10.6 IP",
    )


@pytest.fixture
def workbook_bytes_with_kpis(shape_en_10_6ip) -> bytes:
    return build_workbook_bytes(shape_en_10_6ip, with_kpis=True)
