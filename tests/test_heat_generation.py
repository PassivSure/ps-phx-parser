"""app.heat_generation on synthetic and real workbooks."""

import pathlib
from io import BytesIO

import pytest
from openpyxl import Workbook, load_workbook

from app.heat_generation import read_heat_generation
from app.parser import load_shape, parse_workbook

DOWNLOADS = pathlib.Path.home() / "Downloads"


def _per_workbook(shape, blocks, unit="BTU/hrW", efficiencies=None):
    """A PER sheet with the efficiency unit row, then each block's heading and
    its rows. +efficiencies+ maps (block, label) to (calculated, user)."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(shape.PER.name)
    per = shape.PER
    col, energy = per.locator_col, per.columns.final_energy
    calc, user = per.columns.calculated_efficiency, per.columns.user_determined_efficiency
    if unit is not None:
        ws[f"{calc}3"] = unit
        ws[f"{user}3"] = unit
    efficiencies = efficiencies or {}
    row = 5
    for name, rows in blocks.items():
        ws[f"{col}{row}"] = getattr(per, name).locator_string_heading
        for i, (label, value) in enumerate(rows, start=1):
            ws[f"{col}{row + i}"] = label
            if value is not None:
                ws[f"{energy}{row + i}"] = value
            q, r = efficiencies.get((name, label), (None, None))
            if q is not None:
                ws[f"{calc}{row + i}"] = q
            if r is not None:
                ws[f"{user}{row + i}"] = r
        row += len(rows) + 4
    buf = BytesIO()
    wb.save(buf)
    return load_workbook(BytesIO(buf.getvalue()), data_only=True, read_only=True)


def test_reads_each_block_between_its_locators():
    shape = load_shape("EN_10_6IP")
    wb = _per_workbook(
        shape,
        {
            "heating": [
                ("Electricity (HP compact unit)", None),
                ("Electricity (heat pump)", 5.2),
                ("Boiler", 0),
                ("Solar thermal system", None),
            ],
            "dhw": [
                ("Electricity (HP compact unit)", None),
                ("Boiler", 3.1),
                ("Solar thermal system", None),
            ],
            "cooling": [
                ("Electricity cooling (HP)", 1.5),
                ("Electricity dehumidification (HP)", None),
            ],
        },
    )
    hg = read_heat_generation(wb, shape)
    assert [(r["label"], r["final_energy"]) for r in hg["heating"]] == [
        ("Electricity (HP compact unit)", None),
        ("Electricity (heat pump)", 5.2),
        ("Boiler", 0.0),
        ("Solar thermal system", None),
    ]
    assert [(r["label"], r["final_energy"]) for r in hg["dhw"]][1] == ("Boiler", 3.1)
    assert hg["cooling"][0] == {
        "label": "Electricity cooling (HP)",
        "final_energy": 1.5,
        "source": "PER!T21",
        "calculated_efficiency": None,
        "user_efficiency": None,
        "calculated_efficiency_source": "PER!Q21",
        "user_efficiency_source": "PER!R21",
    }


DHW_DIRECT = {
    "dhw": [
        ("Electricity (HP compact unit)", None),
        ("Electricity (heat pump)", None),
        ("Electricity (direct)", 0.9),
        ("Solar thermal system", None),
    ]
}


def test_ip_efficiencies_are_converted_to_w_per_w():
    """10.x IP states efficiency in BTU/hrW: PHPP's direct electricity (1 W/W)
    reads 3.412..., and a COP of 2.84 typed from a separate HPWH workbook
    reads 9.69."""
    shape = load_shape("EN_10_6IP")
    wb = _per_workbook(
        shape,
        DHW_DIRECT,
        efficiencies={("dhw", "Electricity (direct)"): (3.412141156488376, 2.84 * 3.412141156488376)},
    )
    row = read_heat_generation(wb, shape)["dhw"][2]
    assert row["label"] == "Electricity (direct)"
    assert row["calculated_efficiency"] == pytest.approx(1.0, abs=1e-12)
    assert row["user_efficiency"] == pytest.approx(2.84)
    assert row["user_efficiency_source"] == "PER!R8"


def test_unitless_efficiencies_are_already_w_per_w():
    shape = load_shape("EN_9_7IP")
    wb = _per_workbook(
        shape,
        {"dhw": [(shape.PER.dhw.locator_string_start, None), ("Electricity (direct)", 0.9),
                 (shape.PER.dhw.locator_string_end, None)]},
        unit="-",
        efficiencies={("dhw", "Electricity (direct)"): (1, 2.5)},
    )
    row = read_heat_generation(wb, shape)["dhw"][1]
    assert (row["calculated_efficiency"], row["user_efficiency"]) == (1.0, 2.5)


@pytest.mark.parametrize("unit", [None, "kWh/kWh?"])
def test_unknown_efficiency_unit_reads_no_efficiency(unit):
    """A missing or unrecognised unit row leaves both efficiencies null: a
    BTU/hrW figure read as W/W would triple a COP."""
    shape = load_shape("EN_10_6IP")
    wb = _per_workbook(
        shape, DHW_DIRECT, unit=unit, efficiencies={("dhw", "Electricity (direct)"): (3.41, 9.7)}
    )
    row = read_heat_generation(wb, shape)["dhw"][2]
    assert (row["calculated_efficiency"], row["user_efficiency"]) == (None, None)
    assert row["final_energy"] == 0.9


def test_block_without_its_end_row_is_not_read():
    """A block that runs off without its end label is null, never a partial
    list: a consumer must not read "nothing else is in use" from it."""
    shape = load_shape("EN_10_6IP")
    wb = _per_workbook(
        shape, {"heating": [("Electricity (HP compact unit)", None), ("Boiler", 2.0)]}
    )
    hg = read_heat_generation(wb, shape)
    assert hg["heating"] is None
    assert hg["dhw"] is None and hg["cooling"] is None


def test_no_per_sheet_is_none():
    wb = Workbook()
    assert read_heat_generation(wb, load_shape("EN_10_6IP")) is None


def _used(rows):
    return {r["label"] for r in rows if (r["final_energy"] or 0) > 0}


@pytest.mark.parametrize(
    "file,version,heating,dhw,cooling",
    [
        (
            "260309 PHPP 6840 E ^th Ave Pkwy v 10.6.xlsx",
            "EN_10_6IP",
            {"Electricity (heat pump)"},
            {"Electricity (heat pump)"},
            {"Electricity cooling (HP)"},
        ),
        (
            "PHPP_EN_V10.6_easyPH_Example.xlsx",
            "EN_10_6",
            {"Electricity (heat pump)"},
            {"Electricity (heat pump)", "Solar thermal system"},
            set(),
        ),
        (
            "250121_55th_IP9_741zhOb_FINAL.xlsx",
            "EN_9_7IP",
            {"Electricity (heat pump)", "Aux. electricity (heating, wintertime ventilation)"},
            {"Electricity (heat pump)", "Aux. electricity (DHW + solar DHW)"},
            {
                "Electricity cooling (heat pump)",
                "Auxiliary electricity cooling, ventilation summer",
            },
        ),
    ],
)
def test_real_workbooks(file, version, heating, dhw, cooling):
    path = DOWNLOADS / file
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    hg = parse_workbook(path.read_bytes(), version)["heat_generation"]
    assert (_used(hg["heating"]), _used(hg["dhw"]), _used(hg["cooling"])) == (heating, dhw, cooling)


def _real_hg(file, version):
    path = DOWNLOADS / file
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    return parse_workbook(path.read_bytes(), version)["heat_generation"]


def _row(rows, label):
    (row,) = [r for r in rows if r["label"] == label]
    return row


def test_seventeen_mile_drive_manual_heating_cop():
    """9.7 IP, "Manual value" 6.1 on the heating heat pump row (PER!R22);
    the DHW heat pump keeps PHPP's own calculated 2.68."""
    hg = _real_hg("250901.2821 17 mile drive.IP9.7.final.xlsx", "EN_9_7IP")
    heating = _row(hg["heating"], "Electricity (heat pump)")
    assert heating["user_efficiency"] == 6.1
    assert heating["user_efficiency_source"] == "PER!R22"
    dhw = _row(hg["dhw"], "Electricity (heat pump)")
    assert dhw["user_efficiency"] is None
    assert dhw["calculated_efficiency"] == pytest.approx(2.6816217433765805)


def test_direct_electric_reference_reads_one_w_per_w():
    hg = _real_hg("PHPP_V10.6_IP_Example_dhw_direct.xlsx", "EN_10_6IP")
    direct = _row(hg["dhw"], "Electricity (direct)")
    assert direct["calculated_efficiency"] == pytest.approx(1.0, abs=1e-12)
    assert direct["user_efficiency"] is None
    assert direct["calculated_efficiency_source"] == "PER!Q44"


def test_user_defined_hpwh_cop_on_the_direct_row():
    """The direct-electric reference with a COP of 2.84 typed into PER!R44 in
    Excel (9.6905 BTU/hrW): PHPP's DHW final energy fell from 0.928 to 0.327
    kBTU/(ft²yr), 1/2.84 of it."""
    hg = _real_hg("PHPP_V10.6_IP_Example_dhw_direct_user_cop.xlsx", "EN_10_6IP")
    direct = _row(hg["dhw"], "Electricity (direct)")
    assert direct["calculated_efficiency"] == pytest.approx(1.0, abs=1e-12)
    assert direct["user_efficiency"] == pytest.approx(2.84, abs=1e-4)
    assert direct["user_efficiency_source"] == "PER!R44"
    assert direct["final_energy"] == pytest.approx(0.928044934017917 / 2.84, rel=1e-4)
