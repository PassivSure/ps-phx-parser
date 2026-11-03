"""app.heat_generation on synthetic and real workbooks."""

import pathlib
from io import BytesIO

import pytest
from openpyxl import Workbook, load_workbook

from app.heat_generation import read_heat_generation
from app.parser import load_shape, parse_workbook

DOWNLOADS = pathlib.Path.home() / "Downloads"


def _per_workbook(shape, blocks):
    """A PER sheet with each block's heading, then its rows from row+2."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(shape.PER.name)
    col, energy = shape.PER.locator_col, shape.PER.columns.final_energy
    row = 5
    for name, rows in blocks.items():
        ws[f"{col}{row}"] = getattr(shape.PER, name).locator_string_heading
        for i, (label, value) in enumerate(rows, start=1):
            ws[f"{col}{row + i}"] = label
            if value is not None:
                ws[f"{energy}{row + i}"] = value
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
    }


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
