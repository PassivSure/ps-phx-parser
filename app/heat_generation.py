"""Which heat generators the PER sheet counts energy for.

PER lists every possible source per use -- heating, DHW, cooling -- one row
each ("Electricity (heat pump)", "Boiler", "District heating", "Solar thermal
system", ...), with the final energy it supplies in the shape's
``final_energy`` column. A source is in use when that figure is above zero.

ps-rails reads this to decide whether the parser's equipment read is the whole
picture: the selected-unit cells in app/equipment.py name a heat pump even in
a building heated by a boiler, so a named unit alone does not say it is used.

The rows between each block's PHX start and end locators are emitted as they
stand, the end row included; v9 ends a block on its auxiliary-electricity row,
v10 on solar thermal, and the consumer decides what each label means.
"""

from __future__ import annotations

from typing import Any

from openpyxl.utils import column_index_from_string
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from PHX.PHPP.phpp_localization import shape_model

BLOCKS = ("heating", "dhw", "cooling")

# How far below a block's heading its start row may sit, and how long a block
# may run, before the read gives up rather than wander into the next block.
HEADING_TO_START_MAX = 4
BLOCK_MAX_ROWS = 12


def read_heat_generation(wb: Workbook, shape: shape_model.PhppShape) -> dict[str, Any] | None:
    per = shape.PER
    ws = _sheet(wb, per.name)
    if ws is None:
        return None

    label_col = column_index_from_string(per.locator_col)
    energy_col = column_index_from_string(per.columns.final_energy)
    out: dict[str, Any] = {}
    for block in BLOCKS:
        locators = getattr(per, block)
        rows = _block_rows(ws, label_col, locators)
        out[block] = (
            None
            if rows is None
            else [
                {
                    "label": _str(ws.cell(row=r, column=label_col).value),
                    "final_energy": _num(ws.cell(row=r, column=energy_col).value),
                    "source": f"{ws.title}!{per.columns.final_energy}{r}",
                }
                for r in rows
                if _str(ws.cell(row=r, column=label_col).value)
            ]
        )
    return out


def _block_rows(ws: Worksheet, col: int, locators: Any) -> list[int] | None:
    """Row numbers from the block's start label to its end label, found below
    its heading. None when any locator is missing, so a consumer can tell
    "not read" from "nothing in use"."""
    heading = _find(ws, col, locators.locator_string_heading, 1, ws.max_row)
    if heading is None:
        return None
    start = _find(
        ws, col, locators.locator_string_start, heading + 1, heading + HEADING_TO_START_MAX
    )
    if start is None:
        return None
    end = _find(ws, col, locators.locator_string_end, start, start + BLOCK_MAX_ROWS)
    if end is None:
        return None
    return list(range(start, end + 1))


def _find(ws: Worksheet, col: int, needle: str | None, first: int, last: int) -> int | None:
    if not needle:
        return None
    target = needle.strip().lower()
    for r in range(first, min(last, ws.max_row) + 1):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str) and v.strip().lower() == target:
            return r
    return None


def _sheet(wb: Workbook, name: str) -> Worksheet | None:
    needle = name.strip().lower()
    for sheet_name in wb.sheetnames:
        if sheet_name.strip().lower() == needle:
            return wb[sheet_name]
    return None


def _str(v: Any) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def _num(v: Any) -> float | None:
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    return float(v)
