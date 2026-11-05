"""Which heat generators the PER sheet counts energy for.

PER lists every possible source per use -- heating, DHW, cooling -- one row
each ("Electricity (heat pump)", "Boiler", "District heating", "Solar thermal
system", ...), with the final energy it supplies in the shape's
``final_energy`` column. A source is in use when that figure is above zero.

ps-rails reads this to decide whether the parser's equipment read is the whole
picture: the selected-unit cells in app/equipment.py name a heat pump even in
a building heated by a boiler, so a named unit alone does not say it is used.

Each row also carries its two efficiencies, in W/W: the one PHPP calculated
(the shape's ``calculated_efficiency`` column, Q) and the one the team typed
(``user_determined_efficiency``, R: "User defined" in 10.x, "Manual value" in
9.7). PHPP uses the typed one whenever it holds a number (PER SI's
``AI = IF(ISNUMBER(R), R, Q)``), which is where a COP worked out in a
separate workbook ends up -- a heat pump water heater modelled as direct
electricity, say. The columns are BTU/hrW on a 10.x IP workbook and unitless
(W/W) otherwise; the unit is read from the header over the columns, and an
unrecognised one leaves both efficiencies null rather than guess.

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

# How far above the first block's heading the efficiency columns' unit row
# may sit (10.6 IP: two rows; 9.7: four).
UNIT_ROW_MAX = 8
# Efficiency header -> divisor to W/W. PHPP's IP sheets multiply the SI
# figure by its cvBTUh, so PER!Q44 at 1 W/W reads 3.412141156488376.
EFFICIENCY_UNITS = {"-": 1.0, "w/w": 1.0, "btu/hrw": 3.412141156488376}


def read_heat_generation(wb: Workbook, shape: shape_model.PhppShape) -> dict[str, Any] | None:
    per = shape.PER
    ws = _sheet(wb, per.name)
    if ws is None:
        return None

    label_col = column_index_from_string(per.locator_col)
    energy_col = column_index_from_string(per.columns.final_energy)
    calc_letter = per.columns.calculated_efficiency
    user_letter = per.columns.user_determined_efficiency
    calc_col = column_index_from_string(calc_letter)
    user_col = column_index_from_string(user_letter)
    headings = [getattr(per, block).locator_string_heading for block in BLOCKS]
    divisor = _efficiency_divisor(ws, label_col, calc_col, headings)

    def efficiency(r: int, col: int) -> float | None:
        value = _num(ws.cell(row=r, column=col).value)
        return None if value is None or divisor is None else value / divisor

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
                    "calculated_efficiency": efficiency(r, calc_col),
                    "user_efficiency": efficiency(r, user_col),
                    "calculated_efficiency_source": f"{ws.title}!{calc_letter}{r}",
                    "user_efficiency_source": f"{ws.title}!{user_letter}{r}",
                }
                for r in rows
                if _str(ws.cell(row=r, column=label_col).value)
            ]
        )
    return out


def _efficiency_divisor(
    ws: Worksheet, label_col: int, calc_col: int, headings: list[str | None]
) -> float | None:
    """The divisor that turns the efficiency columns into W/W, from the unit
    row above the first block; None when that row is missing or unknown."""
    found = [r for r in (_find(ws, label_col, h, 1, ws.max_row) for h in headings) if r]
    if not found:
        return None
    heading = min(found)
    for r in range(heading - 1, max(heading - UNIT_ROW_MAX, 1) - 1, -1):
        unit = _str(ws.cell(row=r, column=calc_col).value)
        if unit is not None:
            return EFFICIENCY_UNITS.get(unit.lower())
    return None


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
