"""The boiler on PHPP's Boiler sheet: type, fuel, design output, efficiencies.

The sheet's inputs sit in one "Project data" column (N on 9.7 and 10.x),
which holds the figure PHPP calculates with: the user-defined value when
"Apply user-defined values" is ticked, else PHPP's standard value. The labels
are in the column to their left (I), and the design output's unit is in the
column after (O: kBTU/hr on IP, kW on SI). PHX's shape has no Boiler locators,
so the rows are found by label.

The sheet is filled whether or not anything uses the boiler (a template
leaves PHPP's standard values in place), so a read here does not say the
boiler is in use: PER's heat-generation rows do. A type of "1-None" or blank
reads as no boiler.
"""

from __future__ import annotations

from typing import Any

from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from PHX.PHPP.phpp_localization import shape_model

from app.heat_generation import _num, _sheet, _str

# Field -> the label that starts its row (matched case-insensitively, by prefix,
# since the standby-loss label carries a degree sign and a temperature).
FIELDS = {
    "boiler_type": "Boiler type",
    "fuel": "Fuel",
    "design_output": "Design output",
    "indoor": "Installation of boiler",
    "efficiency_30": "Boiler efficiency at 30% load",
    "efficiency_100": "Boiler efficiency at nominal output",
    "standby_loss_70": "Standby heat loss boiler at 70",
}
TEXT_FIELDS = ("boiler_type", "fuel")
VALUE_HEADER = "Project data"
SEARCH_ROWS = 60
SEARCH_COLS = 20


def read_boiler(wb: Workbook, shape: shape_model.PhppShape) -> dict[str, Any] | None:
    ws = _sheet(wb, shape.BOILER.name)
    if ws is None:
        return None
    grid = [
        list(row)
        for row in ws.iter_rows(
            min_row=1, max_row=SEARCH_ROWS, max_col=SEARCH_COLS, values_only=True
        )
    ]
    rows = _label_rows(grid)
    value_col = _value_col(grid)
    if value_col is None or "boiler_type" not in rows:
        return None

    def at(row: int, col: int) -> Any:
        line = grid[row - 1] if row - 1 < len(grid) else []
        return line[col - 1] if col - 1 < len(line) else None

    boiler_type = _str(at(rows["boiler_type"], value_col))
    if boiler_type is None or boiler_type.startswith("1-"):
        return None

    out: dict[str, Any] = dict.fromkeys(FIELDS)
    sources: dict[str, str] = {}
    for field, row in rows.items():
        value = at(row, value_col)
        out[field] = _str(value) if field in TEXT_FIELDS else _num(value)
        sources[field] = f"{ws.title}!{get_column_letter(value_col)}{row}"
    if out["fuel"] == "-":
        out["fuel"] = None
    if out["indoor"] is not None:
        out["indoor"] = bool(out["indoor"])
    out["design_output_unit"] = (
        _str(at(rows["design_output"], value_col + 1)) if "design_output" in rows else None
    )
    out["sources"] = sources
    return out


def _label_rows(grid: list[list[Any]]) -> dict[str, int]:
    """The first row each field's label starts (1-based)."""
    found: dict[str, int] = {}
    for r, line in enumerate(grid, start=1):
        for v in line:
            if not isinstance(v, str):
                continue
            text = v.strip().lower()
            for field, label in FIELDS.items():
                if field not in found and text.startswith(label.lower()):
                    found[field] = r
    return found


def _value_col(grid: list[list[Any]]) -> int | None:
    """The "Project data" column (1-based): the figures PHPP calculates with."""
    for line in grid:
        for c, v in enumerate(line, start=1):
            if isinstance(v, str) and v.strip().lower() == VALUE_HEADER.lower():
                return c
    return None
