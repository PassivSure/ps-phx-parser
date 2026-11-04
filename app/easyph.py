"""The heat pumps and cooling unit typed on easyPH's simplified form.

easyPH lets each of three services (cooling, space heating, DHW) take its
unit either from the form on the easyPH tab ("Simplified efficiency
verification": manufacturer, product name, nominal power and one rating) or
from the Components sheet ("ALTERNATIVELY: Detailed proof of efficiency").
An "x" in column O on one of those two rows says which. On the simplified
path PHPP fills the HP and cooling selections with placeholders ("easyPH
heating", "easyPH cooling"), so the real unit is only on the form.

The rows move between files (Holmes sits two rows lower than PHI's easyPH
Example), and "Manufacturer" appears once per service, so each service is a
block found by its own header and read by label inside it:
  - cooling: column I "Cooling units", read only when "Mechanical cooling
    available?" is ticked (column P);
  - space heating: column O "Space heating";
  - DHW: column O "DHW".
A block runs to the next block or the next section marker (column D "❌").

Inside a block the labels are in P and the values in Q. Nominal power's unit
is in R ("kW" on SI, "kBTU/hr" on IP). The rating sits in whichever of Q, R
or S holds a number on the "Efficiency characteristic value" row, and its
metric is the header above that column ("SEER", "SEER(2)", "SCOP",
"HSPF(2)"). PHPP's "(2)" covers both HSPF and HSPF2 (SEER and SEER2); it is
reported as HSPF2 / SEER2, the AHRI 210/240 metric the form names, and the
spelling ps-rails reads.
"""

from __future__ import annotations

from typing import Any

from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from app.heat_generation import _num, _sheet, _str

SHEET = "easyPH"
SEARCH_ROWS = 400
SEARCH_COLS = 28  # A..AB

COL_D, COL_I, COL_O, COL_P, COL_Q, COL_R, COL_S = 4, 9, 15, 16, 17, 18, 19
SECTION_MARK = "❌"

# Block -> (column, header text) that starts it.
BLOCKS = {
    "cooling": (COL_I, "Cooling units"),
    "heating": (COL_O, "Space heating"),
    "dhw": (COL_O, "DHW"),
}
SIMPLIFIED = "Simplified efficiency verification"
DETAILED = "ALTERNATIVELY"
COOLING_AVAILABLE = "Mechanical cooling available?"
RATING_ROW = "Efficiency characteristic value"
METRICS = {"SCOP": "SCOP", "HSPF(2)": "HSPF2", "SEER": "SEER", "SEER(2)": "SEER2"}


def read_easyph(wb: Workbook) -> dict[str, dict[str, Any]] | None:
    """Each service's form ("simplified", "detailed" or None) and, on the
    simplified form, its unit. None when the workbook has no easyPH sheet."""
    ws = _sheet(wb, SHEET)
    if ws is None:
        return None
    grid = [
        list(row)
        for row in ws.iter_rows(
            min_row=1, max_row=SEARCH_ROWS, max_col=SEARCH_COLS, values_only=True
        )
    ]

    def at(row: int, col: int) -> Any:
        line = grid[row - 1] if 0 < row <= len(grid) else []
        return line[col - 1] if col - 1 < len(line) else None

    starts = {name: _find(grid, at, col, text) for name, (col, text) in BLOCKS.items()}
    marks = [r for r in range(1, len(grid) + 1) if _text(at(r, COL_D)) == SECTION_MARK]
    out: dict[str, dict[str, Any]] = {}
    for name, start in starts.items():
        if start is None:
            out[name] = {"form": None, "unit": None}
            continue
        ends = [s for s in starts.values() if s is not None and s > start]
        ends += [m for m in marks if m > start]
        end = min(ends, default=len(grid) + 1)
        out[name] = _block(ws.title, at, name, start, end)
    return out


def _block(title: str, at, name: str, start: int, end: int) -> dict[str, Any]:
    rows = range(start, end)
    labels = {}
    for r in rows:
        for col in (COL_P, COL_O):
            label = _text(at(r, col))
            if label and label not in labels:
                labels[label] = r

    if name == "cooling":
        available = labels.get(COOLING_AVAILABLE)
        if available is None or not _ticked(at(available, COL_P)):
            return {"form": None, "unit": None}

    simplified = _row_starting(labels, SIMPLIFIED)
    detailed = _row_starting(labels, DETAILED)
    if simplified is not None and _ticked(at(simplified, COL_O)):
        form = "simplified"
    elif detailed is not None and _ticked(at(detailed, COL_O)):
        form = "detailed"
    else:
        return {"form": None, "unit": None}
    if form == "detailed":
        return {"form": form, "unit": None}

    def cell(row: int, col: int) -> str:
        return f"{title}!{get_column_letter(col)}{row}"

    unit: dict[str, Any] = {"sources": {}}
    for field, label in (("manufacturer", "Manufacturer"), ("product", "Product name")):
        row = labels.get(label)
        unit[field] = _str(at(row, COL_Q)) if row else None
        if row:
            unit["sources"][field] = cell(row, COL_Q)

    power_row = labels.get("Nominal power")
    unit["nominal_power"] = _num(at(power_row, COL_Q)) if power_row else None
    unit["nominal_power_unit"] = _str(at(power_row, COL_R)) if power_row else None
    if power_row:
        unit["sources"]["nominal_power"] = cell(power_row, COL_Q)

    unit["efficiency_value"] = unit["efficiency_metric"] = None
    rating_row = labels.get(RATING_ROW)
    if rating_row:
        for col in (COL_Q, COL_R, COL_S):
            value = _num(at(rating_row, col))
            if value is None:
                continue
            unit["efficiency_value"] = value
            unit["efficiency_metric"] = _metric_above(at, col, start, rating_row)
            unit["sources"]["efficiency_value"] = cell(rating_row, col)
            break
    return {"form": form, "unit": unit}


def _metric_above(at, col: int, start: int, row: int) -> str | None:
    """The metric named over +col+, nearest the rating row first."""
    for r in range(row - 1, start - 1, -1):
        metric = METRICS.get(_text(at(r, col)) or "")
        if metric:
            return metric
    return None


def _find(grid, at, col: int, text: str) -> int | None:
    for r in range(1, len(grid) + 1):
        if _text(at(r, col)) == text:
            return r
    return None


def _row_starting(labels: dict[str, int], prefix: str) -> int | None:
    for label, row in labels.items():
        if label.startswith(prefix):
            return row
    return None


def _ticked(value: Any) -> bool:
    return isinstance(value, str) and value.strip().lower() == "x"


def _text(value: Any) -> str | None:
    return value.strip() if isinstance(value, str) and value.strip() else None
