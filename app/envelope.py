"""Envelope reads from a PHPP workbook (PAS-62 P2.3).

Emits the ``envelope`` subtree of the v1.0.0 output schema. P2.3 covers:

  components  — walls / roofs / floors / doors from AREAS surface_rows,
                plus thermal bridges from AREAS thermal_bridge_rows
  assemblies  — each assembly the surfaces use: U-value and layer stack
                (app.assemblies, PAS-79). Each surface carries its
                ``assembly_id`` as the join key.
  windows     — one component per Windows-sheet row (P2.3.c, PAS-80): its
                area (quantity included), Uw INSTALLED (the installation
                bridges included, as PHPP's heat balance and ps-rails use it,
                never the bare Uw), g-value, and glazing / frame ids.
  airtightness.n50_ach — from VENTILATION (one-cell read)

Patterns:
  * iterate-until-blank from a locator header ("Area input", "Thermal
    bridge input"); termination when column L (description) is None.
  * group_number classification — column M holds e.g. "8-External wall -
    ambient"; we parse the integer prefix and map to the schema enum.
"""

from __future__ import annotations

import re
from typing import Any

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from PHX.PHPP.phpp_localization import shape_model

from app.assemblies import read_assemblies, split_assembly_ref

# PHPP group_number → schema component enum. Pinned 2026-05-04 against the
# canonical 10.6 IP example. Groups 0 and 1 are aggregates (footprint, TFA);
# 2-6 are window placeholders that point at the Windows sheet (handled in
# P2.3.c). 7-11 are user-entered surface types.
GROUP_TO_COMPONENT = {
    7: "door",
    8: "wall",  # external wall - ambient
    9: "wall",  # ground wall (in contact with ground)
    10: "roof",
    11: "slab_on_grade",
}

_GROUP_PREFIX = re.compile(r"^\s*(\d+)\s*-")

# The output keys are imperial (area_ft2, length_ft, psi_value_Btuh_ftF).
# An SI workbook's Areas sheet is metric -- the shape says so -- and its
# figures were emitted unconverted under those keys until this conversion.
_TO_IMPERIAL = {
    "FT2": 1.0, "M2": 10.763910416709722,
    "FT": 1.0, "M": 3.280839895013123,
    "BTU/HR-FT-F": 1.0, "W/MK": 0.5777893165676,
    "BTU/HR-FT2-F": 1.0, "W/M2K": 0.17611018,
}


def read_envelope(wb: Workbook, shape: shape_model.PhppShape) -> dict[str, Any]:
    """Read the envelope subtree. Always returns the structure; each
    sub-piece (components list, airtightness) handles its own missing-
    sheet case by emitting an empty/null variant."""
    rows = _surface_rows(wb, shape.AREAS)
    assemblies = read_assemblies(wb, [r for r in rows if r["group"] not in _NOT_ASSEMBLY_GROUPS])
    surfaces = [_public(r) for r in rows if r["component"] is not None]
    return {
        "components": (
            surfaces + read_windows(wb, shape) + _thermal_bridge_components(wb, shape.AREAS)
        ),
        "assemblies": assemblies,
        "airtightness": _airtightness(wb, shape),
    }


# Footprint and TFA totals, and window placeholders (the windows come from
# the Windows sheet, read_windows).
_NOT_ASSEMBLY_GROUPS = {0, 1, 2, 3, 4, 5, 6}

# The g-value column has no PHX input; it is found by its header label.
# v10 heads it "g-value glazing", v9.7 "SHGC" (a row above the "Quan-tity" row).
_G_LABEL = re.compile(r"^\s*(g-value|shgc)\b", re.IGNORECASE)


# The Windows sheet is read once, inside these bounds. A real v10 sheet ends its
# rows at 177 and uses nothing past column BA; the bound keeps a workbook that
# declares a huge sheet (or has no end marker) from costing a scan per cell.
_WINDOWS_MAX_ROW = 400
_WINDOWS_MAX_COL = 80  # CB


def read_windows(wb: Workbook, shape: shape_model.PhppShape) -> list[dict[str, Any]]:
    """One window component per Windows-sheet row with an area, between the
    "Quan-tity" header and the "Unhide ..." end marker. A row PHPP leaves
    empty (no area, or zero) is skipped."""
    windows = shape.WINDOWS
    ws = _sheet(wb, windows.name)
    if ws is None:
        return []
    grid = _Grid(ws, _WINDOWS_MAX_ROW, _WINDOWS_MAX_COL)

    rows = windows.window_rows
    header_row = grid.find_row_starting(rows.locator_col_entry, rows.locator_string_entry)
    if header_row is None:
        return []
    end_row = grid.find_row_starting(windows.window_rows_end.locator_col_entry,
                                     windows.window_rows_end.locator_string_entry, after=header_row)
    last = (end_row - 1) if end_row else grid.max_row

    cols = rows.inputs
    # An edition may map no column for an input (v9.7 has no frame column): it reads as absent.
    names = ("quantity", "description", "glazing_id", "frame_id", "window_area", "u_w_installed")
    col = {name: column_index_from_string(getattr(cols, name).column)
           for name in names if getattr(cols, name).column}
    if "window_area" not in col or "u_w_installed" not in col:
        return []
    g_col = _g_column(grid, header_row)
    area_unit, u_unit = cols.window_area.unit, cols.u_w_installed.unit

    out: list[dict[str, Any]] = []
    for r in range(header_row + 2, last + 1):
        area = _num(grid.value(r, col["window_area"]))
        if area and area > 0:
            out.append(_window_row(grid, windows.name, r, col, g_col, area, area_unit, u_unit))
    return out


class _Grid:
    """A bounded, values-only copy of a sheet's top-left corner, read in one pass."""

    def __init__(self, ws: Worksheet, max_row: int, max_col: int):
        self.rows = [tuple(row) for row in ws.iter_rows(min_row=1, max_row=max_row,
                                                         max_col=max_col, values_only=True)]
        self.max_row = len(self.rows)
        self.max_col = max_col

    def value(self, r: int, c: int) -> Any:
        if 1 <= r <= self.max_row:
            row = self.rows[r - 1]
            if 1 <= c <= len(row):
                return row[c - 1]
        return None

    def find_row_starting(self, col_letter: str, needle: str, after: int = 0) -> int | None:
        """The first row after +after+ whose cell in the column starts with +needle+
        ("Quan-" for the cell "Quan-\\ntity")."""
        col_idx = column_index_from_string(col_letter)
        target = needle.strip().lower()
        for r in range(after + 1, self.max_row + 1):
            v = self.value(r, col_idx)
            if isinstance(v, str) and v.strip().lower().startswith(target):
                return r
        return None


def _window_row(grid: _Grid, sheet: str, r: int, col: dict[str, int], g_col: int | None,
                area: float, area_unit: str | None, u_unit: str | None) -> dict[str, Any]:
    def ref(c: int) -> str:
        return f"{sheet}!{get_column_letter(c)}{r}"

    def cell(name: str) -> Any:
        return grid.value(r, col[name]) if name in col else None

    desc = cell("description")
    quantity = _num(cell("quantity"))
    if quantity is not None and quantity.is_integer():
        quantity = int(quantity)
    window = {
        "component": "window",
        "label": str(desc).strip() if desc is not None else f"Window row {r}",
        "quantity": quantity,
        "area_ft2": _to_imperial(area, area_unit),
        "u_value_Btuh_ft2F": _to_imperial(_num(cell("u_w_installed")), u_unit),
        "shgc": _num(grid.value(r, g_col)) if g_col else None,
        "glazing_id": split_assembly_ref(cell("glazing_id"))[0],
        "frame_id": split_assembly_ref(cell("frame_id"))[0],
        "source_area": ref(col["window_area"]),
        "source_u_value": ref(col["u_w_installed"]),
    }
    if g_col:
        window["source_shgc"] = ref(g_col)
    return window


def _g_column(grid: _Grid, header_row: int) -> int | None:
    for r in (header_row, header_row - 1):
        for c in range(1, grid.max_col + 1):
            v = grid.value(r, c)
            if isinstance(v, str) and _G_LABEL.match(v):
                return c
    return None


def _public(row: dict[str, Any]) -> dict[str, Any]:
    """A surface row as the schema emits it: the assembly as its bare id."""
    out = {k: v for k, v in row.items() if k not in ("group", "assembly_id_raw")}
    out["assembly_id"], _ = split_assembly_ref(row["assembly_id_raw"])
    return out


def _to_imperial(value: float | None, unit: str | None) -> float | None:
    if value is None:
        return None
    factor = _TO_IMPERIAL.get((unit or "").upper())
    if factor is None:
        raise ValueError(f"no imperial conversion for unit {unit!r}")
    return value * factor


def _surface_components(wb: Workbook, areas: shape_model.Areas) -> list[dict[str, Any]]:
    """The surfaces of a type the schema has (see GROUP_TO_COMPONENT)."""
    return [_public(r) for r in _surface_rows(wb, areas) if r["component"] is not None]


def _surface_rows(wb: Workbook, areas: shape_model.Areas) -> list[dict[str, Any]]:
    """Every surface row, with its group number and raw assembly reference,
    and ``component`` None for a group the schema has no type for (a wall to
    temperature zone X, say). The assemblies count those too: they are
    envelope, just not a type this emits."""
    ws = _sheet(wb, areas.name)
    if ws is None:
        return []

    rows = areas.surface_rows
    header_row = _find_header_row(
        ws, rows.locator_col_header, rows.locator_string_header
    )
    if header_row is None:
        return []

    desc_col = column_index_from_string(rows.inputs.description.column)
    group_col = column_index_from_string(rows.inputs.group_number.column)
    assembly_col = column_index_from_string(rows.inputs.assembly_id.column)

    # The shape's `area` input is the USER-OVERRIDE column ("User determined
    # [ft2]" on v9.7, "User-defined calculation [ft2]" on v10) -- empty for
    # every surface PHPP computes from its own a x b entry, which is most of
    # them. PHPP's result is a separate column headed "Area [ft2]", so prefer
    # that and keep the override as the fallback.
    computed_area_col = _computed_area_col(ws, header_row + 1)
    area_col = computed_area_col or column_index_from_string(rows.inputs.area.column)
    area_letter = get_column_letter(area_col)

    out: list[dict[str, Any]] = []
    # Skip the header label row (header_row + 1 holds column labels like
    # "Building assembly description" — not data) and start at +2.
    for r in range(header_row + 2, ws.max_row + 1):
        desc = ws.cell(row=r, column=desc_col).value
        if desc is None or (isinstance(desc, str) and not desc.strip()):
            break  # iterate-until-blank termination

        group_raw = ws.cell(row=r, column=group_col).value
        group = _group_number(group_raw)
        component_type = _classify(group_raw)
        if group is None:
            continue

        area = _num(ws.cell(row=r, column=area_col).value)
        if computed_area_col and area is None:
            # PHPP seeds the table with one placeholder row per area group, and
            # their area formula is `IF(ISNUMBER(<quantity>), ..., "")` -- so an
            # unused group yields a blank rather than a zero. Groups 0-6 are
            # dropped by _classify anyway; group 7 maps to `door`, so its
            # placeholder used to surface as an exterior door with no area on
            # any building that has none. PHPP totals such a group as 0, so
            # skipping the row agrees with the workbook rather than departing
            # from it.
            #
            # Gated on having found PHPP's computed column, because on the
            # fallback path the value read is the user OVERRIDE, where blank is
            # the normal case and means "not overridden" -- skipping there
            # would drop most of the envelope.
            continue

        out.append(
            {
                "component": component_type,
                "label": str(desc),
                "area_ft2": _to_imperial(area, rows.inputs.area.unit),
                "source_area": f"{areas.name}!{area_letter}{r}",
                "group": group,
                "assembly_id_raw": ws.cell(row=r, column=assembly_col).value,
            }
        )
    return out


def _thermal_bridge_components(
    wb: Workbook, areas: shape_model.Areas
) -> list[dict[str, Any]]:
    ws = _sheet(wb, areas.name)
    if ws is None:
        return []

    tbs = areas.thermal_bridge_rows
    header_row = _find_header_row(
        ws, tbs.locator_col_header, tbs.locator_string_header
    )
    if header_row is None:
        return []

    desc_col = column_index_from_string(tbs.inputs.description.column)
    length_col = column_index_from_string(tbs.inputs.length.column)
    psi_col = column_index_from_string(tbs.inputs.psi_value.column)

    out: list[dict[str, Any]] = []
    for r in range(header_row + 2, ws.max_row + 1):
        desc = ws.cell(row=r, column=desc_col).value
        if desc is None or (isinstance(desc, str) and not desc.strip()):
            break
        length = _num(ws.cell(row=r, column=length_col).value)
        psi = _num(ws.cell(row=r, column=psi_col).value)
        if length is None and psi is None:
            continue  # placeholder row with description only — skip
        out.append(
            {
                "component": "thermal_bridge",
                "label": str(desc),
                "length_ft": _to_imperial(length, tbs.inputs.length.unit),
                "psi_value_Btuh_ftF": _to_imperial(psi, tbs.inputs.psi_value.unit),
                "source_area": f"{areas.name}!{tbs.inputs.length.column}{r}",
            }
        )
    return out


def _airtightness(
    wb: Workbook, shape: shape_model.PhppShape
) -> dict[str, Any]:
    """n50 lives at VENTILATION via the same locator-string pattern
    VERIFICATION uses for num_of_units (col I → input col M)."""
    vent = shape.VENTILATION
    n50 = vent.airtightness_n50

    ws = _sheet(wb, vent.name)
    if ws is None:
        return {"n50_ach": None, "source": None}

    locator_col = column_index_from_string(n50.locator_col)
    input_col = column_index_from_string(n50.input_column)
    needle = n50.locator_string.strip().lower()

    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=locator_col).value
        if isinstance(v, str) and v.strip().lower().startswith(needle):
            return {
                "n50_ach": _num(ws.cell(row=r, column=input_col).value),
                "source": f"{vent.name}!{n50.input_column}{r}",
            }

    return {"n50_ach": None, "source": None}


# --- helpers ---------------------------------------------------------------


def _classify(group_raw: Any) -> str | None:
    """Map an area-group cell to the schema component enum.

    Returns None for groups we intentionally skip (footprint/TFA/window
    placeholders) and for anything unrecognised, so the skip behaviour is the
    same whichever form the cell takes.
    """
    number = _group_number(group_raw)
    if number is None:
        return None
    return GROUP_TO_COMPONENT.get(number)


def _group_number(group_raw: Any) -> int | None:
    """The group number, from either form PHPP writes.

    v10 stores the labelled dropdown member as text -- ``"8-External wall -
    ambient"``. **v9.7 stores the bare number**, and openpyxl hands it back as
    an int. Accepting only the text form silently skipped every surface on
    every v9.7 file: the components list came back holding thermal bridges
    alone, which looks like a building with no walls rather than like a bug.
    """
    # bool is an int subclass, and a checkbox is not a group number.
    if isinstance(group_raw, bool):
        return None
    if isinstance(group_raw, int):
        return group_raw
    if isinstance(group_raw, float):
        return int(group_raw) if group_raw.is_integer() else None
    if isinstance(group_raw, str):
        match = _GROUP_PREFIX.match(group_raw)
        if match:
            return int(match.group(1))
        stripped = group_raw.strip()
        if stripped.isdigit():
            return int(stripped)
    return None


def _computed_area_col(ws: Worksheet, label_row: int) -> int | None:
    """Index of the column headed ``Area [<unit>]`` on the surface table's
    label row, or None.

    Located by label rather than by offset: it sits at AB on v9.7 and Z on
    v10, and the unit in the header varies with the workbook. The other
    area-bearing columns are all qualified ("User determined [ft2]",
    "Subtraction window areas [ft2]"), so the bare prefix is unambiguous --
    verified as a unique match on three real workbooks.
    """
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=label_row, column=col).value
        if isinstance(value, str) and " ".join(value.split()).lower().startswith(
            "area ["
        ):
            return col
    return None


def _find_header_row(ws: Worksheet, col_letter: str, needle: str) -> int | None:
    col_idx = column_index_from_string(col_letter)
    target = needle.strip().lower()
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=col_idx).value
        if isinstance(v, str) and v.strip().lower() == target:
            return r
    return None


def _sheet(wb: Workbook, name: str) -> Worksheet | None:
    needle = name.strip().lower()
    for sheet_name in wb.sheetnames:
        if sheet_name.strip().lower() == needle:
            return wb[sheet_name]
    return None


def _num(v: Any) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None
