"""Opaque building assemblies: U-value and layer stack, per assembly (PAS-79).

Every surface on the Areas sheet names its assembly (column AA, "01ud-Wall 1").
This reads each assembly those surfaces use, once:

  u_value_w_m2k   the Components sheet's "Building assemblies" table, which
                  lists every assembly id -- the user-defined ones from the
                  U-values sheet AND PHPP's built-in ones ("83ud-PH External
                  wall"), which have no U-values block.
  layers          the assembly's block on the U-values sheet, when it has one:
                  each layer's thickness and up to three area sections
                  (material and conductivity), plus the share of sections 2
                  and 3 (the stud fraction of a framed layer).
  component       the type covering most of its area, from the surface rows
                  app.envelope already read, with their total area and their
                  Areas group numbers. A surface in a group the schema has no
                  type for (14, a wall to temperature zone X) still counts
                  towards the area and the groups; an assembly used only there
                  has no type.

Transcribed, not derived: no R, no unit conversion, no insulation
classification. ps-rails does that arithmetic, as it does for the extraction
it replaces. Everything is read from the metric pane through the defined
names, which point there on an IP workbook too, so every figure is SI and its
key says so.

Both tables are found from defined names and checked by their labels. A
layout that moved reads nothing rather than the wrong cells.
"""

from __future__ import annotations

import logging
import re
from collections import defaultdict
from typing import Any

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook import Workbook

from app.named_ranges import _REF

logger = logging.getLogger(__name__)

MAX_BLOCKS = 60
_ID = re.compile(r"^\s*([0-9]+[a-z]+[0-9]*)\s*-\s*(.*)$")

# U-values block: names and the label rows they are checked against.
_BLOCK_NAME = "U_Werte_Bauteil_{n}_Bezeichnung"
_BLOCK_LAYERS = "U_Werte_Bauteil_{n}_Teilflaechen"
_BLOCK_THICKNESS = "U_Werte_Bauteil_{n}_Summe_Breite_Dicke"
_BLOCK_SECTION_2 = "U_Werte_Bauteil_{n}_Flaechenanteil_2"
_BLOCK_SECTION_3 = "U_Werte_Bauteil_{n}_Flaechenanteil_3"
_BLOCK_HEADERS = {"description of building assembly", "building assembly description"}
_BLOCK_ID_LABEL = "assembly no."

# Components "Building assemblies" table.
_TABLE_HEADING = "PHPP_Komponenten_Ueberschrift_Bauteilaufbauten"
_TABLE_ROWS = 600  # ud rows, then PHPP's own catalogue
_TABLE_LABELS = ("id", "building assembly", "total thickness", "u-value")

# Tolerance for the U a block computes against the one the Components table
# lists for the same id. They are the same figure in two places.
_U_AGREE = 1e-6


def read_assemblies(wb: Workbook, surfaces: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """One entry per assembly a surface uses, in first-use order. [] when no
    surface names one."""
    used: dict[str, dict[str, Any]] = {}
    for surface in surfaces:
        assembly_id, name = split_assembly_ref(surface.get("assembly_id_raw"))
        if assembly_id is None:
            continue
        entry = used.setdefault(assembly_id, {
            "name_from_areas": name, "area_ft2": 0.0, "surfaces": 0,
            "by_type": defaultdict(float), "groups": set(), "area_missing": False,
        })
        entry["groups"].add(surface["group"])
        entry["surfaces"] += 1
        area = surface.get("area_ft2")
        if area is None:
            entry["area_missing"] = True
        else:
            entry["area_ft2"] += area
            if surface["component"] is not None:
                entry["by_type"][surface["component"]] += area
    if not used:
        return []

    table = _components_table(wb)
    blocks = _blocks(wb)

    out = []
    for assembly_id, entry in used.items():
        listed = table.get(assembly_id, {})
        block = blocks.get(assembly_id)
        u_value = listed.get("u_value_w_m2k")
        if block and u_value is not None and block["u_value_w_m2k"] is not None \
                and abs(block["u_value_w_m2k"] - u_value) > _U_AGREE:
            logger.warning(
                "[assemblies] %s: U-values block gives U %.6f, Components lists %.6f",
                assembly_id, block["u_value_w_m2k"], u_value,
            )

        out.append({
            "id": assembly_id,
            "name": (block and block["name"]) or listed.get("name") or entry["name_from_areas"],
            "component": _dominant(entry["by_type"]),
            "groups": sorted(entry["groups"]),
            "area_ft2": None if entry["area_missing"] else entry["area_ft2"],
            "surface_count": entry["surfaces"],
            "u_value_w_m2k": u_value,
            "total_thickness_mm": (block and block["total_thickness_mm"])
            or listed.get("total_thickness_mm"),
            "section_2_fraction": block and block["section_2_fraction"],
            "section_3_fraction": block and block["section_3_fraction"],
            "layers": block["layers"] if block else [],
            "source_u_value": listed.get("source"),
            "source_layers": block and block["source"],
        })
    return out


def split_assembly_ref(raw: Any) -> tuple[str | None, str | None]:
    """'01ud-Wall 1- Basement' -> ('01ud', 'Wall 1- Basement')."""
    if not isinstance(raw, str):
        return None, None
    match = _ID.match(raw)
    if match is None:
        return None, None
    return match.group(1), match.group(2).strip() or None


def _dominant(by_type: dict[str, float]) -> str | None:
    """The type covering the most area. A type is never invented: an assembly
    whose surfaces were all unclassifiable has none."""
    if not by_type:
        return None
    return max(sorted(by_type), key=lambda t: by_type[t])


# ── Components table ────────────────────────────────────────────────────


def _components_table(wb: Workbook) -> dict[str, dict[str, Any]]:
    """id -> U (W/m²K), name and thickness, from the Components sheet."""
    ref = _ref(wb, _TABLE_HEADING)
    if ref is None:
        return {}
    sheet, first_col, heading_row = ref
    ws = wb[sheet]
    rows = list(ws.iter_rows(min_row=heading_row, max_row=heading_row + _TABLE_ROWS,
                             min_col=first_col, max_col=first_col + 10, values_only=True))

    # The label row ("ID", "Building assembly", "U-value") and the unit row
    # under it ("W/(m²K)" -- the metric pane) sit a few rows below the heading.
    found = next(((i, [_label(v) for v in row]) for i, row in enumerate(rows[:8])
                  if {"id", "u-value"} <= {_label(v) for v in row}), None)
    if found is None:
        logger.warning("[assemblies] no Components assemblies table under %s", _TABLE_HEADING)
        return {}
    offset, labels = found
    cols = {k: labels.index(k) for k in _TABLE_LABELS if k in labels}

    unit_row = rows[offset + 1]
    if "w/(m²k)" not in _label(unit_row[cols["u-value"]]):
        logger.warning("[assemblies] Components U-value column is not in W/(m²K); not read")
        return {}

    u_letter = get_column_letter(first_col + cols["u-value"])
    table: dict[str, dict[str, Any]] = {}
    for i, row in enumerate(rows[offset + 2:], start=heading_row + offset + 2):
        assembly_id = row[cols["id"]]
        if not isinstance(assembly_id, str) or not assembly_id.strip():
            continue
        assembly_id = assembly_id.strip()
        if assembly_id in table:
            continue
        thickness_m = _num(row[cols["total thickness"]]) if "total thickness" in cols else None
        name = row[cols["building assembly"]] if "building assembly" in cols else None
        table[assembly_id] = {
            "u_value_w_m2k": _num(row[cols["u-value"]]),
            "name": name.strip() if isinstance(name, str) and name.strip() else None,
            "total_thickness_mm": None if thickness_m is None else thickness_m * 1000,
            "source": f"{sheet}!{u_letter}{i}",
        }
    return table


# ── U-values blocks ─────────────────────────────────────────────────────


def _blocks(wb: Workbook) -> dict[str, dict[str, Any]]:
    """id -> the assembly's U-values block, for every block with a name.

    The blocks span ~1,270 rows of one sheet, and in read-only mode every
    iter_rows call streams the sheet from the top, so the sheet is read once
    into a grid rather than once per block.
    """
    refs = []
    for n in range(1, MAX_BLOCKS + 1):
        name_ref = _ref(wb, _BLOCK_NAME.format(n=n))
        layers_ref = _ref(wb, _BLOCK_LAYERS.format(n=n), whole=True)
        if name_ref is not None and layers_ref is not None and name_ref[0] == layers_ref[0]:
            refs.append((n, name_ref, layers_ref))
    if not refs:
        return {}

    sheet = refs[0][1][0]
    last_row = max(layers_ref[4] for _, _, layers_ref in refs) + 8
    grid = _Grid(wb[sheet], last_row, _GRID_COLS)

    out: dict[str, dict[str, Any]] = {}
    for n, name_ref, layers_ref in refs:
        if name_ref[0] != sheet:
            continue
        block = _block(wb, grid, n, name_ref, layers_ref)
        if block is not None and block["id"] not in out:
            out[block["id"]] = block
    return out


_GRID_COLS = 26  # A..Z covers every block column in v9.7 and v10


class _Grid:
    """A sheet's values, read once; cell(row, col) is 1-based like openpyxl."""

    def __init__(self, ws, max_row: int, max_col: int):
        self.rows = list(ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col,
                                      values_only=True))

    def cell(self, row: int, col: int) -> Any:
        if 1 <= row <= len(self.rows) and 1 <= col <= len(self.rows[row - 1]):
            return self.rows[row - 1][col - 1]
        return None

    def row(self, row: int) -> tuple:
        return self.rows[row - 1] if 1 <= row <= len(self.rows) else ()

    def named(self, wb: Workbook, name: str) -> list[Any]:
        ref = _ref(wb, name, whole=True)
        if ref is None:
            return []
        _, c1, r1, c2, r2 = ref
        return [self.cell(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def _block(wb: Workbook, grid: _Grid, n: int, name_ref, layers_ref) -> dict[str, Any] | None:
    sheet, name_col, name_row = name_ref

    # The header row above the name holds the description label and
    # "Assembly no."; the id sits under the latter, on the name row. v10 says
    # "Description of building assembly", v9.7 "Building assembly description".
    labels = [_label(v) for v in grid.row(name_row - 1)]
    if not _BLOCK_HEADERS.intersection(labels) or _BLOCK_ID_LABEL not in labels:
        return None
    name = grid.cell(name_row, name_col)
    if not isinstance(name, str) or not name.strip():
        return None  # an unused block keeps its id and a meaningless U
    assembly_id = grid.cell(name_row, labels.index(_BLOCK_ID_LABEL) + 1)
    if not isinstance(assembly_id, str) or not assembly_id.strip():
        return None

    _, first_col, first_row, last_col, last_row = layers_ref
    thickness = grid.named(wb, _BLOCK_THICKNESS.format(n=n))
    layers = []
    for i, r in enumerate(range(first_row, last_row + 1)):
        row = [grid.cell(r, c) for c in range(first_col, last_col + 1)]
        sections = [_section(row[k], row[k + 1]) for k in (0, 2, 4)]
        mm = _num(thickness[i]) if i < len(thickness) else None
        if all(s is None for s in sections) and not mm:
            continue
        layers.append({
            "thickness_mm": mm,
            "material": sections[0] and sections[0]["material"],
            "conductivity_w_mk": sections[0] and sections[0]["conductivity_w_mk"],
            "section_2_material": sections[1] and sections[1]["material"],
            "section_2_conductivity_w_mk": sections[1] and sections[1]["conductivity_w_mk"],
            "section_3_material": sections[2] and sections[2]["material"],
            "section_3_conductivity_w_mk": sections[2] and sections[2]["conductivity_w_mk"],
        })

    total_mm = sum(layer["thickness_mm"] or 0 for layer in layers) or None
    return {
        "id": assembly_id.strip(),
        "name": name.strip(),
        "layers": layers,
        "total_thickness_mm": total_mm,
        "section_2_fraction": _num(_first(grid.named(wb, _BLOCK_SECTION_2.format(n=n)))),
        "section_3_fraction": _num(_first(grid.named(wb, _BLOCK_SECTION_3.format(n=n)))),
        "u_value_w_m2k": _block_u(grid, last_row),
        "source": f"{sheet}!{get_column_letter(first_col)}{first_row}:"
                  f"{get_column_letter(last_col)}{last_row}",
    }


def _block_u(grid: _Grid, last_layer_row: int) -> float | None:
    """The U the block computes, for the cross-check against the Components
    table: the number right of the "U-value" label a few rows under the
    layers ("U-value [W/(m²K)]:" on v10, "U-value:" on v9.7; never the
    "U-value supplement")."""
    for r in range(last_layer_row + 1, last_layer_row + 9):
        row = grid.row(r)
        for c, value in enumerate(row):
            label = _label(value)
            if label.startswith("u-value") and "supplement" not in label:
                return next((_num(v) for v in row[c + 1:] if _num(v) is not None), None)
    return None


def _section(material: Any, conductivity: Any) -> dict[str, Any] | None:
    material = material.strip() if isinstance(material, str) and material.strip() else None
    conductivity = _num(conductivity)
    if material is None and not conductivity:
        return None
    return {"material": material, "conductivity_w_mk": conductivity}


# ── helpers ─────────────────────────────────────────────────────────────


def _ref(wb: Workbook, name: str, whole: bool = False):
    """(sheet, col, row) of a defined name's first cell -- or with +whole+,
    (sheet, first col, first row, last col, last row). None when absent."""
    defined = wb.defined_names.get(name)
    if defined is None:
        return None
    match = _REF.match(str(defined.value).strip())
    if match is None or match.group(1) not in wb.sheetnames:
        return None
    sheet, first, last = match.groups()
    c1, r1 = _cell(first)
    if not whole:
        return sheet, c1, r1
    c2, r2 = _cell(last or first)
    return sheet, c1, r1, c2, r2


def _cell(ref: str) -> tuple[int, int]:
    letters, digits = re.match(r"\$?([A-Z]+)\$?(\d+)", ref).groups()
    return column_index_from_string(letters), int(digits)


def _first(values: list[Any]) -> Any:
    return values[0] if values else None


def _label(value: Any) -> str:
    return " ".join(value.split()).lower() if isinstance(value, str) else ""


def _num(value: Any) -> float | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return float(value)
