"""Project info reads from a PHPP workbook (PAS-62 P2.5).

Emits the ``project_info`` subtree of the v1.0.0 output schema. Mixes two
sources because PHX's shape doesn't map all the fields ps-rails consumes:

  PHX-mapped (locator pattern):
    project_name     — OVERVIEW.basic_data.address_project_name (Overview!C11)
    occupancy_type   — VERIFICATION.phi_building_use_type
                        (locator R="Building use", value T at row+1)

  Direct cell reads (PHX gap):
    postal_code      — Verification!K7 (strip surrounding quotes)
    location_string  — synthesized from Verification!K6/L7/K8/K7
    organizations    — Overview rows 31-36 + 48: column B has the role label,
                        column C has the firm/person name

  Deferred:
    verification_complete — needs a heuristic over Verification's certifier
                              fields. Filed as a follow-up; emits False here.
"""

from __future__ import annotations

import re
from typing import Any

from openpyxl.utils import column_index_from_string
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from PHX.PHPP.phpp_localization import shape_model

# Overview rows where label/name pairs for project organizations live.
# Pinned 2026-05-04 against the canonical 10.6 IP file.
ORGANIZATION_ROWS = (31, 33, 34, 35, 36, 48)
ORGANIZATION_LABEL_COL = "B"
ORGANIZATION_NAME_COL = "C"

# Verification sheet header cells that hold project address fields.
# These are not in the PHX shape; they're the standard locations PHPP uses
# in its Verification-sheet header.
ADDRESS_CELLS = {
    "street": "K6",
    "postal_code": "K7",
    "city": "L7",
    "state_country": "K8",
}


def read_project_info(wb: Workbook, shape: shape_model.PhppShape) -> dict[str, Any]:
    return {
        "project_name": _project_name(wb, shape),
        "postal_code": _postal_code(wb, shape),
        "location_string": _location_string(wb, shape),
        "occupancy_type": _occupancy_type(wb, shape),
        "verification_complete": False,  # deferred — see module docstring
        "organizations": _organizations(wb, shape),
    }


def _project_name(wb: Workbook, shape: shape_model.PhppShape) -> str | None:
    ws = _sheet(wb, shape.OVERVIEW.name)
    if ws is None:
        return None
    address = shape.OVERVIEW.basic_data.address_project_name
    return _str(ws[address].value)


def _postal_code(wb: Workbook, shape: shape_model.PhppShape) -> str | None:
    ws = _sheet(wb, shape.VERIFICATION.name)
    if ws is None:
        return None
    raw = _str(ws[ADDRESS_CELLS["postal_code"]].value)
    if raw is None:
        return None
    # PHPP wraps zips in literal quotes (e.g. '"03049"') to keep leading
    # zeros from being parsed as numbers.
    return raw.strip().strip('"').strip("'") or None


def _location_string(wb: Workbook, shape: shape_model.PhppShape) -> str | None:
    """Synthesize "<street>, <city>, <state> <zip>" from the Verification
    header. Returns None if street + city + state are all missing."""
    ws = _sheet(wb, shape.VERIFICATION.name)
    if ws is None:
        return None

    street = _str(ws[ADDRESS_CELLS["street"]].value)
    city = _str(ws[ADDRESS_CELLS["city"]].value)
    state = _str(ws[ADDRESS_CELLS["state_country"]].value)
    zip_ = _postal_code(wb, shape)

    parts: list[str] = []
    if street:
        parts.append(street)
    if city and state:
        parts.append(f"{city}, {state}{f' {zip_}' if zip_ else ''}")
    elif city:
        parts.append(city)
    elif state:
        parts.append(state)

    return ", ".join(parts) if parts else None


def _occupancy_type(wb: Workbook, shape: shape_model.PhppShape) -> str | None:
    """Look up phi_building_use_type via the PHX locator pattern. Returns
    the raw PHX-formatted string (e.g. ``"21-Non-res building: School ..."``);
    the Mapper does the human-friendly translation against OccupancyType."""
    ws = _sheet(wb, shape.VERIFICATION.name)
    if ws is None:
        return None

    item = shape.VERIFICATION.phi_building_use_type
    locator_col = column_index_from_string(item.locator_col)
    needle = item.locator_string.strip().lower()

    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=locator_col).value
        if isinstance(v, str) and v.strip().lower() == needle:
            target_row = r + item.input_row_offset
            target_col = column_index_from_string(item.input_column)
            return _str(ws.cell(row=target_row, column=target_col).value)
    return None


def _organizations(
    wb: Workbook, shape: shape_model.PhppShape
) -> list[dict[str, str]]:
    """Each pinned row in Overview has a role label in column B and a
    firm/person name in column C. Empty-name rows are skipped."""
    ws = _sheet(wb, shape.OVERVIEW.name)
    if ws is None:
        return []

    out: list[dict[str, str]] = []
    label_col = column_index_from_string(ORGANIZATION_LABEL_COL)
    name_col = column_index_from_string(ORGANIZATION_NAME_COL)
    for row in ORGANIZATION_ROWS:
        label = _str(ws.cell(row=row, column=label_col).value)
        name = _str(ws.cell(row=row, column=name_col).value)
        if not name:
            continue
        out.append({"label": _clean_org_label(label) or "", "name": name})
    return out


_LABEL_TRAIL = re.compile(r"\s*[/]\s*.*$")  # strip "/ E-mail" / "/ Last name"


def _clean_org_label(label: str | None) -> str | None:
    """Overview labels are formatted ``"Architect name / E-mail"``. Keep
    the role prefix (everything before the slash) so the schema's `label`
    field carries something meaningful."""
    if not label:
        return None
    # Drop "/ <secondary>" tail
    role = _LABEL_TRAIL.sub("", label).strip()
    # Drop trailing " name"
    return re.sub(r"\s+name\s*$", "", role, flags=re.IGNORECASE) or None


# --- helpers ---------------------------------------------------------------


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
