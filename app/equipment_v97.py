"""Mechanical equipment on PHPP 9.7 IP, which names it differently from 10.x.

app.equipment reads 10.x through defined names that 9.7 does not carry (no
"x" box per heat pump, no cooling-unit selection). The same facts sit in 9.7
under other names and beside them, so this module reads them there and emits
the same items (app.equipment._item):

- Ventilation: the selection and design airflow keep their 10.x names. The
  unit's heat recovery and humidity recovery are on the selection's own row,
  three and four columns right of it under the "Unit hHR" header, not on the
  Components sheet.
- Heat pumps: WP_Heizung_Art_Waermepumpe (heating) and
  WP_Warmwasser_Art_Waermepumpe (DHW) hold "<n>-<name>", where n is the slot
  of the unit's data block, WP_WP<n>_Name and WP_WP<n>_Leistungsdaten (rows of
  source °C, sink °C, heating capacity kW, COP). The rating is the test point
  with the largest capacity, as on 10.x; 9.7 shows the same unitless COP on
  both panes. A unit is emitted only when the HP sheet's labelled share of
  the service it covers ("Covered fraction of space heating demand" / "...
  of DHW demand", fed from PER) is above zero, and the DHW unit only when the
  sheet counts two heat pumps: with one, the heating unit heats the water and
  nothing here says so in a form ps-rails reads, so DHW is left to it.
  WP_WP<n>_Nennleistung is not the unit's capacity: it is the building's load
  scaled by the table's relative-power column (3.02 kW on every slot of one
  file).
- Cooling: only recirculation cooling, when its box is ticked. 9.7 names no
  unit, so the row label "Recirculation cooling" is its name. Capacity in kW;
  the seasonal ratio _JAZ is W/W, emitted as SEER (x 3.412 Btu/Wh), which is
  the figure the IP pane shows. Supply-air and surface cooling emit nothing,
  as on 10.x.

Checked on two real 9.7 IP workbooks (55th, 17 Mile); EN_9_6A is not read.
"""

from __future__ import annotations

import re
from typing import Any

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.workbook import Workbook

from app.named_ranges import cells_of, resolve, sheet_of

M3H_TO_CFM = 0.588578
BTU_PER_WH = 3.412141156488376

HEATING_SELECTION = "WP_Heizung_Art_Waermepumpe"
DHW_SELECTION = "WP_Warmwasser_Art_Waermepumpe"
HEAT_PUMP_COUNT = "WP_Anzahl"
VENT_SELECTION = "Lueftung_Auswahl_Lueftungsgeraet"
VENT_AIRFLOW = "Lueftung_Auslegungsvolumenstrom"
COOLING_BOX = "Kuehlgeraete_Umluft_Kuehlung_Ankreuzen"
COOLING_NAME = "Recirculation cooling"

# Labels in column I of the HP sheet, value in M.
HEATING_SHARE = "Covered fraction of space heating demand"
DHW_SHARE = "Covered fraction of DHW demand"
_SHARE_ROWS = range(1, 30)

_SLOT = re.compile(r"^\s*(\d+)-(.+?)\s*$")


def read_equipment_v97(wb: Workbook) -> list[dict[str, Any]]:
    from app.equipment import _item

    return _ventilation(wb, _item) + _cooling(wb, _item) + _heat_pumps(wb, _item)


def _num(value: Any) -> float | None:
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def _first(value: Any) -> Any:
    return value[0] if isinstance(value, list) and value else value


def _ventilation(wb: Workbook, item) -> list[dict[str, Any]]:
    from app.equipment import _classify_ventilation, _strip_prefix

    raw = _first(resolve(wb, VENT_SELECTION))
    name = _strip_prefix(raw)
    if name is None:
        return []
    heat, humidity, cell = _vent_recovery(wb)
    airflow = _num(_first(resolve(wb, VENT_AIRFLOW)))
    source = VENT_SELECTION
    source += (
        f"; recovery from {cell}"
        if cell
        else "; heat/humidity recovery not found beside the selection"
    )
    return [
        item(
            equipment_type=_classify_ventilation(humidity),
            name=name,
            heat_recovery_efficiency_pct=heat,
            airflow_m3h=airflow,
            airflow_cfm=None if airflow is None else airflow * M3H_TO_CFM,
            source=source,
        )
    ]


def _vent_recovery(wb: Workbook) -> tuple[float | None, float | None, str | None]:
    """Heat and humidity recovery (%) on the selection's row, checked by the
    "Unit hHR" header above the first of them."""
    where = cells_of(wb, VENT_SELECTION)
    if where is None:
        return None, None, None
    sheet, first, _last = where
    col, row = coordinate_from_string(first)
    hr_col = column_index_from_string(col) + 3
    rows = list(
        wb[sheet].iter_rows(
            min_row=row - 1, max_row=row, min_col=hr_col, max_col=hr_col + 1, values_only=True
        )
    )
    header, values = rows[0], rows[1]
    if not (isinstance(header[0], str) and header[0].strip() == "Unit hHR"):
        return None, None, None
    heat, humidity = (None if _num(v) is None else _num(v) * 100.0 for v in values)
    return heat, humidity, f"{sheet}!{get_column_letter(hr_col)}{row}"


def _cooling(wb: Workbook, item) -> list[dict[str, Any]]:
    box = _first(resolve(wb, COOLING_BOX))
    if not (isinstance(box, str) and box.strip().lower() == "x"):
        return []
    capacity = _num(_first(resolve(wb, "Kuehlgeraete_Umluft_Kuehlung_Leistung")))
    ratio = _num(_first(resolve(wb, "Kuehlgeraete_Umluft_Kuehlung_JAZ")))
    airflow = _num(_first(resolve(wb, "Kuehlgeraete_Umluft_Kuehlung_Volumenstrom")))
    seer = ratio * BTU_PER_WH if ratio and ratio > 0 else None
    return [
        item(
            equipment_type="cooling_unit",
            name=COOLING_NAME,
            capacity=capacity if capacity and capacity > 0 else None,
            capacity_unit="kW" if capacity and capacity > 0 else None,
            efficiency_value=seer,
            efficiency_type="SEER" if seer else None,
            airflow_m3h=airflow,
            airflow_cfm=None if airflow is None else airflow * M3H_TO_CFM,
            source=f"{COOLING_BOX}; Kuehlgeraete_Umluft_Kuehlung_Leistung, _JAZ (W/W x 3.412)",
        )
    ]


def _heat_pumps(wb: Workbook, item) -> list[dict[str, Any]]:
    sheet = sheet_of(wb, HEAT_PUMP_COUNT)
    if sheet is None:
        return []
    shares = _shares(wb, sheet)
    count = _num(_first(resolve(wb, HEAT_PUMP_COUNT)))
    out = []
    services = [(HEATING_SELECTION, "heat_pump", shares.get(HEATING_SHARE))]
    if count == 2:
        services.append((DHW_SELECTION, "dhw_heat_pump", shares.get(DHW_SHARE)))
    for selection, equipment_type, share in services:
        if share is None or share <= 0:
            continue
        unit = _slot_unit(wb, selection)
        if unit is None:
            continue
        name, slot, rating = unit
        source = f"{selection} -> WP_WP{slot}"
        if rating:
            source += (
                f"; test point {rating['point']} "
                f"({rating['source']} °C source, {rating['sink']} °C sink)"
            )
        else:
            source += "; test points not found"
        out.append(
            item(
                equipment_type=equipment_type,
                name=name,
                capacity=rating and rating["capacity"],
                capacity_unit=rating and "kW",
                efficiency_value=rating and rating["cop"],
                efficiency_type=rating and "COP",
                source=source,
            )
        )
    return out


def _shares(wb: Workbook, sheet: str) -> dict[str, float]:
    found = {}
    for row in wb[sheet].iter_rows(
        min_row=_SHARE_ROWS.start,
        max_row=_SHARE_ROWS.stop - 1,
        min_col=column_index_from_string("I"),
        max_col=column_index_from_string("M"),
        values_only=True,
    ):
        label, value = row[0], _num(row[4])
        if (
            isinstance(label, str)
            and label.strip() in (HEATING_SHARE, DHW_SHARE)
            and value is not None
        ):
            found.setdefault(label.strip(), value)
    return found


def _slot_unit(wb: Workbook, selection: str) -> tuple[str, int, dict[str, Any] | None] | None:
    """(name, slot, rating) for a selection, or None when its slot's name is
    not the selected unit's (a layout that moved reads nothing)."""
    match = _SLOT.match(str(_first(resolve(wb, selection)) or ""))
    if match is None:
        return None
    slot, name = int(match.group(1)), match.group(2)
    slot_name = _first(resolve(wb, f"WP_WP{slot}_Name"))
    if not isinstance(slot_name, str) or slot_name.strip() != name:
        return None
    return name, slot, _rating(wb, f"WP_WP{slot}_Leistungsdaten")


def _rating(wb: Workbook, name: str) -> dict[str, Any] | None:
    where = cells_of(wb, name)
    if where is None:
        return None
    sheet, first, last = where
    (c1, r1), (c2, r2) = coordinate_from_string(first), coordinate_from_string(last)
    best = None
    for n, (source, sink, capacity, cop) in enumerate(
        wb[sheet].iter_rows(
            min_row=r1,
            max_row=r2,
            min_col=column_index_from_string(c1),
            max_col=column_index_from_string(c2),
            values_only=True,
        )
    ):
        capacity, cop = _num(capacity), _num(cop)
        if capacity is None or cop is None or capacity <= 0 or cop <= 0:
            continue
        if best is None or capacity > best["capacity"]:
            best = {
                "point": n + 1,
                "source": source,
                "sink": sink,
                "capacity": capacity,
                "cop": cop,
            }
    return best
