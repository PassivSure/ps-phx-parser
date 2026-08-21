# frozen_string_literal equivalent not needed in Python
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from app.named_ranges import resolve


def _wb_with_names():
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventilation SI"
    ws["K91"] = "01ud-Swegon Casa R7"
    ws["L91"] = None
    ws["M91"] = "extra"
    ws["B4"] = 306.924
    wb.defined_names.add(DefinedName("single", attr_text="'Ventilation SI'!$B$4"))
    wb.defined_names.add(DefinedName("multi", attr_text="'Ventilation SI'!$K$91:$M$91"))
    wb.defined_names.add(DefinedName("bad_sheet", attr_text="'No Such Sheet'!$A$1"))
    wb.defined_names.add(DefinedName("unparseable", attr_text="#REF!"))
    return wb


def test_single_cell_returns_the_scalar():
    assert resolve(_wb_with_names(), "single") == 306.924


def test_multi_cell_returns_all_non_empty_values_not_just_the_first():
    # The defect this guards: taking the first cell is right often enough
    # to look correct, and wrong wherever the useful value is not first.
    assert resolve(_wb_with_names(), "multi") == ["01ud-Swegon Casa R7", "extra"]


def test_absent_name_returns_none():
    assert resolve(_wb_with_names(), "nope") is None


def test_missing_sheet_returns_none_rather_than_raising():
    assert resolve(_wb_with_names(), "bad_sheet") is None


def test_unparseable_reference_returns_none_rather_than_raising():
    assert resolve(_wb_with_names(), "unparseable") is None


def test_multi_cell_that_is_entirely_empty_returns_none():
    wb = _wb_with_names()
    wb["Ventilation SI"]["K91"] = None
    wb["Ventilation SI"]["M91"] = None
    assert resolve(wb, "multi") is None
