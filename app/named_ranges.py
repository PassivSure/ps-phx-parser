"""Resolve a workbook's defined names to their values.

PHPP addresses most of its equipment data by defined name rather than by cell,
and those names have a property the cell maps elsewhere in this service do not:
they always resolve to the METRIC pane. On an IP workbook the unprefixed names
point at `Ventilation SI`, `HP SI` and so on; on an SI workbook they point at
the unsuffixed sheets holding the same data. Either way the caller gets metric.

That is why equipment reads go through here rather than through a per-edition
cell map -- there is no pane to choose, and choosing the wrong one is the defect
class that produced four separate fixes in the EN_9_7IP work.
"""

from __future__ import annotations

import logging
import re
from typing import Any

from openpyxl.workbook import Workbook

logger = logging.getLogger(__name__)

# "'Ventilation SI'!$K$91:$M$91" or "Ventilation!$B$4"
_REF = re.compile(r"^'?([^'!]+)'?!(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?$")


def resolve(wb: Workbook, name: str) -> Any | None:
    """The value(s) behind a defined name, or None.

    Single cell -> the scalar. Multi-cell -> the list of non-empty values,
    deliberately NOT the first one: roughly half of PHPP's equipment names are
    ranges, and a resolver that took the first cell would be right often enough
    to pass a casual test and wrong wherever the value sits elsewhere in the
    range.

    Returns None rather than raising for every failure mode -- an absent name,
    a sheet the workbook does not have, an unparseable reference such as #REF!.
    Equipment is one subtree of five and must never take a parse down with it.
    """
    defined = wb.defined_names.get(name)
    if defined is None:
        return None

    match = _REF.match(str(defined.value).strip())
    if match is None:
        return None

    sheet_name, first, last = match.groups()
    if sheet_name not in wb.sheetnames:
        # An absent NAME is ordinary -- most of PHPP's names are unused in any
        # given file, and logging those would be pure noise. A name that exists
        # and points at a sheet the workbook does not have is not ordinary: it
        # means the mapping and the file disagree, which is worth seeing.
        logger.warning(
            "[equipment] defined name %r points at missing sheet %r", name, sheet_name
        )
        return None
    ws = wb[sheet_name]

    if last is None:
        return ws[first.replace("$", "")].value

    values = [
        cell.value
        for row in ws[f"{first}:{last}"]
        for cell in row
        if cell.value is not None
    ]
    return values or None
