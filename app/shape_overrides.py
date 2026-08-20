"""Local corrections applied over PHX's bundled PHPP shapes.

PHX ships the shape files this service reads (`available_versions()` globs its
`phpp_localization` directory). When one of them is wrong for our purposes, the
alternatives are forking PHX — which this project already knows the cost of, the
easyPH version fix having had to be written twice — or correcting it here.

This is the second option: a *patch*, not a copy. The upstream shape is loaded
normally and only the named keys are replaced, so upstream fixes to the other
1,500 lines still flow through. An override outlives the bug it fixes only until
someone checks; each entry therefore records what is wrong and how to tell when
it can go.
"""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any

# version stem -> {section: {field: value}}
SHAPE_OVERRIDES: dict[str, dict[str, dict[str, Any]]] = {
    # EN_9_7IP reads its energy sections off the SI worksheets.
    #
    #   EN_9_7IP    HEATING name='Heating SI'  unit='KWH'
    #   EN_10_6IP   HEATING name='Heating'     unit='KBTU'
    #
    # An IP PHPP ships both panes — unsuffixed IP sheets and " SI" variants —
    # so each section is *internally* consistent (SI sheet, SI unit) and
    # nothing is corrupted. The defect is that the shape is MIXED: AREAS,
    # UVALUES, VENTILATION and VERIFICATION point at the unsuffixed IP sheets
    # while the energy sections point at SI. One shape therefore reports floor
    # area in ft2 from one section and m2 from another, and its IP sibling
    # answers the identical question the other way. PER's `IP_PE_*` named
    # ranges sitting under name='PER SI' show how it happened: SI sheet names
    # pasted over an IP-derived shape.
    #
    # Overridden here are the four sections that are consumed AND verified.
    # Verified means: the same address on the IP and SI panes holds the same
    # quantity, related by exactly PHPP's own conversion constant — checked on
    # two real v9.7 workbooks (250121_55th_IP9_741zhOb_FINAL.xlsx and
    # 250901.2821 17 mile drive.IP9.7.final.xlsx), both columns, all four
    # sections. Demand ratio 3.154591186 (kWh/m2a per kBtu/ft2yr); peak-load
    # ratio 3.412141156 (Btu/hr per W). Repointing a name onto a differently
    # laid-out sheet would move every number silently, so it is not assumed.
    #
    # End to end, the override makes tfa read 2082.38 ft2 and heating demand
    # 3.6166669768806825 kBtu/ft2yr — matching Phpp::Extract's independently
    # verified canonical cell reads exactly.
    #
    # NOT overridden, deliberately:
    #   PER      — its "Total energy demand" locator matches no row on v9.7 at
    #              all (the label differs), so source_eui/pe_demand are already
    #              null and repointing the sheet would not change that. The IP
    #              layout cannot be verified while the locator does not resolve.
    #   OVERVIEW — self-consistent (name='Overview SI' with M2/M3 units, which
    #              is what that pane holds). The only field read from it is the
    #              project name, identical text on both panes. Repointing the
    #              name alone would make it inconsistent.
    #
    # REMOVE WHEN: upstream PHX ships EN_9_7IP pointing these sections at the
    # unsuffixed sheets with IP units. Checked against PHX 1.56.88.
    "EN_9_7IP": {
        "HEATING_DEMAND": {"name": "Heating", "unit": "KBTU"},
        "COOLING_DEMAND": {"name": "Cooling", "unit": "KBTU"},
        "HEATING_PEAK_LOAD": {"name": "Heating load", "unit": "BTU/HR"},
        "COOLING_PEAK_LOAD": {"name": "Cooling load", "unit": "BTU/HR"},
    },
}


def apply_overrides(version: str, shape_data: Mapping[str, Any]) -> dict[str, Any]:
    """Return `shape_data` with any override for `version` merged in.

    Merges one level into each named section, so an override states only the
    fields it corrects and leaves the rest of that section untouched.
    """
    override = SHAPE_OVERRIDES.get(version)
    if not override:
        return dict(shape_data)

    patched = dict(shape_data)
    for section, fields in override.items():
        if section not in patched:
            # The section vanished upstream — louder than silently inventing it,
            # because an override that no longer matches reality is a bug.
            raise KeyError(
                f"shape override for {version!r} targets section {section!r}, "
                "which the upstream shape no longer has"
            )
        patched[section] = {**patched[section], **fields}
    return patched
