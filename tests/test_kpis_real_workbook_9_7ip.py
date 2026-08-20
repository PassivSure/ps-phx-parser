"""Ground-truth tests for PHPP 9.7 IP, guarding the EN_9_7IP shape override.

Skipped when the files aren't present, like the 10.6 IP suite. Two workbooks,
not one: the override repoints four sections from the SI panes to the IP ones,
and a coordinate confirmed on a single file has been this project's largest
defect source. Both files must agree.

Ground truth is `Phpp::Extract#read_canonical_kpis` in ps-rails — deterministic
cell reads, independently verified, in a different language. Two
implementations agreeing on one workbook is the strongest check available here.
"""

import pathlib
from math import isclose

import pytest

from app.parser import parse_workbook

DOWNLOADS = pathlib.Path.home() / "Downloads"

# Values pinned 2026-08-19.
WORKBOOKS = {
    "55th": {
        "path": DOWNLOADS / "250121_55th_IP9_741zhOb_FINAL.xlsx",
        "tfa": 2082.38,  # == canonical Verification!I34
        "heating_demand": 3.6166669768806825,  # == canonical Verification!I35
        # Canonical reads Verification!I38 = 4.216018735957633, the *overall*
        # specific cooling demand. The parser reads Cooling!Q88, PHPP's
        # *sensible* annual demand — a different quantity, 0.04% lower here
        # because this building has a latent component. Not a discrepancy to
        # chase: on the 17-mile file the two cells are bit-identical, which is
        # what a zero-latent building looks like.
        "cooling_demand": 4.214346069950888,
        "heating_peak": 6874.135536961924,  # max(P87=6874.14, R87=6298.17)
        "cooling_peak": 7716.21386353818,  # max(P60=7716.21, R60=5979.83)
        # PER!W18 / Y18 — the v9 summary, which sits ABOVE the breakdown and
        # carries no "Total energy demand" label. Null before that was handled.
        "source_eui": 14.51462012718278,
        "pe_demand": 33.39526371320298,
    },
    "17mile": {
        "path": DOWNLOADS / "250901.2821 17 mile drive.IP9.7.final.xlsx",
        "tfa": 2457.4500000000003,
        "heating_demand": 2.9081858730439696,
        "cooling_demand": 1.4560853821638844,  # == canonical Verification!I38
        "heating_peak": 5724.443906475532,
        "cooling_peak": 3729.8902616382657,
        "source_eui": 7.205476675544691,
        "pe_demand": 15.655987902420204,
    },
}


@pytest.fixture(scope="module", params=sorted(WORKBOOKS), ids=sorted(WORKBOOKS))
def case(request):
    spec = WORKBOOKS[request.param]
    if not spec["path"].is_file():
        pytest.skip(f"v9.7 IP fixture not found at {spec['path']}")
    kpis = parse_workbook(spec["path"].read_bytes(), "EN_9_7IP")["kpis"]
    return spec, kpis


def _close(actual, expected, tol=1e-6):
    return actual is not None and isclose(actual, expected, rel_tol=tol)


def test_tfa_is_ip(case):
    """The regression this guards: as shipped, EN_9_7IP read Cooling SI!O8 and
    returned 193.46 m2 for this building."""
    spec, kpis = case
    assert _close(kpis["tfa"]["value"], spec["tfa"])
    assert kpis["tfa"]["unit"] == "ft2"


def test_heating_demand_is_ip(case):
    spec, kpis = case
    assert _close(kpis["heating_demand"]["value"], spec["heating_demand"])
    assert kpis["heating_demand"]["unit"] == "kBtu/ft2yr"


def test_cooling_demand_is_ip(case):
    spec, kpis = case
    assert _close(kpis["cooling_demand"]["value"], spec["cooling_demand"])
    assert kpis["cooling_demand"]["unit"] == "kBtu/ft2yr"


def test_peak_loads_are_ip(case):
    spec, kpis = case
    assert _close(kpis["peak_loads"]["heating"]["value"], spec["heating_peak"])
    assert _close(kpis["peak_loads"]["cooling"]["value"], spec["cooling_peak"])
    assert kpis["peak_loads"]["heating"]["unit"] == "Btu/h"
    assert kpis["peak_loads"]["cooling"]["unit"] == "Btu/h"


def test_per_totals_are_found_and_ip(case):
    """v9.7 has no "Total energy demand" row, so both of these were null on
    every v9.7 file until the summary was located by its in-column header."""
    spec, kpis = case
    assert _close(kpis["source_eui"]["value"], spec["source_eui"])
    assert _close(kpis["pe_demand"]["value"], spec["pe_demand"])
    assert kpis["source_eui"]["unit"] == "kBtu/ft2yr"
    assert kpis["pe_demand"]["unit"] == "kBtu/ft2yr"
    assert kpis["source_eui"]["source"] == "PER!W18"
    assert kpis["pe_demand"]["source"] == "PER!Y18"


def test_reads_name_the_ip_panes(case):
    """A value can be right by accident; the source string proves which pane
    it came from."""
    _, kpis = case
    assert kpis["heating_demand"]["source"].startswith("Heating!")
    assert kpis["cooling_demand"]["source"].startswith("Cooling!")
    assert kpis["peak_loads"]["heating"]["source"].startswith("Heating load!")
    assert kpis["peak_loads"]["cooling"]["source"].startswith("Cooling load!")
