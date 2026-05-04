"""Ground-truth integration tests against a real PHPP file.

Skipped when the file isn't present (so CI stays clean). Run locally to
catch drift between our reads and what the canonical 10.6 IP example
actually contains. Values pinned 2026-05-04 from
``~/Downloads/PHPP_V10.6_IP_Example.xlsx``.
"""

import pathlib
from math import isclose

import pytest

from app.parser import parse_workbook

REAL_WORKBOOK = pathlib.Path.home() / "Downloads" / "PHPP_V10.6_IP_Example.xlsx"

pytestmark = pytest.mark.skipif(
    not REAL_WORKBOOK.is_file(),
    reason=f"Canonical PHPP fixture not found at {REAL_WORKBOOK}",
)


# Pinned 2026-05-04 — values dumped directly from the workbook.
GROUND_TRUTH = {
    "tfa": 9057.74975951061,
    "heating_demand_specific": 1.6973902774843397,
    "cooling_demand_specific": 2.6761328160753366,
    "heating_peak": 24959.76432206965,  # max(w1=24185.58, w2=24959.76)
    "cooling_peak": 21853.312457761673,  # w1 == w2 in this file
    "source_eui": 7.342632299379407,  # PER!V68 (PER total)
    "pe_demand": 15.243693837362619,  # PER!X68 (PE total)
}


@pytest.fixture(scope="module")
def real_kpis() -> dict:
    return parse_workbook(REAL_WORKBOOK.read_bytes(), "EN_10_6IP")["kpis"]


def _close(actual: float | None, expected: float, tol: float = 0.001) -> bool:
    return actual is not None and isclose(actual, expected, rel_tol=tol)


def test_real_tfa(real_kpis):
    assert _close(real_kpis["tfa"]["value"], GROUND_TRUTH["tfa"])
    assert real_kpis["tfa"]["unit"] == "ft2"


def test_real_heating_demand(real_kpis):
    assert _close(
        real_kpis["heating_demand"]["value"],
        GROUND_TRUTH["heating_demand_specific"],
    )


def test_real_cooling_demand(real_kpis):
    assert _close(
        real_kpis["cooling_demand"]["value"],
        GROUND_TRUTH["cooling_demand_specific"],
    )


def test_real_heating_peak(real_kpis):
    assert _close(real_kpis["peak_loads"]["heating"]["value"], GROUND_TRUTH["heating_peak"])


def test_real_cooling_peak(real_kpis):
    assert _close(real_kpis["peak_loads"]["cooling"]["value"], GROUND_TRUTH["cooling_peak"])


def test_real_source_eui(real_kpis):
    assert _close(real_kpis["source_eui"]["value"], GROUND_TRUTH["source_eui"])


def test_real_pe_demand(real_kpis):
    assert _close(real_kpis["pe_demand"]["value"], GROUND_TRUTH["pe_demand"])
