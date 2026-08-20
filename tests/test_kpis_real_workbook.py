"""Ground-truth integration tests for PHPP 10.6 IP, against a real workbook.

Skipped when the file isn't present, so CI stays clean. Run locally to catch
drift between our reads and what a real 10.6 IP file actually contains.

**On the fixture, because it changed and should not change back.** These tests
used to point at ``PHPP_V10.6_IP_Example.xlsx``. That file was replaced in place
on 2026-08-05 with a build whose SI source sheets are blank, so its IP sheets --
which are pure conversions of them -- resolve to nothing:

    Heating!Q78       formula '=IF(ISNUMBER(...), ... *cvKBTUsf,"")'  cached None
    Verification!I39  cached '-'   <- the else branch: SI source is not a number

Every value the old ground truth pinned is genuinely absent from it now, not
merely relocated, so it cannot serve as a fixture at all. Re-pinning against it
would pin nulls and assert nothing.

The replacement is a populated project file, and it is a **stronger** fixture
than the original: it is production artifact 18, whose KPIs were verified during
PAS-84 against ``Phpp::Extract#read_canonical_kpis`` in ps-rails -- two
implementations, two languages, one workbook, agreeing to 12+ significant
figures. The old values were a single dump with nothing to check them against.
"""

import pathlib
from math import isclose

import pytest

from app.parser import parse_workbook

# Production artifact 18 ("Donovan Residence"). The odd "^th" is the filename
# as it exists; do not silently "correct" it or the fixture stops resolving.
REAL_WORKBOOK = (
    pathlib.Path.home() / "Downloads" / "260309 PHPP 6840 E ^th Ave Pkwy v 10.6.xlsx"
)

pytestmark = pytest.mark.skipif(
    not REAL_WORKBOOK.is_file(),
    reason=f"PHPP 10.6 IP fixture not found at {REAL_WORKBOOK}",
)


# Pinned 2026-08-19. Cross-verified against Phpp::Extract's canonical cell reads
# (PAS-84) rather than dumped from this parser alone -- a value that agrees only
# with the code that produced it is not ground truth.
GROUND_TRUTH = {
    "tfa": 3388.416556,
    "heating_demand": 7.387608892298333,
    "cooling_demand": 1.780666641931918,
    "source_eui": 16.262043710960263,  # PER!V68
    "pe_demand": 31.63115223023126,  # PER!X68
    "heating_peak": 16588.439212284404,  # max(P88, R88)
    "cooling_peak": 6174.1819636443715,  # max(P64, R64)
}

# These are deterministic cell reads, so anything beyond floating-point noise is
# a real change. The previous 1e-3 would have let a 0.1% drift through silently.
TOL = 1e-9


@pytest.fixture(scope="module")
def real_kpis() -> dict:
    return parse_workbook(REAL_WORKBOOK.read_bytes(), "EN_10_6IP")["kpis"]


def _close(actual: float | None, expected: float) -> bool:
    return actual is not None and isclose(actual, expected, rel_tol=TOL)


def test_real_tfa(real_kpis):
    assert _close(real_kpis["tfa"]["value"], GROUND_TRUTH["tfa"])
    assert real_kpis["tfa"]["unit"] == "ft2"


def test_real_heating_demand(real_kpis):
    assert _close(real_kpis["heating_demand"]["value"], GROUND_TRUTH["heating_demand"])
    assert real_kpis["heating_demand"]["unit"] == "kBtu/ft2yr"


def test_real_cooling_demand(real_kpis):
    assert _close(real_kpis["cooling_demand"]["value"], GROUND_TRUTH["cooling_demand"])
    assert real_kpis["cooling_demand"]["unit"] == "kBtu/ft2yr"


def test_real_heating_peak(real_kpis):
    assert _close(
        real_kpis["peak_loads"]["heating"]["value"], GROUND_TRUTH["heating_peak"]
    )
    assert real_kpis["peak_loads"]["heating"]["unit"] == "Btu/h"


def test_real_cooling_peak(real_kpis):
    assert _close(
        real_kpis["peak_loads"]["cooling"]["value"], GROUND_TRUTH["cooling_peak"]
    )
    assert real_kpis["peak_loads"]["cooling"]["unit"] == "Btu/h"


def test_real_source_eui(real_kpis):
    assert _close(real_kpis["source_eui"]["value"], GROUND_TRUTH["source_eui"])


def test_real_pe_demand(real_kpis):
    assert _close(real_kpis["pe_demand"]["value"], GROUND_TRUTH["pe_demand"])


def test_reads_the_ip_panes(real_kpis):
    """A value can be right by accident; the source string proves which pane it
    came from. This is the EN_10_6IP half of the check that caught EN_9_7IP
    reading the SI sheets."""
    assert real_kpis["heating_demand"]["source"].startswith("Heating!")
    assert real_kpis["cooling_demand"]["source"].startswith("Cooling!")
    assert real_kpis["peak_loads"]["heating"]["source"].startswith("Heating load!")
    assert real_kpis["peak_loads"]["cooling"]["source"].startswith("Cooling load!")
