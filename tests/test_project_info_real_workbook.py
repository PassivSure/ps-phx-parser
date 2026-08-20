"""project_info read from real workbooks, across both layouts.

Skipped when the files aren't present. Covers what synthetic fixtures cannot:
that the label vocabulary matches what PHPP actually writes. Two v9.7 files and
one v10.6, because the blocks this module reads differ between the versions in
both position and wording.
"""

import pathlib

import pytest

from app.parser import parse_workbook

DOWNLOADS = pathlib.Path.home() / "Downloads"

CASES = {
    "55th_v97": {
        "file": "250121_55th_IP9_741zhOb_FINAL.xlsx",
        "version": "EN_9_7IP",
        "project_name": "Younis/Pang Residence",
        "postal_code": "95819",
        # "Caloifornia" is a typo in the source workbook. Pinned as-is: the
        # parser reports what the file says and does not correct it.
        "location_string": "643 55th Street, Sacramento, Caloifornia 95819",
        "organizations": [
            ("Home owner", "Maria Pang/Laith Younis"),
            ("Architect", "Bronwyn Barry/Passive House BB"),
            ("Building services", "Essential Air"),
        ],
    },
    "17mile_v97": {
        "file": "250901.2821 17 mile drive.IP9.7.final.xlsx",
        "version": "EN_9_7IP",
        "project_name": "Asnis Residence",
        "postal_code": "93953",
        "location_string": (
            "2821 Seventeen Mile Drive, Pebble Beach, California 93953"
        ),
        "organizations": [
            ("Home owner", "Anna and Ilya Asnis"),
            ("Architect", "Passive House BB"),
            ("Building services", "N/A"),
        ],
    },
    "6840_v106": {
        "file": "260309 PHPP 6840 E ^th Ave Pkwy v 10.6.xlsx",
        "version": "EN_10_6IP",
        "project_name": "Donovan Residence",
        "postal_code": "80220",
        # The province cell holds "80211" in this workbook — a data-entry
        # error by its author, read faithfully. Unchanged by the move to
        # label location, which is the point of pinning it.
        "location_string": "6840 E. 6th Ave. Pky, Denver, 80211 80220",
        "organizations": [
            ("Home owner", "Billy and Nicole Donovan"),
            ("Architect", "Jewkes"),
            ("Mechanical engineer", "Point 6 LLC"),
            ("Energy consultant", "Point 6 LLC"),
            ("Certification body", "E Mod Studios"),
        ],
    },
}


@pytest.fixture(scope="module", params=sorted(CASES), ids=sorted(CASES))
def case(request):
    spec = CASES[request.param]
    path = DOWNLOADS / spec["file"]
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    info = parse_workbook(path.read_bytes(), spec["version"])["project_info"]
    return spec, info


def test_project_name(case):
    spec, info = case
    assert info["project_name"] == spec["project_name"]


def test_postal_code_is_a_postcode(case):
    """On v9.7 this returned the province — "Caloifornia" — because the
    address block sits a row higher than v10's."""
    spec, info = case
    assert info["postal_code"] == spec["postal_code"]
    assert info["postal_code"].replace("-", "").isdigit()


def test_location_string(case):
    spec, info = case
    assert info["location_string"] == spec["location_string"]


def test_organizations(case):
    """On v9.7 this returned [("", "20"), ("", "2.36"), ("", "Specific
    demand")] — an interior-temperature setpoint and a heat-gain figure read
    as firms."""
    spec, info = case
    assert [
        (o["label"], o["name"]) for o in info["organizations"]
    ] == spec["organizations"]


def test_no_organization_name_is_numeric(case):
    """The shape of the old failure, guarded independently of the pinned
    list: a firm name is never a bare number."""
    _, info = case
    for org in info["organizations"]:
        assert not org["name"].replace(".", "").replace("-", "").isdigit()
