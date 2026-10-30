"""Assemblies on real v10 workbooks, skipped when the files are absent.

U-values read by hand off each file's Components SI table (column K) and
layer counts off its U-values SI blocks; areas reconcile with what the
production extraction totals for the same assemblies (Donovan: 03ud 320.73 ft2
over groups 8 and 14). PAS-79 recorded PHI's example independently as R 41.6
(wall) and R 110.5 (roof) in hr·ft²·°F/Btu.
"""

import pathlib

import pytest

from app.parser import parse_workbook

DOWNLOADS = pathlib.Path.home() / "Downloads"
W_M2K_PER_BTU = 5.678263  # W/(m²K) per Btu/(h·ft²·°F)

CASES = {
    # v9.7 labels its blocks "Building assembly description"; the names and
    # the Components table carry over.
    "55th_9_7_ip": ("250121_55th_IP9_741zhOb_FINAL.xlsx", "EN_9_7IP", {
        "01ud": ("wall", 0.197, 4, 2614.92),
        "02ud": ("roof", 0.1283, 4, 1679.68),
        "03ud": ("slab_on_grade", 0.2786, 2, 1315.25),
        "05ud": ("roof", 0.1188, 4, 215.40),
        "89ud": ("door", 0.5, 0, 0.0),          # PHPP's own "PH External Door"
    }),
    "donovan_10_6_ip": ("260309 PHPP 6840 E ^th Ave Pkwy v 10.6.xlsx", "EN_10_6IP", {
        # id: (component, u_value_w_m2k, layer count, area_ft2)
        "01ud": ("slab_on_grade", 0.1485, 2, 1667.97),
        "02ud": ("wall", 0.1542, 4, 1508.23),
        "03ud": ("wall", 0.1939, 5, 320.73),
        "04ud": ("roof", 0.1093, 7, 1129.50),
        "05ud": ("roof", 0.1043, 7, 793.10),
        "06ud": ("roof", 0.0695, 7, 215.35),
        "84ud": ("roof", 0.15, 0, 38.72),
        "10ud": ("wall", 0.1247, 5, 3336.53),   # past row 200: Claude's read gets no layers
        "83ud": ("wall", 0.15, 0, 24.0),        # PHPP's own "PH External wall": no block
        "88ud": (None, 0.15, 0, 12.63),         # group 14 only: no type
    }),
    "phi_10_6_ip_example": ("PHPP_V10.6_IP_Example.xlsx", "EN_10_6IP", {
        "01ud": ("wall", 0.1364, 5, 5517.58),
        "02ud": ("roof", 0.0513, 6, 9855.61),
        "03ud": ("slab_on_grade", 0.0969, 2, 9056.11),
    }),
    "easyph_example": ("PHPP_EN_V10.6_easyPH_Example.xlsx", "EN_10_6", {
        "01ud": ("wall", 0.1375, 4, 1983.53),
        "02ud": ("roof", 0.1075, 3, 897.84),
        "03ud": ("slab_on_grade", 0.1308, 6, 871.07),
    }),
    "holmes_easyph": ("2536 Holmes_easyPH_260512.xlsx", "EN_10_6", {
        "01ud": ("slab_on_grade", 0.1842, 2, 2095.83),
        "02ud": ("wall", 0.2019, 5, 1433.55),
        "03ud": ("wall", 0.1234, 6, 3655.96),
        "04ud": ("roof", 0.0541, 7, 2096.29),
    }),
}


@pytest.fixture(scope="module", params=sorted(CASES), ids=sorted(CASES))
def case(request):
    file, version, expected = CASES[request.param]
    path = DOWNLOADS / file
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    return expected, parse_workbook(path.read_bytes(), version)["envelope"]


def test_reads_every_assembly_the_surfaces_use(case):
    expected, envelope = case
    got = {a["id"]: a for a in envelope["assemblies"]}
    assert set(got) == set(expected)
    for assembly_id, (component, u_value, layers, area) in expected.items():
        a = got[assembly_id]
        assert a["component"] == component, assembly_id
        assert a["u_value_w_m2k"] == pytest.approx(u_value, abs=5e-5), assembly_id
        assert len(a["layers"]) == layers, assembly_id
        assert a["area_ft2"] == pytest.approx(area, abs=0.01), assembly_id


def test_every_opaque_surface_points_at_an_assembly_it_read(case):
    _, envelope = case
    ids = {a["id"] for a in envelope["assemblies"]}
    opaque = [c for c in envelope["components"] if c["component"] != "thermal_bridge"]
    linked = [c for c in opaque if c["assembly_id"]]
    assert linked and {c["assembly_id"] for c in linked} <= ids


def test_layers_carry_thickness_and_conductivity(case):
    _, envelope = case
    for a in envelope["assemblies"]:
        for layer in a["layers"]:
            assert layer["thickness_mm"], (a["id"], layer)
            assert layer["material"] and layer["conductivity_w_mk"], (a["id"], layer)


def test_phi_example_matches_pas_79s_r_values():
    path = DOWNLOADS / "PHPP_V10.6_IP_Example.xlsx"
    if not path.is_file():
        pytest.skip(f"fixture not found at {path}")
    envelope = parse_workbook(path.read_bytes(), "EN_10_6IP")["envelope"]
    got = {a["id"]: a for a in envelope["assemblies"]}
    assert W_M2K_PER_BTU / got["01ud"]["u_value_w_m2k"] == pytest.approx(41.6, abs=0.1)
    assert W_M2K_PER_BTU / got["02ud"]["u_value_w_m2k"] == pytest.approx(110.5, abs=0.3)
