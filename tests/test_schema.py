"""Contract tests for schema/output.schema.json.

The schema is the long-lived contract between this parser and ps-rails.
Phase 2 (PAS-62) parser modules will populate kpis/envelope/hvac_equipment;
v1 is permissive so partial responses validate during the build.
"""

import json
import pathlib

import httpx
import pytest
import respx
from fastapi.testclient import TestClient
from jsonschema import Draft202012Validator, validate

from app.main import app
from tests.conftest import parse_and_wait

REPO_ROOT = pathlib.Path(__file__).resolve().parent.parent
SCHEMA_PATH = REPO_ROOT / "schema" / "output.schema.json"
WORKBOOK_URL = "https://example.com/signed/phpp.xlsx"

client = TestClient(app)


@pytest.fixture(scope="module")
def schema() -> dict:
    return json.loads(SCHEMA_PATH.read_text())


def test_schema_is_valid_draft_2020_12(schema):
    """The schema doc itself is well-formed under Draft 2020-12."""
    Draft202012Validator.check_schema(schema)


@respx.mock
def test_parse_response_validates_against_schema(workbook_bytes_with_version, schema):
    """Live /parse response (basic fixture) conforms to the v1 envelope."""
    respx.get(WORKBOOK_URL).mock(
        return_value=httpx.Response(200, content=workbook_bytes_with_version)
    )

    settled = parse_and_wait(client, WORKBOOK_URL)
    assert settled["status"] == "done", settled.get("detail")

    validate(instance=settled["result"], schema=schema)


@respx.mock
def test_parse_response_with_all_subtrees_validates(shape_en_10_6ip, schema):
    """Live /parse against a fixture exercising kpis + envelope + project_info.
    Catches schema drift introduced by any P2.x ticket — would have flagged
    the source/null mismatch we hit during P2.2."""
    from tests.conftest import build_workbook_bytes

    body = build_workbook_bytes(
        shape_en_10_6ip,
        with_kpis=True,
        with_envelope=True,
        with_project_info=True,
        data_sheet_version="10.6 IP",
    )
    respx.get(WORKBOOK_URL).mock(return_value=httpx.Response(200, content=body))

    settled = parse_and_wait(client, WORKBOOK_URL)
    assert settled["status"] == "done", settled.get("detail")

    body = settled["result"]
    assert body["kpis"]["tfa"]["value"] == 2400.0
    assert len(body["envelope"]["components"]) > 0
    assert body["envelope"]["airtightness"]["n50_ach"] == 0.6
    assert body["project_info"]["project_name"] == "Test Passive House"
    assert body["project_info"]["postal_code"] == "94028"
    assert len(body["project_info"]["organizations"]) == 5

    validate(instance=body, schema=schema)


def test_schema_accepts_full_phase2_response(schema):
    """A hand-built response with every Phase-2 subtree populated still validates.

    This is a forward-looking smoke test — the parser doesn't emit these yet,
    but it pins the shape we're committing to so P2.2-P2.5 don't drift.
    """
    full_response = {
        "schema_version": "1.0.0",
        "parser": {
            "name": "ps-phx-parser",
            "version": "ps-phx-parser/0.1.0+PHX-1.56.51",
            "phpp_version": "EN_10_6IP",
            "parsed_at": "2026-05-01T12:00:00+00:00",
            "file_hash_sha256": None,
        },
        "units": {
            "system": "IP",
            "energy": "kBtu",
            "area": "ft2",
            "power": "Btu/h",
            "heat_flow": "Btu/h",
            "temperature": "F",
        },
        "kpis": {
            "tfa": {"value": 2400.0, "unit": "ft2", "source": "Verification!I35"},
            "heating_demand": {"value": 4.2, "unit": "kBtu/ft2yr"},
            "cooling_demand": {"value": 1.1, "unit": "kBtu/ft2yr"},
            "site_eui": {"value": 24.0, "unit": "kBtu/ft2yr"},
            "source_eui": {"value": 38.0, "unit": "kBtu/ft2yr"},
            "pe_demand": {"value": 38.0, "unit": "kBtu/ft2yr"},
            "peak_loads": {
                "heating": {
                    "value": 24186,
                    "unit": "Btu/h",
                    "value_per_area": 10.1,
                    "unit_per_area": "Btu/hft2",
                },
                "cooling": {"value": None, "unit": "Btu/h"},
            },
        },
        "envelope": {
            "components": [
                {
                    "component": "wall",
                    "label": "North wall",
                    "area_ft2": 800.0,
                    "u_value_Btuh_ft2F": 0.04,
                    "r_value_hr_ft2F_Btu": 25.0,
                },
                {"component": "window", "label": "South glazing", "area_ft2": 120.0,
                 "u_value_Btuh_ft2F": 0.18},
                {"component": "thermal_bridge", "label": "Slab edge",
                 "length_ft": 110.0, "psi_value_Btuh_ftF": 0.05},
            ],
            "airtightness": {"n50_ach": 0.6, "source": "Verification!N32"},
        },
        "ventilation": {"system": {"hrv_efficiency_pct": 84.0}},
        "hvac_equipment": [
            {"equipment_type": "heat_pump", "name": "Mitsubishi MUZ-FH",
             "manufacturer": "Mitsubishi", "capacity": 12000, "capacity_unit": "Btu/h",
             "efficiency_value": 3.2, "efficiency_type": "COP"},
            {"equipment_type": "hrv", "name": "Zehnder Q450", "airflow_cfm": 200.0,
             "heat_recovery_efficiency_pct": 84.0},
        ],
        "project_info": {
            "project_name": "Sample House",
            "postal_code": "94028",
            "location_string": "378 Balsam Ave, Sunnyvale, CA 94028",
            "verification_complete": True,
            "occupancy_type": "Single-family residential",
            "organizations": [
                {"label": "Architect", "name": "Acme Architects"},
            ],
        },
        "errors": [],
    }
    validate(instance=full_response, schema=schema)
