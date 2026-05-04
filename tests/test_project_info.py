"""Unit tests for app.project_info on synthetic workbooks."""

from io import BytesIO

from openpyxl import load_workbook

from app.parser import load_shape
from app.project_info import read_project_info


def _info(workbook_bytes: bytes) -> dict:
    wb = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        return read_project_info(wb, load_shape("EN_10_6IP"))
    finally:
        wb.close()


def test_project_name_from_overview(workbook_bytes_with_project_info):
    info = _info(workbook_bytes_with_project_info)
    assert info["project_name"] == "Test Passive House"


def test_postal_code_strips_pphpp_quote_wrapping(workbook_bytes_with_project_info):
    """PHPP wraps zips in literal quotes ('"94028"') to keep leading zeros.
    The parser must strip them so the consumer gets a clean 5-digit string."""
    info = _info(workbook_bytes_with_project_info)
    assert info["postal_code"] == "94028"


def test_location_string_synthesized_from_address_cells(
    workbook_bytes_with_project_info,
):
    info = _info(workbook_bytes_with_project_info)
    assert info["location_string"] == "100 Main St, Sunnyvale, CA 94028"


def test_occupancy_type_via_phx_locator(workbook_bytes_with_project_info):
    """phi_building_use_type is read via PHX locator pattern. The raw
    PHX-formatted value (with `10-` prefix) is emitted as-is — the Mapper
    handles the human-friendly translation."""
    info = _info(workbook_bytes_with_project_info)
    assert info["occupancy_type"] == "10-Residential building: Residential"


def test_organizations_label_strips_role_suffix(workbook_bytes_with_project_info):
    """Overview labels look like 'Architect name / E-mail'. We keep the
    role prefix so the schema's `label` field carries something meaningful."""
    info = _info(workbook_bytes_with_project_info)
    by_label = {o["label"]: o["name"] for o in info["organizations"]}
    assert by_label["Architect"] == "Smith Architects"
    assert by_label["Mechanical engineer"] == "Cool MEP Inc"
    assert by_label["Energy consultant"] == "EnergyPro"
    assert by_label["Home owner"] == "Acme Owners LLC"
    assert by_label["Certification body"] == "PHIUS"


def test_organizations_skips_empty_rows(shape_en_10_6ip):
    """Real PHPP files have placeholder rows in the org block (e.g. unused
    Civil engineer slot). Empty-name rows must be dropped, not emitted as
    {label: '...', name: ''}."""
    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    ov_ws = wb.create_sheet(shape_en_10_6ip.OVERVIEW.name)
    # Only fill in row 33 (Architect)
    ov_ws["B33"] = "Architect name / E-mail"
    ov_ws["C33"] = "Solo Architect"

    buf = BytesIO()
    wb.save(buf)

    info = _info(buf.getvalue())
    assert info["organizations"] == [{"label": "Architect", "name": "Solo Architect"}]


def test_verification_complete_defaults_false(workbook_bytes_with_project_info):
    """P2.5 doesn't implement the verification_complete heuristic — emits
    False so the Mapper treats projects as design-stage by default. The
    follow-up ticket replaces this with a real Verification-sheet check."""
    info = _info(workbook_bytes_with_project_info)
    assert info["verification_complete"] is False


def test_handles_missing_overview_and_verification(workbook_bytes):
    """Workbook with neither Overview nor populated Verification — emit
    a fully-null project_info rather than crashing."""
    info = _info(workbook_bytes)
    assert info["project_name"] is None
    assert info["postal_code"] is None
    assert info["location_string"] is None
    assert info["occupancy_type"] is None
    assert info["verification_complete"] is False
    assert info["organizations"] == []
