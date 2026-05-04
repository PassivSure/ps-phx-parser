import pytest

from app.parser import (
    ParseError,
    available_versions,
    load_shape,
    parse_workbook,
)


def test_available_versions_includes_en_10_6ip():
    versions = available_versions()
    assert "EN_10_6IP" in versions
    assert "EN_10_6" in versions


def test_load_shape_caches_result():
    a = load_shape("EN_10_6IP")
    b = load_shape("EN_10_6IP")
    assert a is b  # lru_cache hit


def test_load_shape_unknown_version_raises():
    with pytest.raises(ParseError, match="Unknown phpp_version"):
        load_shape("XX_99_9")


def test_parse_workbook_returns_kpis_subtree(workbook_bytes_with_kpis):
    """parse_workbook emits the kpis subtree with all measurements present."""
    result = parse_workbook(workbook_bytes_with_kpis, "EN_10_6IP")
    kpis = result["kpis"]

    assert set(kpis.keys()) == {
        "tfa",
        "heating_demand",
        "cooling_demand",
        "peak_loads",
        "source_eui",
        "pe_demand",
    }
    assert set(kpis["peak_loads"].keys()) == {"heating", "cooling"}


def test_parse_workbook_returns_nulls_when_sheets_missing(workbook_bytes):
    """Workbook without Heating/Cooling/PER sheets — measurements emit
    value=None but keep their unit so the consumer knows what was tried."""
    result = parse_workbook(workbook_bytes, "EN_10_6IP")
    kpis = result["kpis"]

    assert kpis["tfa"]["value"] is None
    assert kpis["heating_demand"]["value"] is None
    assert kpis["cooling_demand"]["value"] is None
    assert kpis["source_eui"]["value"] is None
    assert kpis["pe_demand"]["value"] is None
    # units still populated from the shape
    assert kpis["tfa"]["unit"] == "ft2"
    assert kpis["heating_demand"]["unit"] == "kBtu/ft2yr"


def test_parse_workbook_rejects_unknown_version(workbook_bytes):
    with pytest.raises(ParseError, match="Unknown phpp_version"):
        parse_workbook(workbook_bytes, "XX_99_9")
