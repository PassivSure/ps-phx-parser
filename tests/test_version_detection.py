from io import BytesIO

import pytest
from openpyxl import Workbook

from app.version_detection import (
    DetectedVersion,
    DetectionError,
    detect_version,
)
from tests.conftest import build_workbook_bytes


class TestShapeStem:
    def test_normalizes_10_6_metric(self):
        v = DetectedVersion(major="10", minor="6", language="EN", raw="10.6")
        assert v.shape_stem == "EN_10_6"

    def test_normalizes_10_6_ip(self):
        v = DetectedVersion(major="10", minor="6 IP", language="EN", raw="10.6 IP")
        assert v.shape_stem == "EN_10_6IP"

    def test_normalizes_v_8_5(self):
        v = DetectedVersion(major="8", minor="5", language="EN", raw="8.5")
        assert v.shape_stem == "EN_8_5"


class TestDetectVersion:
    def test_detects_10_6_ip(self, shape_en_10_6ip):
        bytes_ = build_workbook_bytes(shape_en_10_6ip, data_sheet_version="10.6 IP")

        detected = detect_version(bytes_)

        assert detected.shape_stem == "EN_10_6IP"
        assert detected.raw == "10.6 IP"
        assert detected.language == "EN"

    def test_detects_10_6_metric(self, shape_en_10_6ip):
        bytes_ = build_workbook_bytes(shape_en_10_6ip, data_sheet_version="10.6")

        detected = detect_version(bytes_)

        assert detected.shape_stem == "EN_10_6"

    def test_defaults_to_en_when_pe_factor_is_absent(self, shape_en_10_6ip):
        bytes_ = build_workbook_bytes(
            shape_en_10_6ip,
            data_sheet_version="8.5",
            data_sheet_pe_factor=None,
        )

        detected = detect_version(bytes_)

        assert detected.language == "EN"
        assert detected.shape_stem == "EN_8_5"

    def test_raises_when_no_data_sheet(self, shape_en_10_6ip):
        bytes_ = build_workbook_bytes(shape_en_10_6ip)  # no Data sheet

        with pytest.raises(DetectionError, match="No Data/Daten/Datos sheet"):
            detect_version(bytes_)

    def test_raises_when_no_phpp_marker(self):
        wb = Workbook()
        wb.active.title = "Data"
        wb.active["A1"] = "Something else"

        buf = BytesIO()
        wb.save(buf)

        with pytest.raises(DetectionError, match="pre-v9 PHPP"):
            detect_version(buf.getvalue())

    def test_accepts_localized_daten_sheet(self, shape_en_10_6ip):
        # Build a workbook with "Daten" (German) instead of "Data"
        wb = Workbook()
        wb.remove(wb.active)
        daten = wb.create_sheet("Daten")
        daten["A5"] = "PHPP Version"
        daten["B5"] = "10.6"
        daten["D5"] = "1-PE-FAKTOREN"

        buf = BytesIO()
        wb.save(buf)

        detected = detect_version(buf.getvalue())
        assert detected.language == "DE"
        assert detected.shape_stem == "DE_10_6"

    def test_raises_on_unexpected_version_cell(self):
        wb = Workbook()
        wb.active.title = "Data"
        wb.active["A5"] = "PHPP Version"
        wb.active["B5"] = "garbage"

        buf = BytesIO()
        wb.save(buf)

        with pytest.raises(DetectionError, match="Unexpected version cell"):
            detect_version(buf.getvalue())
