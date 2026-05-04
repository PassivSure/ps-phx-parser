"""Detect a PHPP workbook's version + language from its Data sheet.

Ports PHX.PHPP.phpp_app.PHPPConnection.get_phpp_version to headless
openpyxl so we can run on Heroku without Excel/xlwings. See
spikes/version-detection.md for the writeup + empirical validation.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

DATA_SHEET_NAMES = ("DATA", "DATEN", "DATOS")
LANGUAGE_PROXIES = {
    "1-PE-FAKTOREN": "DE",
    "1-FACTORES EP": "ES",
    "1-PE-FACTORS": "EN",
}


class DetectionError(Exception):
    """Raised when the workbook's version can't be determined."""


@dataclass(frozen=True)
class DetectedVersion:
    major: str
    minor: str
    language: str
    raw: str  # the raw "10.6 IP" string read from the workbook

    @property
    def shape_stem(self) -> str:
        """PHX's `{LANG}_{major}_{minor}` filename stem.

        Matches PHX.PHPP.phpp_model.version.PHPPVersion.clean_input:
        upper-case, strip, remove spaces, replace dots with underscores.
        """
        major = _clean(self.major)
        minor = _clean(self.minor)
        return f"{self.language}_{major}_{minor}"


def _clean(value: str) -> str:
    return value.upper().strip().replace(" ", "").replace(".", "_")


def _find_data_sheet(wb) -> Worksheet:
    sheets_upper = {s.upper(): s for s in wb.sheetnames}
    for candidate in DATA_SHEET_NAMES:
        if candidate in sheets_upper:
            return wb[sheets_upper[candidate]]
    raise DetectionError(
        f"No Data/Daten/Datos sheet. Sheets: {wb.sheetnames[:8]}"
    )


def _find_phpp_row(ws: Worksheet, col: int = 1, max_row: int = 10) -> int:
    for r in range(1, max_row + 1):
        v = ws.cell(row=r, column=col).value
        if v and str(v).upper().strip().replace(" ", "").startswith("PHPP"):
            return r
    raise DetectionError(
        f"No 'PHPP' marker in col A rows 1-{max_row} of sheet {ws.title!r}. "
        "Likely a pre-v9 PHPP — not supported."
    )


def detect_version(workbook_bytes: bytes) -> DetectedVersion:
    wb = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        ws = _find_data_sheet(wb)
        row = _find_phpp_row(ws)
        row_vals = [ws.cell(row=row, column=c).value for c in range(1, 20)]
        non_blank = [v for v in row_vals if v not in (None, "")]

        if len(non_blank) < 2:
            raise DetectionError(f"PHPP row {row} has no version cell: {non_blank}")

        raw_version = str(non_blank[1])
        if "." not in raw_version:
            raise DetectionError(f"Unexpected version cell {raw_version!r}")
        major, minor = raw_version.split(".", 1)

        language = _detect_language(non_blank)

        return DetectedVersion(
            major=major, minor=minor, language=language, raw=raw_version
        )
    finally:
        wb.close()


def _detect_language(row_non_blank: list) -> str:
    """Language defaults to EN.

    PHPP v10+ Data row has a PE-factor proxy as the last non-blank cell,
    which we match against LANGUAGE_PROXIES. Pre-v10 files lack it — since
    we only ship EN shape files, defaulting to EN is the pragmatic fallback.
    """
    last = str(row_non_blank[-1]).upper().strip()
    for needle, lang in LANGUAGE_PROXIES.items():
        if needle in last:
            return lang
    return "EN"
