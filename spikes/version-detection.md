# Spike: PHPP version detection

**Linear:** [PAS-57](https://linear.app/passivsure/issue/PAS-57)
**Date:** 2026-04-24
**Status:** Implemented. Production code lives in `app/version_detection.py` and is wired into `POST /detect-version` and `POST /parse` (auto-detects when `phpp_version` is omitted). The prototype that was at `spikes/version_detection.py` has been deleted.

## Summary

PHPP v9+ files encode their version in the `Data` sheet (localized: `Daten` DE, `Datos` ES). Scan col A rows 1–10 for the first cell starting with "PHPP"; on that row, cell 2 is the version string (`"10.6"` or `"10.6 IP"`), cell 3+ optionally names the language, and the last cell's PE-factor string (`"1-PE-factors …"` EN, `"1-PE-FAKTOREN"` DE, `"1-FACTORES EP"` ES) is a language proxy for pre-v10 files. Normalize to `{LANG}_{major}_{minor}` with spaces stripped and dots→underscores (PHX's `PHPPVersion.clean_input`), and you get the shape filename stem (`EN_10_6`, `EN_10_6IP`, `EN_9_7IP`). Detection happens inside the parser service — ps-rails does not need to implement its own detector.

## How PHX does it

In [`PHX/PHPP/phpp_app.py::PHPPConnection.get_phpp_version`](https://github.com/PH-Tools/PHX/blob/main/PHX/PHPP/phpp_app.py):

1. Find the Data worksheet (`DATA` / `DATEN` / `DATOS` — case-insensitive).
2. In col A, rows 1–10, locate the first cell whose upper-stripped-despaced value starts with `"PHPP"`.
3. On that row, filter out blanks. Cell index 1 is the raw version (e.g. `"10.6 IP"`). Split on `.` to get `major`, `minor`.
4. Language detection: scan the last non-blank cell in that row. `"1-PE-FAKTOREN"` → DE, `"1-FACTORES EP"` → ES, `"1-PE-FACTORS"` → EN. Raises if unrecognised.
5. Build `PHPPVersion(major, minor, language)`. Its `clean_input` uppercases, strips, removes spaces, and replaces `.` with `_`, so `"10.6 IP"` becomes `"10_6IP"`.
6. [`phpp_localization/load.py::phpp_version_as_file_name`](https://github.com/PH-Tools/PHX/blob/main/PHX/PHPP/phpp_localization/load.py) then formats `f"{language}_{number_major}_{number_minor}"` → `"EN_10_6IP"` → loads `EN_10_6IP.json`.

The IP-vs-metric distinction rides for free on the version cell. There is no separate units detection step.

## Empirical findings (sample PHPPs in `~/Downloads/`)

| Sample | Data row | Row values (non-blank) |
|---|---|---|
| `PHPP_V10.6_IP_Example.xlsx` | 5 | `['PHPP Version', '10.6 IP', 'Language', '1-PE-factors (non-renewable) PHI Certification']` |
| `PHPP_EN_V10.6_easyPH_Example.xlsx` | 5 | `['PHPP Version', '10.6', 'Language', 'EN ', '1-PE-factors …']` |
| `260309 PHPP 6840 E 7th v 10.6.xlsx` | 5 | `['PHPP Version', '10.6 IP', 'Language', '1-PE-factors …']` |
| `PHPP_EN_V8.5_Cudz…xlsm` | 3 | `['PHPP Version', '8.5']` ← no language cell, no PE factor |
| `PHPP_EN_V8.5_Huang…xlsm` | 3 | `['PHPP Version', '8.5']` ← same |
| `PHPP US 2007_Moore…xlsx` | — | No "PHPP" row in col A rows 1–10. Data sheet exists but layout differs. |
| `OSH …2012.11.04.xlsx` | — | Same — v2007-era, different Data-sheet layout. |

Prototype at `spikes/version_detection.py` reproduces the full PHX logic in ~40 lines of openpyxl; reads all three v10.6 samples correctly.

## Recommendation

**Detect inside the parser service** (`ps-phx-parser`), not in ps-rails.

- Add `POST /detect-version` (or call detection as step 1 of `POST /parse`). Returns `{"phpp_version": "EN_10_6IP"}` on success or `422 {"detail": "...", "raw": "..."}` on unsupported input.
- Ps-rails stores the returned `phpp_version` string on `artifacts.phpp_version` — no Ruby-side parsing of the .xlsx. This collapses PAS-58 to a trivial "add the column, store what the parser returned."
- For PAS-59's MVP, the caller passes `phpp_version` explicitly. After this spike lands, wire auto-detection so ps-rails can omit it.

## Shape coverage + gating

Bundled shape files (English only):

```
EN_10_3   EN_10_4A   EN_10_4IP   EN_10_6   EN_10_6IP   EN_9_6A   EN_9_7IP
```

Everything below v9.6 (v8.5, v2007) is out of scope — no shape exists — and falls back to the existing Claude/rubyXL extractor (PAS-61's `Phpp::ExtractRouter`).

## Edge cases + open questions

- **`EN_10_4A` naming.** PHX's `PHPPVersion` has no "A" path; a stock 10.4 metric file would compute to `EN_10_4` (no suffix) and fail FileNotFoundError. Upstream inconsistency. Flagging as an open question — may need a `{"10_4": "10_4A"}` rewrite map, or upstream a rename PR (ties into PAS-63's "at least one upstream PHX contribution").
- **Pre-v10 language detection.** v8.5 Data row lacks the PE-factor string, so PHX's language probe throws. We have no non-EN shapes anyway, so default language to `EN` when the probe returns None rather than erroring. Narrow fallback, low risk.
- **v2007 + pre-Data-sheet layouts.** No PHPP marker in col A rows 1–10. Gate cleanly with `422 {"detail": "Pre-v9 PHPP not supported"}` and let the Claude fallback handle it.
- **`easyPH` flag.** PHX's `is_easyPh()` checks for an `"easyPH"` sheet. Affects write behavior, not read paths — MVP ignores it.
- **Localized Data sheet names beyond EN/DE/ES.** French/Italian/etc. would need new entries; deferred until we see one.
