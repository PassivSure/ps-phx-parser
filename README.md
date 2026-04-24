# ps-phx-parser

HTTP service that extracts structured data from PHPP (Passive House Planning Package) workbooks using the [PHX library](https://github.com/PH-Tools/PHX)'s bundled field mappings.

Runs headless (no Excel, no xlwings). Used by [ps-rails](https://github.com/megaohms/ps-rails) as the primary PHPP reader for version 10.x files, with the existing Claude-based extractor as fallback.

## How it works

1. ps-rails generates a short-lived signed URL for an uploaded PHPP `.xlsx` and calls `POST /parse` with the URL and the PHPP version.
2. This service fetches the workbook, loads the matching pydantic shape from `PHX.PHPP.phpp_localization`, reads cells via `openpyxl` (`data_only=True, read_only=True`) using the locator pattern (`locator_col` + `locator_string` → `input_column` + `input_row_offset`), and returns structured JSON.
3. PHX's xlwings / `PHPPConnection` surface is never touched — we use only the shape model JSON.

## Development

Requires Python 3.13 and [uv](https://docs.astral.sh/uv/).

```bash
uv sync                          # install deps + create .venv
uv run uvicorn app.main:app --reload  # start dev server on :8000
uv run pytest                    # run tests
uv run ruff check .              # lint
```

## Deployment

Heroku, separate pipeline (review apps / staging / prod) from the ps-rails app.

Heroku's Python buildpack detects `uv.lock` and runs `uv sync --locked --no-dev` during the build — no `requirements.txt` is needed and shipping one alongside `uv.lock` triggers a "multiple package managers detected" build failure. To update deps:

```bash
uv add <package>        # or `uv add --dev <package>` for test-only deps
git add pyproject.toml uv.lock
```

`.python-version` pins the Python version. The `Procfile` runs `uvicorn app.main:app --host 0.0.0.0 --port $PORT` against the venv the buildpack activates.

### Configuration

| Env var | Required | Purpose |
|---|---|---|
| `PHX_PARSER_AUTH_TOKEN` | yes (in prod) | Shared secret. `/parse` requires `Authorization: Bearer <token>`. Must match `PHX_PARSER_AUTH_TOKEN` on the ps-rails side. Generate with `openssl rand -hex 32`. If unset, auth is disabled (a startup warning is logged) — fine for tests/CI, **unsafe for any publicly reachable host**. |
| `PORT` | (auto) | Heroku injects. |

`/health` and `/versions` are always open (Heroku's router needs `/health`).

## License

GPL-3.0-or-later. This project imports [PHX](https://github.com/PH-Tools/PHX) (GPL-3.0-or-later) at runtime, so the service is itself GPL. The HTTP boundary between ps-rails and this service keeps ps-rails proprietary — legal review gate on PAS-63 before production cutover.

See [LICENSE](LICENSE) for full terms.