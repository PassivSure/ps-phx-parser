# ps-phx-parser

HTTP service that extracts structured data from PHPP (Passive House Planning Package) workbooks using the [PHX library](https://github.com/PH-Tools/PHX)'s bundled field mappings.

Runs headless (no Excel, no xlwings). Reads are fully deterministic — no LLM calls.

**Status: this is a shadow path, not the primary reader.**
[ps-rails](https://github.com/megaohms/ps-rails) enqueues a parse here on *every* PHPP
upload, alongside its Claude-based extractor (`EnergyModel::ExtractRouter`, default
`engine: BOTH`). But `Phpp::PhxPollJob` writes only `phpp_version` and
`phpp_mapped_attrs`, and **never creates a `Performance`** — the output feeds an offline
comparison (PAS-84) whose gate has not been flipped. Nothing user-facing waits on it, and
no rendered number comes from this service today. See `app/jobs.py` for why that shapes the
in-process job store.

Two consequences worth knowing before you add a field here:

- ps-rails must be taught to consume a subtree before it has any effect. `hvac_equipment`
  has been emitted since the equipment-extraction work landed, and `Phpp::Phx::Mapper`
  still hardcodes `hvac_equipment: nil`, so it is currently discarded on arrival.
- `schema/output.schema.json` sets `additionalProperties: false` and is enforced only in
  tests, so a new field needs the JSON Schema, the pydantic response model *and* the
  ps-rails mapper updated together.

## How it works

1. ps-rails generates a short-lived signed URL for an uploaded PHPP `.xlsx` and calls `POST /parse` with the URL. `phpp_version` is optional — omit it and the version is detected from the workbook.
2. `/parse` returns **`202` with a `job_id`** and runs the parse in the background; the caller polls `GET /parse/{job_id}` until `status` is `done` or `failed`. It is async because a real parse takes ~46s against Heroku's hard 30s router cap.
3. The parse fetches the workbook, loads the matching pydantic shape from `PHX.PHPP.phpp_localization`, reads cells via `openpyxl` (`data_only=True, read_only=True`) using the locator pattern (`locator_col` + `locator_string` → `input_column` + `input_row_offset`), and returns structured JSON.
4. PHX's xlwings / `PHPPConnection` surface is never touched — we use only the shape model JSON.

> Job state is an in-process dict, so **`--workers 1` in the `Procfile` is load-bearing**, not tidiness. Read the comment at the top of `app/jobs.py` before changing the worker count or the dyno size.

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