# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
pip install -r requirements.txt   # streamlit, pandas, openpyxl, requests, python-dotenv
streamlit run app.py              # run the app (primary entry point)
python download_dxc_report.py     # legacy standalone CLI (see Architecture)
python build_final_report.py      # legacy standalone CLI (see Architecture)
```

Requires a `.env` file in the project root with `MIS_USERNAME` and `MIS_PASSWORD` (Authbridge MIS Query Browser credentials).

There is no test suite, linter, or build step configured in this repo.

## Architecture

This is a Streamlit tool that pulls case/check data from the Authbridge MIS Query Browser (an internal HTTP reporting system at `mis.authbridge.com/export_query`) for client "DXC V", filters it, and produces a formatted Excel "Non Green Checks" report.

**`app.py` + `pipeline.py` are the current, maintained pipeline** — `app.py` is a Streamlit UI whose widgets drive functions imported from `pipeline.py` (`import pipeline as p`). All shared logic (HTTP calls, filtering, Excel building) lives in `pipeline.py`; `app.py` only handles UI state and wiring. `st.session_state.result` holds the last run's output so results survive Streamlit reruns.

**`download_dxc_report.py` and `build_final_report.py` are older standalone CLI scripts** that predate `app.py`. They do **not** import `pipeline.py` — each redeclares its own copy of constants like `QUERIES`, styling, and closure-date logic. Treat them as a legacy/parallel implementation: if you change shared behavior in `pipeline.py` (e.g. query definitions, closure-date rules, report styling), check whether the same fix needs to be duplicated there, since nothing keeps the two in sync. `build_final_report.py` doesn't hit the network at all — it re-processes a manually-produced local file, `DXC Vietnam - Amber_Red - Selected Columns.xlsx`, into the final polished report.

### The MIS Query Browser flow (`pipeline.py`)

1. `login()` — POSTs credentials to `login.php`, checks the response for `logout.php`/`searchForm` to confirm success (there's no structured success/failure response from this API).
2. `fetch_report_zip()` — POSTs to `process.php` with a fixed `hostname`/`database` (Bridge Live / `checkpoint_live`) plus a **saved query** selected from `QUERIES`: `daily` (csv_query `909`, single day) or `tracker` (csv_query `1077`, arbitrary date range — used for historical pulls). Each query has its own `to_time` used to build the end-of-range timestamp; `date1`/`date2` are always `YYYY-MM-DD HH:MM:SS` strings. The response is a zip file (validated by content-type / `PK` magic bytes) containing one CSV.
3. `extract_csv_rows()` unzips and parses that CSV into `(headers, rows)`.
4. `filter_severity()` keeps only rows whose `check_severity` is in the user-selected set (Red/Amber/Green/etc).
5. `select_and_rename()` projects down to `SELECTED_COLUMNS` — a fixed, ordered list of (source column → output label) pairs. Note `Company_name` intentionally appears twice, matching a specific requested report layout.
6. `get_target_dates()` computes the **closure-date filter**: normally just yesterday, but if today is Monday it returns Friday/Saturday/Sunday (to cover the weekend, since the daily pull cadence assumes Mon–Fri runs). This same rule is independently reimplemented in `build_final_report.py`.
7. `filter_by_closure_date()` keeps rows whose `CHECK_CLOSURE_DATE` (date part only) is in that target-date set.
8. `build_report_workbook()` / `workbook_to_bytes()` render an openpyxl workbook: title/subtitle banner rows, styled header row, banded row colors, Red/Amber severity cell highlighting, autosized columns, frozen header, and an autofilter — returned as in-memory bytes for `st.download_button` (no file is written to disk by `app.py`).

### Data pull range vs. closure-date filter (easy to confuse)

`app.py` exposes two separate, unrelated date controls:
- **"Data Pull Range"** (`from_date`/`to_date`) — how much raw data to request from the MIS query itself.
- **"Closure Date Filter"** (`target_dates`) — which rows survive `filter_by_closure_date()` after the pull, for the final report. This defaults to `get_target_dates()`'s dynamic yesterday/weekend logic but can be overridden manually.

A wide pull range with a narrow closure-date filter is normal (e.g. pulling all history via the `tracker` query, then filtering down to just yesterday's closures).
