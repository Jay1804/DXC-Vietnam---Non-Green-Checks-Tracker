# CLAUDE.md

## Commands

```bash
pip install -r requirements.txt
python report.py --dry-run            # build the Excel locally, no email
python report.py                      # pull data and email the report
python report.py --today 2026-10-12   # test a different "today" (e.g. a Monday)
```

`.env` (gitignored) needs: `DB_HOST/DB_PORT/DB_NAME/DB_USER/DB_PASSWORD` (MySQL, checkpoint_live),
`SMTP_HOST/SMTP_PORT/EMAIL_SENDER/EMAIL_PASSWORD`, and `REPORT_TO` (+ optional `REPORT_CC`), comma-separated.

No test suite or linter is configured.

## Architecture

`app.py` is the Streamlit UI (`streamlit run app.py`); `report.py` holds all logic and connects directly
to the database (the old MIS Query Browser pipeline was removed).

1. `get_closure_window()` - `check_closure_date >= start AND < end`: yesterday, or Fri-Sun when today is Monday.
2. `QUERY_CHECKS` - non-green checks for `company_name LIKE '%DXC Vietnam%'`. Ported from Postgres/Redshift
   syntax to MySQL (`CONCAT_WS`, `GROUP_CONCAT`, no `public.` schema).
3. `QUERY_FIELDS` - client case-field definitions. `field_id` N maps to `case_flex_fieldN`, so flex columns
   are renamed to their `field_name` (e.g. "Recruiter Name", "Case initiator Email Id") at runtime.
4. `FINAL_COLUMNS` - the only columns kept in the report.
5. Excel is built in memory and emailed via SMTP. The email is always sent; with no rows it is a
   "no records" email with a header-only workbook attached.

Schedule `python report.py` daily (e.g. Windows Task Scheduler, weekdays); the Monday logic is built in.
