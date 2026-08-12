"""
Logs into the Authbridge MIS Query Browser and downloads MIS reports for
client "DXC V", for Bridge Live / checkpoint_live.

Two modes:
  Daily (default): uses yesterday's date, except when today is Monday, in
  which case it uses last Friday's date. Saved to downloads/.

  Range: pull a custom date range in one shot (e.g. for historical/report
  building). Saved directly into the project root.

Available queries (--query):
  daily   - "Yesterday Done checks with other details as per client name -
            Bridge" (Morning Data - Daily Dump)
  tracker - "Client wise all cases with all checks with unique check
            name-Bridge On Received" (Tracker - Query old)

Usage:
    python download_dxc_report.py
    python download_dxc_report.py --from 2026-05-10 --to 2026-08-09
    python download_dxc_report.py --months 3
    python download_dxc_report.py --query tracker --from 2020-01-01 --to 2026-08-11

Credentials can be overridden with environment variables:
    MIS_USERNAME, MIS_PASSWORD
"""

import argparse
import io
import os
import sys
import zipfile
from datetime import datetime, timedelta

import requests
from dotenv import load_dotenv

load_dotenv()

BASE_URL = "https://mis.authbridge.com/export_query"
LOGIN_URL = f"{BASE_URL}/login.php"
PROCESS_URL = f"{BASE_URL}/process.php"

USERNAME = os.environ.get("MIS_USERNAME")
PASSWORD = os.environ.get("MIS_PASSWORD")

HOSTNAME_ID = "4"              # Bridge Live
DATABASE = "checkpoint_live"   # Bridge Live database
CLIENT_NAME = "DXC V"

QUERIES = {
    "daily": {
        "access_time": "Morning Data - Daily Dump",
        "csv_query": "909",
        "query_days_range": "0",
        "to_time": "23:59:59",
    },
    "tracker": {
        "access_time": "Tracker - Query old",
        "csv_query": "1077",
        "query_days_range": "",
        "to_time": "00:00:00",
    },
}

DOWNLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

HEADERS = {"User-Agent": "Mozilla/5.0"}


def get_report_date():
    """Yesterday's date, or last Friday if today is Monday."""
    today = datetime.now()
    if today.weekday() == 0:  # Monday
        return today - timedelta(days=3)
    return today - timedelta(days=1)


def login(session):
    if not USERNAME or not PASSWORD:
        raise RuntimeError("MIS_USERNAME / MIS_PASSWORD not found. Add them to a .env file in the project folder.")
    resp = session.post(
        LOGIN_URL,
        data={"username": USERNAME, "password": PASSWORD, "login": "Login"},
        headers=HEADERS,
    )
    resp.raise_for_status()
    if "logout.php" not in resp.text and "searchForm" not in resp.text:
        raise RuntimeError("Login failed - check credentials.")


def download_report(session, query_key, from_dt, to_dt):
    query = QUERIES[query_key]
    date1 = from_dt.strftime("%Y-%m-%d 00:00:00")
    date2 = to_dt.strftime(f"%Y-%m-%d {query['to_time']}")

    resp = session.post(
        PROCESS_URL,
        data={
            "hostname": HOSTNAME_ID,
            "database": DATABASE,
            "access_time": query["access_time"],
            "csv_query": query["csv_query"],
            "query_days_range": query["query_days_range"],
            "date1": date1,
            "date2": date2,
            "client1": CLIENT_NAME,
        },
        headers=HEADERS,
    )
    resp.raise_for_status()

    content_type = resp.headers.get("Content-Type", "")
    if "zip" not in content_type and not resp.content.startswith(b"PK"):
        raise RuntimeError(f"Unexpected response (not a zip file): {resp.text[:500]}")

    return resp.content


def save_report(zip_bytes, out_dir, suffix):
    os.makedirs(out_dir, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        extracted_paths = []
        for name in zf.namelist():
            base, ext = os.path.splitext(name)
            out_name = f"{base} - {suffix}{ext}"
            out_path = os.path.join(out_dir, out_name)
            with zf.open(name) as src, open(out_path, "wb") as dst:
                dst.write(src.read())
            extracted_paths.append(out_path)
    return extracted_paths


def run_daily(query_key):
    report_date = get_report_date()
    print(f"Fetching report for date: {report_date.strftime('%Y-%m-%d')} "
          f"({report_date.strftime('%A')})")

    session = requests.Session()
    login(session)
    print("Login successful.")

    zip_bytes = download_report(session, query_key, report_date, report_date)
    paths = save_report(zip_bytes, DOWNLOAD_DIR, report_date.strftime("%Y-%m-%d"))

    for path in paths:
        print(f"Saved: {path}")


def run_range(query_key, from_dt, to_dt):
    print(f"Fetching report for range: {from_dt.strftime('%Y-%m-%d')} "
          f"to {to_dt.strftime('%Y-%m-%d')} (query: {query_key})")

    session = requests.Session()
    login(session)
    print("Login successful.")

    zip_bytes = download_report(session, query_key, from_dt, to_dt)
    suffix = f"{from_dt.strftime('%Y-%m-%d')}_to_{to_dt.strftime('%Y-%m-%d')}"
    paths = save_report(zip_bytes, PROJECT_ROOT, suffix)

    for path in paths:
        print(f"Saved: {path}")


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--query", choices=list(QUERIES.keys()), default="daily",
                         help="Which saved query to run (default: daily)")
    parser.add_argument("--from", dest="from_date", help="Range start date (YYYY-MM-DD)")
    parser.add_argument("--to", dest="to_date", help="Range end date (YYYY-MM-DD)")
    parser.add_argument("--months", type=int, help="Pull the last N months up to yesterday")
    return parser.parse_args()


def main():
    args = parse_args()

    if args.from_date or args.to_date or args.months:
        yesterday = datetime.now() - timedelta(days=1)
        to_dt = datetime.strptime(args.to_date, "%Y-%m-%d") if args.to_date else yesterday
        if args.from_date:
            from_dt = datetime.strptime(args.from_date, "%Y-%m-%d")
        elif args.months:
            from_dt = to_dt - timedelta(days=30 * args.months)
        else:
            raise ValueError("Provide --from, or --months, along with an optional --to.")
        run_range(args.query, from_dt, to_dt)
    else:
        run_daily(args.query)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
