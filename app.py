"""
Streamlit UI for the DXC Vietnam Non-Green Tracker pipeline:

Login -> Pull MIS query -> Filter severity -> Select/rename columns ->
Filter by closure date (dynamic: yesterday, or Fri/Sat/Sun if today is
Monday) -> Polished Excel report.

Run with:
    streamlit run app.py
"""

import os
from datetime import datetime, timedelta

import pandas as pd
import requests
import streamlit as st
from dotenv import load_dotenv

import pipeline as p

load_dotenv()
MIS_USERNAME = os.environ.get("MIS_USERNAME")
MIS_PASSWORD = os.environ.get("MIS_PASSWORD")

st.set_page_config(page_title="DXC Vietnam Non-Green Tracker", page_icon="📋", layout="wide")

st.title("📋 DXC Vietnam - Non Green Checks Tracker")
st.caption("Authbridge MIS Query Browser → filtered, formatted Final Report")

if "result" not in st.session_state:
    st.session_state.result = None

if not MIS_USERNAME or not MIS_PASSWORD:
    st.error("MIS_USERNAME / MIS_PASSWORD not found. Add them to a .env file in the project folder.")
    st.stop()

with st.sidebar:
    st.header("⚙️ Query Settings")
    query_key = st.selectbox(
        "Saved query",
        options=list(p.QUERIES.keys()),
        format_func=lambda k: f"{k} — {p.QUERIES[k]['label']}",
        index=list(p.QUERIES.keys()).index("tracker"),
    )
    client_name = st.text_input("Client Name", value="DXC V")

    st.header("📅 Data Pull Range")
    range_mode = st.radio("Range mode", ["Last N months", "Custom range"], horizontal=True)
    today = datetime.now()
    if range_mode == "Last N months":
        months = st.number_input("Months back", min_value=1, max_value=84, value=3)
        from_date = today - timedelta(days=30 * months)
        to_date = today
        st.caption(f"From **{from_date.date()}** to **{to_date.date()}**")
    else:
        from_date = st.date_input("From date", value=today - timedelta(days=90))
        to_date = st.date_input("To date", value=today.date())
        from_date = datetime.combine(from_date, datetime.min.time())
        to_date = datetime.combine(to_date, datetime.max.time())

    st.header("🎯 Severity Filter")
    severities = st.multiselect(
        "Keep only these severities",
        options=["Red", "Amber", "Green", "Green_Approved", "Clear"],
        default=["Amber", "Red"],
    )

    st.header("🗓️ Closure Date Filter")
    auto_dates = p.get_target_dates()
    use_auto = st.checkbox(
        f"Auto (dynamic): {sorted(auto_dates)}",
        value=True,
        help="Yesterday's date, or Friday/Saturday/Sunday if today is Monday.",
    )
    if not use_auto:
        manual_dates = st.date_input(
            "Manual closure date(s)",
            value=[datetime.now().date() - timedelta(days=1)],
        )
        if not isinstance(manual_dates, (list, tuple)):
            manual_dates = [manual_dates]
        target_dates = {d.strftime("%Y-%m-%d") for d in manual_dates}
    else:
        target_dates = auto_dates

    run = st.button("🚀 Run Pipeline", type="primary", use_container_width=True)

if run:
    try:
        with st.spinner("Logging in to MIS..."):
            session = requests.Session()
            p.login(session, MIS_USERNAME, MIS_PASSWORD)

        with st.spinner(f"Pulling '{query_key}' data from {from_date.date()} to {to_date.date()} "
                         f"for client '{client_name}'... (can take a while for wide ranges)"):
            zip_bytes = p.fetch_report_zip(session, query_key, from_date, to_date, client_name)
            headers, raw_rows = p.extract_csv_rows(zip_bytes)

        st.success(f"Pulled {len(raw_rows)} total rows.")

        severity_rows = p.filter_severity(raw_rows, severities)
        st.info(f"After severity filter ({', '.join(severities) or 'none'}): {len(severity_rows)} rows.")

        sel_headers, sel_rows = p.select_and_rename(severity_rows)

        final_rows = p.filter_by_closure_date(sel_headers, sel_rows, target_dates)
        st.info(f"After closure-date filter ({', '.join(sorted(target_dates))}): {len(final_rows)} rows.")

        wb = p.build_report_workbook(sel_headers, final_rows, target_dates,
                                      severity_label=", ".join(severities) or "None")
        xlsx_bytes = p.workbook_to_bytes(wb)

        st.session_state.result = {
            "headers": sel_headers,
            "rows": final_rows,
            "xlsx_bytes": xlsx_bytes,
            "target_dates": sorted(target_dates),
        }
    except p.PipelineError as exc:
        st.error(str(exc))
    except Exception as exc:
        st.error(f"Something went wrong: {exc}")

result = st.session_state.result
if result:
    st.divider()
    st.subheader(f"Final Report — Closure Date(s): {', '.join(result['target_dates'])}")

    if result["rows"]:
        df = pd.DataFrame(result["rows"], columns=result["headers"])

        def highlight_severity(row):
            color = ""
            if row.get("check_severity") == "Red":
                color = "background-color: #F8CBAD"
            elif row.get("check_severity") == "Amber":
                color = "background-color: #FFE699"
            return [color] * len(row)

        st.dataframe(
            df.style.apply(highlight_severity, axis=1),
            use_container_width=True,
            height=min(600, 45 + 35 * len(df)),
        )
    else:
        st.warning("No records found for the selected date(s). The report contains headers only.")

    st.download_button(
        "⬇️ Download Final Report (Excel)",
        data=result["xlsx_bytes"],
        file_name="DXC Vietnam - Final Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
