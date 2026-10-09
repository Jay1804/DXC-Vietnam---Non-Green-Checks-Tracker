"""Streamlit UI for the DXC Vietnam Non Green Checks report. Run: streamlit run app.py"""

from datetime import date, timedelta

import streamlit as st

import report as r

st.set_page_config(page_title="DXC Vietnam - Non Green Checks", page_icon="📋", layout="wide")
st.title("DXC Vietnam - Non Green Checks")

# Default closure dates: yesterday, or Fri-Sun when today is Monday.
default_start, default_end = r.get_closure_window(date.today())
default_last_day = default_end - timedelta(days=1)

picked = st.date_input(
    "Check closure date(s)",
    value=(default_start, default_last_day),
    max_value=date.today(),
    help="Defaults to yesterday, or Friday-Sunday when today is Monday.",
)
# date_input returns a 1-tuple while the user is still choosing the end date.
if isinstance(picked, (tuple, list)):
    start = picked[0] if picked else default_start
    last_day = picked[1] if len(picked) > 1 else start
else:
    start = last_day = picked

if st.button("Generate report", type="primary"):
    with st.spinner("Fetching data from the database..."):
        try:
            rows, flex_map = r.fetch_data(start, last_day + timedelta(days=1))
            headers, final_rows, missing = r.build_final_rows(rows, flex_map)
            wb, window = r.build_workbook(headers, final_rows, start, last_day + timedelta(days=1))
            st.session_state.result = {
                "headers": headers,
                "rows": final_rows,
                "missing": missing,
                "window": window,
                "xlsx": r.workbook_bytes(wb),
                "filename": f"DXC_Vietnam_Non_Green_Checks_{start:%Y%m%d}_{last_day:%Y%m%d}.xlsx",
            }
        except Exception as e:  # surface DB errors in the UI
            st.session_state.pop("result", None)
            st.error(f"Failed to generate report: {e}")

res = st.session_state.get("result")
if res:
    if res["missing"]:
        st.warning(f"Not found in the client field mapping (left blank): {', '.join(res['missing'])}")
    st.success(f"Closure date: {res['window']} - {len(res['rows'])} record(s)")
    st.download_button(
        "Download Excel report",
        data=res["xlsx"],
        file_name=res["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    if res["rows"]:
        st.dataframe([dict(zip(res["headers"], row)) for row in res["rows"]], width="stretch")
    else:
        st.info("No non-green checks found for the selected date(s). The downloaded file will be blank.")
