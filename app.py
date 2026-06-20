
from __future__ import annotations

import hashlib
from pathlib import Path

import streamlit as st

from schedule_core import (
    build_macro_workbook,
    extract_schedule,
    read_workbook_to_matrices,
    suggested_batch_tabs,
    minutes_to_text,
)

APP_DIR = Path(__file__).resolve().parent
VBA_PROJECT = APP_DIR / "assets" / "vbaProject.bin"

st.set_page_config(
    page_title="Weekly Schedule Checker",
    page_icon="📅",
    layout="wide",
)


def require_password():
    """
    Optional access password.
    Leave APP_PASSWORD absent in Streamlit secrets to run without a password.
    """
    try:
        expected = st.secrets.get("APP_PASSWORD", "")
    except Exception:
        expected = ""

    if not expected:
        return True

    if st.session_state.get("authenticated"):
        return True

    st.title("Weekly Schedule Checker")
    st.caption("Enter the team password to continue.")
    entered = st.text_input("Password", type="password")

    if st.button("Continue", type="primary"):
        if entered == expected:
            st.session_state["authenticated"] = True
            st.rerun()
        st.error("Incorrect password.")

    return False


def reset_uploaded_data():
    for key in ("file_hash", "matrices", "file_name", "selected_tabs", "result"):
        st.session_state.pop(key, None)


def load_upload(upload):
    data = upload.getvalue()
    digest = hashlib.sha256(data).hexdigest()

    if digest == st.session_state.get("file_hash"):
        return

    reset_uploaded_data()
    matrices = read_workbook_to_matrices(data, upload.name)

    st.session_state["file_hash"] = digest
    st.session_state["matrices"] = matrices
    st.session_state["file_name"] = upload.name
    st.session_state["selected_tabs"] = suggested_batch_tabs(list(matrices.keys()))


if not require_password():
    st.stop()

st.title("📅 Weekly Schedule Extractor & Clash Checker")
st.caption("Upload the weekly schedule, select batch-wise tabs, and download the final macro-enabled schedule.")

with st.expander("What happens after download?", expanded=False):
    st.markdown(
        """
1. Open the downloaded `.xlsm` file in Microsoft Excel.
2. Click **Enable Content**.
3. Press **Alt + F8** and run **Create_All_Schedules**.
4. The macro creates:
   - `Teacherwise Schedule`
   - `Batchwise Schedule`
   - `<original excel name>_teacher schedule` folder with one PDF for every teacher
   - `<original excel name>_Batchwise schedule` folder with one PDF for every batch

The PDF folders are created beside the downloaded `.xlsm` file on the user's own computer.
        """
    )

st.info("Select only batch-wise timetable tabs. Do not select Faculty/Teacher/Team tabs unless their layout is the same.")

upload = st.file_uploader(
    "Upload weekly schedule workbook",
    type=["xlsx", "xlsm", "xls"],
    help="The source workbook can be XLSX, XLSM or XLS. The generated file will always be XLSM.",
)

if upload is None:
    st.stop()

try:
    load_upload(upload)
except Exception as exc:
    st.error(f"Could not read the workbook: {exc}")
    st.stop()

sheet_names = list(st.session_state["matrices"].keys())
selected_tabs = st.multiselect(
    "Select batch-wise tabs",
    options=sheet_names,
    default=st.session_state.get("selected_tabs", []),
)
st.session_state["selected_tabs"] = selected_tabs

left, middle, right = st.columns([1, 1, 4])
left.metric("Tabs available", len(sheet_names))
middle.metric("Tabs selected", len(selected_tabs))
right.caption("All processing is temporary. The app does not save your uploaded workbook after the session ends.")

if st.button("Create Combined XLSM", type="primary", disabled=not selected_tabs):
    try:
        with st.spinner("Reading timetable blocks, standardising time, checking teacher clashes and preparing XLSM..."):
            sessions, clashes, reviews, skipped = extract_schedule(
                st.session_state["matrices"],
                selected_tabs,
            )

            source_excel_base = Path(st.session_state["file_name"]).stem
            output_bytes = build_macro_workbook(
                sessions=sessions,
                clash_pairs=clashes,
                review_rows=reviews,
                selected_sheets=selected_tabs,
                skipped_blocks=skipped,
                source_excel_base=source_excel_base,
                vba_project_path=VBA_PROJECT,
            )

            st.session_state["result"] = {
                "bytes": output_bytes,
                "file_name": f"Combined_Schedule_{source_excel_base}.xlsm",
                "sessions": sessions,
                "clashes": clashes,
                "reviews": reviews,
                "skipped": skipped,
            }
    except Exception as exc:
        st.error(f"Could not create the schedule: {exc}")

result = st.session_state.get("result")
if result:
    st.success("Your macro-enabled combined schedule is ready.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Sessions extracted", len(result["sessions"]))
    c2.metric("Teacher clash pairs", len(result["clashes"]))
    c3.metric("Review rows", len(result["reviews"]))
    c4.metric("Faculty/team blocks skipped", result["skipped"])

    st.download_button(
        "Download Combined Schedule (.xlsm)",
        data=result["bytes"],
        file_name=result["file_name"],
        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        type="primary",
    )

    if result["clashes"]:
        st.subheader("Teacher clashes found")
        st.dataframe(
            [
                {
                    "Date": item["date"].strftime("%d.%m.%Y"),
                    "Teacher": item["teacher"],
                    "Batch 1": item["batch_1"],
                    "Time 1": f"{minutes_to_text(item['start_1'])} - {minutes_to_text(item['end_1'])}",
                    "Batch 2": item["batch_2"],
                    "Time 2": f"{minutes_to_text(item['start_2'])} - {minutes_to_text(item['end_2'])}",
                    "Overlap": f"{item['overlap_minutes']} min",
                }
                for item in result["clashes"]
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.success("No overlapping teacher sessions were found in the selected tabs.")

    if result["reviews"]:
        st.subheader("Cells needing review")
        st.caption("The full list is also included in the `Review Needed` sheet of the downloaded XLSM.")
        st.dataframe(
            [
                {
                    "Source Sheet": item[0],
                    "Source Cell": item[1],
                    "Date": item[2].strftime("%d.%m.%Y") if hasattr(item[2], "strftime") else item[2],
                    "Batch": item[4],
                    "Original Entry": item[6],
                    "Reason": item[7],
                }
                for item in result["reviews"][:100]
            ],
            use_container_width=True,
            hide_index=True,
        )
