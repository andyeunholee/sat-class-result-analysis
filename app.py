"""
Elite Prep — SAT Class Results Analysis (Streamlit web app)
===========================================================
Upload all students' SAT/DSAT score-report PDFs for one test and download
a class-wide Word report:  "<TEST CODE> Result Analysis Teacher Report.docx"

Run locally:   streamlit run app.py
"""

import datetime
import io
import tempfile
from pathlib import Path

import streamlit as st

from generate_class_report import (
    ClassStats,
    build_report,
    history_rows_to_csv,
    make_trend_chart,
    parse_pdf_stream,
    parse_text,
    read_history_csv,
    sanitize_branding,
    upsert_history_rows,
)

NAVY = "#1F3864"
BLUE = "#2E74B5"

st.set_page_config(
    page_title="Elite Prep — SAT Class Results Analysis",
    page_icon="📊",
    layout="centered",
)

st.markdown(
    f"""
    <style>
    .brand-bar {{
        border-bottom: 4px solid {NAVY};
        padding-bottom: 6px;
        margin-bottom: 18px;
    }}
    .brand-title {{
        color: {NAVY};
        font-size: 2.0rem;
        font-weight: 800;
        letter-spacing: 0.5px;
    }}
    .brand-sub {{
        color: {BLUE};
        font-size: 1.0rem;
        font-weight: 600;
    }}
    div.stButton > button, div.stDownloadButton > button {{
        background-color: {NAVY};
        color: white;
        font-weight: 600;
        border: none;
    }}
    div.stButton > button:hover, div.stDownloadButton > button:hover {{
        background-color: {BLUE};
        color: white;
    }}
    </style>
    <div class="brand-bar">
      <span class="brand-title">ELITE PREP</span>
      <span class="brand-sub">&nbsp;|&nbsp; College Admissions &amp; Test Prep</span>
    </div>
    """,
    unsafe_allow_html=True,
)

st.title("SAT Class Results Analysis Report")
st.caption(
    "Upload every student's SAT/DSAT practice-test score-report PDF for one "
    "test. The app analyzes the whole class and generates the Word report "
    "(most-missed questions per section, skill-area accuracy, class "
    "averages, and the score-trend graph on page 1)."
)

# ---------------------------------------------------------------------------
# 1) Upload
# ---------------------------------------------------------------------------

st.header("1. Upload student score reports")
uploads = st.file_uploader(
    "Student score-report PDFs (one per student — select all at once)",
    type=["pdf", "txt"],
    accept_multiple_files=True,
)

# ---------------------------------------------------------------------------
# 2) Test information
# ---------------------------------------------------------------------------

st.header("2. Test information")
col1, col2 = st.columns(2)
with col1:
    test_code_in = st.text_input(
        "Test code (leave blank to auto-detect)", placeholder="e.g. DSAT-02-A")
with col2:
    test_date_in = st.text_input(
        "Test date (leave blank to auto-detect)",
        placeholder="e.g. June 8, 2026")

# ---------------------------------------------------------------------------
# 3) Score history for the page-1 trend graph
# ---------------------------------------------------------------------------

st.header("3. Score history (page-1 trend graph)")
st.caption(
    "The report's first page shows the class-average score curve from the "
    "first practice test to this one. Upload the `score_history.csv` you "
    "downloaded after the previous test to continue the same curve. "
    "If this is the first test, leave it empty."
)
history_upload = st.file_uploader("Previous score_history.csv (optional)",
                                  type=["csv"])
include_trend = st.checkbox("Include the trend graph on page 1", value=True)

if "history_rows" not in st.session_state:
    st.session_state.history_rows = None

# ---------------------------------------------------------------------------
# Generate
# ---------------------------------------------------------------------------

st.header("4. Generate the report")

if st.button("Generate Report", type="primary", use_container_width=True):
    if not uploads:
        st.error("Please upload at least one student score-report PDF first.")
        st.stop()

    students, failed = [], []
    progress = st.progress(0.0, text="Parsing score reports...")
    for i, up in enumerate(sorted(uploads, key=lambda u: u.name)):
        try:
            if up.name.lower().endswith(".txt"):
                s = parse_text(up.getvalue().decode("utf-8", "replace"),
                               up.name)
            else:
                s = parse_pdf_stream(io.BytesIO(up.getvalue()), up.name)
        except Exception as e:
            failed.append((up.name, f"could not read file: {e}"))
            s = None
        if s is not None and s.parsed_ok:
            students.append(s)
        elif s is not None:
            failed.append((up.name, "no question-level data recognized"))
        progress.progress((i + 1) / len(uploads),
                          text=f"Parsing score reports... ({i + 1}/{len(uploads)})")
    progress.empty()

    if failed:
        st.warning("Some files could not be analyzed and were skipped:\n\n"
                   + "\n".join(f"- **{n}** — {why}" for n, why in failed))
    if not students:
        st.error("No student data could be parsed from the uploaded files. "
                 "Please check that these are SAT score-report PDFs.")
        st.stop()

    stats = ClassStats(students)

    test_code = sanitize_branding(
        test_code_in.strip()
        or next((s.test_code for s in students if s.test_code), None)
        or "SAT Practice Test").strip()
    test_date = (test_date_in.strip()
                 or next((s.test_date for s in students if s.test_date), None)
                 or datetime.date.today().strftime("%B %d, %Y"))

    # --- history / trend graph ---------------------------------------------
    trend_png = None
    history_csv_text = None
    if include_trend:
        if history_upload is not None:
            rows = read_history_csv(
                history_upload.getvalue().decode("utf-8-sig", "replace"))
        elif st.session_state.history_rows:
            rows = list(st.session_state.history_rows)
        else:
            rows = []
        rows = upsert_history_rows(rows, test_code, test_date, stats)
        st.session_state.history_rows = rows
        history_csv_text = history_rows_to_csv(rows)

        tmp_png = Path(tempfile.gettempdir()) / "eliteprep_trend_app.png"
        if make_trend_chart(rows, tmp_png):
            trend_png = tmp_png

    # --- build the Word report ----------------------------------------------
    out_name = f"{test_code} Result Analysis Teacher Report.docx"
    out_name = "".join(c if c not in '\\/:*?"<>|' else "-" for c in out_name)
    tmp_docx = Path(tempfile.gettempdir()) / out_name
    build_report(stats, test_code, test_date, tmp_docx, trend_png)
    docx_bytes = tmp_docx.read_bytes()

    st.success(f"Report generated for **{stats.n} student"
               f"{'s' if stats.n != 1 else ''}** — Test Code **{test_code}**, "
               f"Test Date **{test_date}**.")

    # --- on-screen preview ----------------------------------------------------
    st.subheader("Preview")

    if trend_png:
        st.image(str(trend_png),
                 caption="Class Average Score Trend (page 1 of the report)")

    c1, c2, c3 = st.columns(3)
    for col, label, key in ((c1, "Total", "total"),
                            (c2, "Reading & Writing", "rw"),
                            (c3, "Math", "math")):
        avg, rng = stats.avg(key), stats.rng(key)
        col.metric(f"{label} (class avg.)",
                   avg if avg is not None else "N/A",
                   help=f"Range: {rng[0]} - {rng[1]}" if rng else None)

    section_order = [("Math", 1, "① Math — Section 1 (Module 1)"),
                     ("Math", 2, "② Math — Section 2 (Module 2)"),
                     ("RW", 1, "③ English (R&W) — Section 1 (Module 1)"),
                     ("RW", 2, "④ English (R&W) — Section 2 (Module 2)")]
    for section, module, title in section_order:
        groups = stats.error_groups(section, module)
        with st.expander(title, expanded=False):
            if not groups:
                st.write("No incorrect answers recorded.")
            else:
                st.table([{
                    "Error Rate": f"{round(100 * c / stats.n)}%",
                    "Students Missed": f"{c} of {stats.n}",
                    "Questions": ", ".join(f"Q{q}" for q in qs),
                } for c, qs in groups])

    # --- downloads -------------------------------------------------------------
    st.subheader("Downloads")
    st.download_button(
        f"⬇️ Download Word report — {out_name}",
        data=docx_bytes,
        file_name=out_name,
        mime=("application/vnd.openxmlformats-officedocument"
              ".wordprocessingml.document"),
        use_container_width=True,
    )
    if history_csv_text:
        st.download_button(
            "⬇️ Download updated score_history.csv "
            "(keep this file — upload it next time to continue the trend curve)",
            data=history_csv_text,
            file_name="score_history.csv",
            mime="text/csv",
            use_container_width=True,
        )

st.markdown(
    f"<hr style='border-top:2px solid {NAVY};'>"
    f"<div style='color:{NAVY};font-weight:700;'>Elite Prep</div>"
    f"<div style='color:{BLUE};'>www.eliteprep.com</div>",
    unsafe_allow_html=True,
)
