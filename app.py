"""
Elite Prep — SAT Class Results Analysis (Streamlit web app)
===========================================================
Upload all students' SAT/DSAT score-report PDFs for one test and download
the class-wide Word report:  "<TEST CODE> Result Analysis Teacher Report.docx"

The test code and test date are read automatically from the PDFs.

Run locally:   streamlit run app.py
"""

import datetime
import io
import tempfile
from collections import defaultdict
from pathlib import Path

import streamlit as st

from generate_class_report import (
    ClassStats,
    build_report,
    parse_date_iso,
    parse_pdf_stream,
    parse_text,
    sanitize_branding,
)

NAVY = "#1F3864"
BLUE = "#2E75B6"

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
    "test, then click Generate. The test code and test date are detected "
    "automatically, and the class-wide Word report is created in the "
    "standard Elite Prep teacher-report format."
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
# 2) Generate
# ---------------------------------------------------------------------------

st.header("2. Generate the report")

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

    # Safety net: if PDFs from more than one test were mixed in, analyze the
    # most recent test and tell the user what was set aside.
    groups = defaultdict(list)
    for s in students:
        groups[sanitize_branding(s.test_code or "").strip()
               or "(unknown test)"].append(s)
    if len(groups) > 1:
        def group_key(code):
            dates = [parse_date_iso(s.test_date) for s in groups[code]
                     if s.test_date]
            dates = [d for d in dates if d]
            return (max(dates) if dates else "", len(groups[code]))
        current_code = max(groups, key=group_key)
        skipped = [c for c in groups if c != current_code]
        st.warning(f"PDFs from more than one test were uploaded "
                   f"({', '.join(sorted(groups))}). The report was generated "
                   f"for **{current_code}**; files from {', '.join(skipped)} "
                   f"were not included.")
        students = groups[current_code]

    stats = ClassStats(students)

    test_code = sanitize_branding(
        next((s.test_code for s in students if s.test_code), None)
        or "SAT Practice Test").strip()
    test_date = (next((s.test_date for s in students if s.test_date), None)
                 or datetime.date.today().strftime("%B %d, %Y"))

    # --- build the Word report ------------------------------------------------
    out_name = f"{test_code} Result Analysis Teacher Report.docx"
    out_name = "".join(c if c not in '\\/:*?"<>|' else "-" for c in out_name)
    tmp_docx = Path(tempfile.gettempdir()) / out_name
    build_report(stats, test_code, test_date, tmp_docx)
    docx_bytes = tmp_docx.read_bytes()

    st.success(f"Report generated for **{stats.n} student"
               f"{'s' if stats.n != 1 else ''}** — Test Code **{test_code}**, "
               f"Test Date **{test_date}** (detected from the PDFs).")

    # --- download ---------------------------------------------------------------
    st.download_button(
        f"⬇️ Download Word report — {out_name}",
        data=docx_bytes,
        file_name=out_name,
        mime=("application/vnd.openxmlformats-officedocument"
              ".wordprocessingml.document"),
        use_container_width=True,
    )

    # --- on-screen preview --------------------------------------------------------
    st.subheader("Preview")

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
        error_groups = stats.error_groups(section, module)
        with st.expander(title, expanded=False):
            if not error_groups:
                st.write("No incorrect answers recorded.")
            else:
                st.table([{
                    "Error Rate": f"{round(100 * c / stats.n)}%",
                    "Students Missed": f"{c} of {stats.n}",
                    "Questions": ", ".join(f"Q{q}" for q in qs),
                } for c, qs in error_groups])

st.markdown(
    f"<hr style='border-top:2px solid {NAVY};'>"
    f"<div style='color:{NAVY};font-weight:700;'>Elite Prep</div>",
    unsafe_allow_html=True,
)
