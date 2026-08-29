"""
Elite Prep — SAT Class Test Result Analysis for Teachers (Streamlit app)
========================================================================
Upload all students' SAT/DSAT score-report PDFs for one test and download
the class-wide Word report:

    "<TEST CODE> SAT Test Result Analysis for Teacher.docx"

The test code and test date are read automatically from the PDFs.

Run locally:   streamlit run app.py
"""

import datetime
import hashlib
import io
import os
import tempfile
from collections import defaultdict
from pathlib import Path

import streamlit as st

# Load ANTHROPIC_API_KEY from a .env file next to this script, if present.
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass

from generate_class_report import (
    ClassStats,
    build_report,
    parse_date_iso,
    apply_skill_map,
    extract_pdf_text,
    load_skill_map,
    parse_skill_map_text,
    parse_text_all,
    report_filename,
    sanitize_branding,
)
from email_draft import (DEFAULT_CC, DEFAULT_TO, build_draft_message,
                         save_draft_to_gmail)

NAVY = "#1F3864"
BLUE = "#2E75B6"

st.set_page_config(
    page_title="Elite Prep — SAT Class Test Result Analysis",
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

st.title("SAT Test Result Analysis for Teachers")
st.caption(
    "Upload every student's SAT/DSAT practice-test score-report PDF for "
    "one test, then click Generate. The test code and test date are "
    "detected automatically, and the class-wide Word report is created "
    "in the standard Elite Prep teacher-report format."
)

# ---------------------------------------------------------------------------
# 1) Upload
# ---------------------------------------------------------------------------

st.header("1. Upload student score reports")
uploads = st.file_uploader(
    "Student score-report PDFs (one per student — select all at once). "
    "A question-to-skill map CSV for the test may be added too.",
    type=["pdf", "txt", "csv"],
    accept_multiple_files=True,
)

# ---------------------------------------------------------------------------
# 2) Generate
# ---------------------------------------------------------------------------

st.header("2. Generate the report")

# Settings come from .env (loaded above) / the environment, else secrets.toml.
def conf(name, default=""):
    value = os.environ.get(name, "")
    if not value:
        try:
            value = st.secrets.get(name, "")
        except Exception:  # no secrets.toml at all - that's fine
            value = ""
    return value or default


# Claude Opus 4.8 writes the commentary whenever an API key is available.
api_key = conf("ANTHROPIC_API_KEY")

if st.button("Generate Report", type="primary", use_container_width=True):
    if not uploads:
        st.error("Please upload at least one student score-report PDF first.")
        st.stop()

    students, failed = [], []
    uploaded_map = {}
    csv_uploads = [u for u in uploads if u.name.lower().endswith(".csv")]
    uploads = [u for u in uploads if not u.name.lower().endswith(".csv")]
    for u in csv_uploads:
        uploaded_map.update(parse_skill_map_text(
            u.getvalue().decode("utf-8-sig", "replace")))
    if not uploads:
        st.error("Please upload the student score-report PDFs as well.")
        st.stop()
    progress = st.progress(0.0, text="Parsing score reports...")
    for i, up in enumerate(sorted(uploads, key=lambda u: u.name)):
        try:
            if up.name.lower().endswith(".txt"):
                text = up.getvalue().decode("utf-8", "replace")
            else:
                text = extract_pdf_text(io.BytesIO(up.getvalue()))
            # Debug copy of the extracted text (local only, git-ignored) so
            # the parser can be adapted to new score-report layouts.
            try:
                dbg = Path(__file__).resolve().parent / "PDFs" / "_debug"
                dbg.mkdir(parents=True, exist_ok=True)
                dbg_file = dbg / (up.name + ".extracted.txt")
                dbg_file.write_text(text, encoding="utf-8")
                st.caption(f"Extracted text saved for parser tuning: "
                           f"`{dbg_file}`")
            except Exception as e:
                st.caption(f"(could not save extracted text: {e})")
            found = parse_text_all(text, up.name)
        except Exception as e:
            failed.append((up.name, f"could not read file: {e}"))
            found = None
        if found:
            students.extend(found)   # a file may hold many students
        elif found is not None:
            failed.append((up.name, "no question-level data recognized"))
        progress.progress(
            (i + 1) / len(uploads),
            text=f"Parsing score reports... ({i + 1}/{len(uploads)})")
    progress.empty()

    if failed:
        st.warning("Some files could not be analyzed and were skipped:\n\n"
                   + "\n".join(f"- **{n}** — {why}" for n, why in failed))
    if not students:
        st.error("No student data could be parsed from the uploaded files. "
                 "Please check that these are SAT score-report PDFs.")
        st.stop()

    # Safety net: if PDFs from more than one test were mixed in, analyze
    # the most recent test and tell the user what was set aside.
    groups = defaultdict(list)
    for s in students:
        code = sanitize_branding(s.test_code or "").strip() or "(unknown test)"
        groups[code].append(s)
    if len(groups) > 1:
        def group_key(code):
            dates = [parse_date_iso(s.test_date) for s in groups[code]
                     if s.test_date]
            dates = [d for d in dates if d]
            return (max(dates) if dates else "", len(groups[code]))
        current_code = max(groups, key=group_key)
        skipped = [c for c in groups if c != current_code]
        st.warning(f"PDFs from more than one test were uploaded "
                   f"({', '.join(sorted(groups))}). The report was "
                   f"generated for **{current_code}**; files from "
                   f"{', '.join(skipped)} were not included.")
        students = groups[current_code]

    test_code = sanitize_branding(
        next((s.test_code for s in students if s.test_code), None)
        or "SAT Practice Test").strip()
    test_date = (next((s.test_date for s in students if s.test_date), None)
                 or datetime.date.today().strftime("%B %d, %Y"))

    # --- question -> skill-area map (skill_maps/<code>.csv or uploaded CSV) --
    smap = load_skill_map(test_code)
    smap.update(uploaded_map)
    mapped = apply_skill_map(students, smap)
    if mapped:
        st.caption(f"Skill-area map applied for {test_code} "
                   f"({len(smap)} questions mapped).")
    elif not any(q.domain for s in students for q in s.questions):
        st.info(f"No question-to-skill map found for **{test_code}**, so "
                "section 2 (Average Accuracy by Skill Area) will show N/A. "
                f"Add `skill_maps/{test_code}.csv` (see "
                "`skill_maps/README.md`) or upload the CSV with the PDFs.")

    stats = ClassStats(students)

    # --- optional AI commentary (anonymized statistics only) -----------------
    narrative = None
    if not api_key:
        st.warning("No Anthropic API key found in .env - using the "
                   "built-in commentary instead.")
    else:
        with st.spinner("Claude Opus 4.8 is writing the commentary..."):
            try:
                from ai_narrative import write_narrative
                narrative = write_narrative(stats, test_code, test_date,
                                            api_key=api_key)
            except Exception as e:
                st.warning(f"AI commentary unavailable ({e}) - using "
                           "the built-in commentary instead.")

    # --- build the Word report ---------------------------------------------
    out_name = report_filename(test_code)
    tmp_docx = Path(tempfile.gettempdir()) / out_name
    build_report(stats, test_code, test_date, tmp_docx, narrative=narrative)
    docx_bytes = tmp_docx.read_bytes()

    st.success(f"Report generated for **{stats.n} student"
               f"{'s' if stats.n != 1 else ''}** — Test Code "
               f"**{test_code}**, Test Date **{test_date}** "
               f"(detected from the PDFs). "
               + ("Commentary written by Claude Opus 4.8. " if narrative
                  else "")
               + "No student names appear in the report.")

    # --- Gmail draft (the app never sends; it only prepares the draft) -------
    gmail_user, gmail_password = conf("GMAIL_ADDRESS"), conf(
        "GMAIL_APP_PASSWORD")
    if gmail_user and gmail_password:
        # Streamlit re-runs this whole script on every click, so key the draft
        # by the report itself - one report, one draft.
        drafted = st.session_state.setdefault("drafts_made", {})
        key = hashlib.sha1(docx_bytes).hexdigest()
        if key not in drafted:
            recipient = conf("REPORT_TO", DEFAULT_TO)
            copied = conf("REPORT_CC", DEFAULT_CC)
            try:
                with st.spinner("Saving the report to Gmail Drafts ..."):
                    save_draft_to_gmail(
                        build_draft_message(
                            gmail_user, recipient, test_code, test_date,
                            stats.n,
                            {k: stats.avg(k)
                             for k in ("total", "rw", "math")},
                            docx_bytes, out_name, cc=copied),
                        gmail_user, gmail_password)
                drafted[key] = (True, f"{recipient} (cc {copied})"
                                if copied else recipient)
            except Exception as e:  # never let mail trouble hide the report
                drafted[key] = (False, str(e))
        ok, detail = drafted[key]
        if ok:
            st.info(f"A draft to **{detail}** is waiting in the Gmail Drafts "
                    f"folder of {gmail_user} - review it there and press "
                    "Send. "
                    "[Open Drafts](https://mail.google.com/mail/u/0/#drafts)")
        else:
            st.warning(f"The Gmail draft could not be created ({detail}) - "
                       "use the download button below instead.")

    # --- download ------------------------------------------------------------
    st.download_button(
        f"⬇️ Download Word report — {out_name}",
        data=docx_bytes,
        file_name=out_name,
        mime=("application/vnd.openxmlformats-officedocument"
              ".wordprocessingml.document"),
        use_container_width=True,
    )

    # --- on-screen preview -----------------------------------------------------
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
