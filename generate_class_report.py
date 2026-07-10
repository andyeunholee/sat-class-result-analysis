#!/usr/bin/env python3
"""
Elite Prep - SAT Class Results Analysis Report Generator
=========================================================
Reads all student SAT/DSAT practice-test score-report PDFs in a folder,
aggregates class-wide statistics, and generates a Word (.docx) report:

    "<TEST CODE> Result Analysis Teacher Report.docx"

Rules enforced (per project instructions):
  * Report text is always in English.
  * The content/format of the report template is preserved; only branding
    styling is applied ("Elite Prep" navy/blue design system).
  * Branding is always written as just "Elite Prep" (never "Elite Prep
    Suwanee"), and there is no "Andy Lee, Director, ..." signature line.
  * No footers on any page.
  * Page 1 includes a score-trend graph (class average from the first
    practice test to the latest), based on a score history that this
    program maintains automatically (score_history.csv).

Usage:
    python generate_class_report.py <folder_with_pdfs>
        [--test-code DSAT-02-A] [--test-date "June 8, 2026"]
        [--output-dir DIR] [--domains domains.csv]
        [--history score_history.csv] [--no-history] [--dump-text]

If --test-code / --test-date are omitted, the script tries to detect them
from the PDFs (and falls back to the folder name / today's date).

--dump-text writes each PDF's raw extracted text to <name>.extracted.txt
so the parser can be adapted if a new score-report layout appears.
"""

import argparse
import csv
import datetime
import re
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

import pdfplumber
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# ---------------------------------------------------------------------------
# Elite Prep design system (navy / blue)
# ---------------------------------------------------------------------------

NAVY = RGBColor(0x1F, 0x38, 0x64)        # primary brand navy
BLUE = RGBColor(0x2E, 0x74, 0xB5)        # accent blue
LIGHT_BLUE_HEX = "DCE6F1"                # table zebra / header tint
NAVY_HEX = "1F3864"
RED = RGBColor(0xC0, 0x00, 0x00)         # alert highlight (kept per template)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x59, 0x59, 0x59)

CHART_NAVY = "#1F3864"
CHART_BLUE = "#2E74B5"
CHART_SKY = "#7FA8D9"

RW_DOMAINS = [
    "Craft and Structure",
    "Information and Ideas",
    "Standard English Conventions",
    "Expression of Ideas",
]
MATH_DOMAINS = [
    "Algebra",
    "Advanced Math",
    "Problem-Solving & Data Analysis",
    "Geometry & Trigonometry",
]

# Variants of domain names that may appear in score reports -> canonical name
DOMAIN_ALIASES = {
    "craft and structure": "Craft and Structure",
    "information and ideas": "Information and Ideas",
    "standard english conventions": "Standard English Conventions",
    "expression of ideas": "Expression of Ideas",
    "algebra": "Algebra",
    "advanced math": "Advanced Math",
    "problem-solving and data analysis": "Problem-Solving & Data Analysis",
    "problem solving and data analysis": "Problem-Solving & Data Analysis",
    "problem-solving & data analysis": "Problem-Solving & Data Analysis",
    "geometry and trigonometry": "Geometry & Trigonometry",
    "geometry & trigonometry": "Geometry & Trigonometry",
}

CIRCLED = {1: "①", 2: "②", 3: "③", 4: "④"}  # ① ② ③ ④


def sanitize_branding(text: str) -> str:
    """Branding must always be written as just 'Elite Prep'."""
    return re.sub(r"Elite\s*Prep\s*Suwanee", "Elite Prep", text, flags=re.I)


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class QuestionResult:
    __slots__ = ("section", "module", "number", "status", "domain")

    def __init__(self, section, module, number, status, domain=None):
        self.section = section      # "RW" or "Math"
        self.module = module        # 1 or 2
        self.number = number        # int question number
        self.status = status        # "correct" | "incorrect" | "omitted"
        self.domain = domain        # canonical domain name or None


class StudentResult:
    def __init__(self, source):
        self.source = source        # file name
        self.total = None           # int or None
        self.rw = None
        self.math = None
        self.questions = []         # list[QuestionResult]
        self.test_code = None
        self.test_date = None

    @property
    def parsed_ok(self):
        return bool(self.questions)


# ---------------------------------------------------------------------------
# PDF parsing
# ---------------------------------------------------------------------------

STATUS_WORDS = {
    "correct": "correct",
    "incorrect": "incorrect",
    "omitted": "omitted",
    "unanswered": "omitted",
    "no answer": "omitted",
    "skipped": "omitted",
}

SECTION_PATTERNS = [
    (re.compile(r"reading\s*(?:and|&)\s*writing", re.I), "RW"),
    (re.compile(r"\bR\s*&\s*W\b", re.I), "RW"),
    (re.compile(r"\bmath\b", re.I), "Math"),
]

MODULE_PATTERN = re.compile(r"module\s*([12])", re.I)

# Row style A: "12 Reading and Writing: Module 1 B A" (+ optional status word)
ROW_A = re.compile(
    r"^\s*(?:Question\s*)?(\d{1,2})[.)]?\s+"
    r"(Reading\s*(?:and|&)\s*Writing|Math)\s*[:\-]?\s*Module\s*([12])\b(.*)$",
    re.I,
)

# Row style B (section/module known from headers): "12 B A Incorrect" or
# "12 Craft and Structure B A" or "12 Incorrect" etc.
ROW_B = re.compile(r"^\s*(?:Question\s*)?(\d{1,2})[.)]?\s+(\S.*)$")

SCORE_PATTERNS = {
    "total": re.compile(r"total\s*(?:sat\s*)?score\s*[:\-]?\s*(\d{3,4})", re.I),
    "rw": re.compile(
        r"reading\s*(?:and|&)\s*writing(?:\s*(?:section)?\s*score)?\s*[:\-]?\s*(\d{3})",
        re.I,
    ),
    "math": re.compile(r"math(?:\s*(?:section)?\s*score)?\s*[:\-]?\s*(\d{3})", re.I),
}

TEST_CODE_PATTERNS = [
    re.compile(r"test\s*code\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9\-_ ]{1,30}?)\s*$", re.I | re.M),
    re.compile(r"\b(DSAT[-_ ]?\d+[-_ ]?[A-Z]?)\b", re.I),
    re.compile(r"\b(SAT\s*Practice\s*(?:Test\s*)?#?\s*\d+)\b", re.I),
    re.compile(r"\b(Practice\s*Test\s*\d+)\b", re.I),
    re.compile(r"\b(Bluebook\s*Practice\s*#?\s*\d+)\b", re.I),
]

DATE_PATTERNS = [
    re.compile(
        r"\b(January|February|March|April|May|June|July|August|September|"
        r"October|November|December)\s+\d{1,2},\s*\d{4}\b"
    ),
    re.compile(r"\b\d{1,2}/\d{1,2}/\d{4}\b"),
]


def extract_status_and_domain(fragment: str):
    """Pull correct/incorrect/omitted and a skill-domain name out of a row
    fragment. Returns (status_or_None, domain_or_None, remainder)."""
    frag = fragment.strip()
    low = frag.lower()

    domain = None
    for alias, canonical in DOMAIN_ALIASES.items():
        if alias in low:
            domain = canonical
            idx = low.find(alias)
            frag = (frag[:idx] + frag[idx + len(alias):]).strip(" -:;|")
            low = frag.lower()
            break

    status = None
    for word, st in STATUS_WORDS.items():
        if re.search(r"\b" + re.escape(word) + r"\b", low):
            status = st
            break

    return status, domain, frag


def eval_fraction(s: str):
    if re.fullmatch(r"-?\d+/\d+", s):
        num, den = s.split("/")
        return float(num) / float(den)
    return float(s)


def infer_status_from_answers(fragment: str):
    """Fallback: compare 'Correct Answer' vs 'Your Answer' tokens.
    Handles rows like 'B A', 'B B', 'C Omitted', '3/4 .75', '104 --'."""
    frag = fragment.strip(" .|")
    if not frag:
        return None
    tokens = frag.split()
    if len(tokens) < 2:
        return None
    correct_ans, your_ans = tokens[-2], tokens[-1]
    ylow = your_ans.lower().strip(".,;")
    if ylow in ("omitted", "unanswered", "--", "-", "blank", "—"):
        return "omitted"

    def norm(a):
        a = a.strip().strip(".,;").upper()
        try:
            return format(eval_fraction(a), ".6g")
        except Exception:
            return a

    return "correct" if norm(correct_ans) == norm(your_ans) else "incorrect"


def parse_text(text: str, source: str) -> StudentResult:
    """Parse the extracted text of one student's score-report PDF."""
    result = StudentResult(source)

    for key, pat in SCORE_PATTERNS.items():
        m = pat.search(text)
        if m:
            val = int(m.group(1))
            if key == "total" and 400 <= val <= 1600:
                result.total = val
            elif key in ("rw", "math") and 200 <= val <= 800:
                setattr(result, key, val)

    for pat in TEST_CODE_PATTERNS:
        m = pat.search(text)
        if m:
            result.test_code = m.group(1).strip()
            break
    for pat in DATE_PATTERNS:
        m = pat.search(text)
        if m:
            result.test_date = m.group(0)
            break

    seen = set()  # (section, module, number) — keep first occurrence
    cur_section, cur_module = None, None

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        m = ROW_A.match(line)
        if m:
            number = int(m.group(1))
            section = "Math" if re.match(r"^\s*math", m.group(2), re.I) else "RW"
            module = int(m.group(3))
            status, domain, rest = extract_status_and_domain(m.group(4))
            if status is None:
                status = infer_status_from_answers(rest)
            if status and 1 <= number <= 40:
                key = (section, module, number)
                if key not in seen:
                    seen.add(key)
                    result.questions.append(
                        QuestionResult(section, module, number, status, domain)
                    )
            continue

        # Header lines set parsing context for style-B rows.
        header_hit = False
        for pat, sec in SECTION_PATTERNS:
            if pat.search(line) and len(line) < 80:
                cur_section = sec
                header_hit = True
                break
        mm = MODULE_PATTERN.search(line)
        if mm and len(line) < 80:
            cur_module = int(mm.group(1))
            header_hit = True
        if header_hit and not re.match(r"^\s*\d", line):
            continue

        if cur_section and cur_module:
            m = ROW_B.match(line)
            if m:
                number = int(m.group(1))
                if not (1 <= number <= 40):
                    continue
                status, domain, rest = extract_status_and_domain(m.group(2))
                if status is None:
                    status = infer_status_from_answers(rest)
                if status:
                    key = (cur_section, cur_module, number)
                    if key not in seen:
                        seen.add(key)
                        result.questions.append(
                            QuestionResult(cur_section, cur_module, number,
                                           status, domain)
                        )

    return result


def extract_pdf_text(source) -> str:
    """Extract text from a PDF path or file-like object (e.g. an upload)."""
    chunks = []
    with pdfplumber.open(source) as pdf:
        for page in pdf.pages:
            chunks.append(page.extract_text() or "")
            # Table rows sometimes extract more cleanly than raw text.
            try:
                for table in page.extract_tables():
                    for row in table:
                        cells = [c.strip() for c in row if c and c.strip()]
                        if cells:
                            chunks.append("  ".join(cells))
            except Exception:
                pass
    return "\n".join(chunks)


def parse_pdf_stream(fileobj, name: str) -> StudentResult:
    """Parse a student's score report from a file-like object (web upload)."""
    return parse_text(extract_pdf_text(fileobj), name)


def parse_pdf(path: Path, dump_text: bool = False) -> StudentResult:
    if path.suffix.lower() == ".txt":  # debugging convenience
        text = path.read_text(encoding="utf-8", errors="replace")
    else:
        text = extract_pdf_text(str(path))

    if dump_text:
        dump_path = path.with_suffix(path.suffix + ".extracted.txt")
        dump_path.write_text(text, encoding="utf-8")
        print(f"  [debug] extracted text written to {dump_path.name}")

    return parse_text(text, path.name)


# ---------------------------------------------------------------------------
# Aggregation
# ---------------------------------------------------------------------------

class ClassStats:
    def __init__(self, students):
        self.students = students
        self.n = len(students)

        def series(attr):
            return [getattr(s, attr) for s in students
                    if getattr(s, attr) is not None]

        self.scores = {k: series(k) for k in ("total", "rw", "math")}

        # missed counts: {(section, module)}[qnum] = students who missed it
        self.missed = defaultdict(lambda: defaultdict(int))
        self.attempted = defaultdict(lambda: defaultdict(int))
        self.domain_correct = defaultdict(int)
        self.domain_total = defaultdict(int)
        self.late_omissions = 0  # omitted questions late in a module

        for s in students:
            module_max = defaultdict(int)
            for q in s.questions:
                module_max[(q.section, q.module)] = max(
                    module_max[(q.section, q.module)], q.number)
            for q in s.questions:
                key = (q.section, q.module)
                self.attempted[key][q.number] += 1
                if q.status != "correct":
                    self.missed[key][q.number] += 1
                if q.domain:
                    self.domain_total[q.domain] += 1
                    if q.status == "correct":
                        self.domain_correct[q.domain] += 1
                if q.status == "omitted":
                    if q.number >= max(1, int(module_max[key] * 2 / 3)):
                        self.late_omissions += 1

    def avg(self, key):
        vals = self.scores[key]
        return round(sum(vals) / len(vals)) if vals else None

    def rng(self, key):
        vals = self.scores[key]
        return (min(vals), max(vals)) if vals else None

    def domain_accuracy(self, domain):
        t = self.domain_total.get(domain, 0)
        if t == 0:
            return None
        return round(100 * self.domain_correct[domain] / t)

    def has_domain_data(self):
        return any(self.domain_total.values())

    def error_groups(self, section, module):
        """Return list of (missed_count, [qnums]) sorted by missed desc."""
        counts = self.missed[(section, module)]
        groups = defaultdict(list)
        for qnum, c in counts.items():
            if c > 0:
                groups[c].append(qnum)
        return [(c, sorted(groups[c])) for c in sorted(groups, reverse=True)]


# ---------------------------------------------------------------------------
# Score history (for the first-page trend graph)
# ---------------------------------------------------------------------------

HISTORY_FIELDS = ["test_code", "test_date", "date_iso", "students",
                  "avg_total", "avg_rw", "avg_math"]


def parse_date_iso(date_str):
    for fmt in ("%B %d, %Y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(date_str, fmt).date().isoformat()
        except (ValueError, TypeError):
            continue
    return None


def upsert_history_rows(rows, test_code, test_date, stats):
    """Merge this test's class averages into a list of history rows and
    return the rows sorted chronologically (oldest first)."""
    rows = [r for r in rows if r.get("test_code") and r["test_code"] != test_code]
    rows.append({
        "test_code": test_code,
        "test_date": test_date,
        "date_iso": parse_date_iso(test_date) or "",
        "students": str(stats.n),
        "avg_total": str(stats.avg("total") or ""),
        "avg_rw": str(stats.avg("rw") or ""),
        "avg_math": str(stats.avg("math") or ""),
    })
    rows.sort(key=lambda r: (r.get("date_iso") or "9999-99-99", r["test_code"]))
    return rows


def history_rows_to_csv(rows) -> str:
    import io
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=HISTORY_FIELDS, extrasaction="ignore")
    w.writeheader()
    w.writerows(rows)
    return buf.getvalue()


def read_history_csv(fileobj_or_text):
    """Read history rows from an open file object or a CSV text string."""
    import io
    if isinstance(fileobj_or_text, str):
        fileobj_or_text = io.StringIO(fileobj_or_text)
    return [r for r in csv.DictReader(fileobj_or_text) if r.get("test_code")]


def update_history(history_path: Path, test_code, test_date, stats):
    """Upsert this test's class averages into the history CSV; return all
    rows sorted chronologically (oldest first)."""
    rows = []
    if history_path.exists():
        with open(history_path, newline="", encoding="utf-8-sig") as f:
            rows = read_history_csv(f)

    rows = upsert_history_rows(rows, test_code, test_date, stats)

    with open(history_path, "w", newline="", encoding="utf-8") as f:
        f.write(history_rows_to_csv(rows))

    return rows


def make_trend_chart(history_rows, out_png: Path):
    """Draw the class-average score curve (first test -> latest test) using
    the Elite Prep navy/blue palette. Returns True if a chart was drawn."""
    pts = [(r["test_code"],
            int(r["avg_total"]) if r.get("avg_total") else None,
            int(r["avg_rw"]) if r.get("avg_rw") else None,
            int(r["avg_math"]) if r.get("avg_math") else None)
           for r in history_rows]
    pts = [p for p in pts if p[1] is not None or p[2] is not None or p[3] is not None]
    if not pts:
        return False

    labels = [p[0] for p in pts]
    x = list(range(len(pts)))

    fig, ax = plt.subplots(figsize=(7.2, 3.1), dpi=200)
    fig.patch.set_facecolor("white")

    series = [
        ("Total (400-1600)", [p[1] for p in pts], CHART_NAVY, "o", 2.4),
        ("Reading & Writing", [p[2] for p in pts], CHART_BLUE, "s", 1.8),
        ("Math", [p[3] for p in pts], CHART_SKY, "^", 1.8),
    ]
    for name, ys, color, marker, lw in series:
        if any(y is not None for y in ys):
            ax.plot(x, ys, color=color, marker=marker, linewidth=lw,
                    markersize=5, label=name)
            for xi, yi in zip(x, ys):
                if yi is not None:
                    ax.annotate(str(yi), (xi, yi), textcoords="offset points",
                                xytext=(0, 7), ha="center", fontsize=7,
                                color=color)

    ax.set_title("Class Average Score Trend (First Test → Latest Test)",
                 fontsize=11, color=CHART_NAVY, fontweight="bold", pad=12)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=8, color="#333333")
    ax.tick_params(axis="y", labelsize=8, colors="#333333")
    ax.grid(axis="y", color="#D9E2F0", linewidth=0.8)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    for spine in ("left", "bottom"):
        ax.spines[spine].set_color("#B7C6DD")
    ax.legend(fontsize=8, frameon=False, loc="best")
    ax.margins(y=0.22)

    fig.tight_layout()
    fig.savefig(str(out_png), facecolor="white")
    plt.close(fig)
    return True


# ---------------------------------------------------------------------------
# DOCX helpers (Elite Prep navy/blue design system)
# ---------------------------------------------------------------------------

def shade_cell(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def set_cell_text(cell, text, bold=False, color=None, align=None, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    run = p.add_run(sanitize_branding(text))
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color


def add_para(doc, text="", bold=False, size=11, align=None, italic=False,
             color=None, space_after=6):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(sanitize_branding(text))
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return p


def add_rich_para(doc, fragments, size=11, align=None, space_after=6):
    """fragments: list of (text, bold) or (text, bold, color)."""
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    for frag in fragments:
        text, bold = frag[0], frag[1]
        color = frag[2] if len(frag) > 2 else None
        run = p.add_run(sanitize_branding(text))
        run.bold = bold
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = color
    return p


def add_heading(doc, text, level=1):
    h = doc.add_heading(sanitize_branding(text), level=level)
    for run in h.runs:
        run.font.color.rgb = NAVY if level == 1 else BLUE
    return h


def add_brand_rule(doc):
    """Thin navy horizontal rule under the letterhead."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), NAVY_HEX)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def make_table(doc, headers, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, h, bold=True, color=WHITE)
        shade_cell(cell, NAVY_HEX)
    if widths:
        for i, w in enumerate(widths):
            for row in table.rows:
                row.cells[i].width = Inches(w)
    return table


def zebra(table):
    """Light-blue tint on alternating body rows for readability."""
    for idx, row in enumerate(table.rows[1:]):
        if idx % 2 == 1:
            for cell in row.cells:
                shade_cell(cell, LIGHT_BLUE_HEX)


def apply_widths(table, widths):
    for row in table.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Inches(w)


def remove_footers(doc):
    for section in doc.sections:
        for footer in (section.footer, section.first_page_footer,
                       section.even_page_footer):
            try:
                footer.is_linked_to_previous = True
                for p in list(footer.paragraphs):
                    p.clear()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Report text helpers
# ---------------------------------------------------------------------------

def ordinal_section_title(i, section, module):
    label = "Math" if section == "Math" else "English (Reading & Writing)"
    return f"{CIRCLED[i]} {label} — Section {module} (Module {module})"


def q_list(qnums):
    return ", ".join(f"Q{q}" for q in qnums)


def q_dots(qnums):
    return "·".join(f"Q{q}" for q in qnums)


def build_takeaway(section, module, groups, n_students):
    if not groups:
        return ("→ No questions in this section were missed by more "
                "than one student.")
    top_count, top_qs = groups[0]
    if top_count >= max(2, round(n_students * 0.8)):
        verb = "were the most challenging questions" if len(top_qs) > 1 \
            else "was the most challenging question"
        return (f"→ {q_dots(top_qs)}, missed by {top_count} of "
                f"{n_students} students, {verb} on this test.")
    hard = sorted({q for c, qs in groups if c / n_students >= 0.5 for q in qs})
    if hard:
        if section == "Math" and all(q >= 14 for q in hard):
            return (f"→ Errors were concentrated in the "
                    f"higher-difficulty latter portion ({q_dots(hard)}).")
        word = "questions" if len(hard) > 1 else "question"
        return (f"→ {len(hard)} {word} ({q_dots(hard)}) showed the "
                f"highest error rates and will be prioritized in class review.")
    return (f"→ The highest error rate in this section was "
            f"{round(100 * top_count / n_students)}% ({q_dots(top_qs)}).")


def weakest_domains(stats, domains, threshold=65):
    pairs = [(d, stats.domain_accuracy(d)) for d in domains]
    pairs = [(d, a) for d, a in pairs if a is not None]
    pairs.sort(key=lambda x: x[1])
    weak = [d for d, a in pairs if a < threshold][:2]
    return weak or [p[0] for p in pairs[:1]]


# ---------------------------------------------------------------------------
# DOCX report
# ---------------------------------------------------------------------------

def build_report(stats: ClassStats, test_code: str, test_date: str,
                 out_path: Path, trend_png: Path = None):
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")

    n = stats.n

    # --- Letterhead (page 1) ------------------------------------------------
    add_rich_para(doc, [("ELITE PREP", True, NAVY),
                        ("  |  College Admissions & Test Prep", False, BLUE)],
                  size=15, space_after=0)
    add_brand_rule(doc)

    # --- Title ---------------------------------------------------------------
    try:
        dt = datetime.datetime.strptime(test_date, "%B %d, %Y")
        month_year = dt.strftime("%B %Y")
    except ValueError:
        month_year = test_date
    add_para(doc, f"{month_year} DSAT Practice Test Results Analysis Report",
             bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER,
             color=NAVY, space_after=8)
    add_para(doc,
             f"For Teachers  ·  Test Code {test_code}  ·  "
             f"Test Date: {test_date}",
             align=WD_ALIGN_PARAGRAPH.CENTER, color=GRAY, space_after=14)

    # --- Score trend graph (first page) --------------------------------------
    if trend_png and trend_png.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(trend_png), width=Inches(6.4))
        add_para(doc, "Class average scores across all practice tests "
                      "administered to date, from the first test to this "
                      "test.", italic=True, size=9, color=GRAY,
                 align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

    # --- Intro ---------------------------------------------------------------
    add_para(doc, "Dear Teachers,")
    add_para(
        doc,
        "Hello, this is Elite Prep. We are pleased to share an analysis of "
        "the class-wide results from the DSAT (Digital SAT) practice test "
        f"administered at our academy ({test_code}). This report summarizes "
        "the overall performance of all students who took the test. For each "
        "individual student's score, please refer to the personal Score "
        "Report provided separately.",
    )
    add_rich_para(doc, [
        ("A total of ", False),
        (f"{n} student{'s' if n != 1 else ''}", True),
        (" sat for this practice test. To protect each student's privacy, "
         "the analysis below is presented entirely as anonymized, aggregate "
         "statistics.", False),
    ])

    # --- 1. Class Performance Overview ----------------------------------------
    add_heading(doc, "1. Class Performance Overview", level=1)
    add_para(doc, "The class averages and score ranges for this test are as "
                  "follows. (SAT total score scale: 400-1600)")

    if any(stats.scores[k] for k in ("total", "rw", "math")):
        table = make_table(doc, ["Category", "Class Average",
                                 "Score Range (Low-High)"])
        for label, key in (("Total", "total"), ("Reading & Writing", "rw"),
                           ("Math", "math")):
            avg, rng = stats.avg(key), stats.rng(key)
            r = table.add_row().cells
            set_cell_text(r[0], label, bold=True, color=NAVY)
            set_cell_text(r[1], str(avg) if avg is not None else "N/A")
            set_cell_text(r[2], f"{rng[0]} - {rng[1]}" if rng else "N/A")
        zebra(table)
        apply_widths(table, [2.0, 1.8, 2.4])

        rw_avg, math_avg = stats.avg("rw"), stats.avg("math")
        if rw_avg is not None and math_avg is not None:
            if math_avg >= rw_avg:
                stronger, weaker = ("Math", math_avg), ("Reading & Writing", rw_avg)
            else:
                stronger, weaker = ("Reading & Writing", rw_avg), ("Math", math_avg)
            doc.add_paragraph()
            add_rich_para(doc, [
                (f"Overall, the {stronger[0]} section (average ", False),
                (str(stronger[1]), True),
                (f") showed somewhat more stable performance than the "
                 f"{weaker[0]} section (average {weaker[1]}). However, "
                 f"because the variation between students is significant, we "
                 f"recommend reviewing these figures together with each "
                 f"student's individual Score Report.", False),
            ])
    else:
        add_para(doc, "Section scores were not present in the uploaded score "
                      "reports, so the score overview is omitted.",
                 italic=True, color=GRAY)

    # --- 2. Average Accuracy by Skill Area -------------------------------------
    add_heading(doc, "2. Average Accuracy by Skill Area", level=1)
    if stats.has_domain_data():
        add_para(doc, "Below are the class-wide average accuracy rates for "
                      "the eight detailed skill areas of the SAT. Areas with "
                      "an average accuracy below 50% are highlighted in red "
                      "so that you can identify, at a glance, the areas "
                      "requiring focused improvement.")
        table = make_table(doc, ["Subject", "Skill Area", "Average Accuracy"])
        for subject, domains in (("Reading & Writing", RW_DOMAINS),
                                 ("Math", MATH_DOMAINS)):
            for i, d in enumerate(domains):
                acc = stats.domain_accuracy(d)
                r = table.add_row().cells
                set_cell_text(r[0], subject if i == 0 else "",
                              bold=True, color=NAVY)
                set_cell_text(r[1], d)
                if acc is None:
                    set_cell_text(r[2], "N/A")
                else:
                    low = acc < 50
                    set_cell_text(r[2], f"{acc}%", bold=low,
                                  color=RED if low else None)
        zebra(table)
        apply_widths(table, [1.9, 2.9, 1.6])

        weak_rw = weakest_domains(stats, RW_DOMAINS)
        weak_math = weakest_domains(stats, MATH_DOMAINS)
        parts = []
        if weak_rw:
            names = "' and '".join(weak_rw)
            parts.append(f"the '{names}' area{'s' if len(weak_rw) > 1 else ''}"
                         f" in Reading & Writing")
        if weak_math:
            names = "' and '".join(weak_math)
            parts.append(f"the '{names}' area{'s' if len(weak_math) > 1 else ''}"
                         f" in Math")
        if parts:
            doc.add_paragraph()
            add_para(doc, "As the table shows, " + ", along with ".join(parts)
                          + ", proved relatively weak. Our academy will "
                            "provide supplementary instruction focused on "
                            "these areas during class.")
    else:
        add_para(doc, "Skill-area (domain) information was not available in "
                      "the uploaded score reports, so per-skill accuracy "
                      "could not be computed for this test.",
                 italic=True, color=GRAY)

    # --- 3. Priority Review Questions by Section --------------------------------
    add_heading(doc, "3. Priority Review Questions by Section", level=1)
    add_para(doc, "For each section, the questions that students most "
                  "frequently answered incorrectly are listed in order of "
                  "error rate. Our academy will prioritize these questions "
                  "for review and concept reinforcement in class. (Questions "
                  "with an error rate of 60% or higher are marked in red.)")

    section_order = [("Math", 1), ("Math", 2), ("RW", 1), ("RW", 2)]
    for i, (section, module) in enumerate(section_order, start=1):
        add_heading(doc, ordinal_section_title(i, section, module), level=2)
        groups = stats.error_groups(section, module)
        if not groups:
            add_para(doc, "No incorrect answers were recorded for this "
                          "section, or no data was found in the score "
                          "reports.", italic=True, color=GRAY)
            continue
        table = make_table(doc, ["Error Rate", "Students Missed", "Questions"])
        for count, qnums in groups:
            rate = round(100 * count / n)
            red = rate >= 60
            r = table.add_row().cells
            set_cell_text(r[0], f"{rate}%", bold=red,
                          color=RED if red else None)
            set_cell_text(r[1], f"{count} of {n}", bold=red,
                          color=RED if red else None)
            set_cell_text(r[2], q_list(qnums), bold=red,
                          color=RED if red else None)
        zebra(table)
        apply_widths(table, [1.3, 1.6, 3.5])
        doc.add_paragraph()
        add_para(doc, build_takeaway(section, module, groups, n),
                 italic=True, color=BLUE, space_after=10)

    # --- 4. Our Instructional Plan -----------------------------------------------
    add_heading(doc, "4. Our Instructional Plan", level=1)

    math_weak = weakest_domains(stats, MATH_DOMAINS) if stats.has_domain_data() else []
    rw_weak = weakest_domains(stats, RW_DOMAINS) if stats.has_domain_data() else []

    math_bullet = ("Math: We will focus intensively on the questions with "
                   "the highest class-wide error rates identified above"
                   + (f", with particular reinforcement of "
                      f"{' and '.join(math_weak)} concepts."
                      if math_weak else
                      ", including the higher-difficulty latter-portion and "
                      "student-produced response (grid-in) questions."))
    if rw_weak:
        rw_names = f"'{rw_weak[0]}'" + (f" and '{rw_weak[1]}'"
                                        if len(rw_weak) > 1 else "")
        rw_focus = f"the {rw_names} question types"
    else:
        rw_focus = "the question types missed most frequently by the class"
    rw_bullet = (f"Reading & Writing: We will repeatedly train solving "
                 f"strategies for {rw_focus} to improve reading and "
                 f"reasoning accuracy.")

    bullets = [math_bullet, rw_bullet]
    if stats.late_omissions > 0:
        bullets.append(
            "Time Management: Since some students showed instances of "
            "unanswered questions in the latter portion (running out of "
            "time), we will provide separate coaching on real-test pacing "
            "and strategies for handling the final questions.")
    for bullet in bullets:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(sanitize_branding(bullet))
        run.font.size = Pt(11)

    # --- Closing -------------------------------------------------------------
    doc.add_paragraph()
    add_para(doc, "We will continue to walk alongside you in support of your "
                  "students' growth. Thank you.")
    doc.add_paragraph()
    add_para(doc, "Elite Prep", bold=True, color=NAVY, space_after=2)
    add_para(doc, "www.eliteprep.com", color=BLUE)

    remove_footers(doc)
    doc.save(str(out_path))


# ---------------------------------------------------------------------------
# Optional domain mapping CSV
# ---------------------------------------------------------------------------

def load_domain_map(csv_path: Path):
    """CSV columns: section,module,question,domain
       section: RW|Math (or English/Reading and Writing);
       domain: canonical name or alias."""
    mapping = {}
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            sec_raw = row["section"].strip().lower()
            sec = "Math" if sec_raw.startswith("m") else "RW"
            dom = DOMAIN_ALIASES.get(row["domain"].strip().lower(),
                                     row["domain"].strip())
            mapping[(sec, int(row["module"]), int(row["question"]))] = dom
    return mapping


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Generate an Elite Prep class-wide SAT results analysis "
                    "Word report from a folder of student score-report PDFs.")
    ap.add_argument("folder", nargs="?", default=".",
                    help="Folder containing the student PDF score reports "
                         "(default: current folder)")
    ap.add_argument("--test-code", help="Test code, e.g. DSAT-02-A "
                                        "(auto-detected if omitted)")
    ap.add_argument("--test-date", help='Test date, e.g. "June 8, 2026" '
                                        "(auto-detected if omitted)")
    ap.add_argument("--output-dir", help="Where to save the Word report "
                                         "(default: the PDF folder)")
    ap.add_argument("--domains", help="Optional CSV mapping questions to "
                                      "skill areas "
                                      "(section,module,question,domain)")
    ap.add_argument("--history",
                    help="Path of the score-history CSV used for the "
                         "first-page trend graph (default: "
                         "score_history.csv next to this script)")
    ap.add_argument("--no-history", action="store_true",
                    help="Skip the score history / trend graph entirely")
    ap.add_argument("--dump-text", action="store_true",
                    help="Also write each PDF's extracted raw text to a .txt "
                         "file (for debugging new report layouts)")
    args = ap.parse_args(argv)

    folder = Path(args.folder).resolve()
    if not folder.is_dir():
        sys.exit(f"ERROR: folder not found: {folder}")

    pdfs = sorted([p for p in folder.iterdir()
                   if p.suffix.lower() in (".pdf", ".txt")
                   and not p.name.endswith(".extracted.txt")])
    if not pdfs:
        sys.exit(f"ERROR: no PDF files found in {folder}")

    print(f"Found {len(pdfs)} file(s) in {folder}\n")

    students, failed = [], []
    for p in pdfs:
        print(f"Parsing {p.name} ...")
        try:
            s = parse_pdf(p, dump_text=args.dump_text)
        except Exception as e:
            print(f"  !! could not read file: {e}")
            failed.append(p.name)
            continue
        if s.parsed_ok:
            n_wrong = sum(1 for q in s.questions if q.status != "correct")
            print(f"  ok: {len(s.questions)} questions "
                  f"({n_wrong} missed), scores: total={s.total} "
                  f"RW={s.rw} Math={s.math}")
            students.append(s)
        else:
            print("  !! no question-level data recognized "
                  "(run with --dump-text and share the .extracted.txt file "
                  "so the parser can be extended)")
            failed.append(p.name)

    if not students:
        sys.exit("\nERROR: no student data could be parsed from any file. "
                 "Re-run with --dump-text to inspect the PDF text layout.")

    # Optional external domain mapping fills in missing skill areas.
    if args.domains:
        dmap = load_domain_map(Path(args.domains))
        for s in students:
            for q in s.questions:
                if q.domain is None:
                    q.domain = dmap.get((q.section, q.module, q.number))

    stats = ClassStats(students)

    test_code = (args.test_code
                 or next((s.test_code for s in students if s.test_code), None)
                 or folder.name)
    test_date = (args.test_date
                 or next((s.test_date for s in students if s.test_date), None)
                 or datetime.date.today().strftime("%B %d, %Y"))
    test_code = sanitize_branding(test_code).strip()

    # Score history + first-page trend graph
    trend_png = None
    if not args.no_history:
        history_path = (Path(args.history).resolve() if args.history
                        else Path(__file__).resolve().parent / "score_history.csv")
        rows = update_history(history_path, test_code, test_date, stats)
        trend_png = Path(tempfile.gettempdir()) / "eliteprep_trend.png"
        if not make_trend_chart(rows, trend_png):
            trend_png = None
        print(f"\nScore history updated: {history_path} "
              f"({len(rows)} test(s) on record)")

    out_dir = Path(args.output_dir).resolve() if args.output_dir else folder
    safe_code = re.sub(r'[\\/:*?"<>|]', "-", test_code)
    out_path = out_dir / f"{safe_code} Result Analysis Teacher Report.docx"

    build_report(stats, test_code, test_date, out_path, trend_png)

    print(f"\nDone. {len(students)} student(s) analyzed"
          + (f", {len(failed)} file(s) skipped: {', '.join(failed)}"
             if failed else "")
          + f"\nReport saved to:\n  {out_path}")


if __name__ == "__main__":
    main()
