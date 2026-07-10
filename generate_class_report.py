#!/usr/bin/env python3
"""
Elite Prep - SAT Class Results Analysis Report Generator
=========================================================
Reads all student SAT/DSAT practice-test score-report PDFs in a folder,
aggregates class-wide statistics, and generates a Word (.docx) report:

    "<TEST CODE> Result Analysis Teacher Report.docx"

The output format matches the "DSAT-05-A Result Analysis Teacher Report"
reference document exactly (Elite Prep navy/blue design system).

Rules enforced (per project instructions):
  * Report text is always in English.
  * Branding is always written as just "Elite Prep" (never "Elite Prep
    Suwanee"), and there is no "Andy Lee, Director, ..." signature line.
  * No footers on any page.
  * Test code and test date are auto-detected from the PDFs.

Usage:
    python generate_class_report.py <folder_with_pdfs>
        [--test-code DSAT-05-A] [--test-date "June 29, 2026"]
        [--output-dir DIR] [--domains domains.csv] [--dump-text]

--dump-text writes each PDF's raw extracted text to <name>.extracted.txt
so the parser can be adapted if a new score-report layout appears.
"""

import argparse
import csv
import datetime
import re
import sys
from collections import defaultdict
from pathlib import Path

import pdfplumber
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches

# ---------------------------------------------------------------------------
# Elite Prep design system (matches the DSAT-05-A reference document)
# ---------------------------------------------------------------------------

NAVY = RGBColor(0x1F, 0x38, 0x64)        # letterhead, title, H2
BLUE = RGBColor(0x2E, 0x75, 0xB6)        # H1 headings
GRAY = RGBColor(0x59, 0x59, 0x59)        # subtitle, arrow commentary
RED = RGBColor(0xC0, 0x00, 0x00)         # alert highlight
BLACK = RGBColor(0x00, 0x00, 0x00)

HEADER_FILL = "DDEBF7"                   # table header row
LABEL_FILL = "F2F2F2"                    # first-column label cells

EN = "–"                            # – en dash
EM = "—"                            # — em dash

RW_DOMAINS = [
    "Craft and Structure",
    "Information and Ideas",
    "Standard English Conventions",
    "Expression of Ideas",
]
MATH_DOMAINS = [
    "Algebra",
    "Advanced Math",
    "Problem-Solving and Data Analysis",
    "Geometry and Trigonometry",
]

# Variants of domain names that may appear in score reports -> canonical name
DOMAIN_ALIASES = {
    "craft and structure": "Craft and Structure",
    "information and ideas": "Information and Ideas",
    "standard english conventions": "Standard English Conventions",
    "expression of ideas": "Expression of Ideas",
    "algebra": "Algebra",
    "advanced math": "Advanced Math",
    "problem-solving and data analysis": "Problem-Solving and Data Analysis",
    "problem solving and data analysis": "Problem-Solving and Data Analysis",
    "problem-solving & data analysis": "Problem-Solving and Data Analysis",
    "geometry and trigonometry": "Geometry and Trigonometry",
    "geometry & trigonometry": "Geometry and Trigonometry",
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


def parse_date_iso(date_str):
    for fmt in ("%B %d, %Y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(date_str, fmt).date().isoformat()
        except (ValueError, TypeError):
            continue
    return None


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

    def top_missed(self, section, module, min_rate=0.5, limit=3):
        """Question numbers with the highest error rates (>= min_rate)."""
        out = []
        for c, qs in self.error_groups(section, module):
            if c / self.n >= min_rate:
                out.extend(qs)
            if len(out) >= limit:
                break
        return out[:limit]


# ---------------------------------------------------------------------------
# DOCX helpers (styled to match the DSAT-05-A reference document)
# ---------------------------------------------------------------------------

def shade_cell(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def set_cell(cell, text, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.CENTER,
             fill=None, size=11):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    run = p.add_run(sanitize_branding(text))
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color if color else BLACK
    if fill:
        shade_cell(cell, fill)


def add_para(doc, text="", bold=False, size=11, italic=False,
             color=None, space_after=8):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(sanitize_branding(text))
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return p


def add_rich_para(doc, fragments, size=11, space_after=8):
    """fragments: list of (text, bold) or (text, bold, color, size)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    for frag in fragments:
        text, bold = frag[0], frag[1]
        color = frag[2] if len(frag) > 2 else None
        fsize = frag[3] if len(frag) > 3 else size
        run = p.add_run(sanitize_branding(text))
        run.bold = bold
        run.font.size = Pt(fsize)
        if color:
            run.font.color.rgb = color
    return p


def add_arrow_note(doc, text):
    """Italic gray commentary line under a table (matches reference)."""
    return add_para(doc, text, italic=True, color=GRAY, size=10.5,
                    space_after=10)


def add_h1(doc, number, text):
    h = doc.add_heading(f"{number}.  {text}", level=1)
    for run in h.runs:
        run.font.color.rgb = BLUE
        run.font.size = Pt(13)
        run.bold = True
    return h


def add_h2(doc, i, text):
    h = doc.add_heading(f"{CIRCLED[i]}  {text}", level=2)
    for run in h.runs:
        run.font.color.rgb = NAVY
        run.font.size = Pt(11.5)
        run.bold = True
    return h


def make_table(doc, headers, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, h in enumerate(headers):
        set_cell(table.rows[0].cells[i], h, bold=True, fill=HEADER_FILL)
    if widths:
        apply_widths(table, widths)
    return table


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
# Narrative text
# ---------------------------------------------------------------------------

def join_names(items):
    if len(items) <= 1:
        return items[0] if items else ""
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f", and {items[-1]}"


def q_names(qnums):
    return join_names([f"Q{q}" for q in qnums])


def overview_note(stats):
    rw_avg, math_avg = stats.avg("rw"), stats.avg("math")
    if rw_avg is None or math_avg is None:
        return None
    if math_avg >= rw_avg:
        a_name, a, b_name, b = "Math", math_avg, "Reading & Writing", rw_avg
    else:
        a_name, a, b_name, b = "Reading & Writing", rw_avg, "Math", math_avg

    def spread(key):
        r = stats.rng(key)
        return (r[1] - r[0]) if r else 0

    wide_key = "math" if spread("math") >= spread("rw") else "rw"
    wide_name = "Math" if wide_key == "math" else "Reading & Writing"
    lo, hi = stats.rng(wide_key)
    verb = "edged out" if abs(a - b) <= 30 else "outpaced"
    return (f"→ On average, the {a_name} section (avg {a}) {verb} "
            f"{b_name} (avg {b}), but the wide spread between students"
            f"{EM}especially in {wide_name} (range {lo}{EN}{hi}){EM}means "
            f"these figures are best read alongside each student's "
            f"individual Score Report.")


def skill_note(stats):
    pairs = [(d, stats.domain_accuracy(d)) for d in RW_DOMAINS + MATH_DOMAINS]
    pairs = [(d, a) for d, a in pairs if a is not None]
    if not pairs:
        return None
    pairs.sort(key=lambda x: x[1])
    below = [(d, a) for d, a in pairs if a < 50]
    strongest = max(pairs, key=lambda x: x[1])
    if below:
        names = join_names([f"{d} ({a}%)" for d, a in below])
        return (f"→ {names} fell below 50% this cycle and will receive "
                f"focused supplementary instruction during class. "
                f"{strongest[0]} ({strongest[1]}%) was the clear strength.")
    weakest = [f"{d} ({a}%)" for d, a in pairs[:3]]
    return (f"→ No skill area fell below 50% this cycle. The relatively "
            f"weaker areas were {join_names(weakest)}; {strongest[0]} "
            f"({strongest[1]}%) was the clear strength. Supplementary "
            f"instruction will target the weaker areas during class.")


def module_note(stats, section, module):
    groups = stats.error_groups(section, module)
    n = stats.n
    if not groups:
        return ("→ No incorrect answers were recorded in this module.")
    top_count, top_qs = groups[0]
    if top_count == n:
        return (f"→ {q_names(top_qs)}, missed by all {n} students, "
                f"{'were' if len(top_qs) > 1 else 'was'} the most "
                f"challenging {'items' if len(top_qs) > 1 else 'item'} "
                f"in this module.")
    second = groups[1] if len(groups) > 1 else None
    if len(top_qs) == 1 and second and top_count / n >= 0.6:
        return (f"→ The highest miss rate fell on Q{top_qs[0]} "
                f"({top_count} of {n} missed), with {q_names(second[1])} "
                f"({second[0]} of {n}) the next toughest.")
    return (f"→ {q_names(top_qs)} carried the highest error rate"
            f"{'s' if len(top_qs) > 1 else ''} here "
            f"({top_count} of {n} missed).")


def plan_bullets(stats):
    def module_tops(section):
        parts = []
        for module in (1, 2):
            qs = stats.top_missed(section, module, min_rate=0.5, limit=3)
            if qs:
                parts.append(f"{q_names(qs)} in Module {module}")
        return join_names(parts) if parts else None

    def weakest(domains):
        pairs = [(d, stats.domain_accuracy(d)) for d in domains]
        pairs = [(d, a) for d, a in pairs if a is not None]
        return min(pairs, key=lambda x: x[1]) if pairs else None

    bullets = []

    math_tops = module_tops("Math")
    math_weak = weakest(MATH_DOMAINS)
    text = "Math: We will focus intensively on the highest-miss items"
    if math_tops:
        text += f"{EM}notably {math_tops}{EM}"
    if math_weak:
        text += (f" and reinforce {math_weak[0]} concepts, the weakest "
                 f"domain ({math_weak[1]}%) this cycle.")
    else:
        text += " and reinforce the concepts behind them in class."
    bullets.append(text)

    rw_tops = module_tops("RW")
    rw_weak = weakest(RW_DOMAINS)
    text = "Reading & Writing: We will repeatedly drill solving strategies for "
    if rw_weak:
        text += f"{rw_weak[0]} ({rw_weak[1]}%)"
        text += (f" and the highest-miss items, especially {rw_tops}," if rw_tops
                 else " and the question types missed most frequently by the class,")
    else:
        text += (f"the highest-miss items, especially {rw_tops}," if rw_tops
                 else "the question types missed most frequently by the class,")
    text += " to sharpen reading and reasoning accuracy."
    bullets.append(text)

    if stats.late_omissions > 0:
        bullets.append(
            "Time Management: Because a few students left questions "
            "unanswered in the latter portion (running out of time), we will "
            "provide separate coaching on real-test pacing and strategies "
            "for handling the final questions.")

    return bullets


# ---------------------------------------------------------------------------
# DOCX report
# ---------------------------------------------------------------------------

def build_report(stats: ClassStats, test_code: str, test_date: str,
                 out_path: Path):
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")

    n = stats.n

    try:
        dt = datetime.datetime.strptime(test_date, "%B %d, %Y")
        month_year = dt.strftime("%B %Y")
        month_day = f"{dt.strftime('%B')} {dt.day}"
    except ValueError:
        month_year = test_date
        month_day = test_date

    # --- Letterhead / title ---------------------------------------------------
    add_rich_para(doc, [
        ("ELITE PREP", True, NAVY, 13),
        ("   |   College Admissions & Test Prep", False, GRAY, 10),
    ], space_after=10)
    add_para(doc, f"{month_year} DSAT Practice Test Results Analysis Report",
             bold=True, size=15, color=NAVY, space_after=4)
    add_para(doc,
             f"For Teachers   ·   Test Code {test_code}   ·   "
             f"Test Date: {test_date}",
             color=GRAY, size=10, space_after=12)

    # --- Intro -----------------------------------------------------------------
    add_para(doc, "Dear Teachers,", bold=True)
    add_para(
        doc,
        "Please find below an analysis of the class-wide results from the "
        "DSAT (Digital SAT) practice test administered at our academy on "
        f"{month_day}. This report summarizes the overall performance of all "
        "students who took the test and highlights the questions and skill "
        "areas that warrant focused review in class. For each individual "
        "student's score, please refer to the personal Score Report provided "
        "separately.",
    )
    add_rich_para(doc, [
        ("A total of ", False),
        (f"{n} student{'s' if n != 1 else ''}", True),
        (" sat for this practice test. To protect each student's privacy, "
         "the analysis below is presented entirely as anonymized, aggregate "
         "statistics.", False),
    ])

    # --- 1. Class Performance Overview -------------------------------------------
    add_h1(doc, 1, "Class Performance Overview")
    add_para(doc, "The class averages and score ranges for this test are as "
                  f"follows. (SAT total score scale: 400{EN}1600)")

    if any(stats.scores[k] for k in ("total", "rw", "math")):
        table = make_table(doc, ["Category", "Class Average",
                                 f"Score Range (Low{EN}High)"])
        for label, key in (("Total", "total"), ("Reading & Writing", "rw"),
                           ("Math", "math")):
            avg, rng = stats.avg(key), stats.rng(key)
            r = table.add_row().cells
            set_cell(r[0], label, bold=True, fill=LABEL_FILL,
                     align=WD_ALIGN_PARAGRAPH.LEFT)
            set_cell(r[1], str(avg) if avg is not None else "N/A", bold=True)
            set_cell(r[2], f"{rng[0]} {EN} {rng[1]}" if rng else "N/A")
        apply_widths(table, [2.0, 1.9, 2.4])
        note = overview_note(stats)
        if note:
            doc.add_paragraph()
            add_arrow_note(doc, note)
    else:
        add_para(doc, "Section scores were not present in the uploaded score "
                      "reports, so the score overview is omitted.",
                 italic=True, color=GRAY)

    # --- 2. Average Accuracy by Skill Area -----------------------------------------
    add_h1(doc, 2, "Average Accuracy by Skill Area")
    if stats.has_domain_data():
        add_para(doc, "Below are the class-wide average accuracy rates for "
                      "the eight detailed skill areas of the SAT. Any area "
                      "with an average accuracy below 50% is highlighted in "
                      "red so that areas requiring focused improvement can "
                      "be identified at a glance.")
        table = make_table(doc, ["Subject", "Skill Area", "Average Accuracy"])
        for subject, domains in (("Reading & Writing", RW_DOMAINS),
                                 ("Math", MATH_DOMAINS)):
            for d in domains:
                acc = stats.domain_accuracy(d)
                r = table.add_row().cells
                set_cell(r[0], subject, bold=True, fill=LABEL_FILL)
                set_cell(r[1], d, align=WD_ALIGN_PARAGRAPH.LEFT)
                if acc is None:
                    set_cell(r[2], "N/A")
                else:
                    low = acc < 50
                    set_cell(r[2], f"{acc}%", bold=low,
                             color=RED if low else None)
        apply_widths(table, [1.9, 2.9, 1.6])
        note = skill_note(stats)
        if note:
            doc.add_paragraph()
            add_arrow_note(doc, note)
    else:
        add_para(doc, "Skill-area (domain) information was not available in "
                      "the uploaded score reports, so per-skill accuracy "
                      "could not be computed for this test.",
                 italic=True, color=GRAY)

    # --- 3. Priority Review Questions by Section -------------------------------------
    add_h1(doc, 3, "Priority Review Questions by Section")
    add_para(doc, "For each section, the questions students most frequently "
                  "answered incorrectly are listed in order of error rate. "
                  "These will be prioritized for review and concept "
                  "reinforcement in class. (Questions with an error rate of "
                  "60% or higher are marked in red.)")

    section_order = [
        ("Math", 1, f"Math {EM} Section 1 (Module 1)"),
        ("Math", 2, f"Math {EM} Section 2 (Module 2)"),
        ("RW", 1, f"English (Reading & Writing) {EM} Section 1 (Module 1)"),
        ("RW", 2, f"English (Reading & Writing) {EM} Section 2 (Module 2)"),
    ]
    for i, (section, module, title) in enumerate(section_order, start=1):
        add_h2(doc, i, title)
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
            color = RED if red else None
            r = table.add_row().cells
            set_cell(r[0], f"{rate}%", bold=red, color=color)
            set_cell(r[1], f"{count} of {n}", bold=red, color=color)
            set_cell(r[2], ", ".join(f"Q{q}" for q in qnums),
                     bold=red, color=color, align=WD_ALIGN_PARAGRAPH.LEFT)
        apply_widths(table, [1.3, 1.7, 3.4])
        doc.add_paragraph()
        add_arrow_note(doc, module_note(stats, section, module))

    # --- 4. Our Instructional Plan ---------------------------------------------------
    add_h1(doc, 4, "Our Instructional Plan")
    for bullet in plan_bullets(stats):
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(sanitize_branding(bullet))
        run.font.size = Pt(11)

    # --- Closing ----------------------------------------------------------------------
    doc.add_paragraph()
    add_para(doc, "We will continue to work closely with you in support of "
                  "our students' growth. Thank you.")

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
    ap.add_argument("--test-code", help="Override the auto-detected test code")
    ap.add_argument("--test-date", help="Override the auto-detected test date")
    ap.add_argument("--output-dir", help="Where to save the Word report "
                                         "(default: the PDF folder)")
    ap.add_argument("--domains", help="Optional CSV mapping questions to "
                                      "skill areas "
                                      "(section,module,question,domain)")
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

    out_dir = Path(args.output_dir).resolve() if args.output_dir else folder
    safe_code = re.sub(r'[\\/:*?"<>|]', "-", test_code)
    out_path = out_dir / f"{safe_code} Result Analysis Teacher Report.docx"

    build_report(stats, test_code, test_date, out_path)

    print(f"\nDone. {len(students)} student(s) analyzed"
          + (f", {len(failed)} file(s) skipped: {', '.join(failed)}"
             if failed else "")
          + f"\nReport saved to:\n  {out_path}")


if __name__ == "__main__":
    main()
