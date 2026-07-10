# Elite Prep — SAT Class Results Analysis Report Generator

Analyzes all students' SAT/DSAT practice-test score-report PDFs for one test
and generates a class-wide Word report in the standard Elite Prep
teacher-report format:

> **`<TEST CODE> Result Analysis Teacher Report.docx`**

The test code and test date are detected automatically from the PDFs —
nothing needs to be typed in.

## What the report contains

1. **Class Performance Overview** — class average and low–high range for
   Total, Reading & Writing, and Math, with a short analysis note.
2. **Average Accuracy by Skill Area** — the 8 SAT skill domains; accuracy
   below 50% is highlighted in **red**.
3. **Priority Review Questions by Section** — for each of the 4 sections
   (Math Module 1/2, English Module 1/2), questions sorted by how many
   students missed them (most-missed first); error rates ≥ 60% in **red**,
   with a takeaway note under each table.
4. **Our Instructional Plan** — auto-written from the weakest skill areas
   and highest-miss questions; a time-management bullet appears when
   late-module unanswered questions are detected.

The layout matches the "DSAT-05-A Result Analysis Teacher Report" reference
document (Elite Prep navy/blue design system).

## Built-in rules (always enforced)

- The report is always written in **English**.
- Branding is always **"Elite Prep"** (never "Elite Prep Suwanee").
- No "Andy Lee, Director, …" signature line.
- **No footers** on any page.

## Web app (Streamlit)

`app.py` — upload all students' PDFs, click **Generate Report**, download
the Word file. Two steps, nothing else to fill in.

```
streamlit run app.py
```

## Command line

1. Put all the students' score-report PDFs for **one test** into one folder.
2. Run:

   ```
   python generate_class_report.py "C:\path\to\pdf-folder"
   ```

   Or drag the folder onto `Generate-Report.bat`.

3. The Word report is saved into the same folder.

### Options

| Option | Meaning |
| --- | --- |
| `--test-code CODE` | Override the auto-detected test code |
| `--test-date "June 29, 2026"` | Override the auto-detected test date |
| `--output-dir DIR` | Save the report somewhere other than the PDF folder |
| `--domains FILE.csv` | Optional mapping of questions → skill areas, if the PDFs don't include domain names. Columns: `section,module,question,domain` (section = `RW` or `Math`) |
| `--dump-text` | Write each PDF's raw extracted text to `<name>.pdf.extracted.txt` — use this if a PDF is not recognized, so the parser can be extended |

### If a PDF is not recognized

The parser understands the common College Board / Bluebook score-report
layouts (question tables with section, module, correct answer, your answer,
and Correct/Incorrect/Omitted). If a file reports
"no question-level data recognized", re-run with `--dump-text` and share the
generated `.extracted.txt` file so the parser can be adapted to that layout.

## Requirements

Python 3.10+ with: `pdfplumber`, `python-docx` (plus `streamlit` for the
web app)

```
pip install -r requirements.txt
```

---

### 한국어 요약

한 반 학생들의 SAT 성적표 PDF를 전부 업로드(또는 한 폴더에 넣고 실행)하면,
시험 코드·날짜를 PDF에서 자동으로 읽어서
가장 많이 틀린 문항 순 정리 + 스킬영역 정답률 + 반 평균/범위 +
수업 계획이 포함된 Word 보고서
(`<시험코드> Result Analysis Teacher Report.docx`)를
표준 Elite Prep 교사용 보고서 양식으로 자동 생성합니다.
