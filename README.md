# Elite Prep — SAT Class Results Analysis Report Generator

Analyzes all students' SAT/DSAT practice-test score-report PDFs in a folder
and generates a class-wide Word report:

> **`<TEST CODE> Result Analysis Teacher Report.docx`**

## What the report contains

1. **Page 1 — Score Trend Graph**: class average scores (Total / Reading &
   Writing / Math) from the **first practice test to the latest**, drawn in
   the Elite Prep navy/blue design system. The trend data is kept
   automatically in `score_history.csv` — every time you run the program,
   the current test's class averages are added, so the curve grows on its own.
2. **Class Performance Overview** — class average and low–high range for
   Total, Reading & Writing, and Math.
3. **Average Accuracy by Skill Area** — the 8 SAT skill domains; accuracy
   below 50% is highlighted in **red**.
4. **Priority Review Questions by Section** — for each of the 4 sections
   (Math Module 1/2, English Module 1/2), questions sorted by how many
   students missed them (most-missed first); error rates ≥ 60% in **red**.
5. **Our Instructional Plan** — auto-written from the weakest skill areas
   and any late-module unanswered questions (time-management coaching note).

## Built-in rules (always enforced)

- The report is always written in **English**.
- Branding is always **"Elite Prep"** (never "Elite Prep Suwanee").
- No "Andy Lee, Director, …" signature line.
- **No footers** on any page.
- Navy/blue Elite Prep design system: navy headings, navy table headers with
  white text, light-blue zebra rows, blue accents.

## How to use

1. Put all the students' score-report PDF files for **one test** into one
   folder (one PDF per student).
2. Run:

   ```
   python generate_class_report.py "C:\path\to\pdf-folder"
   ```

   Or simply **drag the folder onto `Generate-Report.bat`**.

3. The Word report is saved into the same folder.

The test code and test date are auto-detected from the PDFs when possible.
You can also set them explicitly:

```
python generate_class_report.py "C:\path\to\pdf-folder" --test-code DSAT-02-A --test-date "June 8, 2026"
```

### Options

| Option | Meaning |
| --- | --- |
| `--test-code CODE` | Test code used in the title and the output file name (auto-detected if omitted; falls back to the folder name) |
| `--test-date "June 8, 2026"` | Test date (auto-detected if omitted; falls back to today) |
| `--output-dir DIR` | Save the report somewhere other than the PDF folder |
| `--history FILE.csv` | Use a specific score-history file (default: `score_history.csv` next to the script). Use one history file per class. |
| `--no-history` | Skip the trend graph entirely |
| `--domains FILE.csv` | Optional mapping of questions → skill areas, if the PDFs don't include domain names. Columns: `section,module,question,domain` (section = `RW` or `Math`) |
| `--dump-text` | Write each PDF's raw extracted text to `<name>.pdf.extracted.txt` — use this if a PDF is not recognized, so the parser can be extended |

### If a PDF is not recognized

The parser understands the common College Board / Bluebook score-report
layouts (question tables with section, module, correct answer, your answer,
and Correct/Incorrect/Omitted). If a file reports
"no question-level data recognized", re-run with `--dump-text` and share the
generated `.extracted.txt` file so the parser can be adapted to that layout.

### Multiple classes

Keep a separate history file per class so each class gets its own trend
curve, e.g.:

```
python generate_class_report.py "C:\ClassA\June" --history "C:\ClassA\score_history.csv"
```

## Requirements

Python 3.10+ with: `pdfplumber`, `python-docx`, `matplotlib`

```
pip install pdfplumber python-docx matplotlib
```

---

### 한국어 요약

한 반 학생들의 SAT 성적표 PDF를 한 폴더에 넣고
`Generate-Report.bat` 위로 폴더를 끌어다 놓거나
`python generate_class_report.py <폴더>` 를 실행하면,
가장 많이 틀린 문항 순 정리 + 스킬영역 정답률 + 반 평균/범위 +
첫 시험부터 최근 시험까지의 평균 점수 추이 그래프(1페이지)가 포함된
Word 보고서(`<시험코드> Result Analysis Teacher Report.docx`)가
네이비/블루 Elite Prep 디자인으로 자동 생성됩니다.
점수 추이는 `score_history.csv`에 자동 누적됩니다 (반별로 별도 관리 권장).
