# Elite Prep — SAT Test Result Analysis for Teachers

Give it all the students' SAT/DSAT practice-test score-report PDFs for
**one class / one test**, and it generates a class-wide Word report in
the exact format of the "Sample-DSAT-06-A Result Analysis Teacher Report"
reference document:

> **`<SAT Practice Test code> SAT Test Result Analysis for Teacher.docx`**

The test code and test date are detected automatically from the PDFs —
nothing needs to be typed in.

## What the report contains

1. **Class Performance Overview** — class average and low–high range for
   Total, Reading & Writing, and Math, with a short analysis note.
2. **Average Accuracy by Skill Area** — the 8 SAT skill domains; accuracy
   below 50% is highlighted in **red**.
3. **Priority Review Questions by Section** — for each of the 4 sections
   (Math Section 1 & 2, English Section 1 & 2), every missed question
   listed **from most-missed to least-missed**; error rates ≥ 60% in
   **red**, with a takeaway note under each table.
4. **Our Instructional Plan** — auto-written from the weakest skill areas
   and highest-miss questions (grid-in questions are called out when
   detected); a time-management bullet appears when late-module
   unanswered questions are found.

## Built-in rules (always enforced)

- The report is always written in **English**.
- Branding is always **"Elite Prep"** (never "Elite Prep Suwanee").
- No "Andy Lee, Director, …" signature line.
- **No footers** on any page.
- **No student names, ever.** The report contains only anonymized class
  statistics. As a safety net, any student name or file name detected in
  the PDFs is scrubbed from the finished document and the document is
  re-checked before it is saved.

## AI-written commentary (Claude Opus 4.8) — on by default

The "→ …" commentary lines and the *Our Instructional Plan* bullets are
written by **Claude Opus 4.8** whenever an Anthropic API key is available;
otherwise the built-in rule-based text is used.

- Web app: nothing to click - the key is read from `.env` automatically.
- Command line: set `ANTHROPIC_API_KEY`, then run as usual (`--no-ai` to skip).

Get an API key at https://console.anthropic.com/settings/keys (sign in →
*API Keys* → *Create Key*), add credit under *Billing*, then copy
`.env.example` to `.env` and paste it in (`ANTHROPIC_API_KEY=sk-ant-...`).
`.streamlit/secrets.toml` works too.

Only anonymized aggregate numbers (averages, ranges, per-skill accuracy,
per-question miss counts) are sent to the model — never names, file names,
or individual scores. If the API is unavailable, the built-in text is used.

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
| `--no-ai` | Skip Claude Opus 4.8 and use the built-in commentary |
| `--dump-text` | Write each PDF's raw extracted text to `<name>.pdf.extracted.txt` — use this if a PDF is not recognized, so the parser can be extended |

### If a PDF is not recognized

The parser understands the common College Board / Bluebook score-report
layouts (question tables with section, module, correct answer, your
answer, and Correct/Incorrect/Omitted). If a file reports
"no question-level data recognized", re-run with `--dump-text` and share
the generated `.extracted.txt` file so the parser can be adapted to that
layout.

## Requirements

Python 3.10+ with: `pdfplumber`, `python-docx` (plus `streamlit` for the
web app and `anthropic` for the optional AI commentary)

```
pip install -r requirements.txt
```

---

### 한국어 요약

한 반 학생들의 SAT 성적표 PDF를 전부 업로드(또는 한 폴더에 넣고 실행)하면,
시험 코드·날짜를 PDF에서 자동으로 읽어서 Math Section 1·2, English
Section 1·2 각 섹션별로 **가장 많이 틀린 순서**로 문항을 정리하고,
스킬영역 정답률 + 반 평균/범위 + 수업 계획이 포함된 Word 보고서
(`<시험코드> SAT Test Result Analysis for Teacher.docx`)를
기준 문서와 동일한 양식으로 자동 생성합니다.
