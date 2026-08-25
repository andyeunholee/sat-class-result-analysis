# Question-to-skill maps

One CSV per SAT practice-test form, named after the test code
(e.g. `DSK2402UU.csv`). It tells the report generator which of the eight
SAT skill areas each question belongs to, so that section 2
("Average Accuracy by Skill Area") can be computed even when the score-report
PDFs do not list a skill area per question.

```
section,module,question,skill
RW,1,1,Craft and Structure
RW,1,2,Information and Ideas
Math,1,1,Algebra
Math,2,17,Advanced Math
```

* `section`: `RW` (Reading & Writing) or `Math`
* `module`: 1 or 2
* `question`: question number within that module
* `skill`: one of
  * Reading & Writing: Craft and Structure / Information and Ideas /
    Standard English Conventions / Expression of Ideas
  * Math: Algebra / Advanced Math / Problem-Solving and Data Analysis /
    Geometry and Trigonometry  ("&" instead of "and" is fine)

Rows with an empty `skill` are ignored. The map is applied automatically
whenever the detected test code matches the file name; a CSV can also be
uploaded together with the PDFs in the web app.
