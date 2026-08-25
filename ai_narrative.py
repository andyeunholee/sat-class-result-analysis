"""
Optional AI-written commentary for the teacher report (Claude Opus 4.8).
========================================================================
Only *anonymized class-level statistics* are sent to the model: averages,
score ranges, per-skill accuracy, and per-question miss counts.  No student
names, file names, or individual scores ever leave the machine.

Returns a dict that build_report() accepts as `narrative=`:

    {
      "overview":  "Overall, ...",      # note under the score table
      "skill":     "As the table shows, ...",   # note under the skill table
      "math_note": "→ In Math, ...",    # takeaway after the four tables
      "rw_note":   "→ In Reading & Writing, ...",
      "plan":      [{"title": "Math — Grid-In Questions:", "text": "..."}, ...]
    }

Requires:  pip install anthropic   and   ANTHROPIC_API_KEY in the environment.
"""

import json
import re

MODEL = "claude-opus-4-8"

SYSTEM = """You write short, professional commentary for an SAT class-results
report that teachers at Elite Prep will read. Write in English only.

You receive anonymized, aggregate class statistics as JSON. Produce ONLY a
JSON object (no markdown fences, no extra text) with exactly these keys:

  "overview":  Two sentences. Start with "Overall, the ... section (average
               N) ..." comparing the Reading & Writing and Math class
               averages, then comment on the spread between students and
               recommend reading the figures alongside each student's
               individual Score Report.
  "skill":     Two or three sentences starting "As the table shows, ...".
               Say which skill areas fell below 50% (or that none did), name
               the lowest area with its percentage and the next one or two
               weaker areas, and end with "We will provide supplementary
               instruction focused on these comparatively weaker areas during
               class." Use the exact skill-area names given.
  "math_note": One or two sentences starting "→ In Math, ..." about the
               most-missed Math items in Module 1 and Module 2, citing
               question numbers like Q22 and counts like "12 of 16". Mention
               student-produced response (grid-in) items when the listed
               questions are flagged as grid-in.
  "rw_note":   One or two sentences starting "→ In Reading & Writing, ..."
               about the most-missed Reading & Writing items in each module,
               with question numbers and counts; relate them to the weakest
               Reading & Writing skill area when the data supports it.
  "plan":      A list of 3 or 4 objects {"title": ..., "text": ...}:
               two for Reading & Writing (titles like "Reading & Writing —
               Grammar & Conventions:" and "Reading & Writing — Reading
               Strategy:"), one for Math (e.g. "Math — Grid-In Questions:"
               or "Math — <weakest area>:"), and, only if late-module
               unanswered questions were reported, "Time Management:". Each
               text is one or two sentences starting with "We will ...".

Rules: never invent numbers - use only the statistics given. Never mention
any student by name or refer to an individual student. Brand is "Elite Prep"
(never "Elite Prep Suwanee"). Keep the tone factual and encouraging.
"""


def _stats_payload(stats, test_code, test_date):
    """Anonymized aggregate numbers only."""
    def sec(section, module):
        groups = stats.error_groups(section, module)
        return [{"missed": c, "of": stats.n,
                 "error_rate_pct": round(100 * c / stats.n),
                 "questions": [f"Q{q}" for q in qs],
                 "grid_in": [f"Q{q}" for q in qs
                             if stats.is_gridin(section, module, q)]}
                for c, qs in groups]

    from generate_class_report import MATH_DOMAINS, RW_DOMAINS
    return {
        "test_code": test_code,
        "test_date": test_date,
        "students": stats.n,
        "scores": {k: {"average": stats.avg(k), "range": stats.rng(k)}
                   for k in ("total", "rw", "math")},
        "skill_accuracy_pct": {d: stats.domain_accuracy(d)
                               for d in RW_DOMAINS + MATH_DOMAINS},
        "weakest_rw_skill": (stats.weakest_domain(RW_DOMAINS) or [None])[0],
        "weakest_math_skill": (stats.weakest_domain(MATH_DOMAINS)
                               or [None])[0],
        "modules": {"Math-1": sec("Math", 1), "Math-2": sec("Math", 2),
                    "RW-1": sec("RW", 1), "RW-2": sec("RW", 2)},
        "late_module_unanswered_count": stats.late_omissions,
    }


def _parse_json(text):
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text)
    start, end = text.find("{"), text.rfind("}")
    return json.loads(text[start:end + 1])


def write_narrative(stats, test_code, test_date, api_key=None):
    """Ask Claude Opus 4.8 for the report commentary. Raises on any failure
    so the caller can fall back to the rule-based text."""
    import anthropic

    client = anthropic.Anthropic(api_key=api_key) if api_key else \
        anthropic.Anthropic()
    payload = _stats_payload(stats, test_code, test_date)

    response = client.messages.create(
        model=MODEL,
        max_tokens=4000,
        thinking={"type": "adaptive"},
        output_config={"effort": "high"},
        system=SYSTEM,
        messages=[{"role": "user", "content": json.dumps(payload)}],
    )
    if response.stop_reason == "refusal":
        raise RuntimeError("model declined the request")
    text = "".join(b.text for b in response.content if b.type == "text")
    data = _parse_json(text)

    out = {}
    for key in ("overview", "skill", "math_note", "rw_note"):
        v = data.get(key)
        if isinstance(v, str) and v.strip():
            out[key] = v.strip()
    plan = data.get("plan")
    if isinstance(plan, list):
        items = []
        for b in plan:
            if isinstance(b, dict) and str(b.get("text", "")).strip():
                items.append({"title": str(b.get("title", "")).strip(),
                              "text": str(b["text"]).strip()})
            elif isinstance(b, str) and ":" in b:
                t, x = b.split(":", 1)
                items.append({"title": t.strip() + ":", "text": x.strip()})
        if items:
            out["plan"] = items
    return out
