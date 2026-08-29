"""
Save the class report to Gmail as a draft (never sends it).
===========================================================

The report .docx is appended to the Drafts mailbox over IMAP, using a Google
app password.  The teacher opens the draft, checks it, and presses Send.

Requires  GMAIL_ADDRESS  and  GMAIL_APP_PASSWORD  (see .env / Streamlit
Secrets).  Nothing here imports Streamlit, so it can be exercised directly.
"""

import imaplib
import time
from email.message import EmailMessage

IMAP_HOST = "imap.gmail.com"
DEFAULT_TO = "sue.kim@eliteprep.com"

# Edit these two lines to change how the draft addresses and signs off.
GREETING = "Dear Sue 원장님,"
SIGNATURE = "Thank you so much.\nAndy Lee / Elite Prep Suwanee"

DOCX_MIME = ("application",
             "vnd.openxmlformats-officedocument.wordprocessingml.document")

EN_DASH, MIDDOT = "\u2014", "\u00b7"


def build_draft_message(sender, to, test_code, test_date, n_students,
                        averages, docx_bytes, filename):
    """Assemble the draft.  averages: {"total": n|None, "rw": .., "math": ..}"""
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = (f"{test_code} SAT Test Result Analysis for Teacher "
                      f"({test_date})")

    plural = "s" if n_students != 1 else ""
    lines = [GREETING, "",
             f"Attached is the class result analysis for {test_code} "
             f"({test_date}), based on {n_students} student score "
             f"report{plural}."]

    shown = [(label, averages.get(key))
             for label, key in (("Total", "total"),
                                ("Reading & Writing", "rw"),
                                ("Math", "math"))]
    shown = [f"{label} {value}" for label, value in shown if value is not None]
    if shown:
        lines += ["", f"  Class average {EN_DASH} "
                      f"{f' {MIDDOT} '.join(shown)}"]

    lines += ["", SIGNATURE]
    msg.set_content("\n".join(lines))  # EmailMessage picks UTF-8 as needed

    msg.add_attachment(docx_bytes, maintype=DOCX_MIME[0],
                       subtype=DOCX_MIME[1], filename=filename)
    return msg


def _drafts_mailbox(imap):
    """Gmail's Drafts folder, found by its \\Drafts flag so that a Korean
    (or any other) Gmail interface language still works."""
    ok, lines = imap.list()
    if ok == "OK":
        for raw in lines or []:
            line = raw.decode("utf-8", "replace") if isinstance(raw, bytes) \
                else str(raw)
            if "\\Drafts" in line and '"/"' in line:
                return line.split('"/"', 1)[1].strip()
    return '"[Gmail]/Drafts"'


def save_draft_to_gmail(msg, address, app_password):
    """Append msg to the account's Drafts mailbox.  Raises on failure."""
    with imaplib.IMAP4_SSL(IMAP_HOST) as imap:
        imap.login(address, app_password)
        mailbox = _drafts_mailbox(imap)
        ok, detail = imap.append(mailbox, r"\Draft",
                                 imaplib.Time2Internaldate(time.time()),
                                 msg.as_bytes())
        if ok != "OK":
            raise RuntimeError(f"Gmail refused the draft: {detail}")
