# Emailing the class report as a Gmail draft

**Date:** 2026-08-29

## Goal

When the Streamlit app finishes building the class report, put that `.docx`
into a Gmail **draft** addressed to the class teacher, so the draft can be
reviewed and sent by hand. The app never sends mail itself.

## Why a draft, not a send

The app has no login. If it could send mail, anyone who learned the URL could
send mail from `andy.lee@eliteprep.com`. A draft-only design makes the worst
case "a stray draft in my own Drafts folder".

## Approach

IMAP `APPEND` to the account's Drafts mailbox, authenticated with a Google
**app password** - the same credential family as SMTP, no OAuth client, no
Google Cloud project, no extra dependency (`imaplib` is stdlib).

Rejected: the Gmail API (`drafts.create`). It is the official route and
survives Workspace policies that block app passwords, but it needs a GCP
project, an OAuth consent screen and a stored refresh token - too much setup
for "put one file in my own Drafts folder". App passwords are available on
this account, so that risk did not apply.

## Components

`email_draft.py` - no Streamlit import, so it can be exercised directly.

- `build_draft_message(...) -> EmailMessage` - assembles recipient, subject,
  body and the `.docx` attachment. Pure: no network, no environment reads.
- `save_draft_to_gmail(msg, address, app_password)` - opens IMAP over SSL,
  locates the Drafts mailbox by its `\Drafts` special-use flag (so a
  non-English Gmail UI still works), and appends the message with the
  `\Draft` flag.

Splitting these two means a later move to the Gmail API replaces only the
second function; subject and body formatting stay put.

`app.py` - calls both after `docx_bytes` exists, before the download button.

## Configuration

Read from `.env` locally and Streamlit Secrets on the cloud, same lookup as
`ANTHROPIC_API_KEY`:

| Name | Meaning |
|---|---|
| `GMAIL_ADDRESS` | account whose Drafts folder receives the draft |
| `GMAIL_APP_PASSWORD` | 16-character app password |
| `REPORT_TO` | recipient (default `sue.kim@eliteprep.com`) |

If `GMAIL_ADDRESS` or `GMAIL_APP_PASSWORD` is missing the feature is silently
inactive and the app behaves exactly as it does today - the same shape as the
existing "no API key" path.

## Draft content

```
Subject: <TEST CODE> SAT Test Result Analysis for Teacher (<test date>)
Attach:  <TEST CODE> SAT Test Result Analysis for Teacher.docx

Dear Sue 원장님,

Attached is the class result analysis for <code> (<date>),
based on <n> student score reports.

  Class average — Total <t> · Reading & Writing <rw> · Math <m>

Best,
Andy
```

The average line is omitted when no averages could be computed. UTF-8
throughout for the Korean greeting.

## Failure handling

Draft creation is wrapped so that no IMAP problem can affect the report. On
failure the app shows a warning naming the cause; the download button is
always rendered, success or not.

## Duplicate drafts

Streamlit re-runs the whole script on every widget interaction - opening a
preview expander or clicking download would otherwise append the same draft
again. The result is keyed in `st.session_state` by test code plus a hash of
the `.docx` bytes, so one report produces exactly one draft.

## Verification

1. Build a message locally and check subject, attachment name/MIME and that
   the Korean greeting survives encoding.
2. Run the app locally against a real score-report PDF and confirm the draft
   arrives in the Drafts folder with a readable attachment.
3. Deploy and confirm Streamlit Cloud allows outbound IMAP (993). If it does
   not, switch `save_draft_to_gmail` to the Gmail API; nothing else changes.
