# Outlook Email Analyzer (Legacy/Classic Outlook)

Reads emails from the locally installed classic Outlook desktop app, analyzes
each one with Claude, and exports a formatted Excel report.

No API key is needed. Analysis runs through the **Claude Code CLI** (`claude -p`)
using your own Claude login. No Microsoft Graph, Azure, database, or add-in.

The script only **reads** emails. It never modifies, deletes, moves, replies
to, or marks emails as read. Claude is called with all tools disabled from an
empty temp folder, so email content can only be analyzed, never acted on.

## Requirements

- Windows with classic desktop Outlook installed and a configured mail profile.
- Python 3.9+.
- Claude Code CLI installed and logged in (run `claude` once and sign in).
  Check with `claude --version` in PowerShell.

## Setup (PowerShell)

1. Create and activate a virtual environment:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

2. Install requirements:

   ```powershell
   pip install -r requirements.txt
   ```

3. Copy `.env.example` to `.env`:

   ```powershell
   Copy-Item .env.example .env
   ```

4. Optional: set `CLAUDE_MODEL` in `.env` (e.g. `claude-sonnet-5`). If empty,
   Claude Code's default model is used. No API key is required.

5. In Outlook, create a folder named `AI` under your Inbox, or change
   `OUTLOOK_SUBFOLDER` in `.env`. Leave it empty to read the Inbox directly.

6. Run:

   ```powershell
   python app.py
   ```

7. Open `email_report.xlsx` when it finishes.

Emails are analyzed in groups (`CLAUDE_BATCH_SIZE`, default 5) per `claude -p`
call, so a run counts against your Claude subscription usage limits. Mails of
the same conversation are always kept in one call, and quoted reply history is
cut only when every quoted message matches an older mail in that same call
(otherwise the full text is sent). If a group returns bad output, its
emails are retried one by one. Token usage is printed at the end of each run.

## Configuration (`.env`)

| Variable            | Description                                          |
|---------------------|------------------------------------------------------|
| `CLAUDE_MODEL`      | Optional model name. Empty = Claude Code default.    |
| `OUTLOOK_MAILBOX`   | Mailbox in your Outlook profile to read (its display name, e.g. an email address). Empty = default mailbox. If not found, the available mailboxes are listed. Use a separate `OUTPUT_FILE` per mailbox. |
| `OUTLOOK_SUBFOLDER` | Subfolder under Inbox. Empty = read Inbox.           |
| `MAX_EMAILS`        | Maximum emails to process (newest first).            |
| `DAYS_BACK`         | Only process emails received within this many days.  |
| `OUTPUT_FILE`       | Output Excel file name.                              |
| `EXCLUDE_CATEGORIES` | Comma-separated Outlook categories ignored when learning (default `Follow up`). |
| `ARCHIVE_FOLDER` / `ARCHIVE_CATEGORY` / `ARCHIVE_MAX_ITEMS` | Archive folder name (default `Archive`), the category name for "can be archived" mail, and how many of its newest mails to learn from (default 3000). |
| `MODEL_TARGET_PRECISION` / `MODEL_ARCHIVE_MAX_FALSE` | Accuracy required before the classifier decides alone (0.9), and the share of important mails the archive detector may wrongly flag (0.01). |
| `PROFILE_MAX_AGE_DAYS` | Rescan categories after this many days (default 7). |
| `LEARN_MAX_ITEMS`   | Maximum mails scanned when learning (default 5000). |
| `CLAUDE_BATCH_SIZE` | Emails per Claude call (default 5; 1 = one call per email). |
| `CLAUDE_WORKERS`    | Parallel Claude calls (default 4; 1 = sequential).   |
| `SKIP_ALREADY_ANALYZED` | Default `true`: emails already in the existing report (without an error) are not sent to Claude again; new rows are merged into the file. Set `false` to re-analyze everything. |
| `APPLY_OUTLOOK_CATEGORIES` | `true` writes an `AI: <Category>` Outlook category onto each analyzed email. Default `false` (read-only). This is the only option that modifies emails. |

## Learned categories

If the mailbox already has mails the owner categorized by hand, the app scans
them locally (no Claude tokens) and builds `category_profile*.json`: per
category the mail count, typical senders and typical subject words. Claude then
must pick one of **those** categories (or `Belirsiz` when none fits) instead of
the generic list. The profile adds roughly 1,000-2,500 tokens to each Claude call.

- **Archive category** (`ARCHIVE_CATEGORY`, default `Can be archived`): the
  newest `ARCHIVE_MAX_ITEMS` mails of the owner's archive folder (`ARCHIVE_FOLDER`,
  default `Archive`) that have no category show which senders and topics never needed
  one. Claude uses this category only for purely informational mail (FYI,
  newsletters, ads, automatic notices). Anything with an action, reply, deadline
  or High priority is forced to `Belirsiz` instead. Mails are never moved; the owner
  filters this category in Excel (or in Outlook with `APPLY_OUTLOOK_CATEGORIES=true`)
  and moves them himself. Clear `ARCHIVE_CATEGORY` to switch the feature off.
- **Owner habits override Claude's category** (local, no tokens): a mail in a
  conversation the owner already filed (most of it under one category) gets that
  category; otherwise a sender whose filed mails (`SENDER_RULE_MIN_MAILS`, default 3,
  at least `SENDER_RULE_MIN_SHARE` = 0.9 of them) went to one category gets it.
  Claude still writes the summary and fields. The run summary shows how many
  categories came from these rules.
- **Measure it:** `python app.py --evaluate` (categorized mails, profile rebuilt
  without the test mails) and `python app.py --evaluate-archive` (uncategorized
  archive mails). Each uses `EVAL_SAMPLE` mails (default 40, ~100k tokens) and
  writes a per-mail detail file with subjects, so keep it private.
- It refreshes itself every `PROFILE_MAX_AGE_DAYS` days. Force a rescan any time
  with `python app.py --learn` (local only, Claude is not called).
- `AI: ...` categories written by this app are never used for learning.
- The Excel gets an `Outlook Categories` column (the mail's existing categories)
  next to `Category` (Claude's choice), and a `Categories` sheet with the profile.
- The profile contains sender names, so it is in `.gitignore`. Treat it like the report.
- After changing the category scheme, use a new `OUTPUT_FILE` or set
  `SKIP_ALREADY_ANALYZED=false` once, otherwise old rows keep their old category.

## Local models (used automatically when scikit-learn is installed)

`python app.py --learn` also trains two small models on the owner's own filed
mails, on the PC itself (seconds, no Claude tokens, nothing leaves the machine):

- **Category classifier** (TF-IDF + logistic regression). It decides the category
  itself only when its cross-validated accuracy at that confidence reaches
  `MODEL_TARGET_PRECISION` (default 0.9); otherwise the category is `Belirsiz`
  and the report lists its top 3 guesses in **Suggested Categories**.
- **Archive detector**: important (filed) mails vs mails the owner archives without
  a category. It flags `ARCHIVE_CATEGORY` only above a threshold chosen so that at
  most `MODEL_ARCHIVE_MAX_FALSE` (default 1%) of important mails are flagged in
  cross-validation, and never for mails with an action, reply or High priority.
- **`USE_CLAUDE=false` runs categories only**: no Claude calls, no `claude` command
  needed, no token use, and no mail text leaves the PC. Summary, priority and the
  other Claude fields stay empty. The archive safety check then relies on the
  sender's "High importance" flag instead of Claude's priority/action fields.
  Set `USE_CLAUDE=true` to add the summaries (rows written without them are
  analyzed again).
- In this mode Claude no longer picks categories; it writes summary, priority,
  action and the other fields. If Claude fails, category and suggestions are still filled.
- **Feedback loop without touching Outlook:** pick a value in the report's
  **Chosen Category** dropdown. The next `--learn` reads these choices (and
  `2026'ya Aktarılabilir` choices as archivable examples) as new training labels.
- Free measurement: `python app.py --evaluate-ml` (models trained without the
  test mails, no Claude calls). Model and profile files contain words from the
  mails; they are in `.gitignore` - keep them private.

## Output

`email_report.xlsx` has one row per email: Received Date, Sender Name, Sender
Email, Subject, Summary, Category, Priority, Action Required, Action, Reply
Required, Deadline, Owner, Organization, Important Data, Outlook Entry ID,
Processing Error.

Bold frozen header, filters, wrapped Summary/Action/Important Data, and
red/yellow/green Priority fills. Rows are sorted newest first. A failed email
gets its error in `Processing Error` and does not stop the run.

## Troubleshooting

- **"'claude' command not found"** — install Claude Code and make sure
  `claude --version` works in PowerShell.
- **"Claude CLI failed"** in Processing Error — run `claude` once manually to
  confirm you are logged in; the message includes the CLI's own error text.
- **Could not connect to Outlook / access Inbox** — install and open classic
  Outlook once so a mail profile exists.
- **Subfolder not found** — create it under Inbox or fix `OUTLOOK_SUBFOLDER`.
- **Excel file won't save** — close `email_report.xlsx` and re-run.

## Optional: Windows Task Scheduler

1. **Task Scheduler** → **Create Basic Task**, with a daily trigger.
2. Action: **Start a program**.
   - Program/script: `C:\path\to\project\.venv\Scripts\python.exe`
   - Arguments: `app.py`
   - Start in: `C:\path\to\project`
3. Run it as your own user, while logged in, so Outlook and your Claude login
   are available.
