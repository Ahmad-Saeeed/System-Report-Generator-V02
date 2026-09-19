---
name: tms-report
description: Parses an uploaded Word test-report document to count Pass/Fail/Blocked/In Progress/Invalid/Not Run test cases, then generates a filled-in TMS (Testing Management System) CR test-results HTML report plus a ready-to-paste "copy into Outlook" email version, from a bundled template. Trigger this skill whenever the user types "/TMS-report" (with or without the leading slash, any case), or asks to generate, create, or make a TMS report, CR test report, or testing progress report for a Change Request — especially when they upload a .docx test-results document alongside the request.
---

# TMS CR Report Generator

Produces, from one uploaded test-results `.docx` and a couple of quick questions:
1. A self-contained HTML testing-progress report (dashboard-style: hero header with logo, CR info panel, status table, donut chart, KPI summary) — `assets/template.html` filled in.
2. A ready-to-send email version — an intro paragraph followed by the same report content, built as one static page the user can open, select all, and copy directly into Outlook (no button click, no download-then-open step, no clipboard-permission issues).

## When to trigger
- User types `/TMS-report` (with or without the leading slash, any case)
- User asks to generate/create/make a TMS report, CR test report, or testing progress report — especially with a `.docx` test-results file attached

## Step 1 — Get the test-results document

The user uploads a `.docx` file (e.g. `CR_XXXX-TestReport.docx`) containing one heading per test case, each screenshotted and ending in a bracketed status tag, e.g.:

```
## Verify that created Oracle-CRs are added successfully to "List Change requests" page within TMS [PASS]
```

If no file is attached yet, ask for it before continuing.

### CR title/number — extracted from the filename, never asked
The uploaded filename itself carries the CR number and title, e.g. `CR#6893-GetSpeedProfileDetailsFTTH-TestReport.docx` or `CR_6838-TestReport.docx`. Derive the CR Number field from it:

```python
import re, os
def extract_cr_number(filename):
    base = os.path.splitext(filename)[0]
    base = re.sub(r'[-_]?TestReport$', '', base, flags=re.IGNORECASE)
    m = re.match(r'^CR[\s_\-#]?(\d.*)$', base, flags=re.IGNORECASE)
    return "CR#" + (m.group(1) if m else base)
```

Examples: `CR#6893-GetSpeedProfileDetailsFTTH-TestReport.docx` → `CR#6893-GetSpeedProfileDetailsFTTH`; `CR_6838-TestReport.docx` → `CR#6838`.

Never ask the user for the CR title/number — it always comes from the filename.

### Test case counts
1. Convert to markdown: `pandoc -t markdown "<uploaded.docx>" -o report.md`
2. For every heading line (`#`, `##`, etc.), check whether it ends with a bracketed status tag matching (case-insensitive, may be bold `**[...]**`, and pandoc may escape the brackets as `\[...\]`): `PASS`, `FAIL`, `BLOCKED`, `IN PROGRESS`, `INVALID`, `NOT RUN`.
3. Count one test case per tagged heading, grouped by tag.
4. **Headings with no status tag are not test cases** — they're reference/workflow screenshots (e.g. "ORC-Pending CAB", "ORC business approval") and must be ignored, not counted as Not Run or anything else.
5. Any of the six statuses with zero tagged headings defaults to a count of 0.

Regex that works against the pandoc markdown output:
```python
import re
pattern = re.compile(r'\[(PASS|FAIL|BLOCKED|IN PROGRESS|INVALID|NOT RUN)\]', re.IGNORECASE)
# search each heading line with literal backslashes stripped: line.replace('\\','')
```

Report back to the user what you counted (e.g. "Found 7 test cases: 7 Pass, 0 Fail...") before moving on, so they can catch a miscount.

## Step 2 — Ask for the remaining report details

CR title/number is never asked — it comes from the filename (Step 1). Ask these three, all via `ask_user_input_v0`:

1. **System** — single-select: `TMS`, `MW CP4I13`, `Other`. If `Other`, ask a follow-up free-text question for the actual value. This list is known to grow — if the user names a system not on it, just use what they typed and treat it as the new "Other" value (don't refuse or push back on an unlisted system).
2. **Environment** — single-select: `Testing`, `IOT`, `E2E`, `Other`. Same `Other` free-text fallback as above.
3. **Report type** — single-select: `E2E`, `IOT`, `Other`. If `Other`, ask a follow-up free-text question for the actual value. This drives the header title (`<value> Testing Progress Report`, e.g. `E2E Testing Progress Report`) and the short form used in the email intro (`<value>`, e.g. `E2E`).
4. **Observations** — ask (single-select `Yes`/`No`) whether there's any observation to add. If `No`, proceed normally (observations list stays empty). If `Yes`, your very next reply must be a visible chat message presenting a fill-in-the-blank style prompt — a label followed by an empty line for them to type into, e.g.:

   ```
   Observation:
   ```

   (blank line right after, for them to fill in — same style as the plain "CR title: / System:" fill-in prompts used elsewhere in this skill). Do not stay silent, do not just wait, do not skip straight to generating the report. Stop and wait for the user's reply. Only once they've actually sent the text should you proceed to use it. If it has multiple lines, treat each line as a separate observation/bullet; a single line is one observation.

Note Environment and Report type are separate fields with separate answers — even though their option lists overlap (both can be `E2E`/`IOT`), the user may pick differently for each (e.g. Environment=`Testing`, Report type=`E2E`).

Fixed every run (do not ask):
- Tester: `Ahmed Saeed (ahmed.saiid)`
- Test Date / email date: today's date **in Cairo local time (Africa/Cairo)** — not the sandbox's UTC clock. Format `DD Mon YYYY` for the field, `Month DD, YYYY at HH:MM AM/PM` for the generated-on timestamp.
- Status: `Completed`
- Logo: `assets/logo.png` (the WE logo), baked in by default — no need to ask or use the upload button.

## Step 3 — Generate the HTML report

1. Copy `assets/template.html` to a working location, e.g. `/home/claude/TMS_Report_<CR-number>.html`.
2. Edit in the copy:
   - **`crInfoRows` array** (`let crInfoRows = [...]`): System (Step 2.1), CR Number (from filename, Step 1), Tester, Test Date, Environment (Step 2.2), Status — as above.
   - **`statusData` array** (`const statusData = [...]`): set `count` for each of the six from Step 1's parsed counts.
   - **Header title** (built from Report type, Step 2.3, as `<value> Testing Progress Report`) in all three places so the theme toggle can't revert it: the static `<span id="reportTitleText">`, the static `<span id="footerTitleText">`, and **both** `light`/`dark` keys of `THEME_TITLES` — all set to the same title string.
   - **`<span id="genDate">`** — current Cairo date/time.
   - **`observations` array** (`let observations = [];`): if the user gave observation text (Step 2.4), replace with `['line 1', 'line 2', ...]` (one entry per line they typed). Otherwise leave as `[]`.
   - **Logo** — set `<img id="heroLogo" class="hero-logo visible" src="data:image/png;base64,<assets/logo.png base64>">`, and set the toolbar's `logoBtnLabel` to `Change Logo` and `removeLogoBtn` to `display:inline-flex;` (so the UI reflects that a logo is already present).
3. Save and present the file.

Everything else (donut chart, KPI cards, percentage bars, badges, Excel upload/download, theme toggle, "Copy as HTML for Email" button) is computed automatically by the page's own JS from `statusData`/`crInfoRows`/`observations` — don't hand-edit rendered DOM/table rows. The donut chart's center-text sizing is already fixed in the template to shrink-to-fit and true-center as a block, so it holds up for large totals too.

## Step 4 — Generate the ready-to-send email

Run `scripts/build_email.py` — it rebuilds the exact same donut chart (via Pillow) and the exact markup the report's own "Copy as HTML for Email" button produces, as a standalone page, plus an intro paragraph, so no browser button or clipboard permission is ever needed:

```bash
python3 scripts/build_email.py \
  --title "<Report type value> Testing Progress Report" --short-title "<Report type value, e.g. E2E/IOT/Other value>" \
  --gen-date "<Cairo date/time, e.g. September 15, 2026 at 09:53 AM>" \
  --logo assets/logo.png \
  --system "<System from Step 2.1>" --cr-number "<CR Number field, from filename>" \
  --tester "Ahmed Saeed (ahmed.saiid)" --test-date "<DD Mon YYYY>" \
  --environment "<Environment from Step 2.2>" --status "Completed" \
  --pass N --fail N --blocked N --inprogress N --invalid N --notrun N \
  --observations "line 1" "line 2" \
  --out /home/claude/TMS_Report_<CR-number>_EMAIL.html
```

Omit `--observations` entirely if the user said no (Step 2.4) — the email will show "No test observations provided", matching the report.

The script prints `SUBJECT: ...` — report that to the user as the email subject line (format: `<System>-<CR Number field value>`, e.g. `TMS-CR#6796 - TMS Project Management requirement Part 3`).

The intro paragraph is chosen automatically by pass rate (`passed / total * 100`, using all six categories as the denominator):
- **100% pass** → `Dears, Kindly check below <IOT/E2E> Report for <CR Number field, full text> has passed testing successfully, TMS updated status "Business UAT Sign Off".`
- **Anything else** → `Dears, Kindly check below <IOT/E2E> and the detailed attached Report for <CR Number field, full text>, TMS updated status "Pending Rework".`

Present the `_EMAIL.html` file and tell the user: open it, Ctrl+A / Cmd+A, Ctrl+C / Cmd+C, paste directly into Outlook — no download-then-click flow needed.

## Output naming

- Report: `TMS_Report_<CR-number>.html`
- Email: `TMS_Report_<CR-number>_EMAIL.html`

(`<CR-number>` = just the digits, e.g. `6796`, for filenames.)
