# gmail-team-days-worked

Track the number of days actually worked per employee, per week, for the current and next calendar month — straight in Google Sheets, based on **Out of Office** events in Google Calendar.

No external tools, no servers — just Google Apps Script connected to the Google Calendar API.

## What it does

- Reads **Out of Office** events (`eventType: outOfOffice`) from each employee's Google Calendar
- Supports full-day and partial absences (half day, quarter day, ...)
- Shows days worked per employee per week (e.g. `3.5 / 5 days`)
- Shows a **Total** per employee for each month, right before that month's weeks
- Handles months that don't start or end on a Monday/Friday — the first/last week of a month only counts the days that actually fall inside it
- Color coding: 🟢 fully present · 🟡 partially present · 🔴 mostly absent
- Each employee's name links straight to their Google Calendar, so anyone reviewing the numbers (Sharon checking a contractor invoice, or the contractor cross-checking it themselves) can jump in and verify
- Teams are configured via a dedicated **Teams** tab — no code changes needed

This is a sibling project of [gmail-team-capacity](../gmail-team-capacity) — same architecture and the same underlying assumption: Out of Office entries in Google Calendar are the source of truth for days not worked. This sheet is purpose-built for a simpler, invoicing-oriented question ("how many days did this person work this month") rather than sprint capacity planning.

## Architecture

```
gmail-team-days-worked/
├── Code.js                    # Google Apps Script — main logic
├── Loading.html               # Pulsing Acme logo dialog shown during refresh
├── appsscript.json            # Apps Script manifest (scopes, timezone)
├── .clasp.json                # Link to the Apps Script project
└── .gitignore
```

## Teams tab

The **Teams** sheet drives everything.

| Team    | Name    | Email                  |
|---------|---------|------------------------|
| Team A  | Pier    | pier@acme.be         |
| Team A  | Pol     | pol@acme.be          |
| Team B  | Marie   | marie@acme.be        |

- **Team**: any grouping label you want (scrum team, department, "Contractors", ...) — purely a display grouping
- **Name**: display name shown in the Days Worked sheet
- **Email**: used to query Google Calendar — must be the person's Google Workspace email

> **Email formula**: emails follow the pattern `firstname@acme.be`. If your organisation uses a different pattern, you can use a Sheets formula in the Email column:
> ```
> =LOWER(A2)&"@acme.be"
> ```
> The script detects any cell containing `@` as a valid email address.

Add or remove rows freely — the script picks them up automatically on the next refresh. New teams get an automatically assigned color.

Unlike `gmail-team-capacity`'s Teams tab, there is no sprint/velocity/Jira configuration here — months are derived automatically from today's date, nothing to configure.

## Sheet layout

For each team, the sheet shows two month blocks side by side: the **current** calendar month, then the **next** calendar month. Each month block is: a **Total** column (days worked / working days that month), followed by that month's weeks.

```
Name | Total (this month) | Week 1 | Week 2 | ... | Total (next month) | Week 1 | Week 2 | ...
```

The Total column sits *before* its month's weeks — this also acts as the visual divider between the two month blocks.

There is no team-level total row — only per-person totals, since the goal is invoice cross-checking per individual, not team capacity.

## Requirements

- Google Workspace account with access to your colleagues' Google Calendars
- Google Apps Script project linked to a Google Sheet
- **Advanced Google Services** enabled in Apps Script: `Google Calendar API`

## Setup

### 1. Create a new Google Sheet

Create a Google Sheet with two tabs:
- **Teams** — structured as described above
- **Days Worked** — leave this empty; the script creates and fills it automatically

### 2. Link Apps Script to the Sheet

1. Open your Sheet → **Extensions → Apps Script**
2. Delete any existing code and paste the contents of `Code.js` and `Loading.html`
3. Enable the Calendar API: **Services → Google Calendar API → Add**
4. Save and close

### 3. Get the Script ID

In Apps Script: **Project settings (⚙️) → Script ID** — copy this value.

### 4. Update `.clasp.json`

Replace `REPLACE_WITH_NEW_SCRIPT_ID` in `.clasp.json` with the value from step 3, then push to Bitbucket — the pipeline deploys the code to your new Apps Script project automatically.

### 5. Auto-refresh on sheet open

Run **📅 Days Worked → ⚙️ Install auto-refresh trigger** once. This installs a per-user installable trigger that shows the refresh dialog automatically when the sheet is opened.

### 6. Set up CI/CD

Add the base64-encoded contents of `~/.clasprc.json` as a repository variable in Bitbucket:

**Repository settings → Repository variables**

| Variable name                     | Value                                        | Secured |
|------------------------------------|----------------------------------------------|---------|
| `CLASPRC_JSON_TEAM_DAYS_WORKED`    | base64-encoded contents of `~/.clasprc.json` | ✅      |

Generate the base64 value:
```bash
# Linux/Mac
base64 -i ~/.clasprc.json

# Windows PowerShell
[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes((Get-Content "$env:USERPROFILE\.clasprc.json" -Raw)))
```

Every push to `main` that touches `gmail-team-days-worked/**` automatically deploys to Apps Script.

## Usage

Open your Google Sheet and use the **📅 Days Worked** menu:
- **🔄 Refresh days worked** — manually trigger a refresh
- **⚙️ Install auto-refresh trigger** — install the on-open trigger (once per user)

### First-time authorization (once per colleague)

Every person who refreshes needs their own Google authorization for this script, since it reads other people's calendars. The **first** time someone clicks **🔄 Refresh days worked**, Google shows a standard authorization prompt (choose account → *Advanced* → *Go to [project] (unsafe)*, since the script isn't published/verified — this is expected for an internal script). Approve it once; every refresh after that works normally.

If a colleague instead sees a generic "permission needed" error inside the loading dialog, it means they skipped that first authorization. Have them run **🔄 Refresh days worked** manually once to trigger the proper authorization screen.

### Sheet sharing vs. Calendar sharing

These are independent:
- **Sheet sharing** (Editor access) only matters for whoever triggers a refresh (needs write access to fill in the Days Worked tab). A person whose days are being counted never needs access to the sheet itself.
- **Calendar sharing** (domain policy or per-calendar sharing) is what actually allows the refreshing user's account to read someone's Out of Office events, completely independent of the sheet.

## CI/CD flow

```
git push origin main
       │
       ▼
Bitbucket Pipelines (on changes in gmail-team-days-worked/**)
       │  clasp push --force
       ▼
Google Apps Script
       │
       ▼
Google Sheet
```

## License

This project is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0). This means you can:

- Share — copy and redistribute the material in any medium or format
- Adapt — remix, transform, and build upon the material

Under the following terms:

- **Attribution** — you must give appropriate credit, provide a link to the license, and indicate if changes were made
- **NonCommercial** — you may not use the material for commercial purposes

---

**Made with ❤️ for efficiency and automation**
