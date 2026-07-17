# ERV Construction Dashboard — Project Handover

**Purpose:** This document lets anyone (or you, on a new machine/GitHub/Claude account) pick up this project from zero and keep working on it — no prior context required.

---

## 1. What this project is

A single-page **construction project management dashboard** for Coral Life's ERV Construction team. It reads project data out of an Excel workbook and renders it as an interactive web dashboard: project status overview, financial breakdowns, a zoomable Gantt-style Master Schedule, site-visit logs, internal work tracking, and a monthly change report — all client-side, no backend/database.

- **Live site:** https://wyuwachotiphan-crypto.github.io/CM-Dashborad/
- **GitHub repo:** https://github.com/wyuwachotiphan-crypto/CM-Dashborad (public)
- **Tech:** one HTML file (vanilla JS + CSS), [Chart.js 4.4.1](https://www.chartjs.org/) + [chartjs-plugin-datalabels](https://chartjs-datalabels.netlify.app/) + [SheetJS/xlsx](https://sheetjs.com/) — all loaded from CDN in the `<head>`. No build step, no npm, no server-side code.

## 2. File inventory (local machine: `C:\Users\HPVICTUS\Desktop\Dashboard\`)

| File | Purpose |
|---|---|
| `index (1).html` | **The entire dashboard** — HTML + CSS + JS in one file. This is the file you edit. |
| `Construction_Portfolio_Template.xlsx` | The live data file (4 sheets — see §4). The dashboard auto-fetches this from GitHub raw on every page load. |
| `Update Dashboard to GitHub.bat` | **Double-click this** to push your local edits (both the HTML and the xlsx) to GitHub. See §6. |
| `update_dashboard.ps1` | The PowerShell script the `.bat` calls. |
| `serve.ps1` | A tiny local static file server (`powershell -File serve.ps1 -Port 8080`) for previewing the dashboard locally before pushing. |
| `.dashboard-repo/` | A persistent local git clone of the GitHub repo, maintained automatically by the update script — don't edit files here directly, they get overwritten. |
| `Construction Manager Playbook.xlsx` | Reference document, not consumed by the dashboard. |

**On GitHub**, the repo root has `index.html` (note: no `(1)` — this is the file GitHub Pages actually serves) and `Construction_Portfolio_Template.xlsx`. The local `index (1).html` is copied to `index.html` on every push — this rename is deliberate so GitHub Pages serves it at the site root.

## 3. How to keep developing this

1. Edit `index (1).html` directly (any text editor, or an AI coding assistant like Claude Code).
2. Preview locally: run `serve.ps1` (e.g. `powershell -File serve.ps1 -Port 8080`) and open `http://localhost:8080` in a browser. Or just open the HTML file directly in a browser (some features like the GitHub-hosted-file refresh need a real HTTP server, not `file://`).
3. When happy, double-click **`Update Dashboard to GitHub.bat`** — it commits and pushes both the HTML and the Excel file to the `main` branch of the GitHub repo. GitHub Pages rebuilds automatically within ~1-2 minutes.

No manual `git` commands needed for day-to-day use — the `.bat` handles everything (see §6 for what it actually does, in case you need to debug it or set it up on a different machine).

## 4. Data model — the Excel workbook

The dashboard reads `Construction_Portfolio_Template.xlsx`, which has up to 4 sheets. **Only `Master Schedule` and `Summary Report` are required**; `Site Visit Log` and `Internal Work` are optional (the corresponding chart/page just shows empty state if missing).

Full column-by-column reference (with example rows) is built into the app itself: open the dashboard → **Upload Excel Data** tab → scroll to **"Required File Structure"**. That section is the source of truth for exact header names — don't guess from this doc, read it there since it may have evolved.

Summary of the model:

- **Master Schedule** (1 row = 1 work phase; normally 2 rows per project — "Frist Fix" / "Second Fix" phases, sheet has a header row then data starting a few rows down): Project Name, Phase, Start/End Date (planned), Actual Start/End Date, % Complete, Status, Delivery Date, Actual Delivery Date, Remarks.
- **Summary Report** (1 row = 1 project): Project Name, Status, Project Type (Internal/External/**Non-Project**), Overall Progress, Budget/Spent/Collected/Cash Advance/Preliminary/VO/VE (all THB), Man Power Hours, Site Visit Days, Site Inspector, Contractor Name, Contract Value, Start/End Date, Project Manager, Remarks, Master Schedule PDF link, BD Data Received Date.
  - **Critical:** `Project Name` in Master Schedule must match **exactly** (case-insensitive, trimmed) with `Project Name` in Summary Report, or that project's schedule/phase data won't link up.
  - A row with `Project Type = Non-Project` is a special row used only to add a lump Cash Advance figure not tied to any specific project (e.g. petty cash pool) — it's excluded from every project-level total but its Cash Advance value is folded into the portfolio-wide Cash Advance total shown on the Contract Variations chart.
- **Site Visit Log** (optional, 1 row = 1 site visit): Visit Date, Project Name, Inspector Name, Purpose, Site Visit Hours.
- **Internal Work** (optional, 1 row = 1 internal task): Topic, Details, Start Date, End Date, Status (In Progress/Completed/Planning/On Hold).

**Header matching is fuzzy** (see `normKey()`/`field()` in the JS): it lowercases, strips parenthetical notes like "(THB)", and does substring matching, so minor header wording differences usually still work. When adding a new field, always use `field(row, 'Exact Header', 'Fallback Header')` rather than a raw object key lookup.

**Date parsing gotcha (already fixed, keep in mind for future date fields):** `XLSX.read(..., {cellDates:true})` can hand back `Date` objects with a few seconds of floating-point drift instead of exact midnight, which — combined with a non-UTC timezone (this data uses Asia/Bangkok, UTC+7) — can flip the calendar day. `excelDateToISO()` guards against this by rounding to the nearest minute before reading local Y/M/D. Any new code that converts a `Date` object to an ISO date string should go through `excelDateToISO()`, not do it inline.

## 5. Known open data-quality issues (as of last audit, 2026-07-14)

These are **data problems in the Excel file itself**, not dashboard bugs — the dashboard renders them correctly given the (wrong) input:

- Several Summary Report rows had **Start Date later than End Date** (looked swapped/mistyped): สำนักงานบริษัท เอส.บี.-ซีร่า จำกัดพระราม2, Ventier courtyard ชั้น 3, บ้านบางบอน5ซอย7, Granpix, บ้านคุณอู BUGAAN พัฒนาการ.
- **Khun Sarut's House** and **25-17_Ekamai 28 (K'Phong)** had identical Start/End dates (2025-03-01 → 2025-03-21) — looked like copy-paste residue from templating one row off the other.
- **คุณทวีชัย สมุทรปราการ** and **The Grand Pinkao** had no Master Schedule rows at all (Summary Report entry exists, but no phase/schedule data) — shows as "No Schedule" risk, which is correct given the gap.

Status of these as of the last conversation: the user said they'd correct the dates themselves and were "waiting on confirmation of start/end dates" from elsewhere — **verify with the user whether this has been resolved** before assuming the underlying data is now clean.

## 6. Deployment mechanics (for reference / porting to a new machine or account)

- GitHub Pages is enabled on the repo: **Settings → Pages → Source: Deploy from a branch → Branch: `main` / `(root)`**. Any push to `main` auto-publishes.
- The commit identity used is `wyuwachotiphan-crypto <w.yuwachotiphan@gmail.com>` (set locally in `.dashboard-repo`'s git config, not global).
- Git push authentication rides on **Git Credential Manager** (Windows' built-in `manager` credential helper) — it already had a cached token on this machine so pushes work non-interactively. On a fresh machine you'd need to `git push` once manually and complete the browser sign-in prompt GCM shows.
- **PowerShell 5.1 gotchas hit while building `update_dashboard.ps1`** (useful if you touch that script again):
  - Don't set `$ErrorActionPreference = "Stop"` in a script that shells out to `git` — git's normal progress output goes to stderr, and PS 5.1 wraps every stderr line from a native command into a terminating `NativeCommandError`, which kills the script even though git succeeded. Check `$LASTEXITCODE` after `git clone`/`git push` instead.
  - Avoid non-ASCII characters (em dashes, curly quotes, Thai text mixed with certain punctuation) in `.ps1` files saved as UTF-8-without-BOM — it can cause a **cascading "missing terminator" parse error** far below the actual offending line, which is very confusing to debug. Stick to plain ASCII in `.ps1` files on this setup.

### If moving to a brand new GitHub account

1. Create the new repo, push this repo's contents to it (`git remote set-url origin <new-url>` then `git push`), and re-enable GitHub Pages on the new repo (Settings → Pages, same as above).
2. Update the hardcoded refresh URL inside `index (1).html`'s Upload tab (search for `raw.githubusercontent.com/wyuwachotiphan-crypto/CM-Dashborad` — there's one occurrence, the default value of the "Refresh from a Hosted File Link" input) to point at the new repo's raw URL.
3. Update `update_dashboard.ps1`'s `$repoUrl` and the git identity lines to match the new account, and delete the old `.dashboard-repo` folder so it re-clones fresh from the new remote.
4. Update the live-site links quoted in this document and anywhere else you've shared them.

### If moving to a new Claude / Claude Code account

There's no special migration step — this repo (or a copy of this folder) plus this `HANDOVER.md` is everything needed. A fresh Claude session reading this file, `index (1).html`, and the Excel file has full context to continue. Point a new session at this handover doc first.

## 7. Feature map (what's been built, for quick orientation)

- **Overview page:** KPI cards (Total/Construction/Completed/On Hold, clickable → project list popup), status donut chart, Financial Chart (Revenue/Amount Collected/Pending Drawdown/Total Contractor Value/Paid to Contractor — click a bar for per-project breakdown with % where applicable), Contract Variations & Preliminary chart, Monthly Site Visit Overview (click a bar for the day-by-day visit log), Contractor Assignments.
- **Year filter** ("Year" dropdown, 2025-2027 + All) on both Overview and Projects & Schedule — narrows everything on the page to projects finishing in that year (End Date, falling back to Delivery Date). The two dropdowns stay in sync.
- **Projects & Schedule page:** Status/Type/Year filter dropdowns + search, project list, and the **Master Schedule** Gantt chart:
  - Year dropdown (2025-2027) + From/To month range filter.
  - Drag-to-scroll horizontally + a zoom slider (pixels-per-day) + "Fit" and "Today" buttons.
  - Month header row + a "Week 1..Week N" sub-header that resets every month.
  - Sticky project-name column while scrolling.
  - Click a bar → popup with First Fix / Second Fix (as date ranges) / Delivery Date for that project.
  - **Monthly Report** section at the bottom: the existing Upload-tab change-log, grouped by month, shown as a collapsible list (latest month expanded by default).
- **Internal Work page**, **Customer Survey page** (mostly static/placeholder), **Upload Excel Data page** (drag-drop or URL-based refresh, plus the full column reference docs and Change History log).
- **Project detail popup:** click any project name anywhere → Budget/Collected %/Paid-to-Contractor %/Pending Drawdown, Delivery Date, Contractor, phase list, link to a Planned-vs-Actual S-curve schedule chart popup.
- Color theme is CSS custom properties in `:root` (see current values in §8) — this has been recolored twice this project's history (once to match an external portal, once to a coolors.co palette) so it's easy to swap again if asked.

## 8. Current color theme (CSS variables, `:root` in `index (1).html`)

```css
--ink:#003049;          /* deep navy — primary dark surface / headings / body text */
--ink-soft:#0F4C6B;     /* mid navy — tab bar background */
--grid:#669BBC;         /* steel blue — gridlines / dashed accents */
--paper:#FFFFFF;        /* card surface */
--canvas:#F7F7F5;       /* page background */
--accent:#C1121F;       /* vivid red — primary accent (active tab underline, eyebrows, etc.) */
--accent-soft:#FDF0D5;  /* cream — tint of accent */
--pulse:#669BBC;        /* steel blue — "live" indicator, today-marker, zoom "Today" button */
--steel:#787774;        /* muted secondary text (unchanged neutral) */
--line:#E3E2DE;         /* hairline borders (unchanged neutral) */
--green:#3E8F63; --red:#780000; --gray:#98A5AD;  /* semantic status colors — kept distinct from --accent on purpose */
```

Note: semantic status colors (Construction=orange `#E2691F` hardcoded in a few JS/CSS spots, Completed=green, On Hold=red/maroon, Planning=gray) are intentionally **not** tied to the theme variables above and were left alone during both recolors, so status badges stay distinguishable from the brand accent color. If a future recolor request seems to also want status colors changed, flag the readability tradeoff before doing it.

## 9. Session conventions worth carrying forward

- User communicates in Thai; respond in Thai. Keep responses concise and confirm what was tested before claiming something works.
- Always test changes via a local preview (browser tools) before pushing — check for console errors, verify computed values/DOM state, not just "it looks like it should work."
- **Always push to GitHub after any dashboard edit** — this was an explicit standing instruction from the user (not something to ask about each time). Both `index (1).html` → `index.html` and the xlsx get pushed together via the `.bat`/git flow.
- The user is comfortable with direct, technical explanations of what was found/fixed — no need to oversimplify.
