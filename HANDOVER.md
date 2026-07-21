# ERV Construction Dashboard — Project Handover

**Purpose:** This document lets anyone (or you, on a new account/machine) pick up this project from zero and keep working on it — no prior context required. Last refreshed 2026-07-21, ahead of an account switch — everything below reflects the live, current state.

---

## 1. What this project is

A single-page **construction project management dashboard** for Coral Life's ERV Construction team. It reads project data out of an Excel workbook and renders it as an interactive web dashboard: project status overview, financial breakdowns, a zoomable Gantt-style Master Schedule, site-visit logs, internal work tracking, and weekly/monthly change reports — all client-side, no backend/database.

- **Live site:** https://wyuwachotiphan-crypto.github.io/CM-Dashborad/
- **GitHub repo:** https://github.com/wyuwachotiphan-crypto/CM-Dashborad (public)
- **Tech:** one HTML file (vanilla JS + CSS), [Chart.js 4.4.1](https://www.chartjs.org/) + [chartjs-plugin-datalabels](https://chartjs-datalabels.netlify.app/) + [SheetJS/xlsx](https://sheetjs.com/) — all loaded from CDN in the `<head>`. No build step, no npm, no server-side code.

## 2. File inventory — canonical local folder

**As of 2026-07-20, the working folder moved off Desktop.** The real location now is:

```
C:\Users\HPVICTUS\OneDrive - Coral Life\Coral Life - BA - Solution Design 1\1_Non-Project\7_Engineer\IAQ Solution\0_ERV Management\2_Construction Management\Dashboard
```

(A stale duplicate may still exist at `C:\Users\HPVICTUS\Desktop\Dashboard` — that copy was left behind when the folder moved and is **not** kept in sync. Don't edit it; verify you're in the OneDrive path above before editing anything.)

| File | Purpose |
|---|---|
| `index (1).html` | **The entire dashboard** — HTML + CSS + JS in one file. This is the file you edit. |
| `Construction_Portfolio_Template.xlsx` | The live data file (4 sheets — see §4). The dashboard auto-fetches this from GitHub raw on every page load. |
| `Update Dashboard to GitHub.bat` | **Double-click this** to push local edits (both the HTML and the xlsx) to GitHub. See §6. Works from any folder location — it uses `$PSScriptRoot`, not a hardcoded path. |
| `update_dashboard.ps1` | The PowerShell script the `.bat` calls. |
| `serve.ps1` | A tiny local static file server (`powershell -File serve.ps1 -Port 8080`) for previewing the dashboard locally before pushing. |
| `.dashboard-repo/` | A persistent local git clone of the GitHub repo, maintained automatically by the update script — don't edit files here directly, they get overwritten. |
| `HANDOVER.md` | This file. |
| `Construction Manager Playbook.xlsx` | Reference document, not consumed by the dashboard. |

**On GitHub**, the repo root has `index.html` (note: no `(1)` — this is the file GitHub Pages actually serves) and `Construction_Portfolio_Template.xlsx`. The local `index (1).html` is copied to `index.html` on every push — this rename is deliberate so GitHub Pages serves it at the site root.

## 3. How to keep developing this

1. Edit `index (1).html` directly (any text editor, or an AI coding assistant like Claude Code).
2. Preview locally: run `serve.ps1` (e.g. `powershell -File serve.ps1 -Port 8080`) and open `http://localhost:8080` in a browser. Or just open the HTML file directly in a browser (some features like the GitHub-hosted-file refresh need a real HTTP server, not `file://`).
3. When happy, double-click **`Update Dashboard to GitHub.bat`** — it commits and pushes both the HTML and the Excel file to the `main` branch of the GitHub repo. GitHub Pages rebuilds automatically within ~1-2 minutes.

No manual `git` commands needed for day-to-day use.

**Known gotcha:** GitHub Pages can silently stop auto-deploying on push (happened 2026-07-17 → 2026-07-20 — several commits, zero new deployments). If the live site seems stuck on old content after a successful push, check whether this has recurred: compare the latest commit SHA in `git log` against the most recent deployment at `https://api.github.com/repos/wyuwachotiphan-crypto/CM-Dashborad/deployments`. If they're out of sync, the fix is: log into github.com → repo Settings → Pages → change Source from `main` to `None` → Save → change back to `main` / `(root)` → Save. This forces GitHub to re-register the deploy trigger. Pushing empty/trivial commits does **not** fix it — confirmed by testing.

## 4. Data model — the Excel workbook

The dashboard reads `Construction_Portfolio_Template.xlsx`, which has 4 sheets. **Only `Master Schedule` and `Summary Report` are required**; `Site Visit Log` and `Internal Work` are optional.

Full column-by-column reference (with example rows) is built into the app itself: open the dashboard → **Upload Excel Data** tab → scroll to **"Required File Structure"**. That's the source of truth for exact header names.

Summary of the model:

- **Master Schedule** (1 row = 1 work phase; normally 2 rows per project — "Frist Fix" / "Second Fix" phases): Project Name, Phase, Start/End Date (planned), Actual Start/End Date, % Complete, Status, Delivery Date, Actual Delivery Date, Remarks.
- **Summary Report** (1 row = 1 project): Project Name, Status, Project Type (Internal/External/**Non-Project**), Overall Progress, Budget/Spent/Collected/Cash Advance/Preliminary/VO/VE (all THB), Man Power Hours, Site Visit Days, Site Inspector, Contractor Name, Contract Value, Start/End Date, Project Manager, Remarks, Master Schedule PDF link, BD Data Received Date.
  - **Critical:** `Project Name` in Master Schedule must match **exactly** (case-insensitive, trimmed) with `Project Name` in Summary Report, or that project's phases silently fail to link — instead the app creates a second, financially-empty "phantom" project under whatever name was typed. **This has happened in practice** (2026-07-21: "บ้านคุณทวีชัย สมุทปราการ" in Master Schedule vs "คุณทวีชัย สมุทรปราการ" in Summary Report — extra "บ้าน" prefix + a dropped ร — created exactly this phantom-duplicate bug). When investigating "data I entered isn't showing up," **check for this class of typo first**: compare the full project-name list from both sheets for near-duplicates.
  - A row with `Project Type = Non-Project` is excluded from every project-level total; only its Cash Advance value folds into the portfolio-wide Cash Advance total on the Contract Variations chart.
- **Site Visit Log** (optional, 1 row = 1 site visit): Visit Date, Project Name, Inspector Name, Purpose, Site Visit Hours.
- **Internal Work** (optional, 1 row = 1 internal task): Topic, Details, Start Date, End Date, Status (In Progress/Completed/Planning/On Hold).

**Header matching is fuzzy** (see `normKey()`/`field()` in the JS): lowercases, strips parenthetical notes like "(THB)", does substring matching. When adding a new field, use `field(row, 'Exact Header', 'Fallback Header')`, never a raw object key lookup.

**Date parsing gotcha (already fixed):** `XLSX.read(..., {cellDates:true})` can hand back `Date` objects with a few seconds of floating-point drift instead of exact midnight — combined with Asia/Bangkok's UTC+7 offset, this can flip the calendar day. `excelDateToISO()` guards against this by rounding to the nearest minute before reading local Y/M/D. Any new code converting a `Date` to an ISO string should go through `excelDateToISO()`, not do it inline.

**Excel file lock:** if a `~$Construction_Portfolio_Template.xlsx` file is present alongside the real one, Excel currently has it open on this machine — don't run COM automation against it (conflicts with the live session); ask the user to close Excel first, or have them make small edits themselves.

## 5. Data-quality issues found so far (status varies — verify before trusting)

These were **data problems in the Excel file**, not dashboard bugs:

- **Fixed by the user (as of last check):** several Summary Report rows previously had Start Date later than End Date (looked swapped): สำนักงานบริษัท เอส.บี.-ซีร่า จำกัดพระราม2, Ventier courtyard ชั้น 3, บ้านบางบอน5ซอย7, Granpix, บ้านคุณอู BUGAAN พัฒนาการ. Also Khun Sarut's House / 25-17_Ekamai 28 (K'Phong) had identical Start/End dates (likely copy-paste residue). **Re-verify these are still correct** — not re-checked since the original 2026-07-14 audit.
- **Fixed:** "The Grand Pinkao" was missing Master Schedule data — now has 2 phases linked correctly.
- **Open as of 2026-07-21:** "คุณทวีชัย สมุทรปราการ" — real project with financials, but Master Schedule row uses a misspelled name ("บ้านคุณทวีชัย สมุทปราการ"), so it's not linked. User was given the exact fix (row 39, Master Schedule, column A) and asked to correct it themselves since Excel was open at the time — **confirm this got fixed**.

**Recommended standing check** when told "data isn't showing up": load the workbook, run `parseWorkbook()`, and diff the resulting `projects` array's names against the raw Summary Report project-name list. Any name in `projects` that isn't an exact match to a raw Summary Report name is a phantom created by a Master Schedule typo — this has been the actual root cause every time so far.

## 6. Deployment mechanics (for reference / porting to a new account)

- GitHub Pages: **Settings → Pages → Source: Deploy from a branch → Branch: `main` / `(root)`**. Push to `main` auto-publishes (see the Pages-stuck gotcha in §3).
- Commit identity: `wyuwachotiphan-crypto <w.yuwachotiphan@gmail.com>` (set locally in `.dashboard-repo`'s git config, not global).
- Push auth rides on **Git Credential Manager** (Windows' built-in `manager` credential helper) — already has a cached token on this machine, pushes work non-interactively.
- **PowerShell 5.1 gotchas** (relevant if you touch `update_dashboard.ps1` again):
  - Don't set `$ErrorActionPreference = "Stop"` in a script that shells out to `git` — git's normal progress output goes to stderr, and PS 5.1 wraps every stderr line from a native command into a terminating `NativeCommandError`, killing the script even though git succeeded. Check `$LASTEXITCODE` after `git clone`/`git push` instead.
  - Avoid non-ASCII characters (em dashes, etc.) in `.ps1` files saved as UTF-8-without-BOM — can cause a cascading "missing terminator" parse error far below the actual offending line.

### If moving to a new GitHub account

1. Create the new repo, push this repo's contents to it, re-enable GitHub Pages on the new repo.
2. Update the hardcoded refresh URL inside `index (1).html`'s Upload tab (search for `raw.githubusercontent.com/wyuwachotiphan-crypto/CM-Dashborad`) to the new repo's raw URL.
3. Update `update_dashboard.ps1`'s `$repoUrl` and git identity lines; delete the old `.dashboard-repo` folder so it re-clones from the new remote.
4. Update the live-site links here and anywhere else shared.

### If moving to a new Windows/local-machine account

- Git Credential Manager's cached token is per-Windows-profile — the new account will likely need to `git push` once manually and complete a browser sign-in the first time.
- Everything else (files, `.dashboard-repo` clone) can just be copied to the new profile's equivalent OneDrive path, or re-cloned from GitHub fresh.

### If moving to a new Claude / Claude Code account

No special migration step — this folder (or a copy of it) plus this `HANDOVER.md` is everything needed. Point a new session at this file first.

## 7. Feature map

- **Tabs (in order):** Overview | Projects & Schedule | Internal Work | **Reports** | Customer Survey | Upload Excel Data.
- **Overview page:** KPI cards (Total/Construction/Completed/On Hold, clickable → project list popup), status donut chart, **Financial Chart** (Revenue → Amount Collected → Pending Drawdown → Total Contractor Value → Paid to Contractor, in that order — click a bar for per-project breakdown with % where applicable), Contract Variations & Preliminary chart, Monthly Site Visit Overview (hours rounded to the nearest half-hour; click a bar for the day-by-day visit log), Contractor Assignments.
  - **Year filter** ("Year" dropdown, 2025/2026/2027/All) — narrows everything on the page to projects finishing in that year (End Date, falling back to Delivery Date). Synced with the same dropdown on Projects & Schedule.
- **Projects & Schedule page:** Status/Type/Year filter dropdowns + search, project list, and the **Master Schedule** Gantt chart:
  - Year dropdown (2025-2027) + From/To month range filter.
  - Drag-to-scroll horizontally + a zoom slider (px/day) + "Fit" and "Today" buttons.
  - Month header + a "Week 1..Week N" sub-header that resets every month.
  - Sticky project-name column while scrolling.
  - Click a bar → popup with First Fix / Second Fix (as date ranges) / Delivery Date for that project.
- **Reports page** (new tab, added 2026-07-20): **Weekly Report** and **Monthly Report**, each showing what changed in the database as a **table**: Date | Project | Amount Collected | Paid to Contractor | Progress (as "70% → 90%"). One row per project per upload; a column is populated only if that specific field changed in that upload, "-" otherwise. Sourced from the same change-log the Upload tab's "Change History" already recorded (via `diffProjects()`).
- **Internal Work page**, **Customer Survey page** (mostly static/placeholder), **Upload Excel Data page** (drag-drop or URL-based refresh, column reference docs, Change History log).
- **Project detail popup:** click any project name → Budget/Collected %/Paid-to-Contractor %/Pending Drawdown, Delivery Date, Contractor, phase list, link to a Planned-vs-Actual S-curve schedule chart popup.
- Color theme is CSS custom properties in `:root` (current values in §8) — recolored twice in this project's history; easy to swap again if asked.

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

Tab bar: `#0F4C6B` background, white text for both active and inactive tabs (changed from an earlier maroon/light-blue-gray combo that read as low-contrast).

Semantic status colors (Construction=orange `#E2691F` hardcoded in a few JS/CSS spots, Completed=green, On Hold=red/maroon, Planning=gray) are intentionally **not** tied to the theme variables above, so status badges stay distinguishable from whatever the brand accent color is. If a future recolor also wants status colors changed, flag the readability tradeoff before doing it.

## 9. Session conventions worth carrying forward

- User communicates in Thai; respond in Thai. Keep responses concise, confirm what was tested before claiming something works.
- Always test changes via a local preview before pushing — check console errors, verify computed values/DOM state, not just "it looks right."
- **Always push to GitHub after any dashboard edit** — standing instruction, don't ask each time. Both `index (1).html` → `index.html` and the xlsx get pushed together.
- When told data "isn't showing up," don't just re-check totals — cross-reference project names between Master Schedule and Summary Report for near-duplicate typos first (see §5); this has been the actual cause every time.
- The user is comfortable with direct, technical explanations — no need to oversimplify.
