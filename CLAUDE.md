# TÜV SÜD Recruitment Portal — Project Context

This file gives Claude Code the background needed to keep working on this
project without re-explaining everything from scratch.

## What this is

A local web app for tracking hiring positions and candidates, built for an
HR assignment (the person building this is trying to land an internship
through this deliverable — quality and correctness matter).

- **Backend:** Flask + SQLite (`app.py`)
- **Frontend:** Single Jinja template (`templates/index.html`) — vanilla JS,
  no build step, no framework
- **Export:** openpyxl generates a two-sheet `.xlsx` (Hiring Tracker +
  Candidates) styled in TÜV SÜD's colors

## Run it

```
pip install -r requirements.txt
python app.py
```
Open `http://127.0.0.1:5000`.

## Architecture

Two SQLite tables, each with full CRUD + a matching UI section:

1. **`positions`** — role-level hiring data (Position, Location, Status,
   Date Opened, Date Closed/DOJ, Days, SF Applications, LinkedIn Views/Apps,
   Recruiter Experience, Notes). Powers the "Hiring Tracker" sheet on export.
   Summary stats (total positions, joined count, etc.) are **computed in
   Python at export time**, not stored — because the row count is dynamic.

2. **`candidates`** — individual candidate records (Name, Role, Location,
   Experience, Certifications, ISO Certified, Source, LinkedIn Sourced,
   Stage, Status, Notes). Powers the "Candidates" sheet on export.

Both have: add form, inline edit (click ✎ → form pre-fills → button becomes
"Save Changes"), delete, filters, and a "Load Sample Data" seed button with
data clearly labeled `[Sample]` so it's never mistaken for real records.

## Candidate intake has three modes (tabs in the UI)

- **Standard** — the plain form.
- **Natural Language** — paste free text, one candidate per line (e.g.
  `"Rohit Verma - QSA - Mumbai - 6 years - ISO 9001 - LinkedIn - final round"`).
  Parsed by `parse_candidate_line()` in `app.py` using **regex + keyword
  matching only — no external API/LLM call**. This was an explicit design
  choice by the user ("don't use API"). Accuracy is decent on structured
  phrasing but not true language understanding — hence the mandatory
  **preview step** before anything is saved to the DB.
- **Excel/CSV Import** — flexible header matching (e.g. "Candidate Name",
  "Name", "Role Applied For" all map to the same field), same preview-before-
  confirm pattern.

## Known design decisions (don't relitigate these without asking)

- **No fabricated candidate data.** Early on, the user asked for fake
  candidate data to be invented and I declined — sample/placeholder rows are
  always labeled `[Sample]` / `[Add Name]` so they can never be mistaken for
  real submissions to HR.
- **TÜV SÜD colors** are the public brand identity (deep blue `#005AA0`,
  dark blue `#003A70`, grey `#53565A`) — not an official Pantone spec, since
  that isn't publicly published. If the user gets real brand hex codes from
  HR, swap the CSS variables at the top of `index.html`.
- **Rule-based NLP, not API-based**, per explicit user instruction.
- Chart.js (via cdnjs) powers the two dashboard charts (status doughnut,
  role bar chart) — no other charting lib in use.

## Things not yet built (mentioned to user as possible next steps)

- Multi-user hosting / shared database (currently single-machine, SQLite)
- Login / auth
- NLP or bulk-import specifically for the Positions table (currently only
  Candidates has NLP + import; Positions only has the standard form)
- Sorting on table columns
- Duplicate-candidate detection

## Style conventions in this codebase

- Flask routes return JSON; no server-side rendering beyond the single
  `index.html` template.
- All new candidate/position fields need to be added in **four places**:
  DB schema (`init_db`), the relevant INSERT/UPDATE queries, the frontend
  form, and the export column list — easy to miss one.
- Keep the export column order matching the original uploaded tracker
  format unless explicitly told to change it.
